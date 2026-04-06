import type { Tool } from "./types";
import { z } from "zod";
import { getZodErrorMessage, toolFailure } from "./wordShared";

const argsSchema = z
  .object({
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    strikethrough: z.boolean().optional(),
    fontSize: z.number().finite().nonnegative().optional(),
    fontName: z.string().optional(),
    fontColor: z.string().optional(),
    highlightColor: z.string().optional(),
  })
  .refine(
    (o) => Object.values(o).some((v) => v !== undefined),
    { message: "At least one formatting property must be supplied." },
  );

/**
 * Normalise a user-provided highlight colour name to the Word API enum
 * value. Returns undefined when the input is unrecognised so callers
 * can skip it gracefully.
 */
function resolveHighlight(raw: string): string | undefined {
  const lookup: Record<string, string> = {
    yellow: "Yellow",
    green: "Green",
    cyan: "Turquoise",
    turquoise: "Turquoise",
    magenta: "Pink",
    pink: "Pink",
    blue: "Blue",
    red: "Red",
    darkblue: "DarkBlue",
    darkcyan: "Teal",
    teal: "Teal",
    darkgreen: "Green",
    darkmagenta: "Violet",
    violet: "Violet",
    darkred: "DarkRed",
    darkyellow: "DarkYellow",
    gray25: "Gray25",
    lightgray: "Gray25",
    gray50: "Gray50",
    darkgray: "Gray50",
    black: "Black",
    white: "White",
    nohighlight: "NoHighlight",
    none: "NoHighlight",
  };
  return lookup[raw.toLowerCase()];
}

/** Build a short human-readable summary of what was changed. */
function describeDelta(opts: Record<string, unknown>): string {
  const parts: string[] = [];

  const toggle = (key: string, on: string, off: string) => {
    if (opts[key] === true) parts.push(on);
    else if (opts[key] === false) parts.push(off);
  };

  toggle("bold", "+bold", "-bold");
  toggle("italic", "+italic", "-italic");
  toggle("underline", "+underline", "-underline");
  toggle("strikethrough", "+strikethrough", "-strikethrough");

  if (opts.fontSize !== undefined) parts.push(`size=${opts.fontSize}pt`);
  if (opts.fontName !== undefined) parts.push(`font=${opts.fontName}`);
  if (opts.fontColor !== undefined) parts.push(`color=#${String(opts.fontColor).replace("#", "")}`);
  if (opts.highlightColor !== undefined) parts.push(`highlight=${opts.highlightColor}`);

  return parts.join(", ");
}

/**
 * Applies inline formatting to the current Word selection.
 * Only the properties that are explicitly provided will be changed;
 * everything else is left untouched.
 */
export const applyStyleToSelection: Tool = {
  name: "apply_style_to_selection",
  description: "Apply formatting styles to the current Word selection.",
  parameters: {
    type: "object",
    properties: {
      bold: { type: "boolean", description: "Toggle bold." },
      italic: { type: "boolean", description: "Toggle italic." },
      underline: { type: "boolean", description: "Toggle underline." },
      strikethrough: { type: "boolean", description: "Toggle strikethrough." },
      fontSize: { type: "number", description: "Size in points." },
      fontName: { type: "string", description: "Font family name." },
      fontColor: {
        type: "string",
        description: "Hex colour without leading # (e.g. FF0000).",
      },
      highlightColor: {
        type: "string",
        description:
          "Named highlight colour (yellow, green, cyan, blue, red, etc.) or noHighlight to clear.",
      },
    },
  },

  handler: async (args) => {
    const parsed = argsSchema.safeParse(args ?? {});
    if (!parsed.success) return toolFailure(getZodErrorMessage(parsed.error));

    const opts = parsed.data;

    try {
      return await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load("text");
        await ctx.sync();

        if (!sel.text?.trim()) {
          return "Nothing is selected — select text first.";
        }

        const f = sel.font;

        if (opts.bold !== undefined) f.bold = opts.bold;
        if (opts.italic !== undefined) f.italic = opts.italic;
        if (opts.strikethrough !== undefined) f.strikeThrough = opts.strikethrough;
        if (opts.underline !== undefined) {
          f.underline = opts.underline
            ? Word.UnderlineType.single
            : Word.UnderlineType.none;
        }
        if (opts.fontSize !== undefined) f.size = opts.fontSize;
        if (opts.fontName !== undefined) f.name = opts.fontName;
        if (opts.fontColor !== undefined) {
          f.color = opts.fontColor.startsWith("#")
            ? opts.fontColor
            : `#${opts.fontColor}`;
        }
        if (opts.highlightColor !== undefined) {
          const resolved = resolveHighlight(opts.highlightColor);
          if (resolved) f.highlightColor = resolved;
        }

        await ctx.sync();
        return `Styled selection: ${describeDelta(opts as Record<string, unknown>)}`;
      });
    } catch (err: unknown) {
      return toolFailure(err);
    }
  },
};
