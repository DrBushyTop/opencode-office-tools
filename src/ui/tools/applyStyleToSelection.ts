import type { Tool } from "./types";
import { z } from "zod";
import { getZodErrorMessage, toolFailure } from "./wordShared";

const applyStyleToSelectionArgsSchema = z.object({
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  strikethrough: z.boolean().optional(),
  fontSize: z.number().finite().positive().optional(),
  fontName: z.string().optional(),
  fontColor: z.string().optional(),
  highlightColor: z.string().optional(),
}).refine(
  ({ bold, italic, underline, strikethrough, fontSize, fontName, fontColor, highlightColor }) => (
    bold !== undefined || italic !== undefined || underline !== undefined || strikethrough !== undefined
    || fontSize !== undefined || fontName !== undefined || fontColor !== undefined || highlightColor !== undefined
  ),
  { message: "No styles specified. Provide at least one style parameter." },
);

export type ApplyStyleToSelectionArgs = z.infer<typeof applyStyleToSelectionArgsSchema>;

export const applyStyleToSelection: Tool = {
  name: "apply_style_to_selection",
  description: `Apply formatting styles to the currently selected text in Word.

All parameters are optional - only specified styles will be applied.

Parameters:
- bold: Set text to bold (true) or remove bold (false)
- italic: Set text to italic (true) or remove italic (false)
- underline: Set text to underline (true) or remove underline (false)
- strikethrough: Set strikethrough (true) or remove it (false)
- fontSize: Font size in points (e.g., 12, 14, 24)
- fontName: Font family name (e.g., "Arial", "Times New Roman", "Calibri")
- fontColor: Text color as hex string (e.g., "FF0000" for red, "0000FF" for blue)
- highlightColor: Highlight/background color. Use Word highlight colors: "yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "gray25", "gray50", "black", or "noHighlight" to remove

Examples:
- Make text bold and red: bold=true, fontColor="FF0000"
- Increase font size: fontSize=16
- Highlight important text: highlightColor="yellow"
- Apply multiple styles: bold=true, italic=true, fontSize=14, fontName="Arial"`,
  parameters: {
    type: "object",
    properties: {
      bold: {
        type: "boolean",
        description: "Set to true for bold, false to remove bold.",
      },
      italic: {
        type: "boolean",
        description: "Set to true for italic, false to remove italic.",
      },
      underline: {
        type: "boolean",
        description: "Set to true for underline, false to remove underline.",
      },
      strikethrough: {
        type: "boolean",
        description: "Set to true for strikethrough, false to remove it.",
      },
      fontSize: {
        type: "number",
        description: "Font size in points.",
      },
      fontName: {
        type: "string",
        description: "Font family name (e.g., 'Arial', 'Calibri').",
      },
      fontColor: {
        type: "string",
        description: "Text color as hex string without # (e.g., 'FF0000' for red).",
      },
      highlightColor: {
        type: "string",
        description: "Highlight color name (e.g., 'yellow', 'green', 'noHighlight').",
      },
    },
    required: [],
  },
  handler: async (args) => {
    const parsedArgs = applyStyleToSelectionArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(getZodErrorMessage(parsedArgs.error));
    }

    const { bold, italic, underline, strikethrough, fontSize, fontName, fontColor, highlightColor } = parsedArgs.data;

    try {
      return await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        if (!selection.text || selection.text.trim().length === 0) {
          return "No text selected. Please select some text first.";
        }

        const font = selection.font;

        // Apply each specified style
        if (bold !== undefined) {
          font.bold = bold;
        }
        if (italic !== undefined) {
          font.italic = italic;
        }
        if (underline !== undefined) {
          font.underline = underline ? Word.UnderlineType.single : Word.UnderlineType.none;
        }
        if (strikethrough !== undefined) {
          font.strikeThrough = strikethrough;
        }
        if (fontSize !== undefined) {
          font.size = fontSize;
        }
        if (fontName !== undefined) {
          font.name = fontName;
        }
        if (fontColor !== undefined) {
          font.color = fontColor.startsWith("#") ? fontColor : `#${fontColor}`;
        }
        if (highlightColor !== undefined) {
          // Map user-friendly names to Word highlight color strings
          const colorMap: { [key: string]: string } = {
            "yellow": "Yellow",
            "green": "Green",
            "cyan": "Turquoise",
            "turquoise": "Turquoise",
            "magenta": "Pink",
            "pink": "Pink",
            "blue": "Blue",
            "red": "Red",
            "darkblue": "DarkBlue",
            "darkcyan": "Teal",
            "darkgreen": "Green",
            "darkmagenta": "Violet",
            "darkred": "DarkRed",
            "darkyellow": "DarkYellow",
            "gray25": "Gray25",
            "lightgray": "Gray25",
            "gray50": "Gray50",
            "darkgray": "Gray50",
            "black": "Black",
            "white": "White",
            "nohighlight": "NoHighlight",
            "none": "NoHighlight",
          };
          const color = colorMap[highlightColor.toLowerCase()];
          if (color !== undefined) {
            font.highlightColor = color;
          }
        }

        await context.sync();

        // Build confirmation message
        const applied: string[] = [];
        if (bold !== undefined) applied.push(bold ? "bold" : "not bold");
        if (italic !== undefined) applied.push(italic ? "italic" : "not italic");
        if (underline !== undefined) applied.push(underline ? "underlined" : "not underlined");
        if (strikethrough !== undefined) applied.push(strikethrough ? "strikethrough" : "no strikethrough");
        if (fontSize !== undefined) applied.push(`${fontSize}pt`);
        if (fontName !== undefined) applied.push(fontName);
        if (fontColor !== undefined) applied.push(`color #${fontColor.replace("#", "")}`);
        if (highlightColor !== undefined) applied.push(`${highlightColor} highlight`);

        return `Applied formatting: ${applied.join(", ")}.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
