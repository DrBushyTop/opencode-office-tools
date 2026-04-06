import type { Tool } from "./types";
import { z } from "zod";
import { getZodErrorMessage, toolFailure } from "./wordShared";

const argsSchema = z.object({
  headingText: z.string().min(1, "headingText is required"),
  includeSubsections: z.boolean().optional().default(true),
});

/** Determine the outline depth implied by a Word paragraph style. */
function outlineDepth(style: string): number | null {
  if (style === "Title") return 0;
  if (style === "Subtitle") return 1;
  const m = /Heading\s*(\d)/i.exec(style);
  return m ? Number(m[1]) : null;
}

/**
 * Read a specific Word document section by heading.
 *
 * Locates a heading whose text contains the given search term
 * (case-insensitive) and returns the HTML for its content span.
 * When `includeSubsections` is true the span extends to the next
 * heading at the same or shallower depth; otherwise it stops at any
 * heading.
 */
export const getDocumentSection: Tool = {
  name: "get_document_section",
  description:
    "Read a specific Word document section by heading. " +
    "Supply the heading text (partial match, case-insensitive). " +
    "Set includeSubsections to false to stop at the first child heading.",
  parameters: {
    type: "object",
    properties: {
      headingText: {
        type: "string",
        description: "Text to match against heading paragraphs.",
      },
      includeSubsections: {
        type: "boolean",
        description:
          "When true (default) content up to the next same-or-higher heading is returned.",
      },
    },
    required: ["headingText"],
  },

  handler: async (args) => {
    const parsed = argsSchema.safeParse(args ?? {});
    if (!parsed.success) return toolFailure(getZodErrorMessage(parsed.error));

    const { headingText, includeSubsections } = parsed.data;
    const needle = headingText.toLowerCase();

    try {
      return await Word.run(async (ctx) => {
        const paras = ctx.document.body.paragraphs;
        paras.load("items");
        await ctx.sync();

        // Batch-load text & style for every paragraph
        for (const p of paras.items) p.load(["text", "style"]);
        await ctx.sync();

        // --- locate the target heading ---
        let anchorIdx = -1;
        let anchorDepth = 0;

        for (let i = 0; i < paras.items.length; i++) {
          const depth = outlineDepth(paras.items[i].style ?? "");
          if (depth === null) continue;
          if ((paras.items[i].text ?? "").toLowerCase().includes(needle)) {
            anchorIdx = i;
            anchorDepth = depth;
            break;
          }
        }

        if (anchorIdx < 0) {
          return `Heading "${headingText}" not found. Run get_document_overview to list available headings.`;
        }

        // --- find where the section ends ---
        let boundaryIdx = paras.items.length;

        for (let j = anchorIdx + 1; j < paras.items.length; j++) {
          const d = outlineDepth(paras.items[j].style ?? "");
          if (d === null) continue;

          if (includeSubsections) {
            // stop only at same-or-shallower depth
            if (d <= anchorDepth) {
              boundaryIdx = j;
              break;
            }
          } else {
            // stop at any heading
            boundaryIdx = j;
            break;
          }
        }

        // --- extract HTML for the range ---
        const first = paras.items[anchorIdx];
        const last = paras.items[boundaryIdx - 1];

        const span = first
          .getRange(Word.RangeLocation.whole)
          .expandTo(last.getRange(Word.RangeLocation.whole));

        const htmlResult = span.getHtml();
        await ctx.sync();

        return htmlResult.value || "(section is empty)";
      });
    } catch (err: unknown) {
      return toolFailure(err);
    }
  },
};
