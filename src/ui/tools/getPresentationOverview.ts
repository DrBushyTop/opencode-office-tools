import type { Tool } from "./types";
import { loadShapeTexts } from "./powerpointText";
import { toolFailure } from "./powerpointShared";

/** Cap text to a max length with an ellipsis marker. */
function truncate(s: string, limit: number): string {
  return s.length > limit ? s.slice(0, limit) + "\u2026" : s;
}

/** Slides processed per PowerPoint.run() context. */
const BATCH = 10;

/** Maximum shapes inspected per slide for the preview. */
const PREVIEW_SHAPES = 5;

/**
 * Collect short text previews for a batch of slides.
 * Only the first few shapes are read to keep the overview lightweight.
 */
async function previewBatch(
  from: number,
  to: number,
): Promise<string[]> {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items");
    await ctx.sync();

    const chunk = slides.items.slice(from, to + 1);
    for (const s of chunk) s.shapes.load("items");
    await ctx.sync();

    const lines: string[] = [];
    for (let i = 0; i < chunk.length; i++) {
      const slideNo = from + i + 1;
      const subset = chunk[i].shapes.items.slice(0, PREVIEW_SHAPES);
      const rawTexts = await loadShapeTexts(ctx, subset);

      const snippets = rawTexts
        .map((t) => t.trim())
        .filter(Boolean)
        .slice(0, 3)
        .map((t) => truncate(t, 90));

      const label =
        snippets.length > 0 ? snippets.join(" | ") : "(no text content)";
      lines.push(`  ${slideNo}. ${label}`);
    }
    return lines;
  });
}

/**
 * Provides a quick structural summary of the open PowerPoint deck:
 * slide count plus a short text preview of each slide.
 */
export const getPresentationOverview: Tool = {
  name: "get_presentation_overview",
  description: "Get an overview of the PowerPoint deck.",
  parameters: { type: "object", properties: {} },

  handler: async () => {
    try {
      const count = await PowerPoint.run(async (ctx) => {
        const s = ctx.presentation.slides;
        s.load("items");
        await ctx.sync();
        return s.items.length;
      });

      if (count === 0) return "Deck is empty (0 slides).";

      const previews: string[] = [];
      for (let cursor = 0; cursor < count; cursor += BATCH) {
        const hi = Math.min(cursor + BATCH - 1, count - 1);
        previews.push(...(await previewBatch(cursor, hi)));
      }

      return [`${count} slide(s):`, "", ...previews].join("\n");
    } catch (err: unknown) {
      return toolFailure(err);
    }
  },
};
