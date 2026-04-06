import type { Tool } from "./types";
import { loadShapeTexts } from "./powerpointText";
import { toolFailure } from "./powerpointShared";
import { z } from "zod";

const argsSchema = z.object({
  slideIndex: z.number().int().min(0).optional(),
  startIndex: z.number().int().min(0).optional(),
  endIndex: z.number().int().min(0).optional(),
});

/** How many slides to process in one PowerPoint.run() batch. */
const BATCH_LIMIT = 10;

/** Extract text from a contiguous range of slides in a single context. */
async function extractBatch(
  from: number,
  to: number,
  total: number,
): Promise<string[]> {
  return PowerPoint.run(async (ctx) => {
    const allSlides = ctx.presentation.slides;
    allSlides.load("items");
    await ctx.sync();

    const batch = allSlides.items.slice(from, to + 1);
    for (const s of batch) s.shapes.load("items");
    await ctx.sync();

    const output: string[] = [];
    for (let idx = 0; idx < batch.length; idx++) {
      const slideNum = from + idx + 1;
      const shapeTexts = await loadShapeTexts(ctx, batch[idx].shapes.items);
      const content = shapeTexts.filter(Boolean).join("\n\n");
      output.push(`--- slide ${slideNum}/${total} ---\n${content || "(no text)"}`);
    }
    return output;
  });
}

/**
 * Read text content from one or more PowerPoint slides.
 *
 * Accepts a single slide index, a start/end range, or no arguments
 * (reads the entire deck). Large decks are fetched in batches to
 * stay within Office runtime limits.
 */
export const getPresentationContent: Tool = {
  name: "get_presentation_content",
  description:
    "Read text content from one or more PowerPoint slides. " +
    "Pass slideIndex for a single slide, startIndex+endIndex for a range, " +
    "or omit all parameters to read every slide.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        type: "number",
        description: "Zero-based index of a single slide to read.",
      },
      startIndex: {
        type: "number",
        description: "Zero-based start of a slide range (inclusive).",
      },
      endIndex: {
        type: "number",
        description: "Zero-based end of a slide range (inclusive).",
      },
    },
  },

  handler: async (args) => {
    const parsed = argsSchema.safeParse(args ?? {});
    if (!parsed.success) {
      return toolFailure(parsed.error.issues[0]?.message ?? "Bad arguments");
    }

    const { slideIndex, startIndex, endIndex } = parsed.data;

    try {
      // Determine total slide count first
      const total = await PowerPoint.run(async (ctx) => {
        const s = ctx.presentation.slides;
        s.load("items");
        await ctx.sync();
        return s.items.length;
      });

      if (total === 0) return "The presentation contains no slides.";

      // Resolve requested range
      let lo: number;
      let hi: number;

      if (slideIndex !== undefined) {
        if (slideIndex >= total) {
          return toolFailure(
            `slideIndex ${slideIndex} is out of range (deck has ${total} slides, indices 0\u2013${total - 1})`,
          );
        }
        lo = slideIndex;
        hi = slideIndex;
      } else if (startIndex !== undefined && endIndex !== undefined) {
        if (startIndex > endIndex) {
          return toolFailure("startIndex must be \u2264 endIndex");
        }
        lo = Math.max(0, startIndex);
        hi = Math.min(total - 1, endIndex);
      } else {
        lo = 0;
        hi = total - 1;
      }

      // Fetch in batches
      const segments: string[] = [];
      for (let cursor = lo; cursor <= hi; cursor += BATCH_LIMIT) {
        const batchEnd = Math.min(cursor + BATCH_LIMIT - 1, hi);
        const batch = await extractBatch(cursor, batchEnd, total);
        segments.push(...batch);
      }

      return segments.join("\n\n");
    } catch (err: unknown) {
      return toolFailure(err);
    }
  },
};
