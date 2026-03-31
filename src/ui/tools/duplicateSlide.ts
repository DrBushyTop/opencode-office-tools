import type { Tool } from "./types";
import { getPowerPointContextSnapshot } from "./powerpointContext";
import { isPowerPointRequirementSetSupported, toolFailure } from "./powerpointShared";
import { z } from "zod";

const duplicateSlideArgsSchema = z.object({
  slideIndex: z.number().optional(),
  sourceIndex: z.number().optional(),
  targetIndex: z.number().optional(),
  formatting: z.enum(["KeepSourceFormatting", "UseDestinationTheme"]).optional(),
});

function isNonNegativeInteger(value: unknown): value is number {
  return Number.isInteger(value) && (value as number) >= 0;
}

export const duplicateSlide: Tool = {
  name: "duplicate_slide",
  description: "Duplicate a slide. By default, the copy is inserted immediately after the source slide.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        type: "number",
        description: "0-based source slide index. Defaults to the active slide when available.",
      },
      sourceIndex: {
        type: "number",
        description: "Alias for slideIndex.",
      },
      targetIndex: {
        type: "number",
        description: "Optional 0-based insertion index for the duplicated slide. Defaults to right after the source slide.",
      },
      formatting: {
        type: "string",
        enum: ["KeepSourceFormatting", "UseDestinationTheme"],
        description: "Optional formatting behavior for the inserted copy. Default KeepSourceFormatting.",
      },
    },
  },
  handler: async (args) => {
    const parsedArgs = duplicateSlideArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }

    const snapshot = getPowerPointContextSnapshot();
    const slideIndex = parsedArgs.data.slideIndex ?? parsedArgs.data.sourceIndex ?? snapshot?.activeSlideIndex;
    const targetIndex = parsedArgs.data.targetIndex;
    const formatting = parsedArgs.data.formatting ?? "KeepSourceFormatting";

    if (slideIndex === undefined) {
      return toolFailure("slideIndex is required when there is no active slide.");
    }
    if (!isNonNegativeInteger(slideIndex)) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }
    if (targetIndex !== undefined && !isNonNegativeInteger(targetIndex)) {
      return toolFailure("targetIndex must be a non-negative integer.");
    }
    if (!isPowerPointRequirementSetSupported("1.8")) {
      return toolFailure("Duplicating slides requires PowerPointApi 1.8.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items/id");
        await context.sync();

        const slideCount = slides.items.length;
        const existingSlideIds = slides.items.map((slide) => slide.id);
        if (slideCount === 0) {
          return toolFailure("Presentation has no slides.");
        }
        if (slideIndex >= slideCount) {
          return toolFailure(`Invalid slideIndex ${slideIndex}. Must be 0-${slideCount - 1}.`);
        }

        const insertionIndex = targetIndex ?? (slideIndex + 1);
        if (insertionIndex < 0 || insertionIndex > slideCount) {
          return toolFailure(`Invalid targetIndex ${insertionIndex}. Must be 0-${slideCount}.`);
        }

        const sourceSlide = slides.items[slideIndex];
        sourceSlide.load("id");
        await context.sync();

        let targetSlideId: string | undefined;
        if (insertionIndex > 0) {
          const targetSlide = slides.items[Math.min(insertionIndex - 1, slideCount - 1)];
          targetSlide.load("id");
          await context.sync();
          targetSlideId = targetSlide.id;
        }

        const exported = sourceSlide.exportAsBase64();
        await context.sync();

        context.presentation.insertSlidesFromBase64(exported.value, {
          formatting,
          ...(targetSlideId ? { targetSlideId } : {}),
        });
        await context.sync();

        slides.load("items/id");
        await context.sync();

        const duplicatedSlide = slides.items.find((slide) => !existingSlideIds.includes(slide.id));
        if (insertionIndex === 0 && duplicatedSlide) {
          duplicatedSlide.moveTo(0);
          await context.sync();
          slides.load("items/id");
          await context.sync();
        }

        const duplicatedSlideId = duplicatedSlide?.id;

        return {
          resultType: "success",
          textResultForLlm: `Duplicated slide ${slideIndex + 1} to position ${insertionIndex + 1}.`,
          sourceIndex: slideIndex,
          sourceSlideId: sourceSlide.id,
          targetIndex: insertionIndex,
          duplicatedSlideId,
          formatting,
          toolTelemetry: {
            sourceIndex: slideIndex,
            targetIndex: insertionIndex,
            formatting,
          },
        };
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
