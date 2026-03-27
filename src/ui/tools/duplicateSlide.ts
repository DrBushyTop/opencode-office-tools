import type { Tool } from "./types";
import { isPowerPointRequirementSetSupported, toolFailure } from "./powerpointShared";

export const duplicateSlide: Tool = {
  name: "duplicate_slide",
  description: `Duplicate an existing slide in the PowerPoint presentation.

Parameters:
- sourceIndex: 0-based index of the slide to duplicate
- targetIndex: Optional 0-based index where the duplicate should be inserted. If omitted, the duplicate is placed immediately after the source slide.
- formatting: Whether to keep source formatting or use the destination theme.

On hosts with PowerPointApi 1.8, duplication preserves content and layout by exporting and reinserting the original slide.`,
  parameters: {
    type: "object",
    properties: {
      sourceIndex: {
        type: "number",
        description: "0-based index of the slide to duplicate.",
      },
      targetIndex: {
        type: "number",
        description: "0-based index where the duplicate should be inserted. Default is after the source slide.",
      },
      formatting: {
        type: "string",
        enum: ["KeepSourceFormatting", "UseDestinationTheme"],
        description: "Formatting behavior for the duplicated slide. Default KeepSourceFormatting.",
      },
    },
    required: ["sourceIndex"],
  },
  handler: async (args) => {
    const {
      sourceIndex,
      targetIndex,
      formatting = "KeepSourceFormatting",
    } = args as { sourceIndex: number; targetIndex?: number; formatting?: "KeepSourceFormatting" | "UseDestinationTheme" };

    if (!isPowerPointRequirementSetSupported("1.8")) {
      return toolFailure("True slide duplication requires PowerPointApi 1.8.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const slideCount = slides.items.length;
        if (slideCount === 0) {
          return "Presentation has no slides.";
        }
        if (sourceIndex < 0 || sourceIndex >= slideCount) {
          return toolFailure(`Invalid sourceIndex ${sourceIndex}. Must be 0-${slideCount - 1}.`);
        }

        const insertionIndex = targetIndex === undefined ? sourceIndex + 1 : targetIndex;
        if (insertionIndex < 0 || insertionIndex > slideCount) {
          return toolFailure(`Invalid targetIndex ${insertionIndex}. Must be 0-${slideCount}.`);
        }

        const sourceSlide = slides.items[sourceIndex];
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

        return `Duplicated slide ${sourceIndex + 1} to position ${insertionIndex + 1} using native slide export/import.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
