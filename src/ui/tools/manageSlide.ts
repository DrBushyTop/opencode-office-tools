import type { Tool } from "./types";
import { isPowerPointRequirementSetSupported, toolFailure } from "./powerpointShared";
import { getPowerPointContextSnapshot } from "./powerpointContext";
import { z } from "zod";

const manageSlideArgsSchema = z.object({
  action: z.enum(["create", "duplicate", "delete", "move", "clear"]),
  slideIndex: z.number().optional(),
  sourceIndex: z.number().optional(),
  targetIndex: z.number().optional(),
  slideMasterId: z.string().optional(),
  layoutId: z.string().optional(),
  formatting: z.enum(["KeepSourceFormatting", "UseDestinationTheme"]).optional(),
});

type ManageSlideArgs = z.infer<typeof manageSlideArgsSchema>;

function isNonNegativeInteger(value: unknown): value is number {
  return Number.isInteger(value) && (value as number) >= 0;
}

export const manageSlide: Tool = {
  name: "manage_slide",
  description: "Create, delete, move, or clear PowerPoint slides with one generic slide-management tool.",
  parameters: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: ["create", "delete", "move", "clear"],
        description: "Slide operation to perform.",
      },
      slideIndex: {
        type: "number",
        description: "0-based target slide index for delete, move, or clear.",
      },
      targetIndex: {
        type: "number",
        description: "0-based destination or insertion index for create or move.",
      },
      slideMasterId: {
        type: "string",
        description: "Optional slide master id for create.",
      },
      layoutId: {
        type: "string",
        description: "Optional layout id for create.",
      },
      formatting: {
        type: "string",
        enum: ["KeepSourceFormatting", "UseDestinationTheme"],
        description: "Formatting behavior hint used when the host needs to insert created content. Default KeepSourceFormatting.",
      },
    },
    required: ["action"],
  },
  handler: async (args) => {
    const parsedArgs = manageSlideArgsSchema.safeParse(args);
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }
    const resolvedArgs: ManageSlideArgs = { ...parsedArgs.data };
    const snapshot = getPowerPointContextSnapshot();
    if (resolvedArgs.slideIndex === undefined && snapshot?.activeSlideIndex !== undefined) {
      resolvedArgs.slideIndex = snapshot.activeSlideIndex;
    }

    const { action, slideIndex, sourceIndex, targetIndex, slideMasterId, layoutId, formatting = "KeepSourceFormatting" } = resolvedArgs;

    if (slideIndex !== undefined && !isNonNegativeInteger(slideIndex)) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }
    if (sourceIndex !== undefined && !isNonNegativeInteger(sourceIndex)) {
      return toolFailure("sourceIndex must be a non-negative integer.");
    }
    if (targetIndex !== undefined && !isNonNegativeInteger(targetIndex)) {
      return toolFailure("targetIndex must be a non-negative integer.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const slideCount = slides.items.length;

        switch (action) {
          case "create": {
            if (!isPowerPointRequirementSetSupported("1.3")) {
              return toolFailure("Creating slides requires PowerPointApi 1.3.");
            }
            if (targetIndex !== undefined && targetIndex > slideCount) {
              return toolFailure(`Invalid targetIndex ${targetIndex}. Must be 0-${slideCount}.`);
            }
            if (targetIndex !== undefined && targetIndex !== slideCount && !isPowerPointRequirementSetSupported("1.8")) {
              return toolFailure("Creating a slide at a specific position requires PowerPointApi 1.8.");
            }

            context.presentation.slides.add({
              ...(slideMasterId ? { slideMasterId } : {}),
              ...(layoutId ? { layoutId } : {}),
            });
            await context.sync();

            slides.load("items");
            await context.sync();

            const createdIndex = slides.items.length - 1;
            const createdSlide = slides.items[createdIndex];

            if (targetIndex !== undefined && targetIndex !== createdIndex) {
              createdSlide.moveTo(targetIndex);
              await context.sync();
              return `Created slide at position ${targetIndex + 1}.`;
            }

            return `Created slide ${createdIndex + 1}.`;
          }

          case "duplicate": {
            if (!isPowerPointRequirementSetSupported("1.8")) {
              return toolFailure("Duplicating slides requires PowerPointApi 1.8.");
            }
            if (sourceIndex === undefined) {
              return toolFailure("sourceIndex is required for duplicate.");
            }
            if (slideCount === 0) {
              return toolFailure("Presentation has no slides.");
            }
            if (sourceIndex >= slideCount) {
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

            return `Duplicated slide ${sourceIndex + 1} to position ${insertionIndex + 1}.`;
          }

          case "delete": {
            if (slideIndex === undefined) {
              return toolFailure("slideIndex is required for delete.");
            }
            const slide = slides.items[slideIndex];
            if (!slide) {
              return toolFailure(`Invalid slideIndex ${slideIndex}.`);
            }
            slide.delete();
            await context.sync();
            return `Deleted slide ${slideIndex + 1}.`;
          }

          case "move": {
            if (!isPowerPointRequirementSetSupported("1.8")) {
              return toolFailure("Moving slides requires PowerPointApi 1.8.");
            }
            if (slideIndex === undefined) {
              return toolFailure("slideIndex is required for move.");
            }
            if (targetIndex === undefined) {
              return toolFailure("targetIndex is required for move.");
            }

            const slide = slides.items[slideIndex];
            if (!slide) {
              return toolFailure(`Invalid slideIndex ${slideIndex}.`);
            }
            if (targetIndex >= slideCount) {
              return toolFailure(`Invalid targetIndex ${targetIndex}. Must be 0-${slideCount - 1}.`);
            }

            slide.moveTo(targetIndex);
            await context.sync();
            return `Moved slide ${slideIndex + 1} to position ${targetIndex + 1}.`;
          }

          case "clear": {
            if (!isPowerPointRequirementSetSupported("1.3")) {
              return toolFailure("Clearing slide shapes requires PowerPointApi 1.3.");
            }
            if (slideIndex === undefined) {
              return toolFailure("slideIndex is required for clear.");
            }
            const slide = slides.items[slideIndex];
            if (!slide) {
              return toolFailure(`Invalid slideIndex ${slideIndex}.`);
            }

            slide.shapes.load("items");
            await context.sync();

            const shapeCount = slide.shapes.items.length;
            for (const shape of slide.shapes.items) {
              shape.delete();
            }
            await context.sync();

            return `Cleared slide ${slideIndex + 1}. Removed ${shapeCount} shape(s).`;
          }

          default:
            return toolFailure(`Unsupported action ${action}.`);
        }
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
