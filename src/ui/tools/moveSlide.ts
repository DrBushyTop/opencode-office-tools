import type { Tool } from "./types";
import { isPowerPointRequirementSetSupported, toolFailure } from "./powerpointShared";

export const moveSlide: Tool = {
  name: "move_slide",
  description: "Move a slide to a new zero-based position in the PowerPoint presentation.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: { type: "number", description: "0-based index of the slide to move." },
      targetIndex: { type: "number", description: "0-based index where the slide should be moved." },
    },
    required: ["slideIndex", "targetIndex"],
  },
  handler: async (args) => {
    const { slideIndex, targetIndex } = args as { slideIndex: number; targetIndex: number };
    if (!isPowerPointRequirementSetSupported("1.8")) {
      return toolFailure("Moving slides requires PowerPointApi 1.8.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        const slide = slides.items[slideIndex];
        if (!slide) return toolFailure(`Invalid slideIndex ${slideIndex}.`);
        if (targetIndex < 0 || targetIndex >= slides.items.length) return toolFailure(`Invalid targetIndex ${targetIndex}.`);
        slide.moveTo(targetIndex);
        await context.sync();
        return `Moved slide ${slideIndex + 1} to position ${targetIndex + 1}.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
