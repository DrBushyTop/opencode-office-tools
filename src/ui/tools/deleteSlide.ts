import type { Tool } from "./types";
import { toolFailure } from "./powerpointShared";

export const deleteSlide: Tool = {
  name: "delete_slide",
  description: "Delete a slide from the PowerPoint presentation.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: { type: "number", description: "0-based slide index to delete." },
    },
    required: ["slideIndex"],
  },
  handler: async (args) => {
    const { slideIndex } = args as { slideIndex: number };
    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const slide = slides.items[slideIndex];
        if (!slide) return toolFailure(`Invalid slideIndex ${slideIndex}.`);
        slide.delete();
        await context.sync();
        return `Deleted slide ${slideIndex + 1}.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
