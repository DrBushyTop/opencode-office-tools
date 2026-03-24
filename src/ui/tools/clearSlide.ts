import type { Tool } from "@github/copilot-sdk";

export const clearSlide: Tool = {
  name: "clear_slide",
  description: "Remove all shapes (text boxes, images, etc.) from a specific slide, leaving it empty. Useful for replacing slide content entirely.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        type: "number",
        description: "0-based slide index. Use 0 for first slide, 1 for second, etc.",
      },
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

        const slideCount = slides.items.length;

        if (slideIndex < 0 || slideIndex >= slideCount) {
          return {
            textResultForLlm: `Invalid slideIndex ${slideIndex}. Must be 0-${slideCount - 1} (current slide count: ${slideCount})`,
            resultType: "failure",
            error: "Invalid slideIndex",
            toolTelemetry: {},
          };
        }

        const slide = slides.items[slideIndex];
        const shapes = slide.shapes;
        shapes.load("items");
        await context.sync();

        // Delete all shapes from the slide
        const shapeCount = shapes.items.length;
        for (const shape of shapes.items) {
          shape.delete();
        }
        await context.sync();

        return `Cleared slide ${slideIndex + 1}. Removed ${shapeCount} shape(s).`;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
