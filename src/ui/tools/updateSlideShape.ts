import type { Tool } from "@github/copilot-sdk";

export const updateSlideShape: Tool = {
  name: "update_slide_shape",
  description: "Update the text content of an existing shape on a slide. Use get_presentation_content first to see existing shapes and their indices.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        type: "number",
        description: "0-based slide index. Use 0 for first slide, 1 for second, etc.",
      },
      shapeIndex: {
        type: "number",
        description: "0-based shape index within the slide. Use get_presentation_content to see available shapes.",
      },
      text: {
        type: "string",
        description: "The new text content for the shape.",
      },
    },
    required: ["slideIndex", "shapeIndex", "text"],
  },
  handler: async (args) => {
    const { slideIndex, shapeIndex, text } = args as { slideIndex: number; shapeIndex: number; text: string };

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

        const shapeCount = shapes.items.length;

        if (shapeIndex < 0 || shapeIndex >= shapeCount) {
          return {
            textResultForLlm: `Invalid shapeIndex ${shapeIndex}. Slide ${slideIndex + 1} has ${shapeCount} shape(s) (indices 0-${shapeCount - 1}).`,
            resultType: "failure",
            error: "Invalid shapeIndex",
            toolTelemetry: {},
          };
        }

        const shape = shapes.items[shapeIndex];
        
        try {
          shape.textFrame.textRange.text = text;
          await context.sync();
        } catch (textError: any) {
          return {
            textResultForLlm: `Shape at index ${shapeIndex} does not support text content: ${textError.message}`,
            resultType: "failure",
            error: textError.message,
            toolTelemetry: {},
          };
        }

        return `Updated shape ${shapeIndex + 1} on slide ${slideIndex + 1} with new text.`;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
