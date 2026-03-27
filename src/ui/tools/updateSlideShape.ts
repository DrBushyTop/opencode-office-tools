import type { Tool } from "./types";
import { loadTextFrames } from "./powerpointText";
import { toolFailure } from "./powerpointShared";

export const updateSlideShape: Tool = {
  name: "update_slide_shape",
  description: "Update the text content of an existing shape on a slide. Use get_slide_shapes first to discover shape ids and indices.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        type: "number",
        description: "0-based slide index. Use 0 for first slide, 1 for second, etc.",
      },
      shapeIndex: {
        type: "number",
        description: "0-based shape index within the slide. Use get_slide_shapes to see available shapes.",
      },
      shapeId: {
        type: "string",
        description: "Optional shape id. Preferred over shapeIndex when available.",
      },
      text: {
        type: "string",
        description: "The new text content for the shape.",
      },
    },
    required: ["slideIndex", "text"],
  },
  handler: async (args) => {
    const { slideIndex, shapeIndex, shapeId, text } = args as { slideIndex: number; shapeIndex?: number; shapeId?: string; text: string };

    if (shapeIndex === undefined && !shapeId) {
      return toolFailure("Provide shapeId or shapeIndex.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const slide = slides.items[slideIndex];
        if (!slide) {
          return toolFailure(`Invalid slideIndex ${slideIndex}.`);
        }

        const shapes = slide.shapes;
        shapes.load("items");
        await context.sync();

        let shape: PowerPoint.Shape;
        if (shapeId) {
          const byId = shapes.getItemOrNullObject(shapeId);
          byId.load("isNullObject");
          await context.sync();
          if (byId.isNullObject) {
            return toolFailure(`Shape ${shapeId} was not found on slide ${slideIndex + 1}.`);
          }
          shape = byId;
        } else {
          const shapeCount = shapes.items.length;
          if ((shapeIndex as number) < 0 || (shapeIndex as number) >= shapeCount) {
            return toolFailure(`Invalid shapeIndex ${shapeIndex}. Slide ${slideIndex + 1} has ${shapeCount} shape(s).`);
          }
          shape = shapes.items[shapeIndex as number];
        }

        const [frame] = await loadTextFrames(context, [shape]);
        if (!frame || frame.isNullObject) {
          return toolFailure("Target shape does not support text.");
        }

        frame.textRange.text = text;
        await context.sync();

        return `Updated shape ${shapeId || shapeIndex} on slide ${slideIndex + 1} with new text.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
