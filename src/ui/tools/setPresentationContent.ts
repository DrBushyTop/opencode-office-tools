import type { Tool } from "./types";
import { loadTextFrames } from "./powerpointText";
import { supportsPowerPointPlaceholders, toolFailure } from "./powerpointShared";

export const setPresentationContent: Tool = {
  name: "set_presentation_content",
  description: "Add or update text on a PowerPoint slide. Can target an existing shape, a placeholder type, or create a new text box. Explicit shape or placeholder targets fail if they cannot be resolved. New slides can optionally use a specific master and layout.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        type: "number",
        description: "0-based slide index. Use the current slide count to add a new slide.",
      },
      text: {
        type: "string",
        description: "The text content to add to the slide.",
      },
      shapeId: {
        type: "string",
        description: "Optional existing shape id to update instead of adding a new text box.",
      },
      placeholderType: {
        type: "string",
        description: "Optional placeholder type to target on the slide, such as Title, Body, Subtitle, Content, Footer, Header, Date, or SlideNumber.",
      },
      slideMasterId: {
        type: "string",
        description: "Optional slide master id to use when creating a new slide.",
      },
      layoutId: {
        type: "string",
        description: "Optional slide layout id to use when creating a new slide.",
      },
      left: { type: "number", description: "Left position for a new text box." },
      top: { type: "number", description: "Top position for a new text box." },
      width: { type: "number", description: "Width for a new text box." },
      height: { type: "number", description: "Height for a new text box." },
    },
    required: ["slideIndex", "text"],
  },
  handler: async (args) => {
    const {
      slideIndex,
      text,
      shapeId,
      placeholderType,
      slideMasterId,
      layoutId,
      left = 50,
      top = 100,
      width = 600,
      height = 400,
    } = args as {
      slideIndex: number;
      text: string;
      shapeId?: string;
      placeholderType?: string;
      slideMasterId?: string;
      layoutId?: string;
      left?: number;
      top?: number;
      width?: number;
      height?: number;
    };

    try {
      return await PowerPoint.run(async (context) => {
        const supportsPlaceholders = supportsPowerPointPlaceholders();
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const slideCount = slides.items.length;
        if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex > slideCount) {
          return toolFailure(`Invalid slideIndex ${slideIndex}. Must be 0-${slideCount}.`);
        }

        if (slideIndex === slideCount) {
          context.presentation.slides.add({
            ...(slideMasterId ? { slideMasterId } : {}),
            ...(layoutId ? { layoutId } : {}),
          });
          await context.sync();
          slides.load("items");
          await context.sync();
        }

        const slide = slides.items[slideIndex];
        if (!slide) {
          return toolFailure(`Invalid slideIndex ${slideIndex}. Must be 0-${slideCount}.`);
        }
        slide.shapes.load("items");
        await context.sync();

        let targetShape: PowerPoint.Shape | null = null;

        if (shapeId) {
          const byId = slide.shapes.getItemOrNullObject(shapeId);
          byId.load("isNullObject");
          await context.sync();
          if (byId.isNullObject) {
            return toolFailure(`Shape ${shapeId} was not found on slide ${slideIndex + 1}.`);
          }
          targetShape = byId;
        }

        if (!targetShape && placeholderType) {
          if (!supportsPlaceholders) {
            return toolFailure("Placeholder targeting requires PowerPointApi 1.8.");
          }
          const placeholders = slide.shapes.items.filter((shape) => shape.type === PowerPoint.ShapeType.placeholder);
          for (const shape of placeholders) {
            shape.placeholderFormat.load("type");
          }
          await context.sync();
          targetShape = placeholders.find((shape) => String(shape.placeholderFormat.type) === placeholderType) || null;
          if (!targetShape) {
            return toolFailure(`Placeholder type ${placeholderType} was not found on slide ${slideIndex + 1}.`);
          }
        }

        if (targetShape) {
          const [frame] = await loadTextFrames(context, [targetShape]);
          if (!frame || frame.isNullObject) {
            return toolFailure("Target shape does not support text.");
          }
          frame.textRange.text = text;
          await context.sync();
          return `Updated text on slide ${slideIndex + 1}${shapeId ? ` for shape ${shapeId}` : placeholderType ? ` using ${placeholderType} placeholder` : ""}.`;
        }

        slide.shapes.addTextBox(text, { left, top, width, height });
        await context.sync();

        return `Added text box to slide ${slideIndex + 1}.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
