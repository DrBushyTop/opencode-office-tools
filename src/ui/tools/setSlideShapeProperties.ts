import type { Tool } from "./types";
import { loadTextFrames } from "./powerpointText";
import { toolFailure } from "./powerpointShared";

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

function validateRange(name: string, value: unknown, { min, max }: { min?: number; max?: number } = {}) {
  if (!isFiniteNumber(value)) return null;
  if (min !== undefined && value < min) return `${name} must be >= ${min}.`;
  if (max !== undefined && value > max) return `${name} must be <= ${max}.`;
  return null;
}

export const setSlideShapeProperties: Tool = {
  name: "set_slide_shape_properties",
  description: `Update PowerPoint shape properties by shape id or shape index.

Supports text, geometry, visibility, rotation, alt text, fill color, and line formatting. Use get_slide_shapes first to discover target ids and indices.`,
  parameters: {
    type: "object",
    properties: {
      slideIndex: { type: "number", description: "0-based slide index." },
      shapeIndex: { type: "number", description: "0-based shape index within the slide." },
      shapeId: { type: "string", description: "Unique PowerPoint shape id. Preferred when available." },
      text: { type: "string", description: "Replacement text for shapes that support text." },
      name: { type: "string" },
      left: { type: "number" },
      top: { type: "number" },
      width: { type: "number" },
      height: { type: "number" },
      rotation: { type: "number" },
      visible: { type: "boolean" },
      altTextTitle: { type: "string" },
      altTextDescription: { type: "string" },
      fillColor: { type: "string" },
      fillTransparency: { type: "number" },
      clearFill: { type: "boolean" },
      lineColor: { type: "string" },
      lineWeight: { type: "number" },
      lineTransparency: { type: "number" },
      lineVisible: { type: "boolean" },
    },
    required: ["slideIndex"],
  },
  handler: async (args) => {
    const update = args as Record<string, unknown> & { slideIndex: number; shapeIndex?: number; shapeId?: string; text?: string };
    if (update.shapeId === undefined && update.shapeIndex === undefined) {
      return toolFailure("Provide shapeId or shapeIndex.");
    }
    if (!Number.isInteger(update.slideIndex) || update.slideIndex < 0) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }
    if (update.shapeIndex !== undefined && (!Number.isInteger(update.shapeIndex) || update.shapeIndex < 0)) {
      return toolFailure("shapeIndex must be a non-negative integer.");
    }

    const validationError = [
      validateRange("width", update.width, { min: 0 }),
      validateRange("height", update.height, { min: 0 }),
      validateRange("fillTransparency", update.fillTransparency, { min: 0, max: 1 }),
      validateRange("lineWeight", update.lineWeight, { min: 0 }),
      validateRange("lineTransparency", update.lineTransparency, { min: 0, max: 1 }),
    ].find(Boolean);
    if (validationError) {
      return toolFailure(validationError);
    }

    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const slide = slides.items[update.slideIndex];
        if (!slide) {
          return toolFailure(`Invalid slideIndex ${update.slideIndex}.`);
        }

        slide.shapes.load("items");
        await context.sync();

        const shape = update.shapeId
          ? slide.shapes.getItemOrNullObject(update.shapeId)
          : slide.shapes.items[update.shapeIndex as number];

        if (!shape) {
          return toolFailure("Target shape not found.");
        }

        if (update.shapeId) {
          shape.load(["isNullObject", "id"]);
          await context.sync();
          if ((shape as PowerPoint.Shape & { isNullObject?: boolean }).isNullObject) {
            return toolFailure(`Shape ${update.shapeId} was not found on slide ${update.slideIndex + 1}.`);
          }
        }

        const target = shape as PowerPoint.Shape;
        if (typeof update.name === "string") target.name = update.name;
        if (typeof update.left === "number") target.left = update.left;
        if (typeof update.top === "number") target.top = update.top;
        if (typeof update.width === "number") target.width = update.width;
        if (typeof update.height === "number") target.height = update.height;
        if (typeof update.rotation === "number") target.rotation = update.rotation;
        if (typeof update.visible === "boolean") target.visible = update.visible;
        if (typeof update.altTextTitle === "string") target.altTextTitle = update.altTextTitle;
        if (typeof update.altTextDescription === "string") target.altTextDescription = update.altTextDescription;

        if (update.clearFill) {
          target.fill.clear();
        } else if (typeof update.fillColor === "string") {
          target.fill.setSolidColor(update.fillColor);
        }
        if (typeof update.fillTransparency === "number") target.fill.transparency = update.fillTransparency;
        if (typeof update.lineColor === "string") target.lineFormat.color = update.lineColor;
        if (typeof update.lineWeight === "number") target.lineFormat.weight = update.lineWeight;
        if (typeof update.lineTransparency === "number") target.lineFormat.transparency = update.lineTransparency;
        if (typeof update.lineVisible === "boolean") target.lineFormat.visible = update.lineVisible;

        if (typeof update.text === "string") {
          const [frame] = await loadTextFrames(context, [target]);
          if (!frame || frame.isNullObject) {
            return toolFailure("Target shape does not support text.");
          }
          frame.textRange.text = update.text;
        }

        await context.sync();
        return `Updated shape ${update.shapeId || update.shapeIndex} on slide ${update.slideIndex + 1}.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
