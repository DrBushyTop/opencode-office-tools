import type { Tool } from "./types";
import { loadTextFrames } from "./powerpointText";
import { resolveSlideShapeByIdWithXmlFallback } from "./powerpointShapeTarget";
import { resolvePowerPointTargetingArgs } from "./powerpointContext";
import {
  formatAvailableShapeTargets,
  invalidSlideIndexMessage,
  isPowerPointRequirementSetSupported,
  supportsPowerPointPlaceholders,
  toolFailure,
} from "./powerpointShared";

type ManageSlideShapesAction = "create" | "update" | "delete" | "group" | "ungroup";
type CreateShapeType = "textBox" | "geometricShape" | "line";
type ConnectorType = "Straight" | "Elbow" | "Curve";
type ParagraphAlignment = "Left" | "Center" | "Right" | "Justify" | "JustifyLow" | "Distributed" | "ThaiDistributed";
type UnderlineStyle = "None" | "Single" | "Double" | "Heavy" | "Dotted" | "DottedHeavy" | "Dash" | "DashHeavy" | "DashLong" | "DashLongHeavy" | "DotDash" | "DotDashHeavy" | "DotDotDash" | "DotDotDashHeavy" | "Wavy" | "WavyHeavy" | "WavyDouble";
type TextAutoSize = "AutoSizeNone" | "AutoSizeTextToFitShape" | "AutoSizeShapeToFitText";
type VerticalAlignment = "Top" | "Middle" | "Bottom" | "TopCentered" | "MiddleCentered" | "BottomCentered";

interface ManageSlideShapesArgs {
  action: ManageSlideShapesAction;
  slideIndex?: number;
  shapeId?: string | number;
  shapeIndex?: number;
  shapeIds?: (string | number)[];
  placeholderType?: string;
  shapeType?: CreateShapeType;
  geometricShapeType?: string;
  connectorType?: ConnectorType;
  text?: string;
  name?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  rotation?: number;
  visible?: boolean;
  altTextTitle?: string;
  altTextDescription?: string;
  fillColor?: string;
  fillTransparency?: number;
  clearFill?: boolean;
  lineColor?: string;
  lineWeight?: number;
  lineTransparency?: number;
  lineVisible?: boolean;
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: UnderlineStyle;
  strikethrough?: boolean;
  allCaps?: boolean;
  smallCaps?: boolean;
  subscript?: boolean;
  superscript?: boolean;
  doubleStrikethrough?: boolean;
  paragraphAlignment?: ParagraphAlignment;
  bulletVisible?: boolean;
  indentLevel?: number;
  textAutoSize?: TextAutoSize;
  wordWrap?: boolean;
  verticalAlignment?: VerticalAlignment;
  marginLeft?: number;
  marginRight?: number;
  marginTop?: number;
  marginBottom?: number;
}

function isNonNegativeInteger(value: unknown): value is number {
  return Number.isInteger(value) && (value as number) >= 0;
}

function normalizeColor(value: string) {
  const trimmed = value.trim();
  return /^(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/.test(trimmed)
    ? `#${trimmed}`
    : trimmed;
}

function validateNonNegativeDimensions(args: Pick<ManageSlideShapesArgs, "width" | "height">) {
  return [
    validateRange("width", args.width, { min: 0 }),
    validateRange("height", args.height, { min: 0 }),
  ].find(Boolean) ?? null;
}

function validateGeometricShapeType(value: string) {
  const supportedValues = Object.values(PowerPoint.GeometricShapeType) as string[];

  if (supportedValues.includes(value)) {
    return null;
  }

  const examples = supportedValues.slice(0, 8).join(", ");
  return `geometricShapeType must be a valid PowerPoint.GeometricShapeType value, such as ${examples}.`;
}

function validateRange(name: string, value: unknown, { min, max }: { min?: number; max?: number } = {}) {
  if (value === undefined) return null;
  if (typeof value !== "number" || !Number.isFinite(value)) return `${name} must be a finite number.`;
  if (min !== undefined && value < min) return `${name} must be >= ${min}.`;
  if (max !== undefined && value > max) return `${name} must be <= ${max}.`;
  return null;
}

function hasTarget(args: ManageSlideShapesArgs) {
  return args.shapeId !== undefined || args.shapeIndex !== undefined || args.placeholderType !== undefined;
}

function requiresTextAccess(args: ManageSlideShapesArgs) {
  return [
    args.text,
    args.fontName,
    args.fontSize,
    args.fontColor,
    args.bold,
    args.italic,
    args.underline,
    args.strikethrough,
    args.allCaps,
    args.smallCaps,
    args.subscript,
    args.superscript,
    args.doubleStrikethrough,
    args.paragraphAlignment,
    args.bulletVisible,
    args.indentLevel,
    args.textAutoSize,
    args.wordWrap,
    args.verticalAlignment,
    args.marginLeft,
    args.marginRight,
    args.marginTop,
    args.marginBottom,
  ].some((value) => value !== undefined);
}

async function resolveTargetShape(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  slideIndex: number,
  args: ManageSlideShapesArgs,
): Promise<PowerPoint.Shape | { error: string }> {
  if (args.shapeId !== undefined) {
    try {
      const resolved = await resolveSlideShapeByIdWithXmlFallback(context, slide, slideIndex, args.shapeId);
      return resolved.shape;
    } catch (error: unknown) {
      return { error: error instanceof Error ? error.message : String(error) };
    }
  }

  slide.shapes.load("items/id,name,type");
  await context.sync();

  if (args.shapeIndex !== undefined) {
    const shape = slide.shapes.items[args.shapeIndex];
    if (!shape) {
      return { error: `Invalid shapeIndex ${args.shapeIndex}. ${formatAvailableShapeTargets(slideIndex, slide.shapes.items)}` };
    }
    return shape;
  }

  if (args.placeholderType) {
    if (!supportsPowerPointPlaceholders()) {
      return { error: "Placeholder targeting requires PowerPointApi 1.8." };
    }
    for (const shape of slide.shapes.items) {
      shape.load("type");
    }
    await context.sync();

    const placeholders = slide.shapes.items.filter((shape) => shape.type === PowerPoint.ShapeType.placeholder);
    for (const shape of placeholders) {
      shape.placeholderFormat.load("type");
    }
    await context.sync();

    const target = placeholders.find((shape) => String(shape.placeholderFormat.type) === args.placeholderType) || null;
    if (!target) {
      const placeholderTypes = placeholders.map((shape) => String(shape.placeholderFormat.type));
      const available = placeholderTypes.length > 0
        ? `Available placeholder types on slide ${slideIndex + 1}: ${placeholderTypes.join(", ")}.`
        : `Slide ${slideIndex + 1} has no placeholder shapes.`;
      return { error: `Placeholder type ${args.placeholderType} was not found on slide ${slideIndex + 1}. ${available}` };
    }
    return target;
  }

  return { error: "Provide shapeId, shapeIndex, or placeholderType." };
}

async function applyShapeMutation(
  context: PowerPoint.RequestContext,
  shape: PowerPoint.Shape,
  args: ManageSlideShapesArgs,
): Promise<string | null> {
  const needsText = requiresTextAccess(args);
  let frame: PowerPoint.TextFrame | null = null;

  if (needsText) {
    if (
      !isPowerPointRequirementSetSupported("1.8")
      && [args.allCaps, args.smallCaps, args.subscript, args.superscript, args.strikethrough, args.doubleStrikethrough].some((value) => value !== undefined)
    ) {
      return "Advanced font effects require PowerPointApi 1.8.";
    }
    if (args.indentLevel !== undefined && !isPowerPointRequirementSetSupported("1.10")) {
      return "indentLevel requires PowerPointApi 1.10.";
    }

    const [loadedFrame] = await loadTextFrames(context, [shape]);
    if (!loadedFrame || loadedFrame.isNullObject) {
      return "Target shape does not support text.";
    }
    frame = loadedFrame;
  }

  if (typeof args.name === "string") shape.name = args.name;
  if (typeof args.left === "number") shape.left = args.left;
  if (typeof args.top === "number") shape.top = args.top;
  if (typeof args.width === "number") shape.width = args.width;
  if (typeof args.height === "number") shape.height = args.height;
  if (typeof args.rotation === "number") shape.rotation = args.rotation;
  if (typeof args.visible === "boolean") shape.visible = args.visible;
  if (typeof args.altTextTitle === "string") shape.altTextTitle = args.altTextTitle;
  if (typeof args.altTextDescription === "string") shape.altTextDescription = args.altTextDescription;

  if (args.clearFill) {
    shape.fill.clear();
  } else if (typeof args.fillColor === "string") {
    shape.fill.setSolidColor(normalizeColor(args.fillColor));
  }
  if (typeof args.fillTransparency === "number") shape.fill.transparency = args.fillTransparency;
  if (typeof args.lineColor === "string") shape.lineFormat.color = normalizeColor(args.lineColor);
  if (typeof args.lineWeight === "number") shape.lineFormat.weight = args.lineWeight;
  if (typeof args.lineTransparency === "number") shape.lineFormat.transparency = args.lineTransparency;
  if (typeof args.lineVisible === "boolean") shape.lineFormat.visible = args.lineVisible;

  if (frame) {
    if (typeof args.text === "string") frame.textRange.text = args.text;

    const font = frame.textRange.font;
    const paragraph = frame.textRange.paragraphFormat;

    if (typeof args.fontName === "string") font.name = args.fontName;
    if (typeof args.fontSize === "number") font.size = args.fontSize;
    if (typeof args.fontColor === "string") font.color = normalizeColor(args.fontColor);
    if (typeof args.bold === "boolean") font.bold = args.bold;
    if (typeof args.italic === "boolean") font.italic = args.italic;
    if (typeof args.underline === "string") font.underline = args.underline;
    if (typeof args.strikethrough === "boolean") font.strikethrough = args.strikethrough;
    if (typeof args.allCaps === "boolean") font.allCaps = args.allCaps;
    if (typeof args.smallCaps === "boolean") font.smallCaps = args.smallCaps;
    if (typeof args.subscript === "boolean") font.subscript = args.subscript;
    if (typeof args.superscript === "boolean") font.superscript = args.superscript;
    if (typeof args.doubleStrikethrough === "boolean") font.doubleStrikethrough = args.doubleStrikethrough;

    if (typeof args.paragraphAlignment === "string") paragraph.horizontalAlignment = args.paragraphAlignment;
    if (typeof args.bulletVisible === "boolean") paragraph.bulletFormat.visible = args.bulletVisible;
    if (typeof args.indentLevel === "number") paragraph.indentLevel = args.indentLevel;

    if (typeof args.textAutoSize === "string") frame.autoSizeSetting = args.textAutoSize;
    if (typeof args.wordWrap === "boolean") frame.wordWrap = args.wordWrap;
    if (typeof args.verticalAlignment === "string") frame.verticalAlignment = args.verticalAlignment;
    if (typeof args.marginLeft === "number") frame.leftMargin = args.marginLeft;
    if (typeof args.marginRight === "number") frame.rightMargin = args.marginRight;
    if (typeof args.marginTop === "number") frame.topMargin = args.marginTop;
    if (typeof args.marginBottom === "number") frame.bottomMargin = args.marginBottom;
  }

  return null;
}

export const manageSlideShapes: Tool = {
  name: "manage_slide_shapes",
  description: "Create, update, delete, group, or ungroup PowerPoint shapes with generic geometry, styling, and text formatting controls.",
  parameters: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: ["create", "update", "delete", "group", "ungroup"],
        description: "Shape operation to perform.",
      },
      slideIndex: { type: "number", description: "0-based slide index." },
      shapeId: {
        anyOf: [{ type: "string" }, { type: "number" }],
        description: "Existing Office shape id, or an exported XML p:cNvPr id after an Open XML slide replacement.",
      },
      shapeIndex: { type: "number", description: "Existing 0-based shape index on the slide." },
      shapeIds: {
        type: "array",
        items: { anyOf: [{ type: "string" }, { type: "number" }] },
        description: "Array of shape ids to group together. Required for the group action. Must contain at least 2 shapes.",
      },
      placeholderType: { type: "string", description: "Optional placeholder type to target for update or delete, such as Title, Body, Subtitle, or Content." },
      shapeType: { type: "string", enum: ["textBox", "geometricShape", "line"], description: "Shape type to create." },
      geometricShapeType: { type: "string", description: "Geometric shape type for create. Must be a valid PowerPoint.GeometricShapeType value such as Rectangle, Ellipse, Chevron, or RightArrow." },
      connectorType: { type: "string", enum: ["Straight", "Elbow", "Curve"], description: "Optional connector type for line creation. Default Straight." },
      text: { type: "string", description: "Text content to set or create." },
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
      fontName: { type: "string" },
      fontSize: { type: "number" },
      fontColor: { type: "string" },
      bold: { type: "boolean" },
      italic: { type: "boolean" },
      underline: {
        type: "string",
        enum: ["None", "Single", "Double", "Heavy", "Dotted", "DottedHeavy", "Dash", "DashHeavy", "DashLong", "DashLongHeavy", "DotDash", "DotDashHeavy", "DotDotDash", "DotDotDashHeavy", "Wavy", "WavyHeavy", "WavyDouble"],
      },
      strikethrough: { type: "boolean" },
      allCaps: { type: "boolean" },
      smallCaps: { type: "boolean" },
      subscript: { type: "boolean" },
      superscript: { type: "boolean" },
      doubleStrikethrough: { type: "boolean" },
      paragraphAlignment: { type: "string", enum: ["Left", "Center", "Right", "Justify", "JustifyLow", "Distributed", "ThaiDistributed"] },
      bulletVisible: { type: "boolean" },
      indentLevel: { type: "number" },
      textAutoSize: { type: "string", enum: ["AutoSizeNone", "AutoSizeTextToFitShape", "AutoSizeShapeToFitText"] },
      wordWrap: { type: "boolean" },
      verticalAlignment: { type: "string", enum: ["Top", "Middle", "Bottom", "TopCentered", "MiddleCentered", "BottomCentered"] },
      marginLeft: { type: "number" },
      marginRight: { type: "number" },
      marginTop: { type: "number" },
      marginBottom: { type: "number" },
    },
    required: ["action", "slideIndex"],
  },
  handler: async (args) => {
    const update = resolvePowerPointTargetingArgs(args as ManageSlideShapesArgs);

    if (!isNonNegativeInteger(update.slideIndex)) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }
    const slideIndex = update.slideIndex;
    if (update.shapeIndex !== undefined && !isNonNegativeInteger(update.shapeIndex)) {
      return toolFailure("shapeIndex must be a non-negative integer.");
    }

    const validationError = [
      validateRange("width", update.width),
      validateRange("height", update.height),
      validateRange("fillTransparency", update.fillTransparency, { min: 0, max: 1 }),
      validateRange("lineWeight", update.lineWeight, { min: 0 }),
      validateRange("lineTransparency", update.lineTransparency, { min: 0, max: 1 }),
      validateRange("fontSize", update.fontSize, { min: 1 }),
      validateRange("indentLevel", update.indentLevel, { min: 0 }),
      validateRange("marginLeft", update.marginLeft, { min: 0 }),
      validateRange("marginRight", update.marginRight, { min: 0 }),
      validateRange("marginTop", update.marginTop, { min: 0 }),
      validateRange("marginBottom", update.marginBottom, { min: 0 }),
    ].find(Boolean);
    if (validationError) {
      return toolFailure(validationError);
    }

    if (update.action === "create" && !isPowerPointRequirementSetSupported("1.4")) {
      return toolFailure("Creating shapes requires PowerPointApi 1.4.");
    }
    if ((update.action === "update" || update.action === "delete") && !isPowerPointRequirementSetSupported("1.3")) {
      return toolFailure("Updating or deleting shapes requires PowerPointApi 1.3.");
    }
    if ((update.action === "group" || update.action === "ungroup") && !isPowerPointRequirementSetSupported("1.8")) {
      return toolFailure("Grouping and ungrouping shapes requires PowerPointApi 1.8.");
    }

    if (update.action === "group") {
      if (!update.shapeIds || !Array.isArray(update.shapeIds) || update.shapeIds.length < 2) {
        return toolFailure("shapeIds must be an array of at least 2 shape ids for the group action.");
      }
    } else if (update.action === "create") {
      if (!update.shapeType) {
        return toolFailure("shapeType is required for create.");
      }
      if (update.shapeType === "geometricShape" && !update.geometricShapeType) {
        return toolFailure("geometricShapeType is required when creating a geometric shape.");
      }
      if (update.shapeType === "geometricShape" && update.geometricShapeType) {
        const geometricShapeTypeError = validateGeometricShapeType(update.geometricShapeType);
        if (geometricShapeTypeError) {
          return toolFailure(geometricShapeTypeError);
        }
      }
      if (update.shapeType === "line" && requiresTextAccess(update)) {
        return toolFailure("Line shapes do not support text or text formatting.");
      }
      if (update.shapeType !== "line") {
        const dimensionError = validateNonNegativeDimensions(update);
        if (dimensionError) {
          return toolFailure(dimensionError);
        }
      }
    } else if (update.action === "ungroup" && !hasTarget(update)) {
      return toolFailure("Provide shapeId or shapeIndex for the group shape to ungroup.");
    } else if ((update.action === "update" || update.action === "delete") && !hasTarget(update)) {
      return toolFailure("Provide shapeId, shapeIndex, or placeholderType.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const slide = slides.items[slideIndex];
        if (!slide) {
          return toolFailure(invalidSlideIndexMessage(slideIndex, slides.items.length));
        }

        if (update.action === "create") {
          const options = {
            ...(typeof update.left === "number" ? { left: update.left } : {}),
            ...(typeof update.top === "number" ? { top: update.top } : {}),
            ...(typeof update.width === "number" ? { width: update.width } : {}),
            ...(typeof update.height === "number" ? { height: update.height } : {}),
          };

          const created = update.shapeType === "textBox"
            ? slide.shapes.addTextBox(update.text || "", options)
            : update.shapeType === "geometricShape"
              ? slide.shapes.addGeometricShape(update.geometricShapeType as PowerPoint.GeometricShapeType, options)
              : slide.shapes.addLine(update.connectorType || "Straight", options);

          created.load(["id", "name"]);
          await context.sync();

          const mutationError = await applyShapeMutation(context, created, update);
          if (mutationError) {
            created.delete();
            await context.sync();
            return toolFailure(mutationError);
          }

          await context.sync();
          return `Created ${update.shapeType} ${created.id} on slide ${slideIndex + 1}.`;
        }

        if (update.action === "group") {
          slide.shapes.load("items/id");
          await context.sync();

          // Resolve shape ids to Shape objects (not strings) to avoid Mac crash in addGroup
          const shapeIdStrings = update.shapeIds!.map((id) => String(id));
          const availableIds = slide.shapes.items.map((s: PowerPoint.Shape) => s.id);
          const missing = shapeIdStrings.filter((id) => !availableIds.includes(id));
          if (missing.length > 0) {
            return toolFailure(`Shape id(s) not found on slide ${slideIndex + 1}: ${missing.join(", ")}. ${formatAvailableShapeTargets(slideIndex, slide.shapes.items)}`);
          }

          // Pass Shape objects rather than string IDs — passing strings crashes PowerPoint on Mac
          // (BUG_IN_CLIENT_OF_LIBMALLOC / POINTER_BEING_FREED_WAS_NOT_ALLOCATED in OLEAutomation)
          const shapeObjects = slide.shapes.items.filter((s: PowerPoint.Shape) => shapeIdStrings.includes(s.id));
          const group = slide.shapes.addGroup(shapeObjects);
          group.load(["id", "name"]);
          await context.sync();

          if (typeof update.name === "string") {
            group.name = update.name;
            await context.sync();
          }

          return `Grouped ${shapeIdStrings.length} shapes into group ${group.id}${update.name ? ` (${JSON.stringify(update.name)})` : ""} on slide ${slideIndex + 1}.`;
        }

        if (update.action === "ungroup") {
          const resolved = await resolveTargetShape(context, slide, slideIndex, update);
          if ("error" in resolved) {
            return toolFailure(resolved.error);
          }

          resolved.load("type");
          await context.sync();

          if (resolved.type !== PowerPoint.ShapeType.group) {
            return toolFailure(`Shape is not a group (type=${resolved.type}). Only group shapes can be ungrouped.`);
          }

          resolved.group.ungroup();
          await context.sync();
          return `Ungrouped shape on slide ${slideIndex + 1}.`;
        }

        const resolved = await resolveTargetShape(context, slide, slideIndex, update);
        if ("error" in resolved) {
          return toolFailure(resolved.error);
        }

        if (update.action === "delete") {
          resolved.delete();
          await context.sync();
          return `Deleted shape on slide ${slideIndex + 1}.`;
        }

        if (update.width !== undefined || update.height !== undefined) {
          resolved.load("type");
          await context.sync();

          if (resolved.type !== PowerPoint.ShapeType.line) {
            const dimensionError = validateNonNegativeDimensions(update);
            if (dimensionError) {
              return toolFailure(dimensionError);
            }
          }
        }

        const mutationError = await applyShapeMutation(context, resolved, update);
        if (mutationError) {
          return toolFailure(mutationError);
        }

        await context.sync();
        return `Updated shape on slide ${slideIndex + 1}.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
