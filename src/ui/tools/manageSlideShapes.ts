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
import { z } from "zod";

type ManageSlideShapesAction = "create" | "update" | "delete" | "group" | "ungroup";
type CreateShapeType = "textBox" | "geometricShape" | "line";
type ConnectorType = "Straight" | "Elbow" | "Curve";
type ParagraphAlignment = "Left" | "Center" | "Right" | "Justify" | "JustifyLow" | "Distributed" | "ThaiDistributed";
type UnderlineStyle = "None" | "Single" | "Double" | "Heavy" | "Dotted" | "DottedHeavy" | "Dash" | "DashHeavy" | "DashLong" | "DashLongHeavy" | "DotDash" | "DotDashHeavy" | "DotDotDash" | "DotDotDashHeavy" | "Wavy" | "WavyHeavy" | "WavyDouble";
type TextAutoSize = "AutoSizeNone" | "AutoSizeTextToFitShape" | "AutoSizeShapeToFitText";
type VerticalAlignment = "Top" | "Middle" | "Bottom" | "TopCentered" | "MiddleCentered" | "BottomCentered";

const createOnlyUpdateKeys = ["shapeType", "geometricShapeType", "connectorType"] as const;
const blankableUpdateKeys = [
  "shapeId",
  "placeholderType",
  "name",
  "altTextTitle",
  "altTextDescription",
  "fillColor",
  "lineColor",
  "fontName",
  "fontColor",
] as const;
const blankableCreateKeys = [
  "shapeId",
  "placeholderType",
  "name",
  "altTextTitle",
  "altTextDescription",
  "fillColor",
  "lineColor",
  "fontName",
  "fontColor",
  "geometricShapeType",
] as const;
const updateMutationKeys = [
  "text",
  "name",
  "left",
  "top",
  "width",
  "height",
  "rotation",
  "visible",
  "altTextTitle",
  "altTextDescription",
  "fillColor",
  "fillTransparency",
  "clearFill",
  "lineColor",
  "lineWeight",
  "lineTransparency",
  "lineVisible",
  "fontName",
  "fontSize",
  "fontColor",
  "bold",
  "italic",
  "underline",
  "strikethrough",
  "allCaps",
  "smallCaps",
  "subscript",
  "superscript",
  "doubleStrikethrough",
  "paragraphAlignment",
  "bulletVisible",
  "indentLevel",
  "textAutoSize",
  "wordWrap",
  "verticalAlignment",
  "marginLeft",
  "marginRight",
  "marginTop",
  "marginBottom",
] as const;
const denseUpdateDefaultValues = {
  left: 0,
  top: 0,
  width: 0,
  height: 0,
  rotation: 0,
  visible: true,
  fillTransparency: 0,
  clearFill: false,
  lineWeight: 0,
  lineTransparency: 0,
  lineVisible: true,
  fontSize: 0,
  bold: false,
  italic: false,
  underline: "None",
  strikethrough: false,
  allCaps: false,
  smallCaps: false,
  subscript: false,
  superscript: false,
  doubleStrikethrough: false,
  paragraphAlignment: "Left",
  bulletVisible: false,
  indentLevel: 0,
  textAutoSize: "AutoSizeNone",
  wordWrap: true,
  verticalAlignment: "Top",
  marginLeft: 0,
  marginRight: 0,
  marginTop: 0,
  marginBottom: 0,
} satisfies Partial<ManageSlideShapesArgs>;

const manageSlideShapesArgsSchema = z.object({
  action: z.enum(["create", "update", "delete", "group", "ungroup"]),
  slideIndex: z.number().optional(),
  shapeId: z.union([z.string(), z.number()]).optional(),
  shapeIndex: z.number().optional(),
  shapeIds: z.array(z.union([z.string(), z.number()])).optional(),
  placeholderType: z.string().optional(),
  shapeType: z.enum(["textBox", "geometricShape", "line"]).optional(),
  geometricShapeType: z.string().optional(),
  connectorType: z.enum(["Straight", "Elbow", "Curve"]).optional(),
  text: z.string().optional(),
  name: z.string().optional(),
  left: z.number().optional(),
  top: z.number().optional(),
  width: z.number().optional(),
  height: z.number().optional(),
  rotation: z.number().optional(),
  visible: z.boolean().optional(),
  altTextTitle: z.string().optional(),
  altTextDescription: z.string().optional(),
  fillColor: z.string().optional(),
  fillTransparency: z.number().optional(),
  clearFill: z.boolean().optional(),
  lineColor: z.string().optional(),
  lineWeight: z.number().optional(),
  lineTransparency: z.number().optional(),
  lineVisible: z.boolean().optional(),
  fontName: z.string().optional(),
  fontSize: z.number().optional(),
  fontColor: z.string().optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.enum(["None", "Single", "Double", "Heavy", "Dotted", "DottedHeavy", "Dash", "DashHeavy", "DashLong", "DashLongHeavy", "DotDash", "DotDashHeavy", "DotDotDash", "DotDotDashHeavy", "Wavy", "WavyHeavy", "WavyDouble"]).optional(),
  strikethrough: z.boolean().optional(),
  allCaps: z.boolean().optional(),
  smallCaps: z.boolean().optional(),
  subscript: z.boolean().optional(),
  superscript: z.boolean().optional(),
  doubleStrikethrough: z.boolean().optional(),
  paragraphAlignment: z.enum(["Left", "Center", "Right", "Justify", "JustifyLow", "Distributed", "ThaiDistributed"]).optional(),
  bulletVisible: z.boolean().optional(),
  indentLevel: z.number().optional(),
  textAutoSize: z.enum(["AutoSizeNone", "AutoSizeTextToFitShape", "AutoSizeShapeToFitText"]).optional(),
  wordWrap: z.boolean().optional(),
  verticalAlignment: z.enum(["Top", "Middle", "Bottom", "TopCentered", "MiddleCentered", "BottomCentered"]).optional(),
  marginLeft: z.number().optional(),
  marginRight: z.number().optional(),
  marginTop: z.number().optional(),
  marginBottom: z.number().optional(),
});

type ManageSlideShapesArgs = z.infer<typeof manageSlideShapesArgsSchema>;

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

function isBlankString(value: unknown): value is string {
  return typeof value === "string" && value.trim().length === 0;
}

function deleteKeys<T extends object>(target: T, keys: readonly string[]) {
  for (const key of keys) {
    delete target[key as keyof T];
  }
}

function removeBlankStrings<T extends object>(target: T, keys: readonly string[]) {
  for (const key of keys) {
    if (isBlankString(target[key as keyof T])) {
      delete target[key as keyof T];
    }
  }
}

function pruneDenseDefaultMutationValues(args: ManageSlideShapesArgs) {
  const mutationFieldCount = updateMutationKeys.filter((key) => args[key] !== undefined).length;
  const defaultFieldCount = Object.entries(denseUpdateDefaultValues).filter(([key, value]) => args[key as keyof ManageSlideShapesArgs] === value).length;
  const looksDense = mutationFieldCount >= 10 && defaultFieldCount >= 8;

  if (!looksDense) {
    return args;
  }

  for (const [key, value] of Object.entries(denseUpdateDefaultValues)) {
    if (args[key as keyof ManageSlideShapesArgs] === value) {
      delete args[key as keyof ManageSlideShapesArgs];
    }
  }

  return args;
}

function normalizeCreateArgs(args: ManageSlideShapesArgs): ManageSlideShapesArgs {
  const next: ManageSlideShapesArgs = { ...args };

  if (Array.isArray(next.shapeIds) && next.shapeIds.length === 0) {
    delete next.shapeIds;
  }

  removeBlankStrings(next, blankableCreateKeys);
  deleteKeys(next, ["shapeId", "shapeIndex", "shapeIds", "placeholderType"]);

  if (next.shapeType !== "geometricShape") {
    delete next.geometricShapeType;
  }
  if (next.shapeType !== "line") {
    delete next.connectorType;
  }

  if (next.clearFill) {
    delete next.fillColor;
    delete next.fillTransparency;
  }

  if (next.lineVisible === false) {
    delete next.lineColor;
    delete next.lineWeight;
    delete next.lineTransparency;
  }

  return pruneDenseDefaultMutationValues(next);
}

function normalizeUpdateArgs(args: ManageSlideShapesArgs): ManageSlideShapesArgs {
  const next: ManageSlideShapesArgs = { ...args };

  if (Array.isArray(next.shapeIds) && next.shapeIds.length === 0) {
    delete next.shapeIds;
  }

  for (const key of blankableUpdateKeys) {
    if (isBlankString(next[key])) {
      delete next[key];
    }
  }

  for (const key of createOnlyUpdateKeys) {
    delete next[key];
  }

  if (next.shapeId !== undefined) {
    delete next.shapeIndex;
    delete next.placeholderType;
    delete next.shapeIds;
  } else if (next.shapeIndex !== undefined) {
    delete next.placeholderType;
    delete next.shapeIds;
  } else if (next.placeholderType !== undefined) {
    delete next.shapeIds;
  }

  return pruneDenseDefaultMutationValues(next);
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
  } else {
    if (typeof args.fillColor === "string" && args.fillColor.trim()) {
      shape.fill.setSolidColor(normalizeColor(args.fillColor));
    }
    if (typeof args.fillTransparency === "number") shape.fill.transparency = args.fillTransparency;
  }

  if (typeof args.lineVisible === "boolean") {
    shape.lineFormat.visible = args.lineVisible;
  }
  if (args.lineVisible !== false) {
    if (typeof args.lineColor === "string" && args.lineColor.trim()) {
      shape.lineFormat.color = normalizeColor(args.lineColor);
    }
    if (typeof args.lineWeight === "number") shape.lineFormat.weight = args.lineWeight;
    if (typeof args.lineTransparency === "number") shape.lineFormat.transparency = args.lineTransparency;
  }

  if (frame) {
    if (typeof args.text === "string") frame.textRange.text = args.text;

    const font = frame.textRange.font;
    const paragraph = frame.textRange.paragraphFormat;

    if (typeof args.fontName === "string" && args.fontName.trim()) font.name = args.fontName;
    if (typeof args.fontSize === "number") font.size = args.fontSize;
    if (typeof args.fontColor === "string" && args.fontColor.trim()) font.color = normalizeColor(args.fontColor);
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
  description: "Create, update, delete, group, or ungroup PowerPoint shapes for geometry, styling, naming, and structure changes. For wording or rich-text edits, prefer read_slide_text with edit_slide_text or edit_slide_xml. Updates are patch-like: pass a target plus only the properties you want to change, and omit unchanged/default values.",
  parameters: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: ["create", "update", "delete", "group", "ungroup"],
        description: "Shape operation to perform. Use update as a sparse patch: include only targeting fields and the properties that should change.",
      },
      slideIndex: { type: "number", description: "0-based slide index. Optional when the active slide can be inferred from the current selection." },
      shapeId: {
        anyOf: [{ type: "string" }, { type: "number" }],
        description: "Existing Office shape id, or an exported XML p:cNvPr id after an Open XML slide replacement. Preferred targeting field for update/delete/ungroup when available.",
      },
      shapeIndex: { type: "number", description: "Existing 0-based shape index on the slide. Use this only when shapeId is unavailable." },
      shapeIds: {
        type: "array",
        items: { anyOf: [{ type: "string" }, { type: "number" }] },
        description: "Array of shape ids to group together. Required for the group action. Must contain at least 2 shapes.",
      },
      placeholderType: { type: "string", description: "Optional placeholder type to target for update or delete, such as Title, Body, Subtitle, or Content. Use this instead of shapeId/shapeIndex when targeting a placeholder by role." },
      shapeType: { type: "string", enum: ["textBox", "geometricShape", "line"], description: "Shape type to create. Create-only; omit for update/delete/group/ungroup." },
      geometricShapeType: { type: "string", description: "Geometric shape type for create. Must be a valid PowerPoint.GeometricShapeType value such as Rectangle, Ellipse, Chevron, or RightArrow. Create-only." },
      connectorType: { type: "string", enum: ["Straight", "Elbow", "Curve"], description: "Optional connector type for line creation. Default Straight. Create-only." },
      text: { type: "string", description: "Text content to create or replace. For update, omit this unless you intend to replace the current text." },
      name: { type: "string", description: "Shape name to set. For update, omit to leave the current name unchanged." },
      left: { type: "number", description: "Left position in points. For update, omit to keep the current position." },
      top: { type: "number", description: "Top position in points. For update, omit to keep the current position." },
      width: { type: "number", description: "Width in points. For update, omit to keep the current width." },
      height: { type: "number", description: "Height in points. For update, omit to keep the current height." },
      rotation: { type: "number", description: "Rotation in degrees. For update, omit to keep the current rotation." },
      visible: { type: "boolean", description: "Whether the shape is visible. For update, omit to leave visibility unchanged." },
      altTextTitle: { type: "string", description: "Alt text title. For update, omit to leave unchanged." },
      altTextDescription: { type: "string", description: "Alt text description. For update, omit to leave unchanged." },
      fillColor: { type: "string", description: "Solid fill color. For update, omit to leave fill unchanged. Do not send empty strings as placeholders." },
      fillTransparency: { type: "number", description: "Fill transparency from 0 to 1. For update, omit to leave unchanged." },
      clearFill: { type: "boolean", description: "Clear the fill entirely. For update, include only when you intentionally want no fill." },
      lineColor: { type: "string", description: "Line color. For update, omit to leave unchanged. Do not send empty strings as placeholders." },
      lineWeight: { type: "number", description: "Line weight in points. For update, omit to leave unchanged." },
      lineTransparency: { type: "number", description: "Line transparency from 0 to 1. For update, omit to leave unchanged." },
      lineVisible: { type: "boolean", description: "Whether the line is visible. For update, omit to leave unchanged." },
      fontName: { type: "string", description: "Font family name. For update, omit to leave unchanged. Do not send empty strings as placeholders." },
      fontSize: { type: "number", description: "Font size in points. For update, omit to leave unchanged." },
      fontColor: { type: "string", description: "Font color. For update, omit to leave unchanged. Do not send empty strings as placeholders." },
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
      marginLeft: { type: "number", description: "Left text margin in points. For update, omit to leave unchanged." },
      marginRight: { type: "number", description: "Right text margin in points. For update, omit to leave unchanged." },
      marginTop: { type: "number", description: "Top text margin in points. For update, omit to leave unchanged." },
      marginBottom: { type: "number", description: "Bottom text margin in points. For update, omit to leave unchanged." },
    },
    required: ["action", "slideIndex"],
  },
  handler: async (args) => {
    const parsedArgs = manageSlideShapesArgsSchema.safeParse(args);
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }
    const resolvedArgs = resolvePowerPointTargetingArgs(parsedArgs.data as ManageSlideShapesArgs);
    const update = resolvedArgs.action === "create"
      ? normalizeCreateArgs(resolvedArgs)
      : resolvedArgs.action === "update"
        ? normalizeUpdateArgs(resolvedArgs)
        : resolvedArgs;

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
      validateRange("fontSize", update.fontSize, { min: 0 }),
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
      const errorText = error instanceof Error ? `${error.message} ${(error as { code?: string }).code || ""}` : String(error);
      if (update.action === "update" && /invalidargument/i.test(errorText)) {
        return toolFailure(
          error,
          "Hint: for update, pass only the target (`shapeId`, `shapeIndex`, or `placeholderType`) plus the specific properties to change. Omit unchanged/default fields like empty strings, placeholder text, create-only fields, or filler values from prior inspection output.",
        );
      }
      return toolFailure(error);
    }
  },
};
