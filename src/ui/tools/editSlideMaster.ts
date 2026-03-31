import type { Tool } from "./types";
import { loadTextFrames } from "./powerpointText";
import {
  isPowerPointRequirementSetSupported,
  loadThemeColors,
  normalizeHexColor,
  parseColor,
  readOfficeValue,
  toolFailure,
} from "./powerpointShared";
import { z } from "zod";

const themeColorKeys = [
  "Dark1",
  "Light1",
  "Dark2",
  "Light2",
  "Accent1",
  "Accent2",
  "Accent3",
  "Accent4",
  "Accent5",
  "Accent6",
  "Hyperlink",
  "FollowedHyperlink",
] as const;

const colorValueSchema = z.string().trim().regex(/^#?[0-9A-Fa-f]{6}$/, "Color must be a 6-digit hex value like #123ABC.");
const nonNegativeNumberSchema = z.number().finite().min(0);

const themeColorsSchema = z.object({
  Dark1: colorValueSchema.optional(),
  Light1: colorValueSchema.optional(),
  Dark2: colorValueSchema.optional(),
  Light2: colorValueSchema.optional(),
  Accent1: colorValueSchema.optional(),
  Accent2: colorValueSchema.optional(),
  Accent3: colorValueSchema.optional(),
  Accent4: colorValueSchema.optional(),
  Accent5: colorValueSchema.optional(),
  Accent6: colorValueSchema.optional(),
  Hyperlink: colorValueSchema.optional(),
  FollowedHyperlink: colorValueSchema.optional(),
}).strict();

const decorativeShapeCreateSchema = z.object({
  action: z.literal("create"),
  shapeType: z.enum(["textBox", "geometricShape"]),
  geometricShapeType: z.string().optional(),
  text: z.string().optional(),
  name: z.string().optional(),
  left: z.number().finite().optional(),
  top: z.number().finite().optional(),
  width: nonNegativeNumberSchema.optional(),
  height: nonNegativeNumberSchema.optional(),
  visible: z.boolean().optional(),
  fillColor: colorValueSchema.optional(),
  lineColor: colorValueSchema.optional(),
  lineWeight: nonNegativeNumberSchema.optional(),
  fontColor: colorValueSchema.optional(),
  fontSize: nonNegativeNumberSchema.optional(),
}).strict().superRefine((value, ctx) => {
  if (value.shapeType === "geometricShape" && !value.geometricShapeType) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "geometricShapeType is required when creating a geometric shape.",
      path: ["geometricShapeType"],
    });
  }

  if (value.shapeType === "textBox" && value.geometricShapeType !== undefined) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "geometricShapeType is only supported when shapeType is geometricShape.",
      path: ["geometricShapeType"],
    });
  }

  if (value.shapeType === "geometricShape" && (value.text !== undefined || value.fontColor !== undefined || value.fontSize !== undefined)) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "Text content and font styling are only supported when creating a textBox decorative shape.",
      path: ["shapeType"],
    });
  }
});

const decorativeShapeUpdateSchema = z.object({
  action: z.literal("update"),
  shapeId: z.string().trim().min(1, "shapeId is required for update and delete."),
  text: z.string().optional(),
  name: z.string().optional(),
  left: z.number().finite().optional(),
  top: z.number().finite().optional(),
  width: nonNegativeNumberSchema.optional(),
  height: nonNegativeNumberSchema.optional(),
  visible: z.boolean().optional(),
  fillColor: colorValueSchema.optional(),
  lineColor: colorValueSchema.optional(),
  lineWeight: nonNegativeNumberSchema.optional(),
  fontColor: colorValueSchema.optional(),
  fontSize: nonNegativeNumberSchema.optional(),
}).strict().superRefine((value, ctx) => {
  const hasMutation = [
    value.text,
    value.name,
    value.left,
    value.top,
    value.width,
    value.height,
    value.visible,
    value.fillColor,
    value.lineColor,
    value.lineWeight,
    value.fontColor,
    value.fontSize,
  ].some((entry) => entry !== undefined);

  if (!hasMutation) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "update requires at least one supported mutation field.",
      path: [],
    });
  }
});

const decorativeShapeDeleteSchema = z.object({
  action: z.literal("delete"),
  shapeId: z.string().trim().min(1, "shapeId is required for update and delete."),
}).strict();

const decorativeShapeMutationSchema = z.discriminatedUnion("action", [
  decorativeShapeCreateSchema,
  decorativeShapeUpdateSchema,
  decorativeShapeDeleteSchema,
]);

const editSlideMasterArgsSchema = z.object({
  slideMasterId: z.string().optional(),
  themeColors: themeColorsSchema.optional(),
  decorativeShapes: z.array(decorativeShapeMutationSchema).min(1).optional(),
}).strict().superRefine((value, ctx) => {
  if (Object.keys(value.themeColors || {}).length === 0 && (value.decorativeShapes?.length || 0) === 0) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "Provide themeColors or decorativeShapes.",
      path: [],
    });
  }

  const targetIds = value.decorativeShapes
    ?.filter((shape): shape is z.infer<typeof decorativeShapeUpdateSchema> | z.infer<typeof decorativeShapeDeleteSchema> => shape.action !== "create")
    .map((shape) => shape.shapeId) || [];
  const duplicates = [...new Set(targetIds.filter((shapeId, index) => targetIds.indexOf(shapeId) !== index))];

  if (duplicates.length > 0) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: `Each existing decorative shape may be targeted at most once per request. Duplicate shapeIds: ${duplicates.join(", ")}.`,
      path: ["decorativeShapes"],
    });
  }
});

type ThemeColorKey = (typeof themeColorKeys)[number];
type EditSlideMasterArgs = z.infer<typeof editSlideMasterArgsSchema>;
type DecorativeShapeMutation = z.infer<typeof decorativeShapeMutationSchema>;
type DecorativeShapeCreate = z.infer<typeof decorativeShapeCreateSchema>;
type DecorativeShapeUpdate = z.infer<typeof decorativeShapeUpdateSchema>;

interface ThemeColorChange {
  color: ThemeColorKey;
  previous: string;
  next: string;
}

interface DecorativeShapeResult {
  action: "create" | "update" | "delete";
  shapeId: string;
  name: string;
  shapeType: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  visible?: boolean;
}

function buildMasterNotFoundMessage(slideMasterId: string, masters: PowerPoint.SlideMaster[]) {
  const available = masters
    .map((master) => `${readOfficeValue(() => master.id, "(missing)")}:${JSON.stringify(readOfficeValue(() => master.name, ""))}`)
    .join(", ");
  return `Slide master ${JSON.stringify(slideMasterId)} was not found. Available slide masters: ${available || "(none)"}.`;
}

function buildShapeNotFoundMessage(master: PowerPoint.SlideMaster, shapeId: string, shapes: PowerPoint.Shape[]) {
  const available = shapes.map((shape) => readOfficeValue(() => shape.id, "(missing)")).join(", ");
  return `Master shape ${JSON.stringify(shapeId)} was not found on slide master ${JSON.stringify(readOfficeValue(() => master.name, ""))} (${readOfficeValue(() => master.id, "(missing)")}). Available shape ids: ${available || "(none)"}.`;
}

function buildTextUnsupportedMessage(master: PowerPoint.SlideMaster, shapeId: string) {
  return `Master shape ${JSON.stringify(shapeId)} on slide master ${JSON.stringify(readOfficeValue(() => master.name, ""))} (${readOfficeValue(() => master.id, "(missing)")}) does not support text.`;
}

function buildCreateShapeOptions(shape: DecorativeShapeCreate) {
  return {
    ...(typeof shape.left === "number" ? { left: shape.left } : {}),
    ...(typeof shape.top === "number" ? { top: shape.top } : {}),
    ...(typeof shape.width === "number" ? { width: shape.width } : {}),
    ...(typeof shape.height === "number" ? { height: shape.height } : {}),
  } satisfies PowerPoint.ShapeAddOptions;
}

function hasTextMutation(shape: DecorativeShapeMutation) {
  return shape.action !== "delete" && (shape.text !== undefined || shape.fontColor !== undefined || shape.fontSize !== undefined);
}

function validateGeometricShapeType(value: string) {
  const enumValues = (globalThis.PowerPoint && PowerPoint.GeometricShapeType)
    ? Object.values(PowerPoint.GeometricShapeType as Record<string, string>)
    : [];

  if (enumValues.length === 0 || enumValues.includes(value)) {
    return null;
  }

  return `geometricShapeType must be a valid PowerPoint.GeometricShapeType value, such as ${enumValues.slice(0, 8).join(", ")}.`;
}

function applyShapeVisualMutation(shape: PowerPoint.Shape, mutation: DecorativeShapeCreate | DecorativeShapeUpdate) {
  if (typeof mutation.name === "string") shape.name = mutation.name;
  if (typeof mutation.left === "number") shape.left = mutation.left;
  if (typeof mutation.top === "number") shape.top = mutation.top;
  if (typeof mutation.width === "number") shape.width = mutation.width;
  if (typeof mutation.height === "number") shape.height = mutation.height;
  if (typeof mutation.visible === "boolean") shape.visible = mutation.visible;
  if (typeof mutation.fillColor === "string") shape.fill.setSolidColor(normalizeHexColor(mutation.fillColor));
  if (typeof mutation.lineColor === "string") shape.lineFormat.color = normalizeHexColor(mutation.lineColor);
  if (typeof mutation.lineWeight === "number") shape.lineFormat.weight = mutation.lineWeight;
}

async function applyCreatedTextBoxMutation(
  context: PowerPoint.RequestContext,
  shape: PowerPoint.Shape,
  mutation: DecorativeShapeCreate,
) {
  if (mutation.shapeType !== "textBox" || !hasTextMutation(mutation)) {
    return null;
  }

  const [frame] = await loadTextFrames(context, [shape]);
  if (!frame || frame.isNullObject) {
    return "Created text box does not support text.";
  }

  if (typeof mutation.text === "string") frame.textRange.text = mutation.text;
  if (typeof mutation.fontColor === "string") frame.textRange.font.color = normalizeHexColor(mutation.fontColor);
  if (typeof mutation.fontSize === "number") frame.textRange.font.size = mutation.fontSize;
  return null;
}

function applyTextMutation(frame: PowerPoint.TextFrame, mutation: DecorativeShapeUpdate) {
  if (typeof mutation.text === "string") frame.textRange.text = mutation.text;
  if (typeof mutation.fontColor === "string") frame.textRange.font.color = normalizeHexColor(mutation.fontColor);
  if (typeof mutation.fontSize === "number") frame.textRange.font.size = mutation.fontSize;
}

function toDecorativeShapeResult(shape: PowerPoint.Shape, action: DecorativeShapeResult["action"]): DecorativeShapeResult {
  return {
    action,
    shapeId: readOfficeValue(() => shape.id, "(missing)"),
    name: readOfficeValue(() => shape.name, "") || "",
    shapeType: readOfficeValue(() => String(shape.type), "Unknown"),
    left: readOfficeValue(() => shape.left, undefined),
    top: readOfficeValue(() => shape.top, undefined),
    width: readOfficeValue(() => shape.width, undefined),
    height: readOfficeValue(() => shape.height, undefined),
    visible: readOfficeValue(() => shape.visible, undefined),
  };
}

function buildSummary(masterName: string, masterId: string, themeColorChanges: ThemeColorChange[], decorativeShapeResults: DecorativeShapeResult[]) {
  const parts = [`Updated slide master ${JSON.stringify(masterName)} (${masterId}).`];
  if (themeColorChanges.length > 0) {
    parts.push(`Changed ${themeColorChanges.length} theme color${themeColorChanges.length === 1 ? "" : "s"}.`);
  }
  if (decorativeShapeResults.length > 0) {
    parts.push(`Applied ${decorativeShapeResults.length} decorative shape operation${decorativeShapeResults.length === 1 ? "" : "s"}.`);
  }
  if (themeColorChanges.length === 0 && decorativeShapeResults.length === 0) {
    parts.push("The request completed, but no effective changes were detected.");
  }
  return parts.join(" ");
}

export const editSlideMaster: Tool = {
  name: "edit_slide_master",
  description: "Edit the supported PowerPoint slide master surface: theme colors and a small explicit set of decorative master shapes.",
  parameters: {
    type: "object",
    properties: {
      slideMasterId: { type: "string", description: "Optional slide master id. Defaults to the first slide master." },
      themeColors: {
        type: "object",
        description: "Optional theme color patch for the chosen slide master.",
        properties: Object.fromEntries(themeColorKeys.map((key) => [key, { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" }])) as Record<string, { type: "string"; pattern: string }>,
        additionalProperties: false,
      },
      decorativeShapes: {
        type: "array",
        description: "Optional explicit decorative master shape operations. Supported shape types: textBox and geometricShape.",
        items: {
          oneOf: [
            {
              type: "object",
              properties: {
                action: { type: "string", enum: ["create"] },
                shapeType: { type: "string", enum: ["textBox", "geometricShape"] },
                geometricShapeType: { type: "string" },
                text: { type: "string" },
                name: { type: "string" },
                left: { type: "number" },
                top: { type: "number" },
                width: { type: "number" },
                height: { type: "number" },
                visible: { type: "boolean" },
                fillColor: { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" },
                lineColor: { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" },
                lineWeight: { type: "number" },
                fontColor: { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" },
                fontSize: { type: "number" },
              },
              required: ["action", "shapeType"],
              additionalProperties: false,
            },
            {
              type: "object",
              properties: {
                action: { type: "string", enum: ["update"] },
                shapeId: { type: "string" },
                text: { type: "string" },
                name: { type: "string" },
                left: { type: "number" },
                top: { type: "number" },
                width: { type: "number" },
                height: { type: "number" },
                visible: { type: "boolean" },
                fillColor: { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" },
                lineColor: { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" },
                lineWeight: { type: "number" },
                fontColor: { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" },
                fontSize: { type: "number" },
              },
              required: ["action", "shapeId"],
              additionalProperties: false,
            },
            {
              type: "object",
              properties: {
                action: { type: "string", enum: ["delete"] },
                shapeId: { type: "string" },
              },
              required: ["action", "shapeId"],
              additionalProperties: false,
            },
          ],
        },
      },
    },
  },
  handler: async (args) => {
    const parsedArgs = editSlideMasterArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }

    const update = parsedArgs.data as EditSlideMasterArgs;

    if (update.themeColors && Object.keys(update.themeColors).length > 0 && !isPowerPointRequirementSetSupported("1.10")) {
      return toolFailure("Editing slide master theme colors requires PowerPointApi 1.10.");
    }

    if (update.decorativeShapes?.some((shape) => shape.action === "create") && !isPowerPointRequirementSetSupported("1.4")) {
      return toolFailure("Creating decorative master shapes requires PowerPointApi 1.4.");
    }

    if (update.decorativeShapes?.some((shape) => shape.action === "update" || shape.action === "delete") && !isPowerPointRequirementSetSupported("1.3")) {
      return toolFailure("Updating or deleting decorative master shapes requires PowerPointApi 1.3.");
    }

    for (const mutation of update.decorativeShapes || []) {
      if (mutation.action === "create" && mutation.shapeType === "geometricShape" && mutation.geometricShapeType) {
        const geometricShapeTypeError = validateGeometricShapeType(mutation.geometricShapeType);
        if (geometricShapeTypeError) {
          return toolFailure(geometricShapeTypeError);
        }
      }
    }

    try {
      return await PowerPoint.run(async (context) => {
        const slideMasters = context.presentation.slideMasters;
        slideMasters.load("items");
        await context.sync();

        if (slideMasters.items.length === 0) {
          return toolFailure("Presentation has no slide masters.");
        }

        for (const master of slideMasters.items) {
          master.load(["id", "name"]);
        }
        await context.sync();

        const master = update.slideMasterId
          ? slideMasters.items.find((item) => readOfficeValue(() => item.id, "") === update.slideMasterId) || null
          : slideMasters.items[0] || null;

        if (!master) {
          return toolFailure(buildMasterNotFoundMessage(update.slideMasterId as string, slideMasters.items));
        }

        const masterId = readOfficeValue(() => master.id, "(missing)");
        const masterName = readOfficeValue(() => master.name, "");
        const themeColorChanges: ThemeColorChange[] = [];
        const decorativeShapeResults: DecorativeShapeResult[] = [];
        const updateTargetShapes = new Map<string, PowerPoint.Shape>();
        const updateTextFrames = new Map<string, PowerPoint.TextFrame>();

        if (update.decorativeShapes?.length) {
          master.shapes.load("items");
          await context.sync();
          for (const shape of master.shapes.items) {
            shape.load(["id", "name", "type", "left", "top", "width", "height", "visible"]);
          }
          await context.sync();

          for (const mutation of update.decorativeShapes) {
            if (mutation.action === "create") continue;
            const target = master.shapes.items.find((shape) => readOfficeValue(() => shape.id, "") === mutation.shapeId) || null;
            if (!target) {
              return toolFailure(buildShapeNotFoundMessage(master, mutation.shapeId, master.shapes.items));
            }
            updateTargetShapes.set(mutation.shapeId, target);
          }

          const textMutations = update.decorativeShapes.filter((mutation): mutation is DecorativeShapeUpdate => mutation.action === "update" && hasTextMutation(mutation));
          if (textMutations.length > 0) {
            const textTargets = textMutations.map((mutation) => updateTargetShapes.get(mutation.shapeId)).filter(Boolean) as PowerPoint.Shape[];
            const frames = await loadTextFrames(context, textTargets);
            textTargets.forEach((shape, index) => {
              const shapeId = readOfficeValue(() => shape.id, "");
              const frame = frames[index];
              if (frame && !frame.isNullObject) {
                updateTextFrames.set(shapeId, frame);
              }
            });

            for (const mutation of textMutations) {
              if (!updateTextFrames.has(mutation.shapeId)) {
                return toolFailure(buildTextUnsupportedMessage(master, mutation.shapeId));
              }
            }
          }
        }

        for (const mutation of update.decorativeShapes || []) {
          if (mutation.action === "create") {
            const created = mutation.shapeType === "textBox"
              ? master.shapes.addTextBox(mutation.text || "", buildCreateShapeOptions(mutation))
              : master.shapes.addGeometricShape(mutation.geometricShapeType as PowerPoint.GeometricShapeType, buildCreateShapeOptions(mutation));

            created.load(["id", "name", "type", "left", "top", "width", "height", "visible"]);
            await context.sync();

            applyShapeVisualMutation(created, mutation);
            const textMutationError = await applyCreatedTextBoxMutation(context, created, mutation);
            if (textMutationError) {
              created.delete();
              await context.sync();
              return toolFailure(textMutationError);
            }

            await context.sync();
            decorativeShapeResults.push(toDecorativeShapeResult(created, "create"));
            continue;
          }

          const target = updateTargetShapes.get(mutation.shapeId) || null;
          if (!target) {
            return toolFailure(buildShapeNotFoundMessage(master, mutation.shapeId, master.shapes.items));
          }

          if (mutation.action === "delete") {
            const summary = toDecorativeShapeResult(target, "delete");
            target.delete();
            await context.sync();
            decorativeShapeResults.push(summary);
            continue;
          }

          applyShapeVisualMutation(target, mutation);
          const frame = updateTextFrames.get(mutation.shapeId);
          if (frame) {
            applyTextMutation(frame, mutation);
          }

          await context.sync();
          decorativeShapeResults.push(toDecorativeShapeResult(target, "update"));
        }

        if (update.themeColors && Object.keys(update.themeColors).length > 0) {
          const before = await loadThemeColors(context, master);

          for (const key of themeColorKeys) {
            const nextValue = update.themeColors[key];
            if (typeof nextValue === "string") {
              master.themeColorScheme.setThemeColor(key, normalizeHexColor(nextValue));
            }
          }
          await context.sync();

          const after = await loadThemeColors(context, master);
          for (const key of themeColorKeys) {
            if (update.themeColors[key] === undefined) continue;
            const previous = parseColor(before[key]);
            const next = parseColor(after[key]);
            if (previous !== next) {
              themeColorChanges.push({ color: key, previous, next });
            }
          }
        }

        return {
          resultType: "success",
          textResultForLlm: buildSummary(masterName, masterId, themeColorChanges, decorativeShapeResults),
          slideMasterId: masterId,
          slideMasterName: masterName,
          changedThemeColors: themeColorChanges,
          decorativeShapeResults,
          toolTelemetry: {
            slideMasterId: masterId,
            themeColorChangeCount: themeColorChanges.length,
            decorativeShapeChangeCount: decorativeShapeResults.length,
          },
        };
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
