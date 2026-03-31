import type { Tool } from "./types";
import { createImageRectangle, fetchImageUrlAsBase64, getShapeBounds, toPowerPointTableValues } from "./powerpointNativeContent";
import { isPowerPointRequirementSetSupported, readOfficeValue, toolFailure } from "./powerpointShared";
import { loadTextFrames } from "./powerpointText";
import { z } from "zod";

const tableCellSchema = z.union([z.boolean(), z.number(), z.string()]);

const bindingTargetFields = {
  placeholderType: z.string().optional(),
  placeholderName: z.string().optional(),
};

const tableDataSchema = z.array(z.array(tableCellSchema).min(1)).min(1).refine(
  (rows) => rows.every((row) => row.length === rows[0]?.length),
  { message: "tableData must be a non-empty rectangular 2D array." },
);

const textBindingSchema = z.object({
  ...bindingTargetFields,
  text: z.string(),
}).strict().refine((value) => value.placeholderType !== undefined || value.placeholderName !== undefined, {
  message: "Each binding must include placeholderType or placeholderName.",
});

const imageBindingSchema = z.object({
  ...bindingTargetFields,
  imageUrl: z.string(),
}).strict().refine((value) => value.placeholderType !== undefined || value.placeholderName !== undefined, {
  message: "Each binding must include placeholderType or placeholderName.",
});

const tableBindingSchema = z.object({
  ...bindingTargetFields,
  tableData: tableDataSchema,
}).strict().refine((value) => value.placeholderType !== undefined || value.placeholderName !== undefined, {
  message: "Each binding must include placeholderType or placeholderName.",
});

const slideBindingSchema = z.union([textBindingSchema, imageBindingSchema, tableBindingSchema]);

const createSlideFromLayoutArgsSchema = z.object({
  layoutId: z.string(),
  slideMasterId: z.string().optional(),
  targetIndex: z.number().optional(),
  bindings: z.array(slideBindingSchema).optional(),
});

type CreateSlideFromLayoutArgs = z.infer<typeof createSlideFromLayoutArgsSchema>;
type SlideBinding = z.infer<typeof slideBindingSchema>;

interface PlaceholderTarget {
  shape: PowerPoint.Shape;
  shapeId: string;
  placeholderName: string;
  placeholderType?: string;
  placeholderContainedType?: string | null;
}

type TextSlideBinding = Extract<SlideBinding, { text: string }>;
type ImageSlideBinding = Extract<SlideBinding, { imageUrl: string }>;
type TableSlideBinding = Extract<SlideBinding, { tableData: Array<Array<boolean | number | string>> }>;

type ResolvedBindingPlan =
  | { binding: TextSlideBinding; placeholder: PlaceholderTarget }
  | { binding: ImageSlideBinding; placeholder: PlaceholderTarget; imageBase64: string }
  | { binding: TableSlideBinding; placeholder: PlaceholderTarget };

function isNonNegativeInteger(value: unknown): value is number {
  return Number.isInteger(value) && (value as number) >= 0;
}

function describeAvailablePlaceholders(placeholders: PlaceholderTarget[]) {
  if (placeholders.length === 0) return "The created slide has no detectable placeholder shapes.";
  return `Available placeholders: ${placeholders.map((placeholder) => `${placeholder.shapeId}/${placeholder.placeholderType || "(unknown)"}:${JSON.stringify(placeholder.placeholderName)}`).join(", ")}.`;
}

async function loadPlaceholderTargets(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
): Promise<PlaceholderTarget[]> {
  slide.shapes.load("items/id,name,type");
  await context.sync();

  const placeholderShapes = slide.shapes.items.filter((shape) => shape.type === PowerPoint.ShapeType.placeholder);
  for (const shape of placeholderShapes) {
    shape.placeholderFormat.load(["type", "containedType"]);
  }
  await context.sync();

  return placeholderShapes.map((shape) => ({
    shape,
    shapeId: readOfficeValue(() => shape.id, "(missing)"),
    placeholderName: readOfficeValue(() => shape.name, ""),
    placeholderType: readOfficeValue(() => (shape.placeholderFormat.type ? String(shape.placeholderFormat.type) : undefined), undefined),
    placeholderContainedType: readOfficeValue(
      () => (shape.placeholderFormat.containedType ? String(shape.placeholderFormat.containedType) : null),
      undefined,
    ),
  }));
}

function resolvePlaceholderTargetIndex(binding: SlideBinding, placeholders: PlaceholderTarget[]) {
  return placeholders.findIndex((placeholder) => {
    if (binding.placeholderName !== undefined && placeholder.placeholderName !== binding.placeholderName) return false;
    if (binding.placeholderType !== undefined && placeholder.placeholderType !== binding.placeholderType) return false;
    return true;
  });
}

async function planBindings(
  context: PowerPoint.RequestContext,
  bindings: SlideBinding[],
  placeholders: PlaceholderTarget[],
): Promise<ResolvedBindingPlan[]> {
  const availablePlaceholders = [...placeholders];
  const plans: ResolvedBindingPlan[] = [];

  for (const binding of bindings) {
    const placeholderIndex = resolvePlaceholderTargetIndex(binding, availablePlaceholders);
    if (placeholderIndex < 0) {
      throw new Error(`Could not find a matching placeholder for binding. ${describeAvailablePlaceholders(placeholders)}`);
    }

    const [placeholder] = availablePlaceholders.splice(placeholderIndex, 1);
    if (!placeholder) {
      throw new Error(`Could not reserve a placeholder for binding. ${describeAvailablePlaceholders(placeholders)}`);
    }

    if ("text" in binding) {
      const [frame] = await loadTextFrames(context, [placeholder.shape]);
      if (!frame || frame.isNullObject) {
        throw new Error(`Placeholder ${JSON.stringify(placeholder.placeholderName)} does not support text.`);
      }
      plans.push({ binding, placeholder });
      continue;
    }

    if ("imageUrl" in binding) {
      const imageBase64 = await fetchImageUrlAsBase64(binding.imageUrl);
      plans.push({ binding, placeholder, imageBase64 });
      continue;
    }

    plans.push({ binding, placeholder });
  }

  return plans;
}

async function applyBinding(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  plan: ResolvedBindingPlan,
) {
  const { binding, placeholder } = plan;

  if ("text" in binding) {
    const [frame] = await loadTextFrames(context, [placeholder.shape]);
    if (!frame || frame.isNullObject) {
      throw new Error(`Placeholder ${JSON.stringify(placeholder.placeholderName)} does not support text.`);
    }
    frame.textRange.text = binding.text;
    await context.sync();
    return {
      bindingType: "text",
      shapeId: placeholder.shapeId,
      placeholderName: placeholder.placeholderName,
      placeholderType: placeholder.placeholderType,
      text: binding.text,
    };
  }

  if ("imageUrl" in binding) {
    if (!("imageBase64" in plan)) {
      throw new Error(`Image binding for ${JSON.stringify(placeholder.placeholderName)} is missing prepared image data.`);
    }
    const bounds = await getShapeBounds(placeholder.shape, context);
    placeholder.shape.delete();
    const created = createImageRectangle(slide, {
      left: bounds.left,
      top: bounds.top,
      width: bounds.width,
      height: bounds.height,
      name: bounds.name || placeholder.placeholderName || "Image",
      imageBase64: plan.imageBase64,
    });
    created.load(["id", "name"]);
    await context.sync();
    return {
      bindingType: "image",
      shapeId: readOfficeValue(() => created.id, "(missing)"),
      replacedShapeId: placeholder.shapeId,
      placeholderName: placeholder.placeholderName,
      placeholderType: placeholder.placeholderType,
      imageUrl: binding.imageUrl,
    };
  }

  const values = toPowerPointTableValues(binding.tableData);
  const bounds = await getShapeBounds(placeholder.shape, context);
  placeholder.shape.delete();
  const created = slide.shapes.addTable(values.length, values[0].length, {
    values,
    left: bounds.left,
    top: bounds.top,
    width: bounds.width,
    height: bounds.height,
  });
  if (bounds.name) created.name = bounds.name;
  created.load(["id", "name"]);
  await context.sync();
  return {
    bindingType: "table",
    shapeId: readOfficeValue(() => created.id, "(missing)"),
    replacedShapeId: placeholder.shapeId,
    placeholderName: placeholder.placeholderName,
    placeholderType: placeholder.placeholderType,
    rowCount: values.length,
    columnCount: values[0]?.length ?? 0,
  };
}

export const createSlideFromLayout: Tool = {
  name: "create_slide_from_layout",
  description: "Create a new slide from an existing layout and optionally bind text, images, or tables into layout placeholders.",
  parameters: {
    type: "object",
    properties: {
      layoutId: {
        type: "string",
        description: "Layout id to create the slide from.",
      },
      slideMasterId: {
        type: "string",
        description: "Optional slide master id when the layout id is not unique across masters.",
      },
      targetIndex: {
        type: "number",
        description: "Optional 0-based insertion index for the new slide. Defaults to the end of the deck.",
      },
      bindings: {
        type: "array",
        description: "Optional placeholder bindings. Each binding targets a placeholder by placeholderType or placeholderName and sets text, imageUrl, or tableData. Image and table bindings replace the matched placeholder using its bounds.",
        items: {
          anyOf: [
            {
              type: "object",
              properties: {
                placeholderType: { type: "string" },
                placeholderName: { type: "string" },
                text: { type: "string" },
              },
              required: ["text"],
            },
            {
              type: "object",
              properties: {
                placeholderType: { type: "string" },
                placeholderName: { type: "string" },
                imageUrl: { type: "string" },
              },
              required: ["imageUrl"],
            },
            {
              type: "object",
              properties: {
                placeholderType: { type: "string" },
                placeholderName: { type: "string" },
                tableData: {
                  type: "array",
                  items: { type: "array", items: { anyOf: [{ type: "string" }, { type: "number" }, { type: "boolean" }] } },
                },
              },
              required: ["tableData"],
            },
          ],
        },
      },
    },
    required: ["layoutId"],
  },
  handler: async (args) => {
    const parsedArgs = createSlideFromLayoutArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }

    const { layoutId, slideMasterId, targetIndex, bindings = [] } = parsedArgs.data as CreateSlideFromLayoutArgs;

    if (targetIndex !== undefined && !isNonNegativeInteger(targetIndex)) {
      return toolFailure("targetIndex must be a non-negative integer.");
    }
    if (!isPowerPointRequirementSetSupported("1.3")) {
      return toolFailure("Creating slides requires PowerPointApi 1.3.");
    }
    if (targetIndex !== undefined && !isPowerPointRequirementSetSupported("1.8")) {
      return toolFailure("Creating a slide at a specific position requires PowerPointApi 1.8.");
    }
    if (bindings.length > 0 && !isPowerPointRequirementSetSupported("1.8")) {
      return toolFailure("Placeholder bindings require PowerPointApi 1.8.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items/id");
        await context.sync();

        const slideCount = slides.items.length;
        if (targetIndex !== undefined && targetIndex > slideCount) {
          return toolFailure(`Invalid targetIndex ${targetIndex}. Must be 0-${slideCount}.`);
        }

        context.presentation.slides.add({
          layoutId,
          ...(slideMasterId ? { slideMasterId } : {}),
        });
        await context.sync();

        slides.load("items/id");
        await context.sync();

        const createdIndex = slides.items.length - 1;
        const createdSlide = slides.items[createdIndex];
        createdSlide.load("id");
        await context.sync();

        if (targetIndex !== undefined && targetIndex !== createdIndex) {
          createdSlide.moveTo(targetIndex);
          await context.sync();
        }

        const finalSlideIndex = targetIndex ?? createdIndex;
        const placeholders = bindings.length > 0 ? await loadPlaceholderTargets(context, createdSlide) : [];
        let plannedBindings: ResolvedBindingPlan[] = [];
        try {
          plannedBindings = await planBindings(context, bindings, placeholders);
        } catch (error: unknown) {
          createdSlide.delete();
          await context.sync();
          return toolFailure(error instanceof Error ? `${error.message} Created slide was rolled back.` : error);
        }

        const appliedBindings = [];
        try {
          for (const plan of plannedBindings) {
            appliedBindings.push(await applyBinding(context, createdSlide, plan));
          }
        } catch (error: unknown) {
          createdSlide.delete();
          await context.sync();
          return toolFailure(error instanceof Error ? `${error.message} Created slide was rolled back.` : error);
        }

        return {
          resultType: "success",
          textResultForLlm: `Created slide ${finalSlideIndex + 1} from layout ${JSON.stringify(layoutId)}${appliedBindings.length ? ` and applied ${appliedBindings.length} binding(s).` : "."}`,
          slideId: readOfficeValue(() => createdSlide.id, "(missing)"),
          slideIndex: finalSlideIndex,
          layoutId,
          slideMasterId,
          appliedBindings,
          toolTelemetry: {
            slideIndex: finalSlideIndex,
            layoutId,
            bindingCount: appliedBindings.length,
          },
        };
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
