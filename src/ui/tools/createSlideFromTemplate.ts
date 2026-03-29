import type { Tool } from "./types";
import { createImageRectangle, fetchImageUrlAsBase64 } from "./powerpointNativeContent";
import { loadShapeSummaries, toolFailure } from "./powerpointShared";

interface TemplateBinding {
  placeholderType?: string;
  placeholderName?: string;
  text?: string;
  imageUrl?: string;
  tableData?: string[][];
}

interface CreateSlideFromTemplateArgs {
  layoutId: string;
  bindings?: TemplateBinding[];
  slideMasterId?: string;
  targetIndex?: number;
}

export async function createSlideWithLayout(
  context: PowerPoint.RequestContext,
  args: CreateSlideFromTemplateArgs,
) {
  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();
  const beforeCount = slides.items.length;

  context.presentation.slides.add({
    ...(args.slideMasterId ? { slideMasterId: args.slideMasterId } : {}),
    layoutId: args.layoutId,
  });
  await context.sync();

  slides.load("items/id");
  await context.sync();
  const createdSlide = slides.items[slides.items.length - 1];
  if (args.targetIndex !== undefined && args.targetIndex !== slides.items.length - 1) {
    createdSlide.moveTo(args.targetIndex);
    await context.sync();
  }

  slides.load("items/id");
  await context.sync();
  const finalSlide = slides.items.find((slide) => !slides.items.slice(0, beforeCount).some((existing) => existing.id === slide.id)) || createdSlide;
  const finalIndex = slides.items.findIndex((slide) => slide.id === finalSlide.id);
  return { slide: finalSlide, slideIndex: finalIndex };
}

async function applyTemplateBindingsToSlide(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  bindings: TemplateBinding[],
) {
  slide.shapes.load("items");
  await context.sync();
  const summaries = await loadShapeSummaries(context, slide.shapes.items, { includeText: true, includeFormatting: false, includeTableValues: false });
  const usedIds = new Set<string>();
  const applied: Array<{ binding: TemplateBinding; shapeId?: string; action: string }> = [];

  for (const binding of bindings) {
    const match = summaries.find((shape) => {
      if (usedIds.has(shape.id)) return false;
      if (binding.placeholderName && shape.name !== binding.placeholderName) return false;
      if (binding.placeholderType && shape.placeholderType !== binding.placeholderType) return false;
      return Boolean(binding.placeholderName || binding.placeholderType);
    }) || summaries.find((shape) => !usedIds.has(shape.id) && shape.placeholderType && !binding.placeholderName);

    if (!match) {
      applied.push({ binding, action: "skipped" });
      continue;
    }

    usedIds.add(match.id);
    const target = slide.shapes.items[match.index];

    if (binding.text !== undefined) {
      const frame = target.getTextFrameOrNullObject();
      frame.load("isNullObject");
      await context.sync();
      if (!frame.isNullObject) {
        frame.textRange.text = binding.text;
        applied.push({ binding, shapeId: match.id, action: "text" });
      }
    }

    if (binding.imageUrl) {
      const imageBase64 = await fetchImageUrlAsBase64(binding.imageUrl);
      createImageRectangle(slide, {
        left: match.left,
        top: match.top,
        width: match.width,
        height: match.height,
        name: match.name,
        imageBase64,
      });
      applied.push({ binding, shapeId: match.id, action: "image" });
    }

    if (binding.tableData?.length) {
      slide.shapes.addTable(binding.tableData.length, binding.tableData[0].length, {
        values: binding.tableData,
        left: match.left,
        top: match.top,
        width: match.width,
        height: match.height,
      });
      applied.push({ binding, shapeId: match.id, action: "table" });
    }
  }

  await context.sync();
  return applied;
}

export const createSlideFromTemplate: Tool = {
  name: "create_slide_from_template",
  description: "Create a PowerPoint slide from a chosen layout and bind text, image, or table content into placeholders.",
  parameters: {
    type: "object",
    properties: {
      slideMasterId: { type: "string", description: "Optional slide master id for creation." },
      layoutId: { type: "string", description: "Required PowerPoint layout id to create the slide from." },
      targetIndex: { type: "number", description: "Optional insertion index for the new slide." },
      bindings: {
        type: "array",
        items: {
          type: "object",
          properties: {
            placeholderType: { type: "string" },
            placeholderName: { type: "string" },
            text: { type: "string" },
            imageUrl: { type: "string" },
            tableData: { type: "array", items: { type: "array", items: { type: "string" } } },
          },
        },
      },
    },
    required: ["layoutId"],
  },
  handler: async (args) => {
    const templateArgs = args as CreateSlideFromTemplateArgs;
    if (!templateArgs.layoutId) {
      return toolFailure("layoutId is required.");
    }
    if (templateArgs.targetIndex !== undefined && (!Number.isInteger(templateArgs.targetIndex) || templateArgs.targetIndex < 0)) {
      return toolFailure("targetIndex must be a non-negative integer.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const { slide, slideIndex } = await createSlideWithLayout(context, templateArgs);
        const appliedBindings = await applyTemplateBindingsToSlide(context, slide, templateArgs.bindings || []);
        slide.load("id");
        await context.sync();
        return {
          resultType: "success",
          textResultForLlm: `Created a slide from layout ${templateArgs.layoutId} at position ${slideIndex + 1}.`,
          slideIndex,
          slideId: slide.id,
          appliedBindings,
          toolTelemetry: {
            slideIndex,
            slideId: slide.id,
            layoutId: templateArgs.layoutId,
            appliedBindingCount: appliedBindings.length,
          },
        };
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};

export { applyTemplateBindingsToSlide };
