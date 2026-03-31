import type { Tool } from "./types";
import { inspectSlideXmlFromOfficeSlide } from "./powerpointSlideXml";
import { resolvePowerPointSlideIndexes } from "./powerpointContext";
import { enrichShapeSummariesWithRefs } from "./powerpointShapeRefs";
import { getSlideByIndex } from "./powerpointNativeContent";
import {
  loadShapeSummaries,
  roundTripSlideRefreshHint,
  shouldAddRoundTripRefreshHint,
  summarizePlainText,
  toolFailure,
} from "./powerpointShared";
import { z } from "zod";

const listSlideShapesArgsSchema = z.object({
  slideIndex: z.number().optional(),
  detail: z.boolean().optional(),
});

async function resolveSlideIndex(
  context: PowerPoint.RequestContext,
  requestedSlideIndex: number | undefined,
): Promise<number> {
  const resolved = resolvePowerPointSlideIndexes(requestedSlideIndex);
  if (typeof resolved === "number") {
    return resolved;
  }

  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();
  if (slides.items.length === 1) {
    return 0;
  }

  throw new Error("slideIndex is required when no active slide can be inferred from the current PowerPoint context.");
}

export const listSlideShapes: Tool = {
  name: "list_slide_shapes",
  description: "Inspect one PowerPoint slide and return stable shape refs based on slide id plus XML cNvPr shape ids.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        type: "number",
        description: "Optional 0-based slide index. If omitted, the active slide is used when it can be inferred safely.",
      },
      detail: {
        type: "boolean",
        description: "When true, include full text instead of summarized previews.",
      },
    },
  },
  handler: async (args) => {
    const parsedArgs = listSlideShapesArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }

    const { slideIndex, detail = false } = parsedArgs.data;
    if (slideIndex !== undefined && (!Number.isInteger(slideIndex) || slideIndex < 0)) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const resolvedSlideIndex = await resolveSlideIndex(context, slideIndex);
        const slide = await getSlideByIndex(context, resolvedSlideIndex);
        slide.shapes.load("items");
        await context.sync();

        const shapeSummaries = await loadShapeSummaries(context, slide.shapes.items, {
          includeText: true,
          includeFormatting: false,
          includeTableValues: false,
        });
        const inspection = await inspectSlideXmlFromOfficeSlide(context, slide);
        const slideId = inspection.slideId || slide.id;
        if (!slideId) {
          throw new Error("Slide XML inspection did not include a stable slide id.");
        }

        const shapes = enrichShapeSummariesWithRefs(slideId, shapeSummaries, inspection.shapes).map((shape, index) => ({
          index: shape.index,
          ref: shape.ref,
          slideId: shape.slideId,
          xmlShapeId: shape.xmlShapeId,
          name: shape.name,
          type: shape.type,
          xmlType: inspection.shapes[index]?.type || "unknown",
          box: {
            left: shape.left,
            top: shape.top,
            width: shape.width,
            height: shape.height,
          },
          hasText: !!inspection.shapes[index]?.textBody,
          ...(detail
            ? { text: shape.text || "" }
            : shape.text !== undefined
              ? { textPreview: summarizePlainText(shape.text || "") }
              : {}),
        }));

        return {
          resultType: "success",
          textResultForLlm: `Listed ${shapes.length} shapes on slide ${resolvedSlideIndex + 1}.`,
          slideIndex: resolvedSlideIndex,
          slideId,
          detail,
          shapes,
          toolTelemetry: {},
        };
      });
    } catch (error: unknown) {
      return toolFailure(error, shouldAddRoundTripRefreshHint(error) ? roundTripSlideRefreshHint() : undefined);
    }
  },
};
