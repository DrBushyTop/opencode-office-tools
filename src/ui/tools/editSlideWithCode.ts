import type { Tool } from "./types";
import { getPowerPointContextSnapshot } from "./powerpointContext";
import { resolveSlideShapeByIdWithXmlFallback } from "./powerpointShapeTarget";
import {
  formatAvailableShapeTargets,
  invalidSlideIndexMessage,
  isPowerPointRequirementSetSupported,
  roundTripRefreshHint,
  shouldAddRoundTripShapeTargetRefreshHint,
  toolFailure,
} from "./powerpointShared";
import { z } from "zod";

const editSlideWithCodeArgsSchema = z.object({
  slideIndex: z.number().optional(),
  shapeId: z.string().optional(),
  shapeIndex: z.number().optional(),
  code: z.string(),
});

type EditSlideWithCodeArgs = z.infer<typeof editSlideWithCodeArgsSchema>;

function isNonNegativeInteger(value: unknown): value is number {
  return Number.isInteger(value) && (value as number) >= 0;
}

export function normalizeEditSlideCode(input: string) {
  return String(input || "")
    .trim()
    .replace(/^```(?:javascript|js|typescript|ts)?\s*/i, "")
    .replace(/\s*```$/, "")
    .trim();
}

export async function runEditSlideCode(code: string, bindings: {
  context: PowerPoint.RequestContext;
  slide: PowerPoint.Slide;
  shapes: PowerPoint.ShapeCollection;
  targetShape?: PowerPoint.Shape;
  targetShapeId?: string;
  targetShapeIndex?: number;
  slideIndex: number;
}) {
  const normalizedCode = normalizeEditSlideCode(code);
  const run = new Function(
    "context",
    "slide",
    "shapes",
    "targetShape",
    "targetShapeId",
    "targetShapeIndex",
    "slideIndex",
    "PowerPoint",
    `return (async () => { ${normalizedCode} })();`,
  ) as (
    context: PowerPoint.RequestContext,
    slide: PowerPoint.Slide,
    shapes: PowerPoint.ShapeCollection,
    targetShape: PowerPoint.Shape | undefined,
    targetShapeId: string | undefined,
    targetShapeIndex: number | undefined,
    slideIndex: number,
    powerPointApi: typeof PowerPoint,
  ) => Promise<unknown>;

  return await run(
    bindings.context,
    bindings.slide,
    bindings.shapes,
    bindings.targetShape,
    bindings.targetShapeId,
    bindings.targetShapeIndex,
    bindings.slideIndex,
    PowerPoint,
  );
}

function resolveArgsWithPowerPointContext(args: EditSlideWithCodeArgs): EditSlideWithCodeArgs {
  const snapshot = getPowerPointContextSnapshot();
  if (!snapshot) return args;

  const next = { ...args };
  if (next.slideIndex === undefined && snapshot.activeSlideIndex !== undefined) {
    next.slideIndex = snapshot.activeSlideIndex;
  }

  if (next.shapeId === undefined && next.shapeIndex === undefined && snapshot.selectedShapeIds.length === 1) {
    next.shapeId = snapshot.selectedShapeIds[0];
  }

  return next;
}

export const editSlideWithCode: Tool = {
  name: "edit_slide_with_code",
  description: `Edit an existing PowerPoint slide in place by running JavaScript against the live Office.js slide object.

Use this for pinpoint edits on the current deck when generic slide tools are too rigid but you still need to preserve the deck's actual layout, placeholders, and formatting.

Injected variables:
- context: current PowerPoint request context
- slide: the targeted live PowerPoint slide
- shapes: slide.shapes for collection access
- targetShape: the targeted shape when shapeId, shapeIndex, or a single selected shape resolves one
- targetShapeId: resolved target shape id when available
- targetShapeIndex: resolved target shape index when available
- slideIndex: resolved 0-based slide index
- PowerPoint: the Office.js PowerPoint namespace

Notes:
- Your code may use await and call await context.sync() when it needs to read loaded values.
- If slideIndex is omitted, the active slide is used when available from the current PowerPoint selection.
- If shapeId and shapeIndex are omitted, a single selected shape is used automatically when available.
- Prefer this for live edits. Use add_slide_from_code only when you need to generate an entirely new or replacement slide from PptxGenJS.

Examples:
1. Replace text on a selected shape:
   targetShape.getTextFrameOrNullObject().textRange.text = "Executive summary";

2. Resize a targeted shape:
   targetShape.width = 320;
   targetShape.height = 120;

3. Add a native shape to the current slide:
   slide.shapes.addTextBox("Follow-up", { left: 36, top: 420, width: 240, height: 28 });`,
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        type: "number",
        description: "Optional 0-based slide index. If omitted, the active slide is used when available from the current PowerPoint context.",
      },
      shapeId: {
        type: "string",
        description: "Optional shape id to target for pinpoint edits. Preferred when available.",
      },
      shapeIndex: {
        type: "number",
        description: "Optional 0-based shape index on the targeted slide. Use when shapeId is unavailable.",
      },
      code: {
        type: "string",
        description: "JavaScript function body that runs against the live slide. The code may use await and can reference context, slide, shapes, targetShape, targetShapeId, targetShapeIndex, slideIndex, and PowerPoint.",
      },
    },
    required: ["code"],
  },
  handler: async (args) => {
    const parsedArgs = editSlideWithCodeArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }

    const resolvedArgs = resolveArgsWithPowerPointContext(parsedArgs.data);
    if (resolvedArgs.slideIndex !== undefined && !isNonNegativeInteger(resolvedArgs.slideIndex)) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }
    if (resolvedArgs.shapeIndex !== undefined && !isNonNegativeInteger(resolvedArgs.shapeIndex)) {
      return toolFailure("shapeIndex must be a non-negative integer.");
    }

    const normalizedCode = normalizeEditSlideCode(resolvedArgs.code);
    if (!normalizedCode) {
      return toolFailure("code cannot be empty.");
    }

    if (!isPowerPointRequirementSetSupported("1.3")) {
      return toolFailure("Editing live slides with code requires PowerPointApi 1.3.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const slideCount = slides.items.length;
        if (slideCount === 0) {
          return toolFailure("Presentation has no slides.");
        }

        if (resolvedArgs.slideIndex === undefined) {
          return toolFailure("slideIndex is required when no active slide can be inferred from the current PowerPoint context.");
        }

        if (resolvedArgs.slideIndex >= slideCount) {
          return toolFailure(invalidSlideIndexMessage(resolvedArgs.slideIndex, slideCount));
        }

        const slide = slides.items[resolvedArgs.slideIndex];
        slide.shapes.load("items/id,name");
        await context.sync();

        let targetShape: PowerPoint.Shape | undefined;
        let targetShapeId: string | undefined;
        let targetShapeIndex: number | undefined;

        if (resolvedArgs.shapeId !== undefined) {
          const resolvedTarget = await resolveSlideShapeByIdWithXmlFallback(context, slide, resolvedArgs.slideIndex, resolvedArgs.shapeId);
          targetShape = resolvedTarget.shape;
          targetShapeId = resolvedTarget.shapeId;
          targetShapeIndex = resolvedTarget.shapeIndex;
        } else if (resolvedArgs.shapeIndex !== undefined) {
          const indexedShape = slide.shapes.items[resolvedArgs.shapeIndex];
          if (!indexedShape) {
            return toolFailure(`Invalid shapeIndex ${resolvedArgs.shapeIndex}. ${formatAvailableShapeTargets(resolvedArgs.slideIndex, slide.shapes.items)}`);
          }
          targetShape = indexedShape;
          targetShapeId = indexedShape.id;
          targetShapeIndex = resolvedArgs.shapeIndex;
        }

        const executionResult = await runEditSlideCode(normalizedCode, {
          context,
          slide,
          shapes: slide.shapes,
          targetShape,
          targetShapeId,
          targetShapeIndex,
          slideIndex: resolvedArgs.slideIndex,
        });
        await context.sync();

        const executionSummary = typeof executionResult === "string" && executionResult.trim()
          ? ` ${executionResult.trim()}`
          : "";

        return {
          resultType: "success",
          textResultForLlm: targetShapeId
            ? `Edited shape ${targetShapeId} on slide ${resolvedArgs.slideIndex + 1}.${executionSummary}`
            : `Edited slide ${resolvedArgs.slideIndex + 1}.${executionSummary}`,
          slideIndex: resolvedArgs.slideIndex,
          targetShapeId,
          targetShapeIndex,
          toolTelemetry: {
            slideIndex: resolvedArgs.slideIndex,
            targetShapeId,
            targetShapeIndex,
          },
        };
      });
    } catch (error: unknown) {
      return toolFailure(
        error,
        shouldAddRoundTripShapeTargetRefreshHint(error) ? roundTripRefreshHint() : undefined,
      );
    }
  },
};
