import type { Tool } from "./types";
import {
  addSlideAnimationBatchInBase64Presentation,
  addSlideAnimationInBase64Presentation,
  replaceSlideWithMutatedOpenXml,
  type SlideAnimationDefinition,
} from "./powerpointOpenXml";
import { resolveSlideShapeByIdWithXmlFallback } from "./powerpointShapeTarget";
import { formatAvailableShapeTargets, invalidSlideIndexMessage, roundTripRefreshHint, shouldAddRoundTripShapeTargetRefreshHint, toolFailure } from "./powerpointShared";

type AnimationArgs = SlideAnimationDefinition & {
  slideIndex: number;
  shapeIndex?: number;
  shapeId?: string | number | (string | number)[];
};

async function resolveShapeTarget(context: PowerPoint.RequestContext, slideIndex: number, shapeId?: string | number, shapeIndex?: number) {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  if (slideIndex < 0 || slideIndex >= slides.items.length) {
    throw new Error(invalidSlideIndexMessage(slideIndex, slides.items.length));
  }

  const slide = slides.items[slideIndex];

  if (shapeId !== undefined) {
    const resolved = await resolveSlideShapeByIdWithXmlFallback(context, slide, slideIndex, shapeId);
    return { shapeId: resolved.shapeId, shapeIndex: resolved.shapeIndex };
  }

  slide.shapes.load("items/id,name");
  await context.sync();

  if (shapeIndex === undefined || !Number.isInteger(shapeIndex) || shapeIndex < 0 || shapeIndex >= slide.shapes.items.length) {
    throw new Error(`Provide a valid shapeId or shapeIndex for slide ${slideIndex + 1}. ${formatAvailableShapeTargets(slideIndex, slide.shapes.items)}`);
  }

  return { shapeId: slide.shapes.items[shapeIndex].id, shapeIndex };
}

export const addSlideAnimation: Tool = {
  name: "add_slide_animation",
  description: "Add a PowerPoint slide animation through an Open XML slide round-trip. Supports motion paths, scale emphasis, and rotation with timing control. Also supports entrance animations: appear (instant), fade (opacity), flyIn (from direction), wipe (reveal), zoomIn (scale in), floatIn (float up with fade), riseUp (rise from bottom), peekIn (fade with slight upward slide), and growAndTurn (fade with bounce from below). Entrance animations make shapes start hidden and reveal them. Emphasis color animations (complementaryColor, changeFillColor, changeLineColor) smoothly transition a shape's fill or line color. Use afterPrevious with delayMs for staggered reveal sequences. This replaces the slide in the deck and may change slide identity.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: { type: "number", description: "0-based slide index." },
      shapeId: {
        anyOf: [
          { type: "string" },
          { type: "number" },
          { type: "array", items: { anyOf: [{ type: "string" }, { type: "number" }] } },
        ],
        description: "Shape id target(s). Pass a single id or an array of ids to animate multiple shapes in one call. When an array is provided, the first shape uses the specified start trigger and the rest use withPrevious.",
      },
      shapeIndex: { type: "number", description: "0-based shape index target if shapeId is unavailable. Only for single-shape animations." },
      type: { type: "string", enum: ["motionPath", "scale", "rotate", "appear", "fade", "flyIn", "wipe", "zoomIn", "floatIn", "riseUp", "peekIn", "growAndTurn", "complementaryColor", "changeFillColor", "changeLineColor"], description: "Animation type. Entrance types (appear, fade, flyIn, wipe, zoomIn, floatIn, riseUp, peekIn, growAndTurn) start shapes hidden and reveal them. Emphasis color types (complementaryColor, changeFillColor, changeLineColor) animate a shape's color." },
      start: { type: "string", enum: ["onClick", "withPrevious", "afterPrevious"], description: "When the animation starts relative to the sequence. Use afterPrevious with delayMs for staggered reveals." },
      durationMs: { type: "number", description: "Optional animation duration in milliseconds. Default 1000. For appear, this is effectively instant." },
      delayMs: { type: "number", description: "Optional start delay in milliseconds. Useful for staggered entrance sequences." },
      repeatCount: { type: "number", description: "Optional repeat count." },
      path: { type: "string", description: "Motion path string such as 'M 0 0 L 0.25 0 E'. Required for motionPath." },
      pathOrigin: { type: "string", enum: ["parent", "layout"], description: "Optional motion-path origin." },
      pathEditMode: { type: "string", enum: ["relative", "fixed"], description: "Optional motion-path edit mode." },
      scaleXPercent: { type: "number", description: "Relative X scale change as a percentage. Example: 150 makes the shape 150% larger." },
      scaleYPercent: { type: "number", description: "Relative Y scale change as a percentage. Defaults to scaleXPercent." },
      angleDegrees: { type: "number", description: "Rotation amount in degrees for rotate animations." },
      direction: { type: "string", enum: ["left", "right", "up", "down"], description: "Direction for flyIn (where the shape flies from) or wipe (reveal direction) entrance animations." },
      toColor: { type: "string", description: "Target color for emphasis color animations. Hex without # (e.g. 'FF0000') or theme scheme name (e.g. 'accent2', 'dk1'). Required for complementaryColor, changeFillColor, changeLineColor." },
      colorSpace: { type: "string", enum: ["hsl", "rgb"], description: "Color interpolation space for emphasis color animations. 'hsl' (default) gives smoother transitions." },
    },
    required: ["slideIndex", "type", "start"],
  },
  handler: async (args) => {
    const animation = args as AnimationArgs;
    if (!Number.isInteger(animation.slideIndex) || animation.slideIndex < 0) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }
    if (animation.durationMs !== undefined && (!Number.isFinite(animation.durationMs) || animation.durationMs < 0)) {
      return toolFailure("durationMs must be a non-negative number.");
    }
    if (animation.delayMs !== undefined && (!Number.isFinite(animation.delayMs) || animation.delayMs < 0)) {
      return toolFailure("delayMs must be a non-negative number.");
    }
    if (animation.repeatCount !== undefined && (!Number.isFinite(animation.repeatCount) || animation.repeatCount < 0)) {
      return toolFailure("repeatCount must be a non-negative number.");
    }
    if (animation.type === "motionPath" && !animation.path) {
      return toolFailure("path is required for motionPath animations.");
    }
    if (animation.type === "scale" && animation.scaleXPercent === undefined && animation.scaleYPercent === undefined) {
      return toolFailure("scaleXPercent or scaleYPercent is required for scale animations.");
    }
    if (animation.type === "rotate" && animation.angleDegrees === undefined) {
      return toolFailure("angleDegrees is required for rotate animations.");
    }
    if (animation.type === "flyIn" && animation.direction && !["left", "right", "up", "down"].includes(animation.direction)) {
      return toolFailure("flyIn direction must be left, right, up, or down.");
    }
    if (animation.type === "wipe" && animation.direction && !["left", "right", "up", "down"].includes(animation.direction)) {
      return toolFailure("wipe direction must be left, right, up, or down.");
    }
    if ((animation.type === "complementaryColor" || animation.type === "changeFillColor" || animation.type === "changeLineColor") && !animation.toColor) {
      return toolFailure("toColor is required for emphasis color animations (complementaryColor, changeFillColor, changeLineColor).");
    }

    // Normalize shapeId to detect batch mode
    const isBatch = Array.isArray(animation.shapeId);
    const shapeIds = isBatch ? animation.shapeId as (string | number)[] : undefined;

    if (isBatch && shapeIds!.length === 0) {
      return toolFailure("shapeId array must not be empty.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        if (shapeIds && shapeIds.length > 1) {
          // Batch mode: resolve all shape IDs, apply all in one round-trip
          const resolved: { shapeId: string; shapeIndex: number }[] = [];
          for (const sid of shapeIds) {
            resolved.push(await resolveShapeTarget(context, animation.slideIndex, sid));
          }
          const roundTrip = await replaceSlideWithMutatedOpenXml(context, animation.slideIndex, (base64) =>
            addSlideAnimationBatchInBase64Presentation(
              base64,
              { ...animation, shapeId: resolved[0].shapeId },
              resolved.map((r) => r.shapeIndex),
            ),
          );
          const refreshedShapeIds = resolved.map((entry) => roundTrip.shapeIdMap?.[entry.shapeId] || entry.shapeId);
          return {
            resultType: "success",
            textResultForLlm: `Added ${animation.type} animation to slide ${roundTrip.finalSlideIndex + 1} targeting ${resolved.length} shapes.`,
            slideIndex: roundTrip.finalSlideIndex,
            slideId: roundTrip.replacementSlideId,
            shapeIds: refreshedShapeIds,
            toolTelemetry: {
              ...roundTrip,
              shapeIds: refreshedShapeIds,
            },
          };
        }

        // Single shape mode
        const singleId = isBatch ? shapeIds![0] : animation.shapeId as string | number | undefined;
        const resolvedTarget = await resolveShapeTarget(context, animation.slideIndex, singleId, animation.shapeIndex);
        const roundTrip = await replaceSlideWithMutatedOpenXml(context, animation.slideIndex, (base64) => addSlideAnimationInBase64Presentation(base64, {
          ...animation,
          shapeId: resolvedTarget.shapeId,
        }, resolvedTarget.shapeIndex));
        return {
          resultType: "success",
          textResultForLlm: `Added a ${animation.type} animation to slide ${roundTrip.finalSlideIndex + 1} targeting shape ${resolvedTarget.shapeId}.`,
          slideIndex: roundTrip.finalSlideIndex,
          slideId: roundTrip.replacementSlideId,
          shapeId: resolvedTarget.shapeId,
          refreshedShapeId: roundTrip.shapeIdMap?.[resolvedTarget.shapeId] || resolvedTarget.shapeId,
          toolTelemetry: {
            ...roundTrip,
            shapeId: resolvedTarget.shapeId,
            refreshedShapeId: roundTrip.shapeIdMap?.[resolvedTarget.shapeId] || resolvedTarget.shapeId,
          },
        };
      });
    } catch (error: unknown) {
      return toolFailure(error, shouldAddRoundTripShapeTargetRefreshHint(error) ? roundTripRefreshHint() : undefined);
    }
  },
};
