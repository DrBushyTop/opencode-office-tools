import type { Tool } from "./types";
import {
  addSlideAnimationBatchToXmlShapeIdsInBase64Presentation,
  addSlideAnimationInBase64Presentation,
  addSlideAnimationToXmlShapeIdInBase64Presentation,
  findSlideShapeIndexByXmlShapeIdInBase64Presentation,
  replaceSlideWithMutatedOpenXml,
  resolveXmlShapeIdByMetadataInBase64Presentation,
  type SlideAnimationDefinition,
} from "./powerpointOpenXml";
import { resolvePowerPointTargetingArgs } from "./powerpointContext";
import { formatAvailableShapeTargets, roundTripRefreshHint, shouldAddRoundTripShapeTargetRefreshHint, toolFailure } from "./powerpointShared";
import { z } from "zod";

type AnimationArgs = SlideAnimationDefinition & {
  slideIndex?: number;
  shapeIndex?: number;
  shapeId?: string | number | (string | number)[];
};

const animationArgsSchema = z.object({
  slideIndex: z.number().optional(),
  shapeId: z.union([z.string(), z.number(), z.array(z.union([z.string(), z.number()]))]).optional(),
  shapeIndex: z.number().optional(),
  type: z.enum(["motionPath", "scale", "rotate", "appear", "fade", "flyIn", "wipe", "zoomIn", "floatIn", "riseUp", "peekIn", "growAndTurn", "complementaryColor", "changeFillColor", "changeLineColor"]),
  start: z.enum(["onClick", "withPrevious", "afterPrevious"]),
  durationMs: z.number().optional(),
  delayMs: z.number().optional(),
  repeatCount: z.number().optional(),
  path: z.string().optional(),
  pathOrigin: z.enum(["parent", "layout"]).optional(),
  pathEditMode: z.enum(["relative", "fixed"]).optional(),
  scaleXPercent: z.number().optional(),
  scaleYPercent: z.number().optional(),
  angleDegrees: z.number().optional(),
  direction: z.enum(["left", "right", "up", "down"]).optional(),
  toColor: z.string().optional(),
  colorSpace: z.enum(["hsl", "rgb"]).optional(),
});

/**
 * Resolve a shape ID against already-loaded shapes, with XML fallback using
 * the exported base64 that's already available in the mutate callback.
 * This avoids any extra `context.sync()` calls — everything is pre-loaded.
 */
function resolveShapeSync(
  shapes: PowerPoint.Shape[],
  slideIndex: number,
  base64: string,
  requestedId: string | number,
): { shapeId: string; shapeIndex: number } {
  const id = String(requestedId);

  // Try direct match against Office.js shape IDs
  const directIdx = shapes.findIndex((s) => s.id === id);
  if (directIdx >= 0) {
    return { shapeId: shapes[directIdx].id, shapeIndex: directIdx };
  }

  // XML fallback: the exported base64 uses different (XML) shape IDs
  const xmlIdx = findSlideShapeIndexByXmlShapeIdInBase64Presentation(base64, id);
  if (xmlIdx >= 0 && xmlIdx < shapes.length) {
    return { shapeId: shapes[xmlIdx].id, shapeIndex: xmlIdx };
  }

  throw new Error(`Shape ${id} was not found on slide ${slideIndex + 1}. ${formatAvailableShapeTargets(slideIndex, shapes)}`);
}

function resolveShapeTargetSync(
  shapes: PowerPoint.Shape[],
  slideIndex: number,
  base64: string,
  requestedId: string | number,
): { shapeId: string; shapeIndex: number; targetXmlShapeId: string } {
  const resolved = resolveShapeSync(shapes, slideIndex, base64, requestedId);
  const targetShape = shapes[resolved.shapeIndex];
  const targetXmlShapeId = resolveXmlShapeIdByMetadataInBase64Presentation(
    base64,
    {
      name: targetShape.name,
      left: targetShape.left,
      top: targetShape.top,
      width: targetShape.width,
      height: targetShape.height,
    },
    resolved.shapeIndex,
  ) || String(requestedId);

  return {
    ...resolved,
    targetXmlShapeId,
  };
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
    required: ["type", "start"],
  },
  handler: async (args) => {
    const parsedArgs = animationArgsSchema.safeParse(args);
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }
    const animation = resolvePowerPointTargetingArgs(parsedArgs.data as AnimationArgs);
    if (!Number.isInteger(animation.slideIndex) || (animation.slideIndex as number) < 0) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }
    const slideIndex = animation.slideIndex as number;
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
          // Batch mode: preload shapes, resolve all in the mutate callback
          let resolvedShapeIds: string[] = [];
          const roundTrip = await replaceSlideWithMutatedOpenXml(
            context,
            slideIndex,
            (base64, sourceSlide) => {
              const shapes = sourceSlide.shapes.items;
              const resolved = shapeIds.map((id) => resolveShapeTargetSync(shapes, slideIndex, base64, id));
              resolvedShapeIds = resolved.map((r) => r.shapeId);
              return addSlideAnimationBatchToXmlShapeIdsInBase64Presentation(
                base64,
                { ...animation, shapeId: resolved[0].shapeId },
                resolved.map((r) => r.targetXmlShapeId),
              );
            },
            { preload: (slide) => slide.shapes.load("items/id,name,left,top,width,height") },
          );
          return {
            resultType: "success",
            textResultForLlm: `Added ${animation.type} animation to slide ${roundTrip.finalSlideIndex + 1} targeting ${resolvedShapeIds.length} shapes.`,
            slideIndex: roundTrip.finalSlideIndex,
            slideId: roundTrip.replacementSlideId,
            shapeIds: resolvedShapeIds,
            toolTelemetry: {
              ...roundTrip,
              shapeIds: resolvedShapeIds,
            },
          };
        }

        // Single shape mode: preload shapes, resolve in the mutate callback
        const singleId = isBatch ? shapeIds![0] : animation.shapeId as string | number | undefined;
        let resolvedShapeId = "";
        const roundTrip = await replaceSlideWithMutatedOpenXml(
          context,
          slideIndex,
          (base64, sourceSlide) => {
            const shapes = sourceSlide.shapes.items;
            let targetShapeIndex: number;

            if (singleId !== undefined) {
              const resolved = resolveShapeTargetSync(shapes, slideIndex, base64, singleId);
              resolvedShapeId = resolved.shapeId;
              targetShapeIndex = resolved.shapeIndex;
              return addSlideAnimationToXmlShapeIdInBase64Presentation(base64, {
                ...animation,
                shapeId: resolvedShapeId,
              }, resolved.targetXmlShapeId);
            } else {
              // shapeIndex-only mode
              if (animation.shapeIndex === undefined || !Number.isInteger(animation.shapeIndex) || animation.shapeIndex < 0 || animation.shapeIndex >= shapes.length) {
                throw new Error(`Provide a valid shapeId or shapeIndex for slide ${slideIndex + 1}. ${formatAvailableShapeTargets(slideIndex, shapes)}`);
              }
              resolvedShapeId = shapes[animation.shapeIndex].id;
              targetShapeIndex = animation.shapeIndex;
            }

            return addSlideAnimationInBase64Presentation(base64, {
              ...animation,
              shapeId: resolvedShapeId,
            }, targetShapeIndex);
          },
          { preload: (slide) => slide.shapes.load("items/id,name,left,top,width,height") },
        );
        return {
          resultType: "success",
          textResultForLlm: `Added a ${animation.type} animation to slide ${roundTrip.finalSlideIndex + 1} targeting shape ${resolvedShapeId}.`,
          slideIndex: roundTrip.finalSlideIndex,
          slideId: roundTrip.replacementSlideId,
          shapeId: resolvedShapeId,
          refreshedShapeId: resolvedShapeId,
          toolTelemetry: {
            ...roundTrip,
            shapeId: resolvedShapeId,
            refreshedShapeId: resolvedShapeId,
          },
        };
      });
    } catch (error: unknown) {
      return toolFailure(error, shouldAddRoundTripShapeTargetRefreshHint(error) ? roundTripRefreshHint() : undefined);
    }
  },
};
