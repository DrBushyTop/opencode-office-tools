import type { Tool } from "./types";
import {
  addSlideAnimationInBase64Presentation,
  replaceSlideWithMutatedOpenXml,
  type SlideAnimationDefinition,
} from "./powerpointOpenXml";
import { resolveSlideShapeByIdWithXmlFallback } from "./powerpointShapeTarget";
import { formatAvailableShapeTargets, invalidSlideIndexMessage, roundTripRefreshHint, shouldAddRoundTripShapeTargetRefreshHint, toolFailure } from "./powerpointShared";

type AnimationArgs = SlideAnimationDefinition & { slideIndex: number; shapeIndex?: number; shapeId?: string | number };

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
  description: "Add a PowerPoint slide animation through an Open XML slide round-trip. Supports motion paths, scale emphasis, and rotation with on-click, with-previous, or after-previous timing. This replaces the slide in the deck and may change slide identity.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: { type: "number", description: "0-based slide index." },
      shapeId: {
        anyOf: [{ type: "string" }, { type: "number" }],
        description: "Preferred Office shape id target, or an exported XML p:cNvPr id after an Open XML slide replacement.",
      },
      shapeIndex: { type: "number", description: "0-based shape index target if shapeId is unavailable." },
      type: { type: "string", enum: ["motionPath", "scale", "rotate"], description: "Animation type." },
      start: { type: "string", enum: ["onClick", "withPrevious", "afterPrevious"], description: "When the animation starts relative to the sequence." },
      durationMs: { type: "number", description: "Optional animation duration in milliseconds. Default 1000." },
      delayMs: { type: "number", description: "Optional start delay in milliseconds." },
      repeatCount: { type: "number", description: "Optional repeat count." },
      path: { type: "string", description: "Motion path string such as 'M 0 0 L 0.25 0 E'. Required for motionPath." },
      pathOrigin: { type: "string", enum: ["parent", "layout"], description: "Optional motion-path origin." },
      pathEditMode: { type: "string", enum: ["relative", "fixed"], description: "Optional motion-path edit mode." },
      scaleXPercent: { type: "number", description: "Relative X scale change as a percentage. Example: 150 makes the shape 150% larger." },
      scaleYPercent: { type: "number", description: "Relative Y scale change as a percentage. Defaults to scaleXPercent." },
      angleDegrees: { type: "number", description: "Rotation amount in degrees for rotate animations." },
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

    try {
      return await PowerPoint.run(async (context) => {
        const resolvedTarget = await resolveShapeTarget(context, animation.slideIndex, animation.shapeId, animation.shapeIndex);
        await replaceSlideWithMutatedOpenXml(context, animation.slideIndex, (base64) => addSlideAnimationInBase64Presentation(base64, {
          ...animation,
          shapeId: resolvedTarget.shapeId,
        }, resolvedTarget.shapeIndex));
        return `Added a ${animation.type} animation to slide ${animation.slideIndex + 1} targeting shape ${resolvedTarget.shapeId} via an Open XML slide round-trip.`;
      });
    } catch (error: unknown) {
      return toolFailure(error, shouldAddRoundTripShapeTargetRefreshHint(error) ? roundTripRefreshHint() : undefined);
    }
  },
};
