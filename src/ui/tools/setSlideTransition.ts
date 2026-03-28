import type { Tool } from "./types";
import { replaceSlideWithMutatedOpenXml, setSlideTransitionInBase64Presentation, type SlideTransitionDefinition } from "./powerpointOpenXml";
import { roundTripSlideRefreshHint, shouldAddRoundTripRefreshHint, toolFailure } from "./powerpointShared";

const EFFECTS = ["none", "cut", "fade", "dissolve", "random", "randomBar", "push", "wipe", "split", "cover", "pull", "zoom"] as const;

type TransitionArgs = SlideTransitionDefinition & {
  slideIndex: number | number[];
};

export const setSlideTransition: Tool = {
  name: "set_slide_transition",
  description: "Set a PowerPoint slide transition by round-tripping the slide through Open XML. This replaces the slide in the deck and may change slide identity. Supports a safe subset of transition effects; object animations remain out of scope for now.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        anyOf: [
          { type: "number" },
          { type: "array", items: { type: "number" } },
        ],
        description: "0-based slide index or array of slide indexes. When an array is provided, the same transition is applied to each slide in sequence.",
      },
      effect: { type: "string", enum: [...EFFECTS], description: "Transition effect to apply. Use none to clear the transition." },
      speed: { type: "string", enum: ["slow", "medium", "fast"], description: "Optional transition speed." },
      advanceOnClick: { type: "boolean", description: "Whether a click advances the slide." },
      advanceAfterMs: { type: "number", description: "Optional auto-advance delay in milliseconds." },
      durationMs: { type: "number", description: "Optional transition duration in milliseconds for clients that support the p14 duration extension." },
      direction: { type: "string", enum: ["left", "right", "up", "down", "horizontal", "vertical", "in", "out"], description: "Optional direction for effects that support it." },
      orientation: { type: "string", enum: ["horizontal", "vertical"], description: "Optional orientation for split transitions." },
      throughBlack: { type: "boolean", description: "Optional cut-through-black flag for the cut effect." },
    },
    required: ["slideIndex", "effect"],
  },
  handler: async (args) => {
    const transition = args as TransitionArgs;
    const slideIndexes = Array.isArray(transition.slideIndex) ? transition.slideIndex : [transition.slideIndex];
    if (slideIndexes.length === 0) {
      return toolFailure("slideIndex array must not be empty.");
    }
    if (slideIndexes.some((slideIndex) => !Number.isInteger(slideIndex) || slideIndex < 0)) {
      return toolFailure("slideIndex must be a non-negative integer or a non-empty array of non-negative integers.");
    }
    if (transition.advanceAfterMs !== undefined && (!Number.isFinite(transition.advanceAfterMs) || transition.advanceAfterMs < 0)) {
      return toolFailure("advanceAfterMs must be a non-negative number.");
    }
    if (transition.durationMs !== undefined && (!Number.isFinite(transition.durationMs) || transition.durationMs < 0)) {
      return toolFailure("durationMs must be a non-negative number.");
    }
    if (["push", "wipe", "cover", "pull"].includes(transition.effect) && transition.direction && !["left", "right", "up", "down"].includes(transition.direction)) {
      return toolFailure(`${transition.effect} transitions only support left, right, up, or down directions.`);
    }
    if (transition.effect === "randomBar" && transition.direction && !["horizontal", "vertical"].includes(transition.direction)) {
      return toolFailure("randomBar transitions only support horizontal or vertical direction.");
    }
    if (transition.effect === "split") {
      if (transition.direction && !["in", "out"].includes(transition.direction)) {
        return toolFailure("split transitions only support in or out for direction.");
      }
      if (transition.orientation && !["horizontal", "vertical"].includes(transition.orientation)) {
        return toolFailure("split transitions only support horizontal or vertical orientation.");
      }
    }

    try {
      return await PowerPoint.run(async (context) => {
        const definition: SlideTransitionDefinition = {
          effect: transition.effect,
          speed: transition.speed,
          advanceOnClick: transition.advanceOnClick,
          advanceAfterMs: transition.advanceAfterMs,
          durationMs: transition.durationMs,
          direction: transition.direction,
          orientation: transition.orientation,
          throughBlack: transition.throughBlack,
        };

        for (const slideIndex of slideIndexes) {
          await replaceSlideWithMutatedOpenXml(context, slideIndex, (base64) =>
            setSlideTransitionInBase64Presentation(base64, definition),
          );
        }

        if (slideIndexes.length === 1) {
          return transition.effect === "none"
            ? `Cleared the transition on slide ${slideIndexes[0] + 1} via an Open XML slide round-trip.`
            : `Set the ${transition.effect} transition on slide ${slideIndexes[0] + 1} via an Open XML slide round-trip.`;
        }

        const slideList = slideIndexes.map((slideIndex) => slideIndex + 1).join(", ");
        return transition.effect === "none"
          ? `Cleared transitions on slides ${slideList} via Open XML slide round-trips.`
          : `Set the ${transition.effect} transition on slides ${slideList} via Open XML slide round-trips.`;
      });
    } catch (error: unknown) {
      return toolFailure(error, shouldAddRoundTripRefreshHint(error) ? roundTripSlideRefreshHint() : undefined);
    }
  },
};
