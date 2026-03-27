import type { Tool } from "./types";
import { replaceSlideWithMutatedOpenXml, setSlideTransitionInBase64Presentation, type SlideTransitionDefinition } from "./powerpointOpenXml";
import { toolFailure } from "./powerpointShared";

const EFFECTS = ["none", "cut", "fade", "dissolve", "random", "randomBar", "push", "wipe", "split", "cover", "pull", "zoom"] as const;

export const setSlideTransition: Tool = {
  name: "set_slide_transition",
  description: "Set a PowerPoint slide transition by round-tripping the slide through Open XML. This replaces the slide in the deck and may change slide identity. Supports a safe subset of transition effects; object animations remain out of scope for now.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: { type: "number", description: "0-based slide index." },
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
    const definition = args as SlideTransitionDefinition & { slideIndex: number };
    if (!Number.isInteger(definition.slideIndex) || definition.slideIndex < 0) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }
    if (definition.advanceAfterMs !== undefined && (!Number.isFinite(definition.advanceAfterMs) || definition.advanceAfterMs < 0)) {
      return toolFailure("advanceAfterMs must be a non-negative number.");
    }
    if (definition.durationMs !== undefined && (!Number.isFinite(definition.durationMs) || definition.durationMs < 0)) {
      return toolFailure("durationMs must be a non-negative number.");
    }
    if (["push", "wipe", "cover", "pull"].includes(definition.effect) && definition.direction && !["left", "right", "up", "down"].includes(definition.direction)) {
      return toolFailure(`${definition.effect} transitions only support left, right, up, or down directions.`);
    }
    if (definition.effect === "randomBar" && definition.direction && !["horizontal", "vertical"].includes(definition.direction)) {
      return toolFailure("randomBar transitions only support horizontal or vertical direction.");
    }
    if (definition.effect === "split") {
      if (definition.direction && !["in", "out"].includes(definition.direction)) {
        return toolFailure("split transitions only support in or out for direction.");
      }
      if (definition.orientation && !["horizontal", "vertical"].includes(definition.orientation)) {
        return toolFailure("split transitions only support horizontal or vertical orientation.");
      }
    }

    try {
      return await PowerPoint.run(async (context) => {
        await replaceSlideWithMutatedOpenXml(context, definition.slideIndex, (base64) => setSlideTransitionInBase64Presentation(base64, definition));
        return definition.effect === "none"
          ? `Cleared the transition on slide ${definition.slideIndex + 1} via an Open XML slide round-trip.`
          : `Set the ${definition.effect} transition on slide ${definition.slideIndex + 1} via an Open XML slide round-trip.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
