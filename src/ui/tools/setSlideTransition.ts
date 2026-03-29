import type { Tool } from "./types";
import { replaceSlideWithMutatedOpenXml, setSlideTransitionInBase64Presentation, type SlideTransitionDefinition } from "./powerpointOpenXml";
import { resolvePowerPointSlideIndexes } from "./powerpointContext";
import { roundTripSlideRefreshHint, shouldAddRoundTripRefreshHint, toolFailure } from "./powerpointShared";
import { z } from "zod";

const EFFECTS = ["none", "cut", "fade", "dissolve", "random", "randomBar", "push", "wipe", "split", "cover", "pull", "zoom"] as const;

type TransitionArgs = SlideTransitionDefinition & {
  slideIndex?: number | number[];
};

const transitionArgsSchema = z.object({
  slideIndex: z.union([z.number(), z.array(z.number())]).optional(),
  effect: z.enum(EFFECTS),
  speed: z.enum(["slow", "medium", "fast"]).optional(),
  advanceOnClick: z.boolean().optional(),
  advanceAfterMs: z.number().optional(),
  durationMs: z.number().optional(),
  direction: z.enum(["left", "right", "up", "down", "horizontal", "vertical", "in", "out"]).optional(),
  orientation: z.enum(["horizontal", "vertical"]).optional(),
  throughBlack: z.boolean().optional(),
});

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
    required: ["effect"],
  },
  handler: async (args) => {
    const parsedArgs = transitionArgsSchema.safeParse(args);
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }
    const transition = parsedArgs.data as TransitionArgs;
    const resolvedSlideIndex = resolvePowerPointSlideIndexes(transition.slideIndex);
    const transitionArgs = { ...transition, slideIndex: resolvedSlideIndex };
    const slideIndexes = Array.isArray(transitionArgs.slideIndex) ? transitionArgs.slideIndex : [transitionArgs.slideIndex];
    if (slideIndexes.length === 0) {
      return toolFailure("slideIndex array must not be empty.");
    }
    if (slideIndexes.some((slideIndex) => slideIndex === undefined)) {
      return toolFailure("slideIndex is required unless the current PowerPoint slide context is available.");
    }
    if (slideIndexes.some((slideIndex) => !Number.isInteger(slideIndex) || (slideIndex as number) < 0)) {
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
          effect: transitionArgs.effect,
          speed: transitionArgs.speed,
          advanceOnClick: transitionArgs.advanceOnClick,
          advanceAfterMs: transitionArgs.advanceAfterMs,
          durationMs: transitionArgs.durationMs,
          direction: transitionArgs.direction,
          orientation: transitionArgs.orientation,
          throughBlack: transitionArgs.throughBlack,
        };

        const results: Array<{ originalSlideId: string; replacementSlideId: string; finalSlideIndex: number }> = [];

        for (const slideIndex of slideIndexes) {
          const result = await replaceSlideWithMutatedOpenXml(context, slideIndex as number, (base64) =>
            setSlideTransitionInBase64Presentation(base64, definition),
          );
          results.push(result);
        }

        if (slideIndexes.length === 1) {
          const [result] = results;
          return {
            resultType: "success",
            textResultForLlm: transitionArgs.effect === "none"
              ? `Cleared the transition on slide ${result.finalSlideIndex + 1}.`
              : `Set the ${transitionArgs.effect} transition on slide ${result.finalSlideIndex + 1}.`,
            slideIndex: result.finalSlideIndex,
            slideId: result.replacementSlideId,
            toolTelemetry: result,
          };
        }

        const slideList = results.map((item) => item.finalSlideIndex + 1).join(", ");
        return {
          resultType: "success",
          textResultForLlm: transitionArgs.effect === "none"
            ? `Cleared transitions on slides ${slideList}.`
            : `Set the ${transitionArgs.effect} transition on slides ${slideList}.`,
          slideIndexes: results.map((item) => item.finalSlideIndex),
          slideIds: results.map((item) => item.replacementSlideId),
          toolTelemetry: { results },
        };
      });
    } catch (error: unknown) {
      return toolFailure(error, shouldAddRoundTripRefreshHint(error) ? roundTripSlideRefreshHint() : undefined);
    }
  },
};
