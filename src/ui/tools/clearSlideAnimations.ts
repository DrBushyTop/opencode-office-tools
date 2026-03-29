import type { Tool } from "./types";
import { clearSlideAnimationsInBase64Presentation, replaceSlideWithMutatedOpenXml } from "./powerpointOpenXml";
import { roundTripSlideRefreshHint, shouldAddRoundTripRefreshHint, toolFailure } from "./powerpointShared";

export const clearSlideAnimations: Tool = {
  name: "clear_slide_animations",
  description: "Remove all animations from a PowerPoint slide through an Open XML slide round-trip. This replaces the slide in the deck and may change slide identity.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: { type: "number", description: "0-based slide index." },
    },
    required: ["slideIndex"],
  },
  handler: async (args) => {
    const { slideIndex } = args as { slideIndex: number };
    if (!Number.isInteger(slideIndex) || slideIndex < 0) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const result = await replaceSlideWithMutatedOpenXml(context, slideIndex, clearSlideAnimationsInBase64Presentation);
        return {
          resultType: "success",
          textResultForLlm: `Cleared all animations from slide ${result.finalSlideIndex + 1}.`,
          slideIndex: result.finalSlideIndex,
          slideId: result.replacementSlideId,
          toolTelemetry: result,
        };
      });
    } catch (error: unknown) {
      return toolFailure(error, shouldAddRoundTripRefreshHint(error) ? roundTripSlideRefreshHint() : undefined);
    }
  },
};
