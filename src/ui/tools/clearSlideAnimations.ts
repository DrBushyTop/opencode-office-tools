import type { Tool } from "./types";
import { clearSlideAnimationsInBase64Presentation, replaceSlideWithMutatedOpenXml } from "./powerpointOpenXml";
import { resolvePowerPointSlideIndexes } from "./powerpointContext";
import { roundTripSlideRefreshHint, shouldAddRoundTripRefreshHint, toolFailure } from "./powerpointShared";

export const clearSlideAnimations: Tool = {
  name: "clear_slide_animations",
  description: "Remove all animations from a PowerPoint slide through an Open XML slide round-trip. This replaces the slide in the deck and may change slide identity.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: { type: "number", description: "0-based slide index." },
    },
  },
  handler: async (args) => {
    const slideIndex = resolvePowerPointSlideIndexes((args as { slideIndex?: number }).slideIndex);
    if (!Number.isInteger(slideIndex) || (slideIndex as number) < 0) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }
    const resolvedSlideIndex = slideIndex as number;

    try {
      return await PowerPoint.run(async (context) => {
        const result = await replaceSlideWithMutatedOpenXml(context, resolvedSlideIndex, (base64, _sourceSlide) => clearSlideAnimationsInBase64Presentation(base64));
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
