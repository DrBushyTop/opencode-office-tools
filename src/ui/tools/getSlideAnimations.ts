import type { Tool } from "./types";
import { exportSlideAsBase64, extractSlideAnimationSummaryFromBase64Presentation } from "./powerpointOpenXml";
import { toolFailure } from "./powerpointShared";

export const getSlideAnimations: Tool = {
  name: "get_slide_animations",
  description: "Inspect animations on a PowerPoint slide by exporting through Open XML and parsing the timing tree. Returns a structured summary of all animations including their type, target shapes, timing, and sequence order.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        type: "number",
        description: "0-based slide index.",
      },
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
        const exported = await exportSlideAsBase64(context, slideIndex);
        const summary = extractSlideAnimationSummaryFromBase64Presentation(exported);
        return JSON.stringify({ slideIndex, ...summary }, null, 2);
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
