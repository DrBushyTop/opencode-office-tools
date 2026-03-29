import type { Tool } from "./types";
import { exportSlideAsBase64, extractSlideTransitionFromBase64Presentation } from "./powerpointOpenXml";
import { toolFailure } from "./powerpointShared";
import { z } from "zod";

const getSlideTransitionArgsSchema = z.object({
  slideIndex: z.number(),
});

export const getSlideTransition: Tool = {
  name: "get_slide_transition",
  description: "Inspect a slide's transition by exporting the slide through an Open XML fallback and reading the transition metadata from the slide XML.",
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
    const parsedArgs = getSlideTransitionArgsSchema.safeParse(args);
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }
    const { slideIndex } = parsedArgs.data;
    if (!Number.isInteger(slideIndex) || slideIndex < 0) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const exported = await exportSlideAsBase64(context, slideIndex);
        const transition = extractSlideTransitionFromBase64Presentation(exported);
        return JSON.stringify({ slideIndex, ...transition }, null, 2);
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
