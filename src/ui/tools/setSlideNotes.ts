import type { Tool } from "./types";
import { replaceSlideWithMutatedOpenXml, setSpeakerNotesInBase64Presentation } from "./powerpointOpenXml";
import { resolvePowerPointSlideIndexes } from "./powerpointContext";
import { roundTripSlideRefreshHint, shouldAddRoundTripRefreshHint, toolFailure } from "./powerpointShared";
import { z } from "zod";

const setSlideNotesArgsSchema = z.object({
  slideIndex: z.number().optional(),
  notes: z.string(),
});

export const setSlideNotes: Tool = {
  name: "set_slide_notes",
  description: "Add or update PowerPoint speaker notes by round-tripping the slide through an Open XML package when the native PowerPoint API does not expose notes editing. This replaces the slide in the deck and may change slide identity.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        type: "number",
        description: "0-based index of the slide to update.",
      },
      notes: {
        type: "string",
        description: "Speaker notes text. Use an empty string to clear the notes body.",
      },
    },
    required: ["notes"],
  },
  handler: async (args) => {
    const parsedArgs = setSlideNotesArgsSchema.safeParse(args);
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }
    const { notes } = parsedArgs.data;
    const slideIndex = resolvePowerPointSlideIndexes(parsedArgs.data.slideIndex);
    if (!Number.isInteger(slideIndex) || (slideIndex as number) < 0) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }
    const resolvedSlideIndex = slideIndex as number;

    try {
      return await PowerPoint.run(async (context) => {
        const result = await replaceSlideWithMutatedOpenXml(context, resolvedSlideIndex, (base64, _sourceSlide) => setSpeakerNotesInBase64Presentation(base64, notes));
        return {
          resultType: "success",
          textResultForLlm: `Updated speaker notes on slide ${result.finalSlideIndex + 1}.`,
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
