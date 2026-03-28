import type { Tool } from "./types";
import { replaceSlideWithMutatedOpenXml, setSpeakerNotesInBase64Presentation } from "./powerpointOpenXml";
import { roundTripSlideRefreshHint, shouldAddRoundTripRefreshHint, toolFailure } from "./powerpointShared";

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
    required: ["slideIndex", "notes"],
  },
  handler: async (args) => {
    const { slideIndex, notes } = args as { slideIndex: number; notes: string };
    if (!Number.isInteger(slideIndex) || slideIndex < 0) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        await replaceSlideWithMutatedOpenXml(context, slideIndex, (base64) => setSpeakerNotesInBase64Presentation(base64, notes));
        return `Updated speaker notes on slide ${slideIndex + 1} via an Open XML slide round-trip.`;
      });
    } catch (error: unknown) {
      return toolFailure(error, shouldAddRoundTripRefreshHint(error) ? roundTripSlideRefreshHint() : undefined);
    }
  },
};
