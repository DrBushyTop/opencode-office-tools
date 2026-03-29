import type { Tool } from "./types";
import { exportSlideAsBase64, extractSpeakerNotesFromBase64Presentation } from "./powerpointOpenXml";
import { toolFailure } from "./powerpointShared";
import { z } from "zod";

const getSlideNotesArgsSchema = z.object({
  slideIndex: z.number().optional(),
});

export const getSlideNotes: Tool = {
  name: "get_slide_notes",
  description: "Read PowerPoint speaker notes by exporting slides to Open XML and inspecting the notes parts when the native PowerPoint API does not expose notes directly.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        type: "number",
        description: "Optional 0-based slide index. If omitted, returns notes for all slides.",
      },
    },
  },
  handler: async (args) => {
    const parsedArgs = getSlideNotesArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }
    const { slideIndex } = parsedArgs.data;

    if (slideIndex !== undefined && (!Number.isInteger(slideIndex) || slideIndex < 0)) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        if (!slides.items.length) {
          return "Presentation has no slides.";
        }

        const indices = slideIndex === undefined
          ? slides.items.map((_, index) => index)
          : [slideIndex];

        const lines: string[] = [];
        for (const index of indices) {
          const exported = await exportSlideAsBase64(context, index);
          const notes = extractSpeakerNotesFromBase64Presentation(exported);
          lines.push(`Slide ${index + 1}: ${notes ? notes : "(no speaker notes)"}`);
        }

        return slideIndex === undefined
          ? `Speaker Notes\n${"━".repeat(40)}\n${lines.join("\n\n")}`
          : lines[0];
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
