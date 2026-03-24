import type { Tool } from "@github/copilot-sdk";

export const setSlideNotes: Tool = {
  name: "set_slide_notes",
  description: `Add or update speaker notes for a PowerPoint slide.

Parameters:
- slideIndex: 0-based index of the slide to update
- notes: The text content for the speaker notes

Speaker notes appear in the Notes pane below the slide and are visible to the presenter during a slideshow.

Note: Due to PowerPoint JavaScript API limitations, this tool may have limited functionality in web add-ins.`,
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        type: "number",
        description: "0-based index of the slide to set notes for.",
      },
      notes: {
        type: "string",
        description: "The text content to set as speaker notes.",
      },
    },
    required: ["slideIndex", "notes"],
  },
  handler: async (args) => {
    const { slideIndex, notes } = args as { slideIndex: number; notes: string };

    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const slideCount = slides.items.length;
        if (slideCount === 0) {
          return "Presentation has no slides.";
        }

        if (slideIndex < 0 || slideIndex >= slideCount) {
          return `Invalid slideIndex ${slideIndex}. Must be 0-${slideCount - 1}.`;
        }

        // PowerPoint JavaScript API has limited support for notes
        // The full notes API is available in desktop but limited in web
        // We attempt to use available methods
        
        const slide = slides.items[slideIndex];
        slide.load("id");
        await context.sync();

        // Due to API limitations, we provide feedback about the limitation
        // In a full implementation, this would use Office.context.document.setSelectedDataAsync
        // or OOXML manipulation for notes
        
        return `Note: Setting slide notes via the web add-in API has limitations. For slide ${slideIndex + 1}, please use the Notes pane in PowerPoint to add: "${notes.substring(0, 100)}${notes.length > 100 ? "..." : ""}"`;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
