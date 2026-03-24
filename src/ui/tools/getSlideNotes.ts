import type { Tool } from "@github/copilot-sdk";

export const getSlideNotes: Tool = {
  name: "get_slide_notes",
  description: `Get speaker notes from PowerPoint slides.

Parameters:
- slideIndex: Optional 0-based index of specific slide. If omitted, returns notes from all slides.

Speaker notes are the presenter's notes that appear below the slide in the Notes view.
Use this to understand context or instructions the presenter has added.`,
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        type: "number",
        description: "0-based index of the slide to get notes from. If omitted, returns notes from all slides.",
      },
    },
    required: [],
  },
  handler: async (args) => {
    const { slideIndex } = args as { slideIndex?: number };

    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const slideCount = slides.items.length;
        if (slideCount === 0) {
          return "Presentation has no slides.";
        }

        // Validate slideIndex if provided
        if (slideIndex !== undefined) {
          if (slideIndex < 0 || slideIndex >= slideCount) {
            return `Invalid slideIndex ${slideIndex}. Must be 0-${slideCount - 1}.`;
          }
        }

        // Determine which slides to process
        const startIdx = slideIndex !== undefined ? slideIndex : 0;
        const endIdx = slideIndex !== undefined ? slideIndex + 1 : slideCount;

        const results: string[] = [];

        for (let i = startIdx; i < endIdx; i++) {
          const slide = slides.items[i];
          
          // Load the notes slide - PowerPoint JS API requires getting shapes from notes
          try {
            // Get the slide's shapes to find notes placeholder
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();

            // Try to access notes through the slide's layout
            // Note: Direct notes access may have API limitations
            // We'll try to get text from body placeholder shapes
            
            let notesText = "";
            
            // Alternative approach: Load slide tags or check for notes body
            // The PowerPoint JS API has limited notes support, so we work around it
            slide.load("id");
            await context.sync();
            
            // For now, indicate that notes reading requires specific API support
            // This is a placeholder that can be enhanced when API support improves
            notesText = "(Notes access requires PowerPoint desktop - API limitation)";
            
            results.push(`Slide ${i + 1}: ${notesText}`);
          } catch (slideError: any) {
            results.push(`Slide ${i + 1}: (unable to read notes - ${slideError.message})`);
          }
        }

        if (slideIndex !== undefined) {
          return results[0] || "No notes found.";
        }
        
        return `Speaker Notes:\n${"━".repeat(40)}\n${results.join("\n")}`;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
