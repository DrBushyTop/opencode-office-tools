import type { Tool } from "@github/copilot-sdk";

export const setPresentationContent: Tool = {
  name: "set_presentation_content",
  description: "Add a text box to a slide in the PowerPoint presentation. To add a new slide, pass slideIndex equal to the current slide count (0-based indexing, so if there are 3 slides, pass slideIndex=3 to add a 4th slide).",
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        type: "number",
        description: "0-based slide index. Use 0 for first slide, 1 for second, etc. To add a new slide, use the current slide count as the index.",
      },
      text: {
        type: "string",
        description: "The text content to add to the slide.",
      },
    },
    required: ["slideIndex", "text"],
  },
  handler: async (args) => {
    const { slideIndex, text } = args as { slideIndex: number; text: string };
    
    try {
      return await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const slideCount = slides.items.length;

        if (slideIndex < 0 || slideIndex > slideCount) {
          return { 
            textResultForLlm: `Invalid slideIndex ${slideIndex}. Must be 0-${slideCount} (current slide count: ${slideCount})`, 
            resultType: "failure", 
            error: "Invalid slideIndex", 
            toolTelemetry: {} 
          };
        }

        // Add new slide if index equals current count
        if (slideIndex === slideCount) {
          context.presentation.slides.add();
          await context.sync();
          slides.load("items");
          await context.sync();
        }

        const slide = slides.items[slideIndex];
        
        slide.shapes.addTextBox(text, {
          left: 50,
          top: 100,
          width: 600,
          height: 400,
        });
        await context.sync();

        return `Added text to slide ${slideIndex + 1}`;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
