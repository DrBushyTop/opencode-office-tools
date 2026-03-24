import type { Tool } from "@github/copilot-sdk";
import pptxgen from "pptxgenjs";

export const addSlideFromCode: Tool = {
  name: "add_slide_from_code",
  description: `Add a slide to the PowerPoint presentation by providing PptxGenJS code.

Your code should be a function body that receives a 'slide' object and calls methods on it to add content.
The slide will be automatically created and inserted into the presentation.

Function signature:
  function(slide) { /* your code here */ }

PptxGenJS API Examples:

1. Add text:
   slide.addText("Hello World", { x: 1, y: 1, w: 8, h: 1, fontSize: 24, bold: true, color: "363636" });

2. Add text with bullets:
   slide.addText([
     { text: "First point", options: { bullet: true } },
     { text: "Second point", options: { bullet: true } },
     { text: "Sub-point", options: { bullet: true, indentLevel: 1 } }
   ], { x: 0.5, y: 1.5, w: 9, h: 3, fontSize: 18 });

3. Add a table:
   slide.addTable([
     [{ text: "Header 1", options: { bold: true, fill: "0088CC", color: "FFFFFF" } }, { text: "Header 2", options: { bold: true, fill: "0088CC", color: "FFFFFF" } }],
     ["Row 1 Cell 1", "Row 1 Cell 2"],
     ["Row 2 Cell 1", "Row 2 Cell 2"]
   ], { x: 0.5, y: 2, w: 9, fontSize: 14, border: { pt: 1, color: "CFCFCF" } });

4. Add a shape:
   slide.addShape("rect", { x: 1, y: 1, w: 3, h: 1, fill: "FF0000" });
   slide.addShape("ellipse", { x: 5, y: 1, w: 2, h: 2, fill: "00FF00" });

5. Add an image from URL:
   slide.addImage({ path: "https://example.com/image.png", x: 1, y: 1, w: 4, h: 3 });

6. Positioning and sizing (all values in inches):
   - x: distance from left edge
   - y: distance from top edge  
   - w: width
   - h: height

7. Text formatting options:
   - fontSize: number (points)
   - bold: true/false
   - italic: true/false
   - underline: true/false
   - color: "RRGGBB" (hex without #)
   - align: "left" | "center" | "right"
   - valign: "top" | "middle" | "bottom"
   - fontFace: "Arial", "Calibri", etc.

8. Complete slide example:
   // Title
   slide.addText("Quarterly Report", { x: 0.5, y: 0.3, w: 9, h: 0.8, fontSize: 32, bold: true, color: "003366" });
   // Subtitle
   slide.addText("Q3 2024 Results", { x: 0.5, y: 1, w: 9, h: 0.5, fontSize: 18, color: "666666" });
   // Bullet points
   slide.addText([
     { text: "Revenue increased 25% YoY", options: { bullet: true } },
     { text: "Customer base grew to 10,000", options: { bullet: true } },
     { text: "Launched 3 new products", options: { bullet: true } }
   ], { x: 0.5, y: 1.8, w: 9, h: 2.5, fontSize: 16 });
`,
  parameters: {
    type: "object",
    properties: {
      code: {
        type: "string",
        description: "JavaScript code (function body) that receives a 'slide' parameter and calls PptxGenJS methods to build the slide content.",
      },
      replaceSlideIndex: {
        type: "number",
        description: "Optional 0-based index of an existing slide to replace. If provided, the slide at this index will be deleted and the new slide inserted in its place. If not provided, the new slide is appended at the end.",
      },
    },
    required: ["code"],
  },
  handler: async (args) => {
    const { code, replaceSlideIndex } = args as { code: string; replaceSlideIndex?: number };

    try {
      // Create presentation and slide
      const pptx = new pptxgen();
      const slide = pptx.addSlide();

      // Execute the provided code with slide in scope
      try {
        const buildSlide = new Function("slide", code);
        buildSlide(slide);
      } catch (codeError: any) {
        return {
          textResultForLlm: `Code execution error: ${codeError.message}\n\nStack: ${codeError.stack}`,
          resultType: "failure",
          error: codeError.message,
          toolTelemetry: {},
        };
      }

      // Generate base64
      let base64: string;
      try {
        base64 = await pptx.write({ outputType: "base64" }) as string;
      } catch (writeError: any) {
        return {
          textResultForLlm: `Failed to generate presentation: ${writeError.message}`,
          resultType: "failure",
          error: writeError.message,
          toolTelemetry: {},
        };
      }

      // Insert into PowerPoint (replace existing or append)
      try {
        await PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();

          const slideCount = slides.items.length;

          // Validate replaceSlideIndex if provided
          if (replaceSlideIndex !== undefined) {
            if (replaceSlideIndex < 0 || replaceSlideIndex >= slideCount) {
              throw new Error(`Invalid replaceSlideIndex ${replaceSlideIndex}. Must be 0-${slideCount - 1} (current slide count: ${slideCount})`);
            }
          }

          const insertOptions: PowerPoint.InsertSlideOptions = {
            formatting: PowerPoint.InsertSlideFormatting.useDestinationTheme,
          };

          if (replaceSlideIndex !== undefined) {
            // Insert before the slide we're replacing, then delete the old one
            if (replaceSlideIndex > 0) {
              const prevSlide = slides.items[replaceSlideIndex - 1];
              prevSlide.load("id");
              await context.sync();
              insertOptions.targetSlideId = prevSlide.id;
            }
            // If replaceSlideIndex is 0, don't set targetSlideId (insert at beginning)

            context.presentation.insertSlidesFromBase64(base64, insertOptions);
            await context.sync();

            // Reload slides to get the updated list
            slides.load("items");
            await context.sync();

            // Delete the old slide (now at replaceSlideIndex + 1 since we inserted before it)
            const oldSlideIndex = replaceSlideIndex + 1;
            if (oldSlideIndex < slides.items.length) {
              slides.items[oldSlideIndex].delete();
              await context.sync();
            }
          } else {
            // Append at the end (original behavior)
            if (slides.items.length > 0) {
              const lastSlide = slides.items[slides.items.length - 1];
              lastSlide.load("id");
              await context.sync();
              insertOptions.targetSlideId = lastSlide.id;
            }

            context.presentation.insertSlidesFromBase64(base64, insertOptions);
            await context.sync();
          }
        });
      } catch (insertError: any) {
        return {
          textResultForLlm: `Failed to insert slide: ${insertError.message}`,
          resultType: "failure",
          error: insertError.message,
          toolTelemetry: {},
        };
      }

      return replaceSlideIndex !== undefined 
        ? `Successfully replaced slide ${replaceSlideIndex + 1} in the presentation.`
        : "Successfully added new slide to the presentation.";
    } catch (e: any) {
      return {
        textResultForLlm: `Unexpected error: ${e.message}\n\nStack: ${e.stack}`,
        resultType: "failure",
        error: e.message,
        toolTelemetry: {},
      };
    }
  },
};
