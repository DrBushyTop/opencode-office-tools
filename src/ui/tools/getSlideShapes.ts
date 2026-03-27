import type { Tool } from "./types";
import { formatShapeSummary, loadShapeSummaries, supportsPowerPointPlaceholders, toolFailure } from "./powerpointShared";

export const getSlideShapes: Tool = {
  name: "get_slide_shapes",
  description: `Inspect PowerPoint shapes on one slide or across the deck.

Returns shape indices, ids, names, types, positions, placeholder types, text previews, and table info so later edits can target the correct shape.`,
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        type: "number",
        description: "Optional 0-based slide index. Omit to inspect all slides.",
      },
      includeFormatting: {
        type: "boolean",
        description: "Include fill and line formatting details. Default true.",
      },
      includeTableValues: {
        type: "boolean",
        description: "Include table cell values for table shapes. Default false.",
      },
      detail: {
        type: "boolean",
        description: "Return full text and table data instead of previews. Default false.",
      },
    },
  },
  handler: async (args) => {
    const {
      slideIndex,
      includeFormatting = true,
      includeTableValues = false,
      detail = false,
    } = args as { slideIndex?: number; includeFormatting?: boolean; includeTableValues?: boolean; detail?: boolean };

    try {
      return await PowerPoint.run(async (context) => {
        const supportsPlaceholders = supportsPowerPointPlaceholders();
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        if (slides.items.length === 0) {
          return "Presentation has no slides.";
        }

        const targetSlides = slideIndex === undefined
          ? slides.items
          : [slides.items[slideIndex]].filter(Boolean);

        if (targetSlides.length === 0) {
          return toolFailure(`Invalid slideIndex ${slideIndex}.`);
        }

        for (const slide of targetSlides) {
          slide.load(["id", "index"]);
          slide.shapes.load("items");
        }
        await context.sync();

        const sections: string[] = [];
        if (!supportsPlaceholders) {
          sections.push("Placeholder metadata is unavailable on this host because it requires PowerPointApi 1.8.");
        }
        for (const slide of targetSlides) {
          const summaries = await loadShapeSummaries(context, slide.shapes.items, { includeText: true, includeFormatting, includeTableValues });
          sections.push([
            `Slide ${slide.index + 1} (${slide.id}) shapes: ${summaries.length}`,
            ...summaries.map((shape) => formatShapeSummary(shape, detail)),
          ].join("\n"));
        }

        return sections.join("\n\n");
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
