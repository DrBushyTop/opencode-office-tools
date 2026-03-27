import type { Tool } from "./types";
import {
  isPowerPointRequirementSetSupported,
  loadShapeSummaries,
  supportsPowerPointPlaceholders,
  summarizePlainText,
  toolFailure,
} from "./powerpointShared";

const themeColors = [
  "Dark1",
  "Light1",
  "Dark2",
  "Light2",
  "Accent1",
  "Accent2",
  "Accent3",
  "Accent4",
  "Accent5",
  "Accent6",
  "Hyperlink",
  "FollowedHyperlink",
] as const;

export const getPresentationStructure: Tool = {
  name: "get_presentation_structure",
  description: `Inspect PowerPoint slide masters, layouts, selection state, backgrounds, and footer-like placeholders.

Returns:
- slide master and layout ids/names
- placeholder inventory on masters and layouts
- theme colors and background info when supported
- selected slide ids and selected shape ids when supported`,
  parameters: {
    type: "object",
    properties: {},
  },
  handler: async () => {
    try {
      return await PowerPoint.run(async (context) => {
        const presentation = context.presentation;
        const slideMasters = presentation.slideMasters;
        const slides = presentation.slides;
        const supportsSelection = isPowerPointRequirementSetSupported("1.5");
        const supportsPlaceholders = supportsPowerPointPlaceholders();
        const supportsThemes = isPowerPointRequirementSetSupported("1.10");

        slideMasters.load("items");
        slides.load("items");
        await context.sync();

        const selectedSlides = supportsSelection ? presentation.getSelectedSlides() : null;
        const selectedShapes = supportsSelection ? presentation.getSelectedShapes() : null;
        if (selectedSlides) selectedSlides.load("items/id");
        if (selectedShapes) selectedShapes.load("items/id");

        for (const master of slideMasters.items) {
          master.load(["id", "name", "layouts/items/id", "layouts/items/name", "layouts/items/type"]);
          master.shapes.load("items");
          if (supportsThemes) {
            master.background.fill.load("type");
          }
        }
        await context.sync();

        const lines: string[] = [
          `Slides: ${slides.items.length}`,
          `Slide masters: ${slideMasters.items.length}`,
          `Placeholder metadata: ${supportsPlaceholders ? "supported" : "not supported on this host"}`,
        ];

        if (selectedSlides) {
          const selectedSlideIds = selectedSlides.items.map((slide) => slide.id || "(unloaded)");
          lines.push(`Selected slides: ${selectedSlideIds.length ? selectedSlideIds.join(", ") : "(none)"}`);
        }
        if (selectedShapes) {
          const selectedShapeIds = selectedShapes.items.map((shape) => shape.id || "(unloaded)");
          lines.push(`Selected shapes: ${selectedShapeIds.length ? selectedShapeIds.join(", ") : "(none)"}`);
        }

        for (const master of slideMasters.items) {
          const masterShapes = await loadShapeSummaries(context, master.shapes.items, { includeText: true, includeFormatting: false, includeTableValues: false });
          const masterPlaceholders = masterShapes.filter((shape) => Boolean(shape.placeholderType));
          lines.push("", `Master ${JSON.stringify(master.name)} (${master.id}):`);
          if (supportsThemes) {
            const colorValues = themeColors.map((color) => ({ color: String(color), value: master.themeColorScheme.getThemeColor(color) }));
            await context.sync();
            const colorPairs = colorValues.map((entry) => `${entry.color}=${entry.value.value}`);
            lines.push(`- Theme colors: ${colorPairs.join(", ")}`);
            lines.push(`- Master background fill: ${master.background.fill.type}`);
          }
          lines.push(`- Master placeholders: ${masterPlaceholders.length ? masterPlaceholders.map((shape) => `${shape.placeholderType}:${shape.id}`).join(", ") : "(none)"}`);

          for (const layout of master.layouts.items) {
            layout.shapes.load("items");
            if (supportsThemes) {
              layout.background.load(["areBackgroundGraphicsHidden", "isMasterBackgroundFollowed"]);
            }
          }
          await context.sync();

          for (const layout of master.layouts.items) {
            const layoutShapes = await loadShapeSummaries(context, layout.shapes.items, { includeText: true, includeFormatting: false, includeTableValues: false });
            const placeholders = layoutShapes.filter((shape) => Boolean(shape.placeholderType));
            lines.push(`  - Layout ${JSON.stringify(layout.name)} (${layout.id}, ${layout.type}): ${placeholders.length ? placeholders.map((shape) => `${shape.placeholderType}:${summarizePlainText(shape.text || shape.name, 40)}`).join(" | ") : "(no placeholders)"}`);
            if (supportsThemes) {
              lines.push(`    background follows master=${layout.background.isMasterBackgroundFollowed ? "yes" : "no"}, graphics hidden=${layout.background.areBackgroundGraphicsHidden ? "yes" : "no"}`);
            }
          }
        }

        return lines.join("\n");
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
