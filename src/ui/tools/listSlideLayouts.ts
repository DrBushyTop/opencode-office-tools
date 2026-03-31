import type { Tool } from "./types";
import { loadShapeSummaries, readOfficeValue, supportsPowerPointPlaceholders, toolFailure } from "./powerpointShared";
import { z } from "zod";

const listSlideLayoutsArgsSchema = z.object({}).strict();

interface LayoutPlaceholderCatalogEntry {
  shapeId: string;
  placeholderName: string;
  placeholderType?: string;
  placeholderContainedType?: string | null;
  text?: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

interface SlideLayoutCatalogEntry {
  slideMasterId: string;
  slideMasterName: string;
  layoutId: string;
  layoutName: string;
  layoutType: string;
  placeholders: LayoutPlaceholderCatalogEntry[];
}

function buildSummary(layouts: SlideLayoutCatalogEntry[], masterCount: number, placeholderMetadataSupported: boolean) {
  const lines = [`Found ${layouts.length} layout${layouts.length === 1 ? "" : "s"} across ${masterCount} slide master${masterCount === 1 ? "" : "s"}.`];
  if (!placeholderMetadataSupported) {
    lines.push("Placeholder metadata is unavailable on this host.");
  }

  for (const layout of layouts) {
    const placeholderSummary = !placeholderMetadataSupported
      ? "(placeholder metadata unavailable on this host)"
      : layout.placeholders.length > 0
      ? layout.placeholders.map((placeholder) => `${placeholder.placeholderType || "(unknown)"}:${JSON.stringify(placeholder.placeholderName)}`).join(", ")
      : "(no placeholders detected)";
    lines.push(`- ${JSON.stringify(layout.layoutName)} (${layout.layoutId}, ${layout.layoutType}) on ${JSON.stringify(layout.slideMasterName)}: ${placeholderSummary}`);
  }

  return lines.join("\n");
}

export const listSlideLayouts: Tool = {
  name: "list_slide_layouts",
  description: "List available slide layouts in a direct catalog with slide master ids, layout ids, names, types, and placeholder inventory.",
  parameters: {
    type: "object",
    properties: {},
  },
  handler: async (args) => {
    const parsedArgs = listSlideLayoutsArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const slideMasters = context.presentation.slideMasters;
        slideMasters.load("items");
        await context.sync();

        for (const master of slideMasters.items) {
          master.load(["id", "name", "layouts/items/id", "layouts/items/name", "layouts/items/type"]);
        }
        await context.sync();

        const placeholderMetadataSupported = supportsPowerPointPlaceholders();
        const layouts: SlideLayoutCatalogEntry[] = [];

        for (const master of slideMasters.items) {
          for (const layout of master.layouts.items) {
            layout.shapes.load("items");
          }
          await context.sync();

          for (const layout of master.layouts.items) {
            const shapeSummaries = await loadShapeSummaries(context, layout.shapes.items, {
              includeText: true,
              includeFormatting: false,
              includeTableValues: false,
            });
            const placeholders = shapeSummaries
              .filter((shape) => placeholderMetadataSupported && Boolean(shape.placeholderType))
              .map<LayoutPlaceholderCatalogEntry>((shape) => ({
                shapeId: shape.id,
                placeholderName: shape.name,
                placeholderType: shape.placeholderType,
                placeholderContainedType: shape.placeholderContainedType,
                text: shape.text,
                left: shape.left,
                top: shape.top,
                width: shape.width,
                height: shape.height,
              }));

            layouts.push({
              slideMasterId: readOfficeValue(() => master.id, "(missing)"),
              slideMasterName: readOfficeValue(() => master.name, ""),
              layoutId: readOfficeValue(() => layout.id, "(missing)"),
              layoutName: readOfficeValue(() => layout.name, ""),
              layoutType: readOfficeValue(() => String(layout.type), "Unknown"),
              placeholders,
            });
          }
        }

        const slideMasterCount = slideMasters.items.length;
        const summary = buildSummary(layouts, slideMasterCount, placeholderMetadataSupported);

        return {
          resultType: "success",
          textResultForLlm: summary,
          slideMasterCount,
          layoutCount: layouts.length,
          placeholderMetadataSupported,
          layouts,
          toolTelemetry: {
            slideMasterCount,
            layoutCount: layouts.length,
          },
        };
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
