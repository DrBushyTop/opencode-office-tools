import type { Tool } from "./types";
import {
  isPowerPointRequirementSetSupported,
  loadShapeSummaries,
  readOfficeValue,
  supportsPowerPointPlaceholders,
  toolFailure,
} from "./powerpointShared";
import {
  loadPresentationLayoutCatalogFromDocument,
  lookupPresentationLayoutMetadata,
  resolveSlideLayoutMetadata,
  type PresentationLayoutCatalog,
} from "./powerpointLayoutCatalog";
import { z } from "zod";

const getSlideLayoutDetailsArgsSchema = z.object({
  layoutId: z.string().min(1),
  slideMasterId: z.string().optional(),
});

interface LayoutPlaceholderDetail {
  index: number;
  shapeId: string;
  placeholderType?: string;
  placeholderContainedType?: string | null;
  left: number;
  top: number;
  width: number;
  height: number;
}

function summarizePlaceholders(placeholders: LayoutPlaceholderDetail[], placeholderMetadataSupported: boolean) {
  if (!placeholderMetadataSupported) return "Placeholder metadata is unavailable on this host.";
  if (!placeholders.length) return "No placeholders detected.";

  const counts = new Map<string, number>();
  for (const placeholder of placeholders) {
    const key = placeholder.placeholderType || "Unknown";
    counts.set(key, (counts.get(key) || 0) + 1);
  }

  return Array.from(counts.entries())
    .map(([type, count]) => `${count} ${type}`)
    .join(", ");
}

export const getSlideLayoutDetails: Tool = {
  name: "get_slide_layout_details",
  description: "Inspect one slide layout and return its resolved name, type, and placeholder geometry.",
  parameters: {
    type: "object",
    properties: {
      layoutId: {
        type: "string",
        description: "Layout id from list_slide_layouts.",
      },
      slideMasterId: {
        type: "string",
        description: "Optional slide master id to disambiguate when the same layout id appears under multiple masters.",
      },
    },
    required: ["layoutId"],
  },
  handler: async (args) => {
    const parsedArgs = getSlideLayoutDetailsArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }

    const { layoutId, slideMasterId } = parsedArgs.data;

    let openXmlCatalog: PresentationLayoutCatalog | null = null;
    try {
      openXmlCatalog = await loadPresentationLayoutCatalogFromDocument();
    } catch {
      openXmlCatalog = null;
    }

    try {
      return await PowerPoint.run(async (context) => {
        const slideMasters = context.presentation.slideMasters;
        const supportsLayoutType = isPowerPointRequirementSetSupported("1.8");
        slideMasters.load("items");
        await context.sync();

        for (const master of slideMasters.items) {
          master.load(supportsLayoutType ? ["id", "name", "layouts/items/id", "layouts/items/name", "layouts/items/type"] : ["id", "name", "layouts/items/id", "layouts/items/name"]);
        }
        await context.sync();

        const matches: Array<{ master: PowerPoint.SlideMaster; layout: PowerPoint.SlideLayout; masterIndex: number; layoutIndex: number }> = [];
        for (const [masterIndex, master] of slideMasters.items.entries()) {
          const currentMasterId = readOfficeValue(() => master.id, "");
          if (slideMasterId && currentMasterId !== slideMasterId) continue;

          for (const [layoutIndex, layout] of master.layouts.items.entries()) {
            if (readOfficeValue(() => layout.id, "") === layoutId) {
              matches.push({ master, layout, masterIndex, layoutIndex });
            }
          }
        }

        if (!matches.length) {
          return toolFailure(`Slide layout ${JSON.stringify(layoutId)} was not found.`);
        }
        if (matches.length > 1 && !slideMasterId) {
          const available = matches
            .map((match) => readOfficeValue(() => match.master.id, "(missing)"))
            .join(", ");
          return toolFailure(`Slide layout ${JSON.stringify(layoutId)} is ambiguous. Retry with slideMasterId. Matching masters: ${available}.`);
        }

        const target = matches[0];
        const placeholderMetadataSupported = supportsPowerPointPlaceholders();
        target.layout.load(supportsLayoutType ? ["id", "name", "type"] : ["id", "name"]);
        target.layout.shapes.load("items");
        await context.sync();

        const shapeSummaries = await loadShapeSummaries(context, target.layout.shapes.items, {
          includeText: false,
          includeFormatting: false,
          includeTableValues: false,
        });
        const placeholders = shapeSummaries
          .filter((shape) => placeholderMetadataSupported && Boolean(shape.placeholderType))
          .map<LayoutPlaceholderDetail>((shape, index) => ({
            index,
            shapeId: shape.id,
            placeholderType: shape.placeholderType,
            placeholderContainedType: shape.placeholderContainedType,
            left: shape.left,
            top: shape.top,
            width: shape.width,
            height: shape.height,
          }));

        const resolvedMasterId = readOfficeValue(() => target.master.id, "(missing)");
        const fallback = lookupPresentationLayoutMetadata(openXmlCatalog, {
          slideMasterId: resolvedMasterId,
          layoutId,
          masterIndex: target.masterIndex,
          layoutIndex: target.layoutIndex,
        });
        const resolvedMasterName = readOfficeValue(() => target.master.name, "") || fallback?.slideMasterName || "";
        const resolvedLayout = resolveSlideLayoutMetadata(
          readOfficeValue(() => target.layout.name, ""),
          readOfficeValue(() => String(target.layout.type), ""),
          fallback,
        );

        return {
          resultType: "success",
          slideMasterId: resolvedMasterId,
          slideMasterName: resolvedMasterName,
          layoutId,
          layoutName: resolvedLayout.layoutName,
          layoutType: resolvedLayout.layoutType,
          placeholderMetadataSupported,
          placeholderCount: placeholders.length,
          placeholders,
          toolTelemetry: {
            placeholderCount: placeholders.length,
          },
        };
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
