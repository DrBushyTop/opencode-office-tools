import type { Tool } from "./types";
import { isPowerPointRequirementSetSupported, readOfficeValue, toolFailure } from "./powerpointShared";
import {
  loadPresentationLayoutCatalogFromDocument,
  lookupPresentationLayoutMetadata,
  resolveSlideLayoutMetadata,
  type PresentationLayoutCatalog,
} from "./powerpointLayoutCatalog";
import { z } from "zod";

const listSlideLayoutsArgsSchema = z.object({}).strict();

interface SlideLayoutOverviewEntry {
  slideMasterId: string;
  slideMasterName: string;
  layoutId: string;
  layoutName: string;
  layoutType: string;
}

interface SlideMasterLayoutOverview {
  slideMasterId: string;
  slideMasterName: string;
  layoutCount: number;
  layouts: SlideLayoutOverviewEntry[];
}

function buildSummary(slideMasters: SlideMasterLayoutOverview[], layoutCount: number) {
  const lines = [`Found ${layoutCount} layout${layoutCount === 1 ? "" : "s"} across ${slideMasters.length} slide master${slideMasters.length === 1 ? "" : "s"}.`];

  for (const master of slideMasters) {
    lines.push(`Master ${JSON.stringify(master.slideMasterName)} (${master.slideMasterId}), ${master.layoutCount} layout${master.layoutCount === 1 ? "" : "s"}:`);
    for (const layout of master.layouts) {
      lines.push(`- ${JSON.stringify(layout.layoutName)} (${layout.layoutId}, ${layout.layoutType})`);
    }
  }

  return lines.join("\n");
}

export const listSlideLayouts: Tool = {
  name: "list_slide_layouts",
  description: "List slide masters and layouts as a concise overview with ids, names, and types.",
  parameters: {
    type: "object",
    properties: {},
  },
  handler: async (args) => {
    const parsedArgs = listSlideLayoutsArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }

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

        for (const master of slideMasters.items) {
          for (const layout of master.layouts.items) {
            layout.load(supportsLayoutType ? ["id", "name", "type"] : ["id", "name"]);
          }
        }
        await context.sync();

        const slideMasterOverviews: SlideMasterLayoutOverview[] = [];
        const layouts: SlideLayoutOverviewEntry[] = [];

        for (const [masterIndex, master] of slideMasters.items.entries()) {
          const slideMasterId = readOfficeValue(() => master.id, "(missing)");
          const fallbackMasterName = lookupPresentationLayoutMetadata(openXmlCatalog, { slideMasterId, masterIndex })?.slideMasterName || "";
          const slideMasterName = readOfficeValue(() => master.name, "") || fallbackMasterName;

          const masterLayouts = master.layouts.items.map<SlideLayoutOverviewEntry>((layout, layoutIndex) => {
            const layoutId = readOfficeValue(() => layout.id, "(missing)");
            const fallback = lookupPresentationLayoutMetadata(openXmlCatalog, {
              slideMasterId,
              layoutId,
              masterIndex,
              layoutIndex,
            });
            const resolved = resolveSlideLayoutMetadata(
              readOfficeValue(() => layout.name, ""),
              readOfficeValue(() => String(layout.type), ""),
              fallback,
            );

            return {
              slideMasterId,
              slideMasterName,
              layoutId,
              layoutName: resolved.layoutName,
              layoutType: resolved.layoutType,
            };
          });

          slideMasterOverviews.push({
            slideMasterId,
            slideMasterName,
            layoutCount: masterLayouts.length,
            layouts: masterLayouts,
          });
          layouts.push(...masterLayouts);
        }

        const slideMasterCount = slideMasters.items.length;
        const layoutCount = layouts.length;

        return {
          resultType: "success",
          textResultForLlm: buildSummary(slideMasterOverviews, layoutCount),
          slideMasterCount,
          layoutCount,
          slideMasters: slideMasterOverviews,
          layouts,
          toolTelemetry: {
            slideMasterCount,
            layoutCount,
          },
        };
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
