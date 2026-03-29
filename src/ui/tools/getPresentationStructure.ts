import type { Tool } from "./types";
import {
  isPowerPointRequirementSetSupported,
  loadShapeSummaries,
  loadThemeColors,
  parseColor,
  readOfficeValue,
  summarizePlainText,
  supportsPowerPointPlaceholders,
  toolFailure,
  type PowerPointThemeColors,
} from "./powerpointShared";
import { z } from "zod";

type StructureFormat = "summary" | "structured" | "both";

const getPresentationStructureArgsSchema = z.object({
  format: z.enum(["summary", "structured", "both"]).optional(),
});

type GetPresentationStructureArgs = z.infer<typeof getPresentationStructureArgsSchema>;

const presentationStructureShapeSummarySchema = z.object({
  id: z.string(),
  name: z.string(),
  type: z.string(),
  text: z.string().optional(),
  placeholderType: z.string().optional(),
  placeholderContainedType: z.string().nullable().optional(),
  left: z.number(),
  top: z.number(),
  width: z.number(),
  height: z.number(),
});

const presentationLayoutSummarySchema = z.object({
  id: z.string(),
  name: z.string(),
  type: z.string(),
  placeholders: z.array(presentationStructureShapeSummarySchema),
  background: z.object({
    followsMaster: z.boolean().optional(),
    graphicsHidden: z.boolean().optional(),
  }).optional(),
});

const presentationMasterSummarySchema = z.object({
  id: z.string(),
  name: z.string(),
  themeColors: z.custom<PowerPointThemeColors>().optional(),
  backgroundFillType: z.string().optional(),
  placeholders: z.array(presentationStructureShapeSummarySchema),
  layouts: z.array(presentationLayoutSummarySchema),
});

type PresentationStructureShapeSummary = z.infer<typeof presentationStructureShapeSummarySchema>;
type PresentationLayoutSummary = z.infer<typeof presentationLayoutSummarySchema>;
type PresentationMasterSummary = z.infer<typeof presentationMasterSummarySchema>;

function toShapeSummary(shape: Awaited<ReturnType<typeof loadShapeSummaries>>[number]): PresentationStructureShapeSummary {
  return {
    id: shape.id,
    name: shape.name,
    type: shape.type,
    text: shape.text,
    placeholderType: shape.placeholderType,
    placeholderContainedType: shape.placeholderContainedType,
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
  };
}

function buildSummaryText(data: {
  slideCount: number;
  slideDimensions?: { widthPoints: number; heightPoints: number; widthInches: number; heightInches: number };
  selectedSlideIds: string[];
  selectedShapeIds: string[];
  masters: PresentationMasterSummary[];
  supportsPlaceholders: boolean;
}) {
  const lines: string[] = [
    `Slides: ${data.slideCount}`,
    ...(data.slideDimensions
      ? [`Slide dimensions: ${data.slideDimensions.widthInches.toFixed(2)}" x ${data.slideDimensions.heightInches.toFixed(2)}" (${data.slideDimensions.widthPoints}pt x ${data.slideDimensions.heightPoints}pt)`]
      : []),
    `Slide masters: ${data.masters.length}`,
    `Placeholder metadata: ${data.supportsPlaceholders ? "supported" : "not supported on this host"}`,
    `Selected slides: ${data.selectedSlideIds.length ? data.selectedSlideIds.join(", ") : "(none)"}`,
    `Selected shapes: ${data.selectedShapeIds.length ? data.selectedShapeIds.join(", ") : "(none)"}`,
  ];

  for (const master of data.masters) {
    lines.push("", `Master ${JSON.stringify(master.name)} (${master.id}):`);
    if (master.themeColors) {
      const colorPairs = Object.entries(master.themeColors).map(([key, value]) => `${key}=${value}`);
      lines.push(`- Theme colors: ${colorPairs.join(", ")}`);
    }
    if (master.backgroundFillType) {
      lines.push(`- Master background fill: ${master.backgroundFillType}`);
    }
    lines.push(`- Master placeholders: ${master.placeholders.length ? master.placeholders.map((shape) => `${shape.placeholderType}:${shape.id}`).join(", ") : "(none)"}`);

    for (const layout of master.layouts) {
      const placeholderText = layout.placeholders.length
        ? layout.placeholders.map((shape) => `${shape.placeholderType}:${summarizePlainText(shape.text || shape.name, 40)}`).join(" | ")
        : "(no placeholders)";
      lines.push(`  - Layout ${JSON.stringify(layout.name)} (${layout.id}, ${layout.type}): ${placeholderText}`);
      if (layout.background) {
        lines.push(`    background follows master=${layout.background.followsMaster ? "yes" : "no"}, graphics hidden=${layout.background.graphicsHidden ? "yes" : "no"}`);
      }
    }
  }

  return lines.join("\n");
}

export const getPresentationStructure: Tool = {
  name: "get_presentation_structure",
  description: `Inspect PowerPoint slide masters, layouts, selection state, backgrounds, and footer-like placeholders.

Returns:
- slide master and layout ids/names
- placeholder inventory on masters and layouts
- theme colors and background info when supported
- selected slide ids and selected shape ids when supported
- optional structured metadata for template-aware generation`,
  parameters: {
    type: "object",
    properties: {
      format: {
        type: "string",
        enum: ["summary", "structured", "both"],
        description: "Response format. summary returns readable text, structured returns machine-usable metadata, both returns both.",
      },
    },
  },
  handler: async (args) => {
    try {
      return await PowerPoint.run(async (context) => {
        const parsedArgs = getPresentationStructureArgsSchema.safeParse(args ?? {});
        if (!parsedArgs.success) {
          return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
        }
        const { format = "summary" } = parsedArgs.data as GetPresentationStructureArgs;
        const presentation = context.presentation;
        const slideMasters = presentation.slideMasters;
        const slides = presentation.slides;
        const supportsSelection = isPowerPointRequirementSetSupported("1.5");
        const supportsPlaceholders = supportsPowerPointPlaceholders();
        const supportsThemes = isPowerPointRequirementSetSupported("1.10");

        slideMasters.load("items");
        slides.load("items/id");
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

        let slideDimensions: { widthPoints: number; heightPoints: number; widthInches: number; heightInches: number } | undefined;
        if (supportsThemes) {
          const pageSetup = presentation.pageSetup;
          pageSetup.load(["slideWidth", "slideHeight"]);
          await context.sync();
          slideDimensions = {
            widthPoints: pageSetup.slideWidth,
            heightPoints: pageSetup.slideHeight,
            widthInches: pageSetup.slideWidth / 72,
            heightInches: pageSetup.slideHeight / 72,
          };
        }

        const selectedSlideIds = selectedSlides ? selectedSlides.items.map((slide) => slide.id || "(unloaded)") : [];
        const selectedShapeIds = selectedShapes ? selectedShapes.items.map((shape) => shape.id || "(unloaded)") : [];

        const masters: PresentationMasterSummary[] = [];
        for (const master of slideMasters.items) {
          const masterShapes = await loadShapeSummaries(context, master.shapes.items, { includeText: true, includeFormatting: false, includeTableValues: false });
          const masterPlaceholders = masterShapes.filter((shape) => Boolean(shape.placeholderType)).map(toShapeSummary);
          const layouts: PresentationLayoutSummary[] = [];

          for (const layout of master.layouts.items) {
            layout.shapes.load("items");
            if (supportsThemes) {
              layout.background.load(["areBackgroundGraphicsHidden", "isMasterBackgroundFollowed"]);
            }
          }
          await context.sync();

          for (const layout of master.layouts.items) {
            const layoutShapes = await loadShapeSummaries(context, layout.shapes.items, { includeText: true, includeFormatting: false, includeTableValues: false });
            layouts.push({
              id: readOfficeValue(() => layout.id, "(missing)"),
              name: readOfficeValue(() => layout.name, ""),
              type: readOfficeValue(() => String(layout.type), "Unknown"),
              placeholders: layoutShapes.filter((shape) => Boolean(shape.placeholderType)).map(toShapeSummary),
              background: supportsThemes
                ? {
                    followsMaster: readOfficeValue(() => layout.background.isMasterBackgroundFollowed, undefined),
                    graphicsHidden: readOfficeValue(() => layout.background.areBackgroundGraphicsHidden, undefined),
                  }
                : undefined,
            });
          }

          masters.push({
            id: readOfficeValue(() => master.id, "(missing)"),
            name: readOfficeValue(() => master.name, ""),
            themeColors: supportsThemes ? await loadThemeColors(context, master) : undefined,
            backgroundFillType: supportsThemes ? readOfficeValue(() => String(master.background.fill.type), undefined) : undefined,
            placeholders: masterPlaceholders,
            layouts,
          });
        }

        const structure = {
          slideCount: slides.items.length,
          slideDimensions,
          selectedSlideIds,
          selectedShapeIds,
          activeSlideId: selectedSlideIds[0],
          activeSlideIndex: selectedSlideIds.length
            ? slides.items.findIndex((slide) => slide.id === selectedSlideIds[0])
            : undefined,
          masters,
        };

        const summary = buildSummaryText({
          slideCount: structure.slideCount,
          slideDimensions: structure.slideDimensions,
          selectedSlideIds,
          selectedShapeIds,
          masters,
          supportsPlaceholders,
        });

        if (format === "structured") {
          return {
            resultType: "success",
            textResultForLlm: summary,
            structure,
            toolTelemetry: {
              slideCount: structure.slideCount,
              masterCount: structure.masters.length,
              selectedSlideIds,
              selectedShapeIds,
            },
          };
        }

        if (format === "both") {
          return {
            resultType: "success",
            textResultForLlm: summary,
            summary,
            structure,
            toolTelemetry: {
              slideCount: structure.slideCount,
              masterCount: structure.masters.length,
              selectedSlideIds,
              selectedShapeIds,
            },
          };
        }

        return summary;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
