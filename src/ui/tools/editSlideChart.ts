import type { Tool } from "./types";
import { buildPowerPointShapeRef, parsePowerPointShapeRef } from "./powerpointShapeRefs";
import { resolvePowerPointTargetingArgs } from "./powerpointContext";
import { getSlideById } from "./powerpointNativeContent";
import {
  createChartInBase64Presentation,
  deleteChartInBase64Presentation,
  slideChartDefinitionSchema,
  slideChartLegendPositionSchema,
  slideChartTypeSchema,
  type SlideChartMutationResult,
  updateChartInBase64Presentation,
} from "./powerpointChartXml";
import { replaceSlideWithMutatedOpenXml } from "./powerpointOpenXml";
import { roundTripRefreshHint, shouldAddRoundTripShapeTargetRefreshHint, toolFailure } from "./powerpointShared";
import { z } from "zod";

const editSlideChartArgsSchema = z.object({
  action: z.enum(["create", "update", "delete"]),
  slideIndex: z.number().optional(),
  ref: z.string().optional(),
  chartType: slideChartTypeSchema.optional(),
  title: z.string().optional(),
  categories: z.array(z.string()).optional(),
  series: z.array(z.object({
    name: z.string(),
    values: z.array(z.number()),
  })).optional(),
  stacked: z.boolean().optional(),
  left: z.number().optional(),
  top: z.number().optional(),
  width: z.number().optional(),
  height: z.number().optional(),
  fontColor: z.string().optional(),
  showDataLabels: z.boolean().optional(),
  showLegend: z.boolean().optional(),
  legendPosition: slideChartLegendPositionSchema.optional(),
});

type EditSlideChartArgs = z.infer<typeof editSlideChartArgsSchema>;

function buildDefinition(args: EditSlideChartArgs) {
  return slideChartDefinitionSchema.parse({
    chartType: args.chartType,
    title: args.title,
    categories: args.categories,
    series: args.series,
    stacked: args.stacked,
    left: args.left,
    top: args.top,
    width: args.width,
    height: args.height,
    fontColor: args.fontColor,
    showDataLabels: args.showDataLabels,
    showLegend: args.showLegend,
    legendPosition: args.legendPosition,
  });
}

export const editSlideChart: Tool = {
  name: "edit_slide_chart",
  description: "Create, update, or delete a real PowerPoint OOXML chart on one slide through a slide-scoped Open XML round-trip.",
  parameters: {
    type: "object",
    properties: {
      action: { type: "string", enum: ["create", "update", "delete"] },
      slideIndex: { type: "number", description: "0-based slide index for create. Defaults to the active slide when available." },
      ref: { type: "string", description: "Stable chart ref in the form slide-id:<slideId>/shape:<xmlShapeId> for update or delete." },
      chartType: { type: "string", enum: slideChartTypeSchema.options },
      title: { type: "string" },
      categories: { type: "array", items: { type: "string" } },
      series: {
        type: "array",
        items: {
          type: "object",
          properties: {
            name: { type: "string" },
            values: { type: "array", items: { type: "number" } },
          },
          required: ["name", "values"],
        },
      },
      stacked: { type: "boolean" },
      left: { type: "number" },
      top: { type: "number" },
      width: { type: "number" },
      height: { type: "number" },
      fontColor: { type: "string", description: "Hex color (e.g. \"FFFFFF\" or \"#FFFFFF\") for all chart text including title, axes, legend, and data labels." },
      showDataLabels: { type: "boolean", description: "Whether to show data value labels on chart series. Defaults to true." },
      showLegend: { type: "boolean", description: "Whether to show the chart legend. Defaults to true." },
      legendPosition: { type: "string", enum: slideChartLegendPositionSchema.options, description: "Legend placement. Defaults to \"top\"." },
    },
    required: ["action"],
  },
  handler: async (args) => {
    const parsedArgs = editSlideChartArgsSchema.safeParse(args);
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }

    const chartArgs = resolvePowerPointTargetingArgs(parsedArgs.data as EditSlideChartArgs);

    if (chartArgs.action === "create") {
      if (!Number.isInteger(chartArgs.slideIndex) || (chartArgs.slideIndex as number) < 0) {
        return toolFailure("slideIndex must be a non-negative integer.");
      }
    } else if (!chartArgs.ref) {
      return toolFailure("ref is required for update and delete.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        if (chartArgs.action === "create") {
          const definition = buildDefinition(chartArgs);
          let mutation: SlideChartMutationResult | undefined;
          const roundTrip = await replaceSlideWithMutatedOpenXml(context, chartArgs.slideIndex as number, (base64) => {
            mutation = createChartInBase64Presentation(base64, definition);
            return mutation.base64;
          });
          if (!mutation) {
            throw new Error("Chart creation did not produce mutation metadata.");
          }

          const xmlShapeId = mutation.xmlShapeId;
          const ref = buildPowerPointShapeRef(roundTrip.replacementSlideId, xmlShapeId);
          return {
            resultType: "success",
            textResultForLlm: `Created a ${definition.chartType} chart on slide ${roundTrip.finalSlideIndex + 1}.`,
            slideIndex: roundTrip.finalSlideIndex,
            slideId: roundTrip.replacementSlideId,
            xmlShapeId,
            ref,
            toolTelemetry: {
              ...roundTrip,
              xmlShapeId,
              ref,
            },
          };
        }

        const parsedRef = parsePowerPointShapeRef(chartArgs.ref as string);
        const { slideIndex } = await getSlideById(context, parsedRef.slideId);

        if (chartArgs.action === "delete") {
          const roundTrip = await replaceSlideWithMutatedOpenXml(context, slideIndex, (base64) =>
            deleteChartInBase64Presentation(base64, parsedRef.xmlShapeId).base64,
          );
          return {
            resultType: "success",
            textResultForLlm: `Deleted chart ${parsedRef.xmlShapeId} from slide ${roundTrip.finalSlideIndex + 1}.`,
            slideIndex: roundTrip.finalSlideIndex,
            slideId: roundTrip.replacementSlideId,
            deletedTarget: parsedRef,
            toolTelemetry: {
              ...roundTrip,
              deletedTarget: parsedRef,
            },
          };
        }

        const definition = buildDefinition(chartArgs);
        let mutation: SlideChartMutationResult | undefined;
        const roundTrip = await replaceSlideWithMutatedOpenXml(context, slideIndex, (base64) => {
          mutation = updateChartInBase64Presentation(base64, parsedRef.xmlShapeId, definition);
          return mutation.base64;
        });
        if (!mutation) {
          throw new Error("Chart update did not produce mutation metadata.");
        }

        const xmlShapeId = mutation.xmlShapeId;
        const ref = buildPowerPointShapeRef(roundTrip.replacementSlideId, xmlShapeId);
        return {
          resultType: "success",
          textResultForLlm: `Updated chart ${xmlShapeId} on slide ${roundTrip.finalSlideIndex + 1}.`,
          slideIndex: roundTrip.finalSlideIndex,
          slideId: roundTrip.replacementSlideId,
          xmlShapeId,
          ref,
          toolTelemetry: {
            ...roundTrip,
            xmlShapeId,
            ref,
            previousRef: parsedRef.ref,
          },
        };
      });
    } catch (error: unknown) {
      return toolFailure(error, shouldAddRoundTripShapeTargetRefreshHint(error) ? roundTripRefreshHint() : undefined);
    }
  },
};
