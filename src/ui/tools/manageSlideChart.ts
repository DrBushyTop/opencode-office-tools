import type { Tool } from "./types";
import { resolvePowerPointTargetingArgs } from "./powerpointContext";
import { defaultChartPalette, getShapeBounds, getSlideByIndex } from "./powerpointNativeContent";
import { resolveSlideShapeByIdWithXmlFallback } from "./powerpointShapeTarget";
import { normalizeHexColor, toolFailure } from "./powerpointShared";

type ChartType = "column" | "bar" | "line" | "pie";
type ManageSlideChartAction = "create" | "update" | "delete";

interface SlideChartPoint {
  label: string;
  value: number;
  color?: string;
}

interface ManageSlideChartArgs {
  action: ManageSlideChartAction;
  slideIndex?: number;
  shapeId?: string | number;
  chartType?: ChartType;
  title?: string;
  data?: SlideChartPoint[];
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  name?: string;
}

function validData(data?: SlideChartPoint[]) {
  return Boolean(data?.length) && data!.every((item) => typeof item.label === "string" && Number.isFinite(item.value));
}

async function buildChartShapes(
  slide: PowerPoint.Slide,
  chartType: ChartType,
  data: SlideChartPoint[],
  frame: { left: number; top: number; width: number; height: number; title?: string; name?: string },
) {
  const created: PowerPoint.Shape[] = [];
  const titleHeight = frame.title ? 28 : 0;
  const plotTop = frame.top + titleHeight + 8;
  const plotHeight = Math.max(frame.height - titleHeight - 24, 60);
  const maxValue = Math.max(...data.map((item) => item.value), 1);

  if (frame.title) {
    const titleShape = slide.shapes.addTextBox(frame.title, { left: frame.left, top: frame.top, width: frame.width, height: titleHeight });
    titleShape.textFrame.textRange.font.bold = true;
    titleShape.textFrame.textRange.font.size = 18;
    titleShape.lineFormat.visible = false;
    created.push(titleShape);
  }

  if (chartType === "column") {
    const gap = 12;
    const labelHeight = 18;
    const barWidth = (frame.width - gap * (data.length - 1)) / data.length;
    data.forEach((item, index) => {
      const barHeight = Math.max((item.value / maxValue) * (plotHeight - labelHeight - 10), 8);
      const left = frame.left + index * (barWidth + gap);
      const top = plotTop + (plotHeight - labelHeight - barHeight);
      const bar = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, { left, top, width: barWidth, height: barHeight });
      bar.fill.setSolidColor(normalizeHexColor(item.color || defaultChartPalette(index)));
      bar.lineFormat.visible = false;
      const label = slide.shapes.addTextBox(item.label, { left, top: plotTop + plotHeight - labelHeight, width: barWidth, height: labelHeight });
      label.textFrame.textRange.font.size = 10;
      label.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
      label.lineFormat.visible = false;
      created.push(bar, label);
    });
  } else if (chartType === "bar" || chartType === "pie") {
    const rowHeight = plotHeight / data.length;
    data.forEach((item, index) => {
      const labelWidth = chartType === "pie" ? frame.width * 0.34 : frame.width * 0.28;
      const barLeft = frame.left + labelWidth + 10;
      const barWidth = Math.max(((frame.width - labelWidth - 18) * item.value) / maxValue, 8);
      const y = plotTop + index * rowHeight;
      const label = slide.shapes.addTextBox(chartType === "pie" ? `${item.label} ${(item.value / data.reduce((sum, entry) => sum + entry.value, 0) * 100).toFixed(0)}%` : item.label, {
        left: frame.left,
        top: y,
        width: labelWidth,
        height: rowHeight - 6,
      });
      label.textFrame.textRange.font.size = 11;
      label.lineFormat.visible = false;
      const bar = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundRectangle, {
        left: barLeft,
        top: y + 5,
        width: barWidth,
        height: Math.max(rowHeight - 12, 10),
      });
      bar.fill.setSolidColor(normalizeHexColor(item.color || defaultChartPalette(index)));
      bar.lineFormat.visible = false;
      created.push(label, bar);
    });
  } else {
    const gap = data.length > 1 ? frame.width / (data.length - 1) : frame.width;
    const points: Array<{ x: number; y: number; label: string; color: string }> = [];
    data.forEach((item, index) => {
      points.push({
        x: frame.left + index * gap,
        y: plotTop + (plotHeight - 18) - (item.value / maxValue) * (plotHeight - 36),
        label: item.label,
        color: normalizeHexColor(item.color || defaultChartPalette(index)),
      });
    });
    for (let index = 0; index < points.length; index += 1) {
      const point = points[index];
      const dot = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, { left: point.x - 5, top: point.y - 5, width: 10, height: 10 });
      dot.fill.setSolidColor(point.color);
      dot.lineFormat.visible = false;
      const label = slide.shapes.addTextBox(point.label, { left: point.x - 28, top: plotTop + plotHeight - 18, width: 56, height: 18 });
      label.textFrame.textRange.font.size = 10;
      label.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
      label.lineFormat.visible = false;
      created.push(dot, label);
      if (index > 0) {
        const previous = points[index - 1];
        const line = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
          left: previous.x,
          top: previous.y,
          width: point.x - previous.x,
          height: point.y - previous.y,
        });
        line.lineFormat.color = point.color;
        line.lineFormat.weight = 2;
        created.push(line);
      }
    }
  }

  const group = slide.shapes.addGroup(created);
  if (frame.name) group.name = frame.name;
  group.load(["id", "name"]);
  return group;
}

export const manageSlideChart: Tool = {
  name: "manage_slide_chart",
  description: "Create, update, or delete editable PowerPoint chart-style business visuals built from native shapes.",
  parameters: {
    type: "object",
    properties: {
      action: { type: "string", enum: ["create", "update", "delete"] },
      slideIndex: { type: "number", description: "0-based slide index. Defaults to the active slide when available." },
      shapeId: { anyOf: [{ type: "string" }, { type: "number" }], description: "Existing chart group shape id for update or delete." },
      chartType: { type: "string", enum: ["column", "bar", "line", "pie"] },
      title: { type: "string" },
      data: {
        type: "array",
        items: {
          type: "object",
          properties: {
            label: { type: "string" },
            value: { type: "number" },
            color: { type: "string" },
          },
          required: ["label", "value"],
        },
      },
      left: { type: "number" },
      top: { type: "number" },
      width: { type: "number" },
      height: { type: "number" },
      name: { type: "string" },
    },
    required: ["action"],
  },
  handler: async (args) => {
    const chart = resolvePowerPointTargetingArgs(args as ManageSlideChartArgs);
    if (!Number.isInteger(chart.slideIndex) || (chart.slideIndex as number) < 0) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }
    if ((chart.action === "create" || chart.action === "update") && (!chart.chartType || !validData(chart.data))) {
      return toolFailure("chartType and data are required for create and update. Each data item needs label and numeric value.");
    }
    if ((chart.action === "update" || chart.action === "delete") && chart.shapeId === undefined) {
      return toolFailure("shapeId is required for update and delete.");
    }

    const slideIndex = chart.slideIndex as number;

    try {
      return await PowerPoint.run(async (context) => {
        const slide = await getSlideByIndex(context, slideIndex);

        if (chart.action === "create") {
          const group = await buildChartShapes(slide, chart.chartType!, chart.data!, {
            left: chart.left ?? 70,
            top: chart.top ?? 90,
            width: chart.width ?? 420,
            height: chart.height ?? 240,
            title: chart.title,
            name: chart.name || `${chart.chartType} chart`,
          });
          await context.sync();
          return {
            resultType: "success",
            textResultForLlm: `Created ${chart.chartType} chart ${group.id} on slide ${slideIndex + 1}.`,
            slideIndex,
            shapeId: group.id,
            toolTelemetry: { slideIndex, shapeId: group.id, chartType: chart.chartType },
          };
        }

        const resolved = await resolveSlideShapeByIdWithXmlFallback(context, slide, slideIndex, chart.shapeId!);
        if (chart.action === "delete") {
          resolved.shape.delete();
          await context.sync();
          return {
            resultType: "success",
            textResultForLlm: `Deleted chart ${resolved.shapeId} from slide ${slideIndex + 1}.`,
            slideIndex,
            shapeId: resolved.shapeId,
            toolTelemetry: { slideIndex, shapeId: resolved.shapeId },
          };
        }

        const bounds = await getShapeBounds(resolved.shape, context);
        resolved.shape.delete();
        const group = await buildChartShapes(slide, chart.chartType!, chart.data!, {
          left: chart.left ?? bounds.left,
          top: chart.top ?? bounds.top,
          width: chart.width ?? bounds.width,
          height: chart.height ?? bounds.height,
          title: chart.title,
          name: chart.name || bounds.name || `${chart.chartType} chart`,
        });
        await context.sync();
        return {
          resultType: "success",
          textResultForLlm: `Updated ${chart.chartType} chart ${resolved.shapeId} on slide ${slideIndex + 1}.`,
          slideIndex,
          shapeId: group.id,
          replacedShapeId: resolved.shapeId,
          toolTelemetry: { slideIndex, replacedShapeId: resolved.shapeId, shapeId: group.id, chartType: chart.chartType },
        };
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
