import { z } from "zod";
import type { Tool } from "./types";
import { getWorksheet, nonNegativeFiniteNumberSchema, parseToolArgs, splitSheetQualifiedRange, toolFailure } from "./excelShared";

const supportedChartTypes = ["column", "bar", "line", "pie", "area", "scatter", "doughnut"] as const;

const manageChartArgsSchema = z.object({
  action: z.enum(["create", "setData", "setProperties", "activate", "delete"]),
  chartName: z.string().optional(),
  sheetName: z.string().optional(),
  dataRange: z.string().optional(),
  chartType: z.enum(supportedChartTypes).optional(),
  title: z.string().optional(),
  newName: z.string().optional(),
  left: nonNegativeFiniteNumberSchema("left must be a non-negative number.").optional(),
  top: nonNegativeFiniteNumberSchema("top must be a non-negative number.").optional(),
  width: nonNegativeFiniteNumberSchema("width must be a non-negative number.").optional(),
  height: nonNegativeFiniteNumberSchema("height must be a non-negative number.").optional(),
  positionStartCell: z.string().optional(),
  positionEndCell: z.string().optional(),
});

export const manageChart: Tool = {
  name: "manage_chart",
  description: "Create or update Excel charts, including placement, resizing, source data changes, title changes, type changes, activation, and deletion.",
  parameters: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: ["create", "setData", "setProperties", "activate", "delete"],
        description: "Chart operation to perform.",
      },
      chartName: {
        type: "string",
        description: "Existing chart name. Required for all actions except create.",
      },
      sheetName: {
        type: "string",
        description: "Worksheet name for create, placement, or data ranges when ranges are sheet-local.",
      },
      dataRange: {
        type: "string",
        description: "Source data range for create or setData, such as 'A1:D10' or 'Sales!A1:D10'.",
      },
      chartType: {
        type: "string",
        enum: [...supportedChartTypes],
        description: "Chart type for create or setProperties.",
      },
      title: {
        type: "string",
        description: "Optional chart title. Use an empty string to clear the title in setProperties.",
      },
      newName: {
        type: "string",
        description: "New chart name for create or setProperties.",
      },
      left: {
        type: "number",
        description: "Left position in points for setProperties.",
      },
      top: {
        type: "number",
        description: "Top position in points for setProperties.",
      },
      width: {
        type: "number",
        description: "Width in points for setProperties.",
      },
      height: {
        type: "number",
        description: "Height in points for setProperties.",
      },
      positionStartCell: {
        type: "string",
        description: "Top-left placement cell for create or setProperties.",
      },
      positionEndCell: {
        type: "string",
        description: "Optional bottom-right placement cell for create or setProperties.",
      },
    },
    required: ["action"],
  },
  handler: async (args) => {
    const parsedArgs = parseToolArgs(manageChartArgsSchema, args);
    if (!parsedArgs.success) return parsedArgs.failure;

    const { action, chartName, sheetName, dataRange, chartType, title, newName, left, top, width, height, positionStartCell, positionEndCell } = parsedArgs.data;

    if (positionEndCell && !positionStartCell) {
      return toolFailure("positionStartCell is required when positionEndCell is provided.");
    }

    try {
      return await Excel.run(async (context) => {
        const chartTypeMap: Record<string, Excel.ChartType> = {
          column: Excel.ChartType.columnClustered,
          bar: Excel.ChartType.barClustered,
          line: Excel.ChartType.line,
          pie: Excel.ChartType.pie,
          area: Excel.ChartType.area,
          scatter: Excel.ChartType.xyscatter,
          doughnut: Excel.ChartType.doughnut,
        };

        if (action === "create") {
          if (!dataRange) return toolFailure("dataRange is required for create.");

          const qualifiedRange = splitSheetQualifiedRange(dataRange);
          const targetSheet = await getWorksheet(context, qualifiedRange?.sheetName || sheetName);
          const targetRange = targetSheet.getRange(qualifiedRange?.rangeAddress || dataRange);
          targetRange.load(["address", "left", "top", "width"]);
          await context.sync();

          const requestedChartType = chartType ? chartTypeMap[chartType.toLowerCase()] : Excel.ChartType.columnClustered;
          const excelChartType = requestedChartType || Excel.ChartType.columnClustered;
          const chart = targetSheet.charts.add(excelChartType, targetRange, Excel.ChartSeriesBy.auto);

          if (newName) chart.name = newName;
          if (title !== undefined) {
            chart.title.text = title;
            chart.title.visible = title.length > 0;
          }

          if (positionStartCell) {
            chart.setPosition(positionStartCell, positionEndCell);
          } else {
            chart.left = targetRange.left + targetRange.width + 20;
            chart.top = targetRange.top;
            chart.width = width ?? 400;
            chart.height = height ?? 300;
          }

          chart.load(["name", "chartType", "width", "height"]);
          await context.sync();
          return `Created chart ${chart.name} (${chart.chartType}) from ${targetRange.address} on ${targetSheet.name}.`;
        }

        if (!chartName) return toolFailure("chartName is required for this action.");

        let chart: Excel.Chart | null = null;
        let chartWorksheet: Excel.Worksheet | null = null;
        const sheets = sheetName
          ? [await getWorksheet(context, sheetName)]
          : (() => {
            const collection = context.workbook.worksheets;
            collection.load("items/name");
            return collection;
          })();

        if (!Array.isArray(sheets)) {
          await context.sync();
        }

        const candidateSheets = Array.isArray(sheets) ? sheets : sheets.items;
        const candidates = candidateSheets.map((sheet) => ({
          sheet,
          chart: sheet.charts.getItemOrNullObject(chartName),
        }));
        for (const candidate of candidates) {
          candidate.chart.load(["isNullObject", "name", "chartType"]);
        }
        await context.sync();

        for (const candidate of candidates) {
          if (!(candidate.chart as Excel.Chart & { isNullObject?: boolean }).isNullObject) {
            chart = candidate.chart;
            chartWorksheet = candidate.sheet;
            break;
          }
        }

        if (!chart || !chartWorksheet) {
          return toolFailure(`Chart ${chartName} was not found${sheetName ? ` on ${sheetName}` : " in the workbook"}.`);
        }

        switch (action) {
          case "setData": {
            if (!dataRange) return toolFailure("dataRange is required for setData.");
            const qualifiedRange = splitSheetQualifiedRange(dataRange);
            const targetSheet = qualifiedRange?.sheetName
              ? await getWorksheet(context, qualifiedRange.sheetName)
              : chartWorksheet;
            const range = targetSheet.getRange(qualifiedRange?.rangeAddress || dataRange);
            range.load("address");
            chart.setData(range, Excel.ChartSeriesBy.auto);
            await context.sync();
            return `Updated chart ${chart.name} to use data range ${range.address}.`;
          }
          case "setProperties":
            if (chartType) {
              chart.chartType = chartTypeMap[chartType.toLowerCase()] || Excel.ChartType.columnClustered;
            }
            if (newName) chart.name = newName;
            if (title !== undefined) {
              chart.title.text = title;
              chart.title.visible = title.length > 0;
            }
            if (left !== undefined) chart.left = left;
            if (top !== undefined) chart.top = top;
            if (width !== undefined) chart.width = width;
            if (height !== undefined) chart.height = height;
            if (positionStartCell) chart.setPosition(positionStartCell, positionEndCell);
            await context.sync();
            return `Updated chart ${newName || chart.name}.`;
          case "activate":
            chart.activate();
            await context.sync();
            return `Activated chart ${chart.name}.`;
          case "delete":
            chart.delete();
            await context.sync();
            return `Deleted chart ${chart.name}.`;
          default:
            return toolFailure(`Unsupported action ${action}.`);
        }
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
