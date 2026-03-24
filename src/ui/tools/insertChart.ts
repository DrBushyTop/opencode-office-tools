import type { Tool } from "@github/copilot-sdk";

export const insertChart: Tool = {
  name: "insert_chart",
  description: `Create a chart from data in an Excel worksheet.

Parameters:
- dataRange: The range containing the data (e.g., "A1:D10", "Sheet1!B2:E20")
- chartType: Type of chart to create:
  - "column" (default), "bar", "line", "pie", "area", "scatter", "doughnut"
- title: Optional chart title
- sheetName: Optional worksheet to place the chart on. Defaults to active sheet.

The chart will be placed to the right of the data range.

Examples:
- Column chart from data: dataRange="A1:C10", chartType="column", title="Sales by Quarter"
- Pie chart: dataRange="A1:B5", chartType="pie", title="Market Share"
- Line chart for trends: dataRange="A1:E12", chartType="line", title="Monthly Revenue"`,
  parameters: {
    type: "object",
    properties: {
      dataRange: {
        type: "string",
        description: "The cell range containing the chart data (e.g., 'A1:D10').",
      },
      chartType: {
        type: "string",
        enum: ["column", "bar", "line", "pie", "area", "scatter", "doughnut"],
        description: "The type of chart to create. Default is 'column'.",
      },
      title: {
        type: "string",
        description: "Optional title for the chart.",
      },
      sheetName: {
        type: "string",
        description: "Optional worksheet name where the chart will be placed. Defaults to active sheet.",
      },
    },
    required: ["dataRange"],
  },
  handler: async (args) => {
    const { dataRange, chartType = "column", title, sheetName } = args as {
      dataRange: string;
      chartType?: string;
      title?: string;
      sheetName?: string;
    };

    try {
      return await Excel.run(async (context) => {
        // Get the target worksheet
        let sheet: Excel.Worksheet;
        if (sheetName) {
          sheet = context.workbook.worksheets.getItem(sheetName);
        } else {
          sheet = context.workbook.worksheets.getActiveWorksheet();
        }
        sheet.load("name");

        // Get the data range
        const range = sheet.getRange(dataRange);
        range.load(["address", "left", "top", "width"]);
        await context.sync();

        // Map chart type string to Excel.ChartType
        const chartTypeMap: { [key: string]: Excel.ChartType } = {
          "column": Excel.ChartType.columnClustered,
          "bar": Excel.ChartType.barClustered,
          "line": Excel.ChartType.line,
          "pie": Excel.ChartType.pie,
          "area": Excel.ChartType.area,
          "scatter": Excel.ChartType.xyscatter,
          "doughnut": Excel.ChartType.doughnut,
        };

        const excelChartType = chartTypeMap[chartType.toLowerCase()] || Excel.ChartType.columnClustered;

        // Calculate chart position (to the right of data)
        const chartLeft = range.left + range.width + 20;
        const chartTop = range.top;
        const chartWidth = 400;
        const chartHeight = 300;

        // Add the chart
        const chart = sheet.charts.add(
          excelChartType,
          range,
          Excel.ChartSeriesBy.auto
        );

        // Position the chart
        chart.left = chartLeft;
        chart.top = chartTop;
        chart.width = chartWidth;
        chart.height = chartHeight;

        // Set title if provided
        if (title) {
          chart.title.text = title;
          chart.title.visible = true;
        }

        await context.sync();

        return `Created ${chartType} chart from range ${dataRange}${title ? ` with title "${title}"` : ""} in worksheet "${sheet.name}".`;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
