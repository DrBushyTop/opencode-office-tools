import { tool } from "@opencode-ai/plugin"
import { excel } from "../lib/office-excel"

export default excel("insert_chart", "Create a chart from data in Excel.", {
  dataRange: tool.schema.string().describe("Range containing chart data."),
  chartType: tool.schema.enum(["column", "bar", "line", "pie", "area", "scatter", "doughnut"]).optional().describe("Type of chart to create."),
  title: tool.schema.string().optional().describe("Optional chart title."),
  sheetName: tool.schema.string().optional().describe("Worksheet name where the chart will be placed."),
})
