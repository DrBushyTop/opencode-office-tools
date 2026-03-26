import { tool } from "@opencode-ai/plugin"
import { excel } from "../lib/office-excel"

export default excel("insert_chart", "Create a chart from data in Excel.", {
  dataRange: tool.schema.string().describe("Data range for the chart."),
  chartType: tool.schema.enum(["column", "bar", "line", "pie", "area", "scatter", "doughnut"]).optional().describe("Chart type."),
  title: tool.schema.string().optional().describe("Optional chart title."),
  sheetName: tool.schema.string().optional().describe("Worksheet name to use."),
})
