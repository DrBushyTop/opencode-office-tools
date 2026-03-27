import { tool } from "@opencode-ai/plugin"
import { excel } from "../lib/office-excel"

export default excel("manage_chart", "Create or update Excel charts, including placement, resizing, source data changes, title changes, type changes, activation, and deletion.", {
  action: tool.schema.enum(["create", "setData", "setProperties", "activate", "delete"]).describe("Chart operation to perform."),
  chartName: tool.schema.string().optional().describe("Existing chart name. Required for all actions except create."),
  sheetName: tool.schema.string().optional().describe("Worksheet name for create, placement, or data ranges when ranges are sheet-local."),
  dataRange: tool.schema.string().optional().describe("Source data range for create or setData."),
  chartType: tool.schema.enum(["column", "bar", "line", "pie", "area", "scatter", "doughnut"]).optional().describe("Chart type for create or setProperties."),
  title: tool.schema.string().optional().describe("Optional chart title. Use an empty string to clear it in setProperties."),
  newName: tool.schema.string().optional().describe("New chart name for create or setProperties."),
  left: tool.schema.number().optional().describe("Left position in points for setProperties."),
  top: tool.schema.number().optional().describe("Top position in points for setProperties."),
  width: tool.schema.number().optional().describe("Width in points for setProperties."),
  height: tool.schema.number().optional().describe("Height in points for setProperties."),
  positionStartCell: tool.schema.string().optional().describe("Top-left placement cell for create or setProperties."),
  positionEndCell: tool.schema.string().optional().describe("Optional bottom-right placement cell for create or setProperties."),
})
