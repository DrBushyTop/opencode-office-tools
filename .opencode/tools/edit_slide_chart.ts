import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("edit_slide_chart", "Create, update, or delete native PowerPoint charts on one slide.", {
  action: tool.schema.enum(["create", "update", "delete"]).describe("Chart operation to perform."),
  slideIndex: tool.schema.number().optional().describe("0-based slide index for create. Defaults to the active slide when available."),
  ref: tool.schema.string().optional().describe("Stable chart ref in the format slide-id:<slideId>/shape:<xmlShapeId> for update or delete."),
  chartType: tool.schema.enum(["column", "bar", "line", "pie", "doughnut", "area", "scatter"]).optional().describe("Chart type for create or update."),
  title: tool.schema.string().optional().describe("Optional chart title."),
  categories: tool.schema.array(tool.schema.string()).optional().describe("Category labels for the chart."),
  series: tool.schema.array(tool.schema.object({
  name: tool.schema.string(),
  values: tool.schema.array(tool.schema.number()),
})).optional().describe("Chart series definitions."),
  stacked: tool.schema.boolean().optional().describe("Whether the chart should use stacked series when supported."),
  left: tool.schema.number().optional().describe("Left position in points."),
  top: tool.schema.number().optional().describe("Top position in points."),
  width: tool.schema.number().optional().describe("Width in points."),
  height: tool.schema.number().optional().describe("Height in points."),
  fontColor: tool.schema.string().optional().describe("Hex color (e.g. \"FFFFFF\" or \"#FFFFFF\") for all chart text including title, axes, legend, and data labels."),
  showDataLabels: tool.schema.boolean().optional().describe("Whether to show data value labels on chart series. Defaults to true."),
  showLegend: tool.schema.boolean().optional().describe("Whether to show the chart legend. Defaults to true."),
  legendPosition: tool.schema.enum(["top", "bottom", "left", "right"]).optional().describe("Legend placement. Defaults to \"top\"."),
})
