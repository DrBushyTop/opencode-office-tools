import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("manage_slide_chart", "Create, update, or delete editable PowerPoint chart-style business visuals built from native shapes.", {
  action: tool.schema.enum(["create", "update", "delete"]),
  slideIndex: tool.schema.number().optional().describe("0-based slide index. Defaults to the active slide when available."),
  shapeId: tool.schema.union([tool.schema.string(), tool.schema.number()]).optional().describe("Existing chart group shape id for update or delete."),
  chartType: tool.schema.enum(["column", "bar", "line", "pie"]).optional(),
  title: tool.schema.string().optional(),
  data: tool.schema.array(tool.schema.object({
  label: tool.schema.string(),
  value: tool.schema.number(),
  color: tool.schema.string().optional(),
})).optional(),
  left: tool.schema.number().optional(),
  top: tool.schema.number().optional(),
  width: tool.schema.number().optional(),
  height: tool.schema.number().optional(),
  name: tool.schema.string().optional(),
})
