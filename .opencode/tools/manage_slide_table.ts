import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("manage_slide_table", "Create, update, or delete editable native PowerPoint tables on a slide.", {
  action: tool.schema.enum(["create", "update", "delete"]),
  slideIndex: tool.schema.number().optional().describe("0-based slide index. Defaults to the active slide when available."),
  shapeId: tool.schema.union([tool.schema.string(), tool.schema.number()]).optional().describe("Existing table shape id for update or delete."),
  values: tool.schema.array(tool.schema.array(tool.schema.union([tool.schema.string(), tool.schema.number(), tool.schema.boolean()]))).optional().describe("2D table values for create or update."),
  left: tool.schema.number().optional(),
  top: tool.schema.number().optional(),
  width: tool.schema.number().optional(),
  height: tool.schema.number().optional(),
  name: tool.schema.string().optional(),
})
