import { tool } from "@opencode-ai/plugin"
import { excel } from "../lib/office-excel"

export default excel("create_named_range", "Create or update a named range in Excel.", {
  name: tool.schema.string().describe("Range name."),
  range: tool.schema.string().describe("Cell range."),
  comment: tool.schema.string().optional().describe("Optional description."),
})
