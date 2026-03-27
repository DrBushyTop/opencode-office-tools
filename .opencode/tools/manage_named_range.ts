import { tool } from "@opencode-ai/plugin"
import { excel } from "../lib/office-excel"

export default excel("manage_named_range", "Create, update, rename, set visibility, or delete Excel named ranges.", {
  action: tool.schema.enum(["create", "update", "rename", "setVisibility", "delete"]).describe("Named range operation to perform."),
  name: tool.schema.string().describe("Existing or new named range name depending on the action."),
  newName: tool.schema.string().optional().describe("New named range name for rename."),
  reference: tool.schema.string().optional().describe("Cell or formula reference such as A1:D10, Sheet1!B2, or =SUM(A:A)."),
  comment: tool.schema.string().optional().describe("Optional description to set when creating or updating."),
  visible: tool.schema.boolean().optional().describe("Whether the named range is visible for setVisibility or update."),
})
