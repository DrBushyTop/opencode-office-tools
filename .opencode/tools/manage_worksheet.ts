import { tool } from "@opencode-ai/plugin"
import { excel } from "../lib/office-excel"

export default excel("manage_worksheet", "Create, rename, delete, move, change visibility, freeze, unfreeze, activate, protect, or unprotect Excel worksheets.", {
  action: tool.schema.enum(["create", "rename", "delete", "move", "setVisibility", "activate", "freeze", "unfreeze", "protect", "unprotect"]).describe("Worksheet operation to perform."),
  sheetName: tool.schema.string().optional().describe("Target worksheet name. Omit for the active sheet when supported."),
  newName: tool.schema.string().optional().describe("New worksheet name for create or rename."),
  targetPosition: tool.schema.number().optional().describe("Zero-based worksheet position for move or create."),
  visibility: tool.schema.enum(["Visible", "Hidden", "VeryHidden"]).optional().describe("Visibility to apply for setVisibility."),
  freezeRows: tool.schema.number().optional().describe("Number of top rows to freeze."),
  freezeColumns: tool.schema.number().optional().describe("Number of left columns to freeze."),
  freezeRange: tool.schema.string().optional().describe("Range to freeze at, such as B2."),
  password: tool.schema.string().optional().describe("Optional protection password for protect or unprotect."),
})
