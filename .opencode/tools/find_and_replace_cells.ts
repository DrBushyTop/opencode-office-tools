import { tool } from "@opencode-ai/plugin"
import { excel } from "../lib/office-excel"

export default excel("find_and_replace_cells", "Find and replace text in Excel cells.", {
  find: tool.schema.string().describe("Text to find."),
  replace: tool.schema.string().describe("Replacement text."),
  sheetName: tool.schema.string().optional().describe("Optional worksheet name."),
  matchCase: tool.schema.boolean().optional().describe("Match case exactly."),
  matchEntireCell: tool.schema.boolean().optional().describe("Match the entire cell value."),
})
