import { tool } from "@opencode-ai/plugin"
import { word } from "../lib/office-word"

export default word("insert_table", "Insert a table at the current Word selection.", {
  data: tool.schema.array(tool.schema.array(tool.schema.string())).describe("Two-dimensional array of table cell values."),
  hasHeader: tool.schema.boolean().optional().describe("Treat the first row as a header row."),
  style: tool.schema.enum(["grid", "striped", "plain"]).optional().describe("Table style."),
})
