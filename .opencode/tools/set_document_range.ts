import { tool } from "@opencode-ai/plugin"
import { word } from "../lib/office-word"

export default word("set_document_range", "Update a generic Word target by address.", {
  address: tool.schema.string().describe("Target address such as selection, bookmark[Clause], content_control[id=12], table[1], or table[1].cell[2,3]."),
  operation: tool.schema.enum(["replace", "insert", "clear"]).optional().describe("Operation to perform."),
  format: tool.schema.enum(["html", "text", "ooxml"]).optional().describe("Content format for replace or insert operations."),
  content: tool.schema.string().optional().describe("Content to write for replace or insert operations. Required unless operation is clear."),
  location: tool.schema.enum(["replace", "before", "after", "start", "end"]).optional().describe("Insertion location for insert operations."),
})
