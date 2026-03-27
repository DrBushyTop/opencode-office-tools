import { tool } from "@opencode-ai/plugin"
import { word } from "../lib/office-word"

export default word("get_document_range", "Read a generic Word target by address.", {
  address: tool.schema.string().describe("Target address such as selection, bookmark[Clause], content_control[id=12], table[1], or table[1].cell[2,3]."),
  format: tool.schema.enum(["summary", "text", "html", "ooxml"]).optional().describe("Preferred response format."),
})
