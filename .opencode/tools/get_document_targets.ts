import { tool } from "@opencode-ai/plugin"
import { word } from "../lib/office-word"

export default word("get_document_targets", "Inspect Word tables, content controls, and bookmarks for later generic targeting.", {
  kind: tool.schema.enum(["all", "tables", "contentControls", "bookmarks"]).optional().describe("Which target family to inspect."),
  maxItems: tool.schema.number().optional().describe("Maximum items to include per family."),
  includeText: tool.schema.boolean().optional().describe("Include short text previews when available."),
})
