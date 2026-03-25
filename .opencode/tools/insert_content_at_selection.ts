import { tool } from "@opencode-ai/plugin"
import { word } from "../lib/office-word"

export default word("insert_content_at_selection", "Insert HTML content at the current Word selection.", {
  html: tool.schema.string().describe("HTML to insert."),
  location: tool.schema.enum(["replace", "before", "after", "start", "end"]).optional().describe("Where to insert relative to the current selection."),
})
