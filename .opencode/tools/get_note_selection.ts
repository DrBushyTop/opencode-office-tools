import { tool } from "@opencode-ai/plugin"
import { onenote } from "../lib/office-onenote"

export default onenote("get_note_selection", "Read the current OneNote selection as plain text or a matrix of values.", {
  format: tool.schema.enum(["text", "matrix"]).optional().describe("Selection format to read. Default is text."),
})
