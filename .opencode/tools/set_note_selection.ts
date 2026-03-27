import { tool } from "@opencode-ai/plugin"
import { onenote } from "../lib/office-onenote"

export default onenote("set_note_selection", "Write text, HTML, or an image to the current OneNote selection.", {
  content: tool.schema.string().describe("Content to insert into the current selection."),
  coercionType: tool.schema.enum(["text", "html", "image"]).optional().describe("How to treat the provided content. Default is text."),
})
