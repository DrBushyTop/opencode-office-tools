import { tool } from "@opencode-ai/plugin"
import { word } from "../lib/office-word"

export default word("set_document_content", "Replace the current Word document with new HTML content.", {
  html: tool.schema.string().describe("HTML to write into the document."),
})
