import { tool } from "@opencode-ai/plugin"
import { onenote } from "../lib/office-onenote"

export default onenote("append_page_content", "Append HTML content to the active OneNote page. Appends to the last outline when possible, or creates a new outline if needed.", {
  html: tool.schema.string().describe("HTML content to append to the active page."),
})
