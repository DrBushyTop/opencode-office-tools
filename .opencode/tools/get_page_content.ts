import { tool } from "@opencode-ai/plugin"
import { onenote } from "../lib/office-onenote"

export default onenote("get_page_content", "Read the active OneNote page.", {
  format: tool.schema.enum(["summary", "text", "json"]).optional().describe("Preferred response format. Default is summary."),
})
