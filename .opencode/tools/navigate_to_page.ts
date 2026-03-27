import { tool } from "@opencode-ai/plugin"
import { onenote } from "../lib/office-onenote"

export default onenote("navigate_to_page", "Navigate OneNote to a specific page by page id or client URL.", {
  pageId: tool.schema.string().optional().describe("Page id from get_notebook_overview. Provide exactly one of pageId or clientUrl."),
  clientUrl: tool.schema.string().optional().describe("Client URL of the page to open. Provide exactly one of pageId or clientUrl."),
})
