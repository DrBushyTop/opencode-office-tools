import { tool } from "@opencode-ai/plugin"
import { onenote } from "../lib/office-onenote"

export default onenote("set_page_title", "Update the title of the active OneNote page.", {
  title: tool.schema.string().describe("New title for the active page."),
})
