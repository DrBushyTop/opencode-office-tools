import { tool } from "@opencode-ai/plugin"
import { onenote } from "../lib/office-onenote"

export default onenote("create_page", "Create a new OneNote page in the active section or before/after the active page, with optional initial HTML content.", {
  title: tool.schema.string().optional().describe("New page title. Default is New Page."),
  html: tool.schema.string().optional().describe("Optional initial HTML content to place on the new page as an outline."),
  location: tool.schema.enum(["sectionEnd", "before", "after"]).optional().describe("Where to create the page. Default is sectionEnd."),
})
