import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("get_slide_layout_details", "Inspect one slide layout and return its resolved name, type, and placeholder geometry.", {
  layoutId: tool.schema.string().describe("Layout id from list_slide_layouts."),
  slideMasterId: tool.schema.string().optional().describe("Optional slide master id to disambiguate when the same layout id appears under multiple masters."),
})
