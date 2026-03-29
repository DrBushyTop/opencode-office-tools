import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("get_presentation_structure", "Inspect PowerPoint slide masters, layouts, backgrounds, themes, and current selection state.", {
  format: tool.schema.enum(["summary", "structured", "both"]).optional().describe("Response format. Default summary."),
})
