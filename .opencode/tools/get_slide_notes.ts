import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("get_slide_notes", "Read speaker notes from PowerPoint slides.", {
  slideIndex: tool.schema.number().optional().describe("Optional single slide index."),
})
