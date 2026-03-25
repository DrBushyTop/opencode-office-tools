import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("set_slide_notes", "Add or update PowerPoint speaker notes.", {
  slideIndex: tool.schema.number().describe("0-based slide index."),
  notes: tool.schema.string().describe("Speaker notes text."),
})
