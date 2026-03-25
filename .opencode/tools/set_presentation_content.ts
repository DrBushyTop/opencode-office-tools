import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("set_presentation_content", "Add text content to a PowerPoint slide.", {
  slideIndex: tool.schema.number().describe("0-based slide index."),
  text: tool.schema.string().describe("Text content to add."),
})
