import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("clear_slide", "Remove all shapes from a PowerPoint slide.", {
  slideIndex: tool.schema.number().describe("0-based slide index."),
})
