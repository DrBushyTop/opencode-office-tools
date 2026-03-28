import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("get_slide_animations", "Inspect animations on a PowerPoint slide by exporting through Open XML and parsing the timing tree. Returns a structured summary of all animations including their type, target shapes, timing, and sequence order.", {
  slideIndex: tool.schema.number().describe("0-based slide index."),
})
