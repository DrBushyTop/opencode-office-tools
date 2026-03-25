import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("get_presentation_content", "Read text content from one or more PowerPoint slides.", {
  slideIndex: tool.schema.number().optional().describe("Single 0-based slide index to read."),
  startIndex: tool.schema.number().optional().describe("Start of a slide range."),
  endIndex: tool.schema.number().optional().describe("End of a slide range, inclusive."),
})
