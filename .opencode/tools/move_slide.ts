import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("move_slide", "Move a slide to a new zero-based position in the PowerPoint presentation.", {
  slideIndex: tool.schema.number().describe("0-based index of the slide to move."),
  targetIndex: tool.schema.number().describe("0-based index where the slide should be moved."),
})
