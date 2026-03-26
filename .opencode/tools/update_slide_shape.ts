import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("update_slide_shape", "Update the text content of a PowerPoint shape.", {
  slideIndex: tool.schema.number().describe("0-based slide index."),
  shapeIndex: tool.schema.number().describe("0-based shape index."),
  text: tool.schema.string().describe("Replacement text."),
})
