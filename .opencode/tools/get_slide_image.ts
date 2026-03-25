import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("get_slide_image", "Capture a slide image from PowerPoint.", {
  slideIndex: tool.schema.number().describe("0-based slide index."),
  width: tool.schema.number().optional().describe("Optional output width in pixels."),
})
