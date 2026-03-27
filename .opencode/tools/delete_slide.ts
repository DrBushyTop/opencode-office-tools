import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("delete_slide", "Delete a slide from the PowerPoint presentation.", {
  slideIndex: tool.schema.number().describe("0-based slide index to delete."),
})
