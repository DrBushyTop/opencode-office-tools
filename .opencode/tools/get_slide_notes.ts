import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("get_slide_notes", "Read PowerPoint speaker notes by exporting slides to Open XML when the native API does not expose notes directly.", {
  slideIndex: tool.schema.number().optional().describe("Optional 0-based slide index."),
})
