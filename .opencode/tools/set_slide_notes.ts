import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("set_slide_notes", "Add or update PowerPoint speaker notes by round-tripping slides through Open XML. This replaces the slide in the deck and may change slide identity.", {
  slideIndex: tool.schema.number().optional().describe("0-based slide index."),
  notes: tool.schema.string().describe("Speaker notes text. Use an empty string to clear notes."),
})
