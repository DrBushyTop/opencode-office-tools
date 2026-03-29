import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("clear_slide_animations", "Remove all animations from a PowerPoint slide through an Open XML slide round-trip. This replaces the slide in the deck and may change slide identity.", {
  slideIndex: tool.schema.number().optional().describe("0-based slide index."),
})
