import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("get_slide_transition", "Inspect a slide transition by exporting the slide through an Open XML fallback and reading the transition metadata.", {
  slideIndex: tool.schema.number().describe("0-based slide index."),
})
