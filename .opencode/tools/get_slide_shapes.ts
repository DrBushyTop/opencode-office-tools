import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("get_slide_shapes", "Inspect PowerPoint shapes on one slide or across the deck.", {
  slideIndex: tool.schema.number().optional().describe("Optional 0-based slide index. Omit to inspect all slides."),
  includeFormatting: tool.schema.boolean().optional().describe("Include fill and line formatting details. Default true."),
  includeTableValues: tool.schema.boolean().optional().describe("Include table cell values for table shapes. Default false."),
  detail: tool.schema.boolean().optional().describe("Return full text and table data instead of previews. Default false."),
})
