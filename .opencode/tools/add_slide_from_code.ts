import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("add_slide_from_code", "Advanced fallback: add or replace a PowerPoint slide from PptxGenJS code.", {
  code: tool.schema.string().describe("PptxGenJS slide-building code."),
  replaceSlideIndex: tool.schema.number().optional().describe("Optional slide index to replace."),
})
