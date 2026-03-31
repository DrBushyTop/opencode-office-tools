import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("duplicate_slide", "Duplicate a slide as a convenience wrapper over native PowerPoint slide operations.", {
  slideIndex: tool.schema.number().optional().describe("0-based source slide index. Defaults to the active slide when available."),
  sourceIndex: tool.schema.number().optional().describe("Alias for slideIndex."),
  targetIndex: tool.schema.number().optional().describe("Optional 0-based insertion index for the duplicated slide. Defaults to right after the source slide."),
  formatting: tool.schema.enum(["KeepSourceFormatting", "UseDestinationTheme"]).optional().describe("Optional formatting behavior for the inserted copy. Default KeepSourceFormatting."),
})
