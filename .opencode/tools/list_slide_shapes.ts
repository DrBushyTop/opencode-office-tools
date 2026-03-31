import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("list_slide_shapes", "List shapes on one PowerPoint slide and return stable shape refs for follow-up text and chart operations.", {
  slideIndex: tool.schema.number().optional().describe("Optional 0-based slide index. If omitted, the active slide is used when it can be inferred safely."),
  detail: tool.schema.boolean().optional().describe("When true, include full text instead of summarized previews."),
})
