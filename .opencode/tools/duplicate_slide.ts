import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("duplicate_slide", "Duplicate a PowerPoint slide.", {
  sourceIndex: tool.schema.number().describe("0-based slide index to duplicate."),
  targetIndex: tool.schema.number().optional().describe("Optional 0-based insertion index."),
})
