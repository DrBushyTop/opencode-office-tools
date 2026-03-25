import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("duplicate_slide", "Duplicate a PowerPoint slide.", {
  sourceIndex: tool.schema.number().describe("0-based source slide index."),
  targetIndex: tool.schema.number().optional().describe("Optional target insert index."),
})
