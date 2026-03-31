import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("read_slide_text", "Read the raw paragraph XML from one PowerPoint text shape using a stable shape ref.", {
  ref: tool.schema.string().describe("Stable shape ref in the format slide-id:<slideId>/shape:<xmlShapeId>."),
})
