import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("read_slide_text", "Read the raw paragraph XML from one PowerPoint text shape using a stable shape ref. Use this before fidelity-sensitive single-shape edits with edit_slide_text, or before preparing broader slide-scoped XML edits with edit_slide_xml.", {
  ref: tool.schema.string().describe("Stable shape ref in the format slide-id:<slideId>/shape:<xmlShapeId>."),
})
