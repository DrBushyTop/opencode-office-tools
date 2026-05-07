import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("edit_slide_text", "Preferred text-editing tool for one existing PowerPoint text shape. Replaces paragraph XML in place while preserving the rest of the slide. If multiple existing shapes on the same slide need edits, prefer edit_slide_xml.", {
  ref: tool.schema.string().describe("Stable shape ref in the format slide-id:<slideId>/shape:<xmlShapeId>."),
  paragraphsXml: tool.schema.array(tool.schema.string()).describe("Replacement raw <a:p> paragraph XML for the targeted shape."),
})
