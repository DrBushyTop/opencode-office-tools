import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("edit_slide_text", "Preferred text-editing tool for one PowerPoint text shape. Replaces paragraph XML in place while preserving the rest of the slide.", {
  ref: tool.schema.string().describe("Stable shape ref in the format slide-id:<slideId>/shape:<xmlShapeId>."),
  paragraphsXml: tool.schema.array(tool.schema.string()).describe("Replacement raw <a:p> paragraph XML for the targeted shape."),
})
