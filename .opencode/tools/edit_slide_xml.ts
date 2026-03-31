import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("edit_slide_xml", "Batch-update paragraph XML for multiple text shapes on one slide in a single native PowerPoint round-trip.", {
  replacements: tool.schema.array(tool.schema.object({
  ref: tool.schema.string().describe("Stable shape ref in the format slide-id:<slideId>/shape:<xmlShapeId>."),
  paragraphsXml: tool.schema.array(tool.schema.string()).describe("Replacement raw <a:p> paragraph XML for the targeted shape."),
})).describe("Text-shape replacements to apply on a single slide."),
})
