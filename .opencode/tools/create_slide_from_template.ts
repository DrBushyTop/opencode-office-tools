import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("create_slide_from_template", "Create a PowerPoint slide from a chosen layout and bind text, image, or table content into placeholders.", {
  slideMasterId: tool.schema.string().optional(),
  layoutId: tool.schema.string().describe("Required PowerPoint layout id to create the slide from."),
  targetIndex: tool.schema.number().optional(),
  bindings: tool.schema.array(tool.schema.object({
  placeholderType: tool.schema.string().optional(),
  placeholderName: tool.schema.string().optional(),
  text: tool.schema.string().optional(),
  imageUrl: tool.schema.string().optional(),
  tableData: tool.schema.array(tool.schema.array(tool.schema.string())).optional(),
})).optional(),
})
