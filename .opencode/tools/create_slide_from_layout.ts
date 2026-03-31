import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("create_slide_from_layout", "Create a new slide from a layout and optionally bind text, images, or tables into layout placeholders.", {
  layoutId: tool.schema.string().describe("Layout id to create the slide from."),
  slideMasterId: tool.schema.string().optional().describe("Optional slide master id when the layout id is not unique across masters."),
  targetIndex: tool.schema.number().optional().describe("Optional 0-based insertion index for the new slide. Defaults to the end of the deck."),
  bindings: tool.schema.array(tool.schema.union([tool.schema.object({
  placeholderType: tool.schema.string().optional(),
  placeholderName: tool.schema.string().optional(),
  text: tool.schema.string(),
}), tool.schema.object({
  placeholderType: tool.schema.string().optional(),
  placeholderName: tool.schema.string().optional(),
  imageUrl: tool.schema.string(),
}), tool.schema.object({
  placeholderType: tool.schema.string().optional(),
  placeholderName: tool.schema.string().optional(),
  tableData: tool.schema.array(tool.schema.array(tool.schema.union([tool.schema.string(), tool.schema.number(), tool.schema.boolean()]))),
})])).optional().describe("Optional placeholder bindings by placeholderType or placeholderName."),
})
