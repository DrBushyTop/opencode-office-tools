import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("manage_slide_media", "Insert, replace, or delete editable PowerPoint image shapes on a slide.", {
  action: tool.schema.enum(["insertImage", "replaceImage", "deleteImage"]),
  slideIndex: tool.schema.number().optional().describe("0-based slide index. Defaults to the active slide when available."),
  shapeId: tool.schema.union([tool.schema.string(), tool.schema.number()]).optional().describe("Existing image shape id for replaceImage or deleteImage."),
  imageUrl: tool.schema.string().optional().describe("Source image URL for insertImage or replaceImage."),
  left: tool.schema.number().optional(),
  top: tool.schema.number().optional(),
  width: tool.schema.number().optional(),
  height: tool.schema.number().optional(),
  name: tool.schema.string().optional(),
})
