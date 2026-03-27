import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("set_slide_shape_properties", "Update PowerPoint shape properties by shape id or shape index.", {
  slideIndex: tool.schema.number().describe("0-based slide index."),
  shapeIndex: tool.schema.number().optional().describe("0-based shape index within the slide."),
  shapeId: tool.schema.string().optional().describe("Unique PowerPoint shape id. Preferred when available."),
  text: tool.schema.string().optional().describe("Replacement text for shapes that support text."),
  name: tool.schema.string().optional(),
  left: tool.schema.number().optional(),
  top: tool.schema.number().optional(),
  width: tool.schema.number().optional(),
  height: tool.schema.number().optional(),
  rotation: tool.schema.number().optional(),
  visible: tool.schema.boolean().optional(),
  altTextTitle: tool.schema.string().optional(),
  altTextDescription: tool.schema.string().optional(),
  fillColor: tool.schema.string().optional(),
  fillTransparency: tool.schema.number().optional(),
  clearFill: tool.schema.boolean().optional(),
  lineColor: tool.schema.string().optional(),
  lineWeight: tool.schema.number().optional(),
  lineTransparency: tool.schema.number().optional(),
  lineVisible: tool.schema.boolean().optional(),
})
