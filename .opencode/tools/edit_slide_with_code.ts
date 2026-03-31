import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("edit_slide_with_code", "Edit an existing PowerPoint slide in place with live Office.js code.", {
  slideIndex: tool.schema.number().optional().describe("Optional 0-based slide index. If omitted, the active slide is used when available."),
  shapeId: tool.schema.string().optional().describe("Optional existing shape id to target for pinpoint edits."),
  shapeIndex: tool.schema.number().optional().describe("Optional 0-based shape index on the targeted slide."),
  code: tool.schema.string().describe("JavaScript function body that runs against the live slide and may reference context, slide, shapes, targetShape, targetShapeId, targetShapeIndex, slideIndex, and PowerPoint."),
})
