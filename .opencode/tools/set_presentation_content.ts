import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("set_presentation_content", "Add or update text on a PowerPoint slide.", {
  slideIndex: tool.schema.number().describe("0-based slide index."),
  text: tool.schema.string().describe("Text to add."),
  shapeId: tool.schema.string().optional().describe("Optional existing shape id to update instead of adding a new text box."),
  placeholderType: tool.schema.string().optional().describe("Optional placeholder type to target, such as Title, Body, Subtitle, Content, Footer, Header, Date, or SlideNumber."),
  slideMasterId: tool.schema.string().optional().describe("Optional slide master id to use when creating a new slide."),
  layoutId: tool.schema.string().optional().describe("Optional slide layout id to use when creating a new slide."),
  left: tool.schema.number().optional().describe("Left position for a new text box."),
  top: tool.schema.number().optional().describe("Top position for a new text box."),
  width: tool.schema.number().optional().describe("Width for a new text box."),
  height: tool.schema.number().optional().describe("Height for a new text box."),
})
