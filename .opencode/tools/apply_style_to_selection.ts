import { tool } from "@opencode-ai/plugin"
import { word } from "../lib/office-word"

export default word("apply_style_to_selection", "Apply formatting styles to the current Word selection.", {
  bold: tool.schema.boolean().optional(),
  italic: tool.schema.boolean().optional(),
  underline: tool.schema.boolean().optional(),
  strikethrough: tool.schema.boolean().optional(),
  fontSize: tool.schema.number().optional(),
  fontName: tool.schema.string().optional(),
  fontColor: tool.schema.string().optional(),
  highlightColor: tool.schema.string().optional(),
})
