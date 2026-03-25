import { tool } from "@opencode-ai/plugin"
import { excel } from "../lib/office-excel"

export default excel("apply_cell_formatting", "Apply formatting to cells in Excel.", {
  range: tool.schema.string().describe("Cell range to format."),
  sheetName: tool.schema.string().optional().describe("Worksheet name. Defaults to the active sheet."),
  bold: tool.schema.boolean().optional(),
  italic: tool.schema.boolean().optional(),
  underline: tool.schema.boolean().optional(),
  fontSize: tool.schema.number().optional(),
  fontColor: tool.schema.string().optional(),
  backgroundColor: tool.schema.string().optional(),
  numberFormat: tool.schema.string().optional(),
  horizontalAlignment: tool.schema.enum(["left", "center", "right"]).optional(),
  borderStyle: tool.schema.enum(["thin", "medium", "thick", "none"]).optional(),
  borderColor: tool.schema.string().optional(),
})
