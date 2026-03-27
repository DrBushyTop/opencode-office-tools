import { tool } from "@opencode-ai/plugin"
import { excel } from "../lib/office-excel"

export default excel("apply_cell_formatting", "Apply formatting to cells in Excel, including alignment, wrapping, merging, borders, and sizing.", {
  range: tool.schema.string().describe("Cell range to format."),
  sheetName: tool.schema.string().optional().describe("Worksheet name. Defaults to the active sheet."),
  bold: tool.schema.boolean().optional(),
  italic: tool.schema.boolean().optional(),
  underline: tool.schema.boolean().optional(),
  fontSize: tool.schema.number().optional(),
  fontColor: tool.schema.string().optional(),
  backgroundColor: tool.schema.string().optional(),
  numberFormat: tool.schema.string().optional(),
  horizontalAlignment: tool.schema.enum(["left", "center", "right", "general", "fill", "justify", "centerAcrossSelection", "distributed"]).optional(),
  verticalAlignment: tool.schema.enum(["top", "center", "bottom", "justify", "distributed"]).optional(),
  wrapText: tool.schema.boolean().optional(),
  merge: tool.schema.boolean().optional(),
  mergeAcross: tool.schema.boolean().optional(),
  borderStyle: tool.schema.enum(["thin", "medium", "thick", "none", "double", "dashed", "dotted"]).optional(),
  borderColor: tool.schema.string().optional(),
  interiorBorders: tool.schema.boolean().optional(),
  rowHeight: tool.schema.number().optional(),
  columnWidth: tool.schema.number().optional(),
  autoFitRows: tool.schema.boolean().optional(),
  autoFitColumns: tool.schema.boolean().optional(),
})
