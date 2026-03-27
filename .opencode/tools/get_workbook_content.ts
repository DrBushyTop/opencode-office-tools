import { tool } from "@opencode-ai/plugin"
import { excel } from "../lib/office-excel"

export default excel("get_workbook_content", "Read values and formulas from an Excel worksheet or range, with optional rich cell metadata.", {
  sheetName: tool.schema.string().optional().describe("Worksheet name to read. Defaults to the active sheet."),
  range: tool.schema.string().optional().describe("Optional cell range such as A1:D10."),
  detail: tool.schema.boolean().optional().describe("Include display text, number formats, validation, merged areas, and table or PivotTable overlap details."),
})
