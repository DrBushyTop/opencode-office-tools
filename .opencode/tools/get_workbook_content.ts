import { tool } from "@opencode-ai/plugin"
import { excel } from "../lib/office-excel"

export default excel("get_workbook_content", "Read content from an Excel worksheet or range.", {
  sheetName: tool.schema.string().optional().describe("Worksheet name to read. Defaults to the active sheet."),
  range: tool.schema.string().optional().describe("Optional cell range such as A1:D10."),
})
