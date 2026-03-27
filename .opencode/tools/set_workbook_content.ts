import { tool } from "@opencode-ai/plugin"
import { excel } from "../lib/office-excel"

export default excel("set_workbook_content", "Write tabular data to an Excel worksheet range.", {
  sheetName: tool.schema.string().optional().describe("Worksheet name to write. Defaults to the active sheet."),
  startCell: tool.schema.string().describe("Starting cell such as A1."),
  data: tool.schema.array(tool.schema.array(tool.schema.union([tool.schema.string(), tool.schema.number(), tool.schema.boolean()]))).describe("Two-dimensional array of cell values."),
  useFormulas: tool.schema.boolean().optional().describe("Treat strings starting with = as formulas."),
  clearMode: tool.schema.enum(["none", "contents", "all"]).optional().describe("Optionally clear the destination before writing."),
  createTable: tool.schema.boolean().optional().describe("Create an Excel table over the written range after writing."),
  tableName: tool.schema.string().optional().describe("Optional table name when createTable is true."),
  hasHeaders: tool.schema.boolean().optional().describe("Whether the written range includes a header row when createTable is true."),
  tableStyle: tool.schema.string().optional().describe("Optional Excel table style when createTable is true."),
})
