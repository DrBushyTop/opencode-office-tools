import { tool } from "@opencode-ai/plugin"
import { excel } from "../lib/office-excel"

export default excel("set_workbook_content", "Write tabular data to an Excel worksheet range.", {
  sheetName: tool.schema.string().optional().describe("Worksheet name to write. Defaults to the active sheet."),
  startCell: tool.schema.string().describe("Starting cell such as A1."),
  data: tool.schema.array(tool.schema.array(tool.schema.union([tool.schema.string(), tool.schema.number(), tool.schema.boolean()]))).describe("Two-dimensional array of cell values."),
})
