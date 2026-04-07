import { tool } from "@opencode-ai/plugin"
import { excel } from "../lib/office-excel"

export default excel("get_range_image", "Render an Excel range as a PNG snapshot. Useful for checking layout, truncation, spacing, wrapping, and readability after formatting changes.", {
  range: tool.schema.string().describe("Target range such as A1:F12."),
  sheetName: tool.schema.string().optional().describe("Optional worksheet name. Defaults to the active sheet."),
})
