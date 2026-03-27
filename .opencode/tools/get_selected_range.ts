import { tool } from "@opencode-ai/plugin"
import { excel } from "../lib/office-excel"

export default excel("get_selected_range", "Read the currently selected Excel range, with optional rich cell metadata.", {
  detail: tool.schema.boolean().optional().describe("Include display text, number formats, validation, merged areas, and table or PivotTable overlap details."),
})
