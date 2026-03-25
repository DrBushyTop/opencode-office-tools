import { tool } from "@opencode-ai/plugin"
import { excel } from "../lib/office-excel"

export default excel("set_selected_range", "Write values or formulas to the currently selected Excel range.", {
  data: tool.schema.array(tool.schema.array(tool.schema.union([tool.schema.string(), tool.schema.number(), tool.schema.boolean()]))).describe("Two-dimensional array of values to write."),
  useFormulas: tool.schema.boolean().optional().describe("Treat strings starting with = as formulas."),
})
