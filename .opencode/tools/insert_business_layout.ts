import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("insert_business_layout", "Insert an editable PowerPoint business layout such as a timeline, process flow, comparison grid, phase plan, or estimate summary.", {
  slideIndex: tool.schema.number().optional().describe("Optional slide index. If omitted, the tool prefers creating a new slide from a matching template layout."),
  layoutType: tool.schema.enum(["timeline", "phase_plan", "process_flow", "comparison_grid", "estimate_summary"]),
  title: tool.schema.string().optional(),
  themeMode: tool.schema.enum(["deck", "custom"]).optional(),
  items: tool.schema.array(tool.schema.object({
  title: tool.schema.string(),
  subtitle: tool.schema.string().optional(),
  value: tool.schema.string().optional(),
  colorToken: tool.schema.string().optional(),
})),
})
