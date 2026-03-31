import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("execute_office_js", "Primary Office.js escape hatch for live PowerPoint automation. Runs custom code inside PowerPoint.run(context) for custom visualizations, advanced shape work, slide creation, positioning, fills, and host operations the higher-level tools cannot express cleanly.", {
  code: tool.schema.string().describe("Async JavaScript function body that runs inside PowerPoint.run with context, presentation, PowerPoint, Office, console, sync(), and setResult(value) in scope."),
})
