import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("execute_office_js", "Office.js escape hatch for live PowerPoint automation. Runs custom code inside PowerPoint.run(context). Use only when the higher-level tools (manage_slide_shapes, edit_slide_xml, edit_slide_text, etc.) cannot express the operation cleanly. Do not use for batch shape creation or text formatting that edit_slide_xml or manage_slide_shapes can handle.", {
  code: tool.schema.string().describe("Async JavaScript function body that runs inside PowerPoint.run with context, presentation, PowerPoint, Office, console, sync(), and setResult(value) in scope."),
})
