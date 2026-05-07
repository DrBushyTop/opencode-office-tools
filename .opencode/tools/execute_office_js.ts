import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("execute_office_js", "Office.js escape hatch for live PowerPoint automation. Runs custom code inside PowerPoint.run(context). Use only when the higher-level tools, especially edit_slide_xml, edit_slide_text, layout/slide tools, and manage_slide_shapes, cannot express the operation cleanly. Do not use this for existing-slide text, formatting, or structure edits that edit_slide_xml can handle.", {
  code: tool.schema.string().describe("Async JavaScript function body that runs inside PowerPoint.run with context, presentation, PowerPoint, Office, console, sync(), and setResult(value) in scope."),
})
