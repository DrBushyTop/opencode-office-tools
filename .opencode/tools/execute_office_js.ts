import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("execute_office_js", "Run Office.js against the live PowerPoint deck. Prefer this for host-native slide authoring, coordinated shape layout, direct positioning/sizing, fills, lines, z-order, and operations that are clearer as one Office.js batch. Keep specialized tools for safer precision workflows such as rich text, OOXML, charts, masters, animations, notes, and transitions.", {
  code: tool.schema.string().describe("Async JavaScript function body that runs inside PowerPoint.run with context, presentation, PowerPoint, Office, console, sync(), and setResult(value) in scope."),
})
