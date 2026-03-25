import { tool } from "@opencode-ai/plugin"
import { word } from "../lib/office-word"

export default word("find_and_replace", "Find and replace text throughout the active Word document.", {
  find: tool.schema.string().describe("Text to find."),
  replace: tool.schema.string().describe("Replacement text."),
  matchCase: tool.schema.boolean().optional().describe("Match case exactly."),
  matchWholeWord: tool.schema.boolean().optional().describe("Only match whole words."),
})
