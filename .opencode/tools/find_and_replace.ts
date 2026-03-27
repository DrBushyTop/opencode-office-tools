import { tool } from "@opencode-ai/plugin"
import { word } from "../lib/office-word"

export default word("find_and_replace", "Find and replace text in Word, optionally scoped to a generic target address.", {
  find: tool.schema.string().describe("Text to find."),
  replace: tool.schema.string().describe("Replacement text."),
  address: tool.schema.string().optional().describe("Optional scope address such as selection, bookmark[Clause], content_control[id=12], table[1], or table[1].cell[2,3]."),
  matchCase: tool.schema.boolean().optional().describe("Match case exactly."),
  matchWholeWord: tool.schema.boolean().optional().describe("Match whole words only."),
})
