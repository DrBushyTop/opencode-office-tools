import { tool } from "@opencode-ai/plugin"
import { word } from "../lib/office-word"

export default word("get_document_section", "Read a specific Word document section by heading.", {
  headingText: tool.schema.string().describe("Heading text to search for."),
  includeSubsections: tool.schema.boolean().optional().describe("Include nested subsections."),
})
