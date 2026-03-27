import { tool } from "@opencode-ai/plugin"
import { word } from "../lib/office-word"

export default word("get_document_part", "Read a structural Word document part using an address.", {
  address: tool.schema.string().describe("Document part address such as headers_footers, section[1], section[1].header.primary, or table_of_contents."),
  format: tool.schema.enum(["summary", "text", "html"]).optional().describe("Preferred response format."),
})
