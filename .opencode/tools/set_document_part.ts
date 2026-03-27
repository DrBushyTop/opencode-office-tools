import { tool } from "@opencode-ai/plugin"
import { word } from "../lib/office-word"

export default word("set_document_part", "Update a structural Word document part using an address.", {
  address: tool.schema.string().describe("Document part address such as section[1].header.primary, section[*], or table_of_contents."),
  operation: tool.schema.enum(["replace", "append", "clear", "insert", "configure"]).optional().describe("Operation to perform."),
  html: tool.schema.string().optional().describe("HTML content to write when updating a body-like part."),
  differentFirstPage: tool.schema.boolean().optional().describe("Use different first-page headers and footers for a section target."),
  oddAndEvenPages: tool.schema.boolean().optional().describe("Use different odd and even page headers and footers for a section target."),
  headerDistance: tool.schema.number().optional().describe("Header distance in points for a section target."),
  footerDistance: tool.schema.number().optional().describe("Footer distance in points for a section target."),
  location: tool.schema.enum(["replace", "before", "after", "start", "end"]).optional().describe("Placement for TOC insertion."),
  upperHeadingLevel: tool.schema.number().optional().describe("Starting heading level for TOC insertion."),
  lowerHeadingLevel: tool.schema.number().optional().describe("Ending heading level for TOC insertion."),
  includePageNumbers: tool.schema.boolean().optional().describe("Include page numbers in TOC insertion."),
  rightAlignPageNumbers: tool.schema.boolean().optional().describe("Right-align page numbers in TOC insertion."),
  useHyperlinksOnWeb: tool.schema.boolean().optional().describe("Use hyperlinks for TOC entries on web publishing."),
})
