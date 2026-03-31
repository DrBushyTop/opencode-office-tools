import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("edit_slide_xml", "General-purpose single-slide XML editor. Exports one slide as a ZIP package, exposes ppt/slides/slide1.xml for DOM-based mutation, and reimports the edited slide in one round-trip.", {
  slideIndex: tool.schema.number().optional().describe("Optional 0-based slide index. Required when no active slide can be inferred safely. Use this for arbitrary single-slide XML edits."),
  code: tool.schema.string().optional().describe("Async JavaScript function body that receives a JSZip-style single-slide package in zip, the parsed slideXml DOM for ppt/slides/slide1.xml, slidePath, DOMParser, XMLSerializer, escapeXml, namespaces, console, parseXml, serializeXml, and setResult(value)."),
  autosize_shape_ids: tool.schema.array(tool.schema.union([tool.schema.string(), tool.schema.number()])).optional().describe("Optional XML cNvPr shape ids that should be reset to AutoSizeShapeToFitText after the edited slide is reimported."),
  replacements: tool.schema.array(tool.schema.object({
  ref: tool.schema.string().describe("Stable shape ref in the format slide-id:<slideId>/shape:<xmlShapeId>."),
  paragraphsXml: tool.schema.array(tool.schema.string()).describe("Replacement raw <a:p> paragraph XML for the targeted shape."),
})).optional().describe("Legacy shorthand for text-only multi-shape updates on one slide. Prefer code for general slide XML edits."),
})
