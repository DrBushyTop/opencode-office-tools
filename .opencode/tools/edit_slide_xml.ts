import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("edit_slide_xml", "General-purpose single-slide XML editor. Exports one slide as a ZIP package, exposes ppt/slides/slide1.xml for DOM-based mutation, and reimports the edited slide in one round-trip. Use for batch text edits across multiple shapes, advanced formatting, structural shape work, shape geometry and style fixes, and any single-slide edit that benefits from full OOXML fidelity. Prefer this over execute_office_js for text editing and formatting work.", {
  slideIndex: tool.schema.number().optional().describe("Optional 0-based slide index. Required when no active slide can be inferred safely. Use this for arbitrary single-slide XML edits."),
  code: tool.schema.string().optional().describe("Async JavaScript function body that receives a JSZip-style single-slide package in zip (supporting both zip.file(path) and zip.files[path] reads), the raw slide XML string in xml, the parsed slide XML DOM in both doc and slideXml for ppt/slides/slide1.xml, slidePath, DOMParser, XMLSerializer, escapeXml, namespaces, console, parseXml, serializeXml, and setResult(value). Returning an XML string replaces ppt/slides/slide1.xml directly. This is the right tool for repairing Office.js-created shapes whose OOXML is wrong, such as `a:prstGeom prst=\"lineInv\"` shapes that ignore fills until you switch the preset to `rect` or `roundRect` and remove `p:style` fill overrides."),
  autosize_shape_ids: tool.schema.array(tool.schema.union([tool.schema.string(), tool.schema.number()])).optional().describe("Optional XML cNvPr shape ids whose current text auto-size settings should be preserved after the edited slide is reimported."),
  replacements: tool.schema.array(tool.schema.object({
  ref: tool.schema.string().describe("Stable shape ref in the format slide-id:<slideId>/shape:<xmlShapeId>."),
  paragraphsXml: tool.schema.array(tool.schema.string()).describe("Replacement raw <a:p> paragraph XML for the targeted shape."),
})).optional().describe("Legacy shorthand for text-only multi-shape updates on one slide. Prefer code for general slide XML edits."),
})
