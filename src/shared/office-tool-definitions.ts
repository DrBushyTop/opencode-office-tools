export type OfficeToolHost = "word" | "powerpoint" | "excel";

export interface OfficeToolDefinition {
  name: string;
  description: string;
  parameters: Record<string, unknown>;
  hosts: OfficeToolHost[];
}

function tool(name: string, description: string, parameters: Record<string, unknown>, hosts: OfficeToolHost[]): OfficeToolDefinition {
  return { name, description, parameters, hosts };
}

export const officeToolDefinitions: Record<string, OfficeToolDefinition> = {
  get_document_overview: tool("get_document_overview", "Get a structural overview of the active Word document.", { type: "object", properties: {} }, ["word"]),
  get_document_content: tool("get_document_content", "Read the current Word document.", { type: "object", properties: {} }, ["word"]),
  get_document_part: tool("get_document_part", "Read a specific Word document part using an address.", { type: "object", properties: { address: { type: "string" }, format: { type: "string", enum: ["summary", "text", "html"] } }, required: ["address"] }, ["word"]),
  get_document_section: tool("get_document_section", "Read a specific Word document section by heading.", { type: "object", properties: { headingText: { type: "string" }, includeSubsections: { type: "boolean" } }, required: ["headingText"] }, ["word"]),
  set_document_content: tool("set_document_content", "Replace the current Word document with new HTML content.", { type: "object", properties: { html: { type: "string" } }, required: ["html"] }, ["word"]),
  set_document_part: tool("set_document_part", "Update a specific Word document part using an address.", { type: "object", properties: { address: { type: "string" }, operation: { type: "string", enum: ["replace", "append", "clear", "insert", "configure"] }, html: { type: "string" }, differentFirstPage: { type: "boolean" }, oddAndEvenPages: { type: "boolean" }, headerDistance: { type: "number" }, footerDistance: { type: "number" }, location: { type: "string", enum: ["replace", "before", "after", "start", "end"] }, upperHeadingLevel: { type: "number" }, lowerHeadingLevel: { type: "number" }, includePageNumbers: { type: "boolean" }, rightAlignPageNumbers: { type: "boolean" }, useHyperlinksOnWeb: { type: "boolean" } }, required: ["address"] }, ["word"]),
  get_selection: tool("get_selection", "Read the current Word selection as OOXML.", { type: "object", properties: {} }, ["word"]),
  get_selection_text: tool("get_selection_text", "Read the current Word selection as plain text.", { type: "object", properties: {} }, ["word"]),
  insert_content_at_selection: tool("insert_content_at_selection", "Insert HTML content at the current Word selection.", { type: "object", properties: { html: { type: "string" }, location: { type: "string", enum: ["replace", "before", "after", "start", "end"] } }, required: ["html"] }, ["word"]),
  find_and_replace: tool("find_and_replace", "Find and replace text throughout the active Word document.", { type: "object", properties: { find: { type: "string" }, replace: { type: "string" }, matchCase: { type: "boolean" }, matchWholeWord: { type: "boolean" } }, required: ["find", "replace"] }, ["word"]),
  insert_table: tool("insert_table", "Insert a table at the current Word selection.", { type: "object", properties: { data: { type: "array" }, hasHeader: { type: "boolean" }, style: { type: "string", enum: ["grid", "striped", "plain"] } }, required: ["data"] }, ["word"]),
  apply_style_to_selection: tool("apply_style_to_selection", "Apply formatting styles to the current Word selection.", { type: "object", properties: { bold: { type: "boolean" }, italic: { type: "boolean" }, underline: { type: "boolean" }, strikethrough: { type: "boolean" }, fontSize: { type: "number" }, fontName: { type: "string" }, fontColor: { type: "string" }, highlightColor: { type: "string" } } }, ["word"]),
  get_workbook_overview: tool("get_workbook_overview", "Get a structural overview of the active Excel workbook.", { type: "object", properties: {} }, ["excel"]),
  get_workbook_info: tool("get_workbook_info", "Get workbook metadata including worksheet names and the active sheet.", { type: "object", properties: {} }, ["excel"]),
  get_workbook_content: tool("get_workbook_content", "Read content from an Excel worksheet or range.", { type: "object", properties: { sheetName: { type: "string" }, range: { type: "string" } } }, ["excel"]),
  set_workbook_content: tool("set_workbook_content", "Write tabular data to an Excel worksheet range.", { type: "object", properties: { sheetName: { type: "string" }, startCell: { type: "string" }, data: { type: "array" } }, required: ["startCell", "data"] }, ["excel"]),
  get_selected_range: tool("get_selected_range", "Read the currently selected Excel range.", { type: "object", properties: {} }, ["excel"]),
  set_selected_range: tool("set_selected_range", "Write values or formulas to the currently selected Excel range.", { type: "object", properties: { data: { type: "array" }, useFormulas: { type: "boolean" } }, required: ["data"] }, ["excel"]),
  find_and_replace_cells: tool("find_and_replace_cells", "Find and replace text in Excel cells.", { type: "object", properties: { find: { type: "string" }, replace: { type: "string" }, sheetName: { type: "string" }, matchCase: { type: "boolean" }, matchEntireCell: { type: "boolean" } }, required: ["find", "replace"] }, ["excel"]),
  insert_chart: tool("insert_chart", "Create a chart from data in Excel.", { type: "object", properties: { dataRange: { type: "string" }, chartType: { type: "string", enum: ["column", "bar", "line", "pie", "area", "scatter", "doughnut"] }, title: { type: "string" }, sheetName: { type: "string" } }, required: ["dataRange"] }, ["excel"]),
  apply_cell_formatting: tool("apply_cell_formatting", "Apply formatting to cells in Excel.", { type: "object", properties: { range: { type: "string" }, sheetName: { type: "string" }, bold: { type: "boolean" }, italic: { type: "boolean" }, underline: { type: "boolean" }, fontSize: { type: "number" }, fontColor: { type: "string" }, backgroundColor: { type: "string" }, numberFormat: { type: "string" }, horizontalAlignment: { type: "string", enum: ["left", "center", "right"] }, borderStyle: { type: "string", enum: ["thin", "medium", "thick", "none"] }, borderColor: { type: "string" } }, required: ["range"] }, ["excel"]),
  create_named_range: tool("create_named_range", "Create or update a named range in Excel.", { type: "object", properties: { name: { type: "string" }, range: { type: "string" }, comment: { type: "string" } }, required: ["name", "range"] }, ["excel"]),
  get_presentation_overview: tool("get_presentation_overview", "Get an overview of the PowerPoint deck.", { type: "object", properties: {} }, ["powerpoint"]),
  get_presentation_content: tool("get_presentation_content", "Read text content from one or more PowerPoint slides.", { type: "object", properties: { slideIndex: { type: "number" }, startIndex: { type: "number" }, endIndex: { type: "number" } } }, ["powerpoint"]),
  get_slide_image: tool("get_slide_image", "Capture a slide image from PowerPoint.", { type: "object", properties: { slideIndex: { type: "number" }, width: { type: "number" } }, required: ["slideIndex"] }, ["powerpoint"]),
  get_slide_notes: tool("get_slide_notes", "Read speaker notes from PowerPoint slides.", { type: "object", properties: { slideIndex: { type: "number" } } }, ["powerpoint"]),
  set_presentation_content: tool("set_presentation_content", "Add text content to a PowerPoint slide.", { type: "object", properties: { slideIndex: { type: "number" }, text: { type: "string" } }, required: ["slideIndex", "text"] }, ["powerpoint"]),
  add_slide_from_code: tool("add_slide_from_code", "Add or replace a PowerPoint slide from PptxGenJS code.", { type: "object", properties: { code: { type: "string" }, replaceSlideIndex: { type: "number" } }, required: ["code"] }, ["powerpoint"]),
  clear_slide: tool("clear_slide", "Remove all shapes from a PowerPoint slide.", { type: "object", properties: { slideIndex: { type: "number" } }, required: ["slideIndex"] }, ["powerpoint"]),
  update_slide_shape: tool("update_slide_shape", "Update the text content of a PowerPoint shape.", { type: "object", properties: { slideIndex: { type: "number" }, shapeIndex: { type: "number" }, text: { type: "string" } }, required: ["slideIndex", "shapeIndex", "text"] }, ["powerpoint"]),
  set_slide_notes: tool("set_slide_notes", "Add or update PowerPoint speaker notes.", { type: "object", properties: { slideIndex: { type: "number" }, notes: { type: "string" } }, required: ["slideIndex", "notes"] }, ["powerpoint"]),
  duplicate_slide: tool("duplicate_slide", "Duplicate a PowerPoint slide.", { type: "object", properties: { sourceIndex: { type: "number" }, targetIndex: { type: "number" } }, required: ["sourceIndex"] }, ["powerpoint"]),
};

export function getOfficeToolNames(host: OfficeToolHost) {
  return Object.values(officeToolDefinitions)
    .filter((item) => item.hosts.includes(host))
    .map((item) => item.name);
}
