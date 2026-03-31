import { createRequire } from "module";
import { describe, expect, it } from "vitest";

const require = createRequire(import.meta.url);
const { validateOfficeToolCall } = require("./officeToolValidation");

describe("office tool validation", () => {
  it("accepts valid Office tool calls", () => {
    expect(() => validateOfficeToolCall("word", "get_document_part", { address: "section[1]" })).not.toThrow();
    expect(() => validateOfficeToolCall("word", "get_document_range", { address: "table[1].cell[2,3]", format: "html" })).not.toThrow();
    expect(() => validateOfficeToolCall("excel", "set_selected_range", { data: [[1, true, "x"]] })).not.toThrow();
    expect(() => validateOfficeToolCall("excel", "manage_range", { action: "sort", range: "A1:C10", sortKey: 0 })).not.toThrow();
    expect(() => validateOfficeToolCall("excel", "manage_range", { action: "filter", range: "A1:C10", filterOn: "Values", filterValues: ["North"], columnIndex: 0 })).not.toThrow();
    expect(() => validateOfficeToolCall("powerpoint", "manage_slide", { action: "clear", slideIndex: 0 })).not.toThrow();
    expect(() => validateOfficeToolCall("powerpoint", "manage_slide", { action: "create", targetIndex: 0 })).not.toThrow();
    expect(() => validateOfficeToolCall("powerpoint", "get_presentation_structure", { format: "both" })).not.toThrow();
    expect(() => validateOfficeToolCall("powerpoint", "set_slide_transition", { effect: "fade" })).not.toThrow();
    expect(() => validateOfficeToolCall("powerpoint", "set_slide_notes", { notes: "Hello" })).not.toThrow();
    expect(() => validateOfficeToolCall("powerpoint", "clear_slide_animations", {})).not.toThrow();
    expect(() => validateOfficeToolCall("powerpoint", "edit_slide_with_code", { slideIndex: 0, shapeId: "shape-1", code: "targetShape.left = 24;" })).not.toThrow();
    expect(() => validateOfficeToolCall("powerpoint", "manage_slide_shapes", { action: "update", slideIndex: 0, shapeId: "shape-1", text: "Hello" })).not.toThrow();
    expect(() => validateOfficeToolCall("powerpoint", "manage_slide_shapes", { action: "update", text: "Hello" })).not.toThrow();
    expect(() => validateOfficeToolCall("powerpoint", "manage_slide_shapes", { action: "create", slideIndex: 0, shapeType: "line", width: -20, height: 10 })).not.toThrow();
    expect(() => validateOfficeToolCall("powerpoint", "manage_slide_media", { action: "insertImage", slideIndex: 0, imageUrl: "https://example.com/a.png" })).not.toThrow();
    expect(() => validateOfficeToolCall("powerpoint", "manage_slide_table", { action: "create", slideIndex: 0, values: [["A", 1, true]] })).not.toThrow();
    expect(() => validateOfficeToolCall("powerpoint", "manage_slide_chart", { action: "create", slideIndex: 0, chartType: "column", data: [{ label: "Jan", value: 10 }] })).not.toThrow();
    expect(() => validateOfficeToolCall("powerpoint", "insert_business_layout", { layoutType: "timeline", items: [{ title: "Kickoff" }] })).not.toThrow();
    expect(() => validateOfficeToolCall("powerpoint", "create_slide_from_template", { layoutId: "layout-1", bindings: [{ placeholderType: "Title", text: "Hello" }] })).not.toThrow();
    expect(() => validateOfficeToolCall("onenote", "get_page_content", { format: "summary" })).not.toThrow();
    expect(() => validateOfficeToolCall("onenote", "set_note_selection", { content: "Hello", coercionType: "text" })).not.toThrow();
    expect(() => validateOfficeToolCall("onenote", "navigate_to_page", { pageId: "page-1" })).not.toThrow();
  });

  it("rejects invalid hosts and payload shapes", () => {
    expect(() => validateOfficeToolCall("excel", "get_document_part", { address: "section[1]" })).toThrow(/not available/);
    expect(() => validateOfficeToolCall("word", "get_document_part", {})).toThrow(/Missing required args.address/);
    expect(() => validateOfficeToolCall("word", "set_document_range", { address: "selection", location: "middle" })).toThrow(/expected one of/);
    expect(() => validateOfficeToolCall("word", "set_document_range", { address: "selection" })).toThrow(/args.content/);
    expect(() => validateOfficeToolCall("excel", "manage_range", { action: "filter", range: "A1:C10", filterOn: "Values", filterValues: ["North"] })).toThrow(/columnIndex/);
    expect(() => validateOfficeToolCall("onenote", "set_note_selection", {})).toThrow(/args.content/);
    expect(() => validateOfficeToolCall("onenote", "set_note_selection", { content: "   " })).toThrow(/cannot be empty/);
    expect(() => validateOfficeToolCall("onenote", "set_page_title", { title: "   " })).toThrow(/cannot be empty/);
    expect(() => validateOfficeToolCall("onenote", "append_page_content", { html: "   " })).toThrow(/cannot be empty/);
    expect(() => validateOfficeToolCall("powerpoint", "edit_slide_with_code", { code: "   " })).toThrow(/cannot be empty/);
    expect(() => validateOfficeToolCall("onenote", "navigate_to_page", {})).toThrow(/exactly one/);
    expect(() => validateOfficeToolCall("onenote", "navigate_to_page", { pageId: "page-1", clientUrl: "https://example.invalid" })).toThrow(/exactly one/);
    expect(() => validateOfficeToolCall("word", "get_document_part", { address: "x", extra: true })).toThrow(/Unexpected args.extra/);
    expect(() => validateOfficeToolCall("powerpoint", "get_presentation_structure", { format: "yaml" })).toThrow(/expected one of/);
    expect(() => validateOfficeToolCall("word", "__proto__", {})).toThrow(/Unknown Office tool/);
  });
});
