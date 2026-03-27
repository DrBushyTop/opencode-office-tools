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
    expect(() => validateOfficeToolCall("powerpoint", "manage_slide", { action: "clear", slideIndex: 0 })).not.toThrow();
    expect(() => validateOfficeToolCall("powerpoint", "manage_slide_shapes", { action: "update", slideIndex: 0, shapeId: "shape-1", text: "Hello" })).not.toThrow();
  });

  it("rejects invalid hosts and payload shapes", () => {
    expect(() => validateOfficeToolCall("excel", "get_document_part", { address: "section[1]" })).toThrow(/not available/);
    expect(() => validateOfficeToolCall("word", "get_document_part", {})).toThrow(/Missing required args.address/);
    expect(() => validateOfficeToolCall("word", "set_document_range", { address: "selection", location: "middle" })).toThrow(/expected one of/);
    expect(() => validateOfficeToolCall("word", "get_document_part", { address: "x", extra: true })).toThrow(/Unexpected args.extra/);
    expect(() => validateOfficeToolCall("word", "__proto__", {})).toThrow(/Unknown Office tool/);
  });
});
