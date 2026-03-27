import { describe, expect, it } from "vitest";
import { getOfficeToolNames } from "./office-tool-definitions";

describe("office tool definitions", () => {
  it("filters tool names by host", () => {
    expect(getOfficeToolNames("word")).toContain("get_document_content");
    expect(getOfficeToolNames("word")).toContain("set_document_part");
    expect(getOfficeToolNames("word")).toContain("get_document_range");
    expect(getOfficeToolNames("word")).toContain("get_document_targets");
    expect(getOfficeToolNames("word")).not.toContain("manage_chart");
    expect(getOfficeToolNames("excel")).toContain("manage_chart");
    expect(getOfficeToolNames("excel")).toContain("manage_named_range");
    expect(getOfficeToolNames("excel")).toContain("manage_range");
    expect(getOfficeToolNames("excel")).not.toContain("manage_slide");
    expect(getOfficeToolNames("powerpoint")).toContain("manage_slide");
    expect(getOfficeToolNames("powerpoint")).toContain("manage_slide_shapes");
    expect(getOfficeToolNames("powerpoint")).not.toContain("duplicate_slide");
    expect(getOfficeToolNames("powerpoint")).not.toContain("get_selection_text");
  });
});
