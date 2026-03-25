import { describe, expect, it } from "vitest";
import { getOfficeToolNames } from "./office-tool-definitions";

describe("office tool definitions", () => {
  it("filters tool names by host", () => {
    expect(getOfficeToolNames("word")).toContain("get_document_content");
    expect(getOfficeToolNames("word")).not.toContain("insert_chart");
    expect(getOfficeToolNames("excel")).toContain("insert_chart");
    expect(getOfficeToolNames("excel")).not.toContain("duplicate_slide");
    expect(getOfficeToolNames("powerpoint")).toContain("duplicate_slide");
    expect(getOfficeToolNames("powerpoint")).not.toContain("get_selection_text");
  });
});
