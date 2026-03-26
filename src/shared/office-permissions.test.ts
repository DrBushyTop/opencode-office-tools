import { describe, expect, it } from "vitest";
import { canAutoApprove } from "./office-permissions";

describe("office permissions", () => {
  it("auto-approves read-only Office tools", () => {
    expect(canAutoApprove({
      id: "1",
      sessionID: "s",
      permission: "tool",
      patterns: ["get_document_content"],
      metadata: { tool: "get_document_content" },
      always: [],
    })).toBe(true);

    expect(canAutoApprove({
      id: "2",
      sessionID: "s",
      permission: "tool",
      patterns: ["get_presentation_overview"],
      metadata: { tool: "get_presentation_overview" },
      always: [],
    })).toBe(true);

    expect(canAutoApprove({
      id: "2b",
      sessionID: "s",
      permission: "tool",
      patterns: ["get_document_headers_footers"],
      metadata: { tool: "get_document_headers_footers" },
      always: [],
    })).toBe(true);
  });

  it("keeps mutating tools interactive", () => {
    expect(canAutoApprove({
      id: "3",
      sessionID: "s",
      permission: "tool",
      patterns: ["set_document_content"],
      metadata: { tool: "set_document_content" },
      always: [],
    })).toBe(false);

    expect(canAutoApprove({
      id: "4",
      sessionID: "s",
      permission: "tool",
      patterns: ["set_section_header_footer"],
      metadata: { tool: "set_section_header_footer" },
      always: [],
    })).toBe(false);
  });
});
