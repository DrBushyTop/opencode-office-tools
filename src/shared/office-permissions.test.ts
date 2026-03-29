import { describe, expect, it } from "vitest";
import { officePermissionRequestSchema } from "./office-metadata";
import { canAutoApprove, permissionKind, permissionTarget } from "./office-permissions";

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
      patterns: ["get_document_part"],
      metadata: { tool: "get_document_part" },
      always: [],
    })).toBe(true);

    expect(canAutoApprove({
      id: "2c",
      sessionID: "s",
      permission: "tool",
      patterns: ["find_document_text"],
      metadata: { tool: "find_document_text" },
      always: [],
    })).toBe(true);

    expect(canAutoApprove({
      id: "2d",
      sessionID: "s",
      permission: "tool",
      patterns: ["get_notebook_overview"],
      metadata: { tool: "get_notebook_overview" },
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
      patterns: ["set_document_part"],
      metadata: { tool: "set_document_part" },
      always: [],
    })).toBe(false);

    expect(canAutoApprove({
      id: "5",
      sessionID: "s",
      permission: "tool",
      patterns: ["set_document_range"],
      metadata: { tool: "set_document_range" },
      always: [],
    })).toBe(false);

    expect(canAutoApprove({
      id: "6",
      sessionID: "s",
      permission: "tool",
      patterns: ["append_page_content"],
      metadata: { tool: "append_page_content" },
      always: [],
    })).toBe(false);
  });

  it("describes non-office permission requests with the right kind and target", () => {
    expect(permissionKind({
      id: "7",
      sessionID: "s",
      permission: "read",
      patterns: ["/tmp/file.txt"],
      metadata: {},
      always: [],
    })).toBe("read");

    expect(permissionTarget({
      id: "8",
      sessionID: "s",
      permission: "task",
      patterns: ["subagents/code/reviewer"],
      metadata: { subagent_type: "subagents/code/reviewer" },
      always: [],
    })).toBe("subagents/code/reviewer");
  });

  it("parses permission requests with the shared schema", () => {
    expect(() => officePermissionRequestSchema.parse({
      id: "9",
      sessionID: "s",
      permission: "tool",
      patterns: ["get_document_content"],
      metadata: { tool: "get_document_content" },
      always: [],
      tool: {
        messageID: "m",
        callID: "c",
      },
    })).not.toThrow();
  });
});
