import { describe, expect, it } from "vitest";
import { makeSessionTitle, mapMessages, sessionHistoryInternals } from "./opencode-session-history";

describe("opencode session history", () => {
  it("builds host-aware titles", () => {
    expect(makeSessionTitle("word", " Draft proposal ")).toBe("Word: Draft proposal");
    expect(makeSessionTitle("excel", "")).toBe("Excel: New chat");
  });

  it("maps OpenCode messages into transcript items", () => {
    const items = mapMessages([
      {
        info: { id: "u1", role: "user", time: { created: 1 } },
        parts: [
          { type: "text", text: "Hello" },
          { type: "file", mime: "image/png", url: "file:///tmp/a.png", filename: "a.png" },
        ],
      },
      {
        info: { id: "a1", role: "assistant", time: { created: 2, completed: 3 } },
        parts: [
          { type: "reasoning", id: "r1", text: "**Inspecting** issue", time: { start: 2 } },
          { type: "tool", id: "t1", tool: "get_document_content", state: { input: { section: "intro" }, time: { start: 2 } } },
          { type: "text", id: "m1", text: "Done", time: { start: 3 } },
        ],
      },
    ]);

    expect(items).toHaveLength(4);
    expect(items[0]).toMatchObject({ sender: "user", text: "Hello" });
    expect(items[1]).toMatchObject({ sender: "thinking", text: "**Inspecting** issue" });
    expect(items[2]).toMatchObject({ sender: "tool", toolName: "get_document_content" });
    expect(items[3]).toMatchObject({ sender: "assistant", text: "Done" });
  });

  it("keeps only image file parts", () => {
    expect(sessionHistoryInternals.images([
      { type: "file", mime: "image/png", url: "a", filename: "a.png" },
      { type: "file", mime: "text/plain", url: "b", filename: "b.txt" },
    ])).toEqual([{ dataUrl: "a", name: "a.png" }]);
  });
});
