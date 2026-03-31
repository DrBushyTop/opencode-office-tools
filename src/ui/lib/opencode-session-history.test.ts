import { describe, expect, it } from "vitest";
import { carry, coalesceTranscriptMessages, makeSessionTitle, mapMessages, sessionHistoryInternals } from "./opencode-session-history";

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
          { type: "tool", id: "t1", tool: "get_document_content", state: { input: { section: "intro" }, output: "Hello", time: { start: 2 } } },
          { type: "text", id: "m1", text: "Done", time: { start: 3 } },
        ],
      },
    ]);

    expect(items).toHaveLength(4);
    expect(items[0]).toMatchObject({ sender: "user", text: "Hello" });
    expect(items[1]).toMatchObject({ sender: "thinking", text: "**Inspecting** issue" });
    expect(items[2]).toMatchObject({ sender: "tool", toolName: "get_document_content", toolResult: "Hello" });
    expect(items[3]).toMatchObject({ sender: "assistant", text: "Done" });
  });

  it("keeps only image file parts", () => {
    expect(sessionHistoryInternals.images([
      { type: "file", mime: "image/png", url: "a", filename: "a.png" },
      { type: "file", mime: "text/plain", url: "b", filename: "b.txt" },
    ])).toEqual([{ dataUrl: "a", name: "a.png" }]);
  });

  it("keeps live tool history when final assistant text omits it", () => {
    expect(carry([
      { id: "t1", text: "{}", sender: "tool", toolName: "read", timestamp: new Date(1) },
      { id: "a1", text: "partial", sender: "assistant", timestamp: new Date(1) },
      { id: "r1", text: "thinking", sender: "thinking", timestamp: new Date(1) },
    ], [
      { id: "m1", text: "Done", sender: "assistant", timestamp: new Date(2) },
    ])).toMatchObject([
      { id: "t1", sender: "tool" },
      { id: "r1", sender: "thinking" },
    ]);
  });

  it("drops live tool rows already present in final assistant parts", () => {
    expect(carry([
      { id: "t1", text: "{}", sender: "tool", toolName: "read", timestamp: new Date(1) },
    ], [
      { id: "t1", text: "{}", sender: "tool", toolName: "read", timestamp: new Date(2) },
      { id: "m1", text: "Done", sender: "assistant", timestamp: new Date(2) },
    ])).toEqual([]);
  });

  it("coalesces duplicate history and live transcript rows by id", () => {
    expect(coalesceTranscriptMessages([
      { id: "t1", text: "{}", sender: "tool", toolName: "task", toolStatus: "running", timestamp: new Date(1) },
      { id: "u1", text: "It should really be a timeline like chart, not just text", sender: "user", timestamp: new Date(2) },
    ], [
      { id: "t1", text: "{}", sender: "tool", toolName: "task", toolStatus: "completed", timestamp: new Date(3) },
      { id: "a1", text: "Updating the slide now.", sender: "assistant", timestamp: new Date(4) },
    ])).toMatchObject([
      { id: "t1", sender: "tool", toolStatus: "completed" },
      { id: "u1", sender: "user", text: "It should really be a timeline like chart, not just text" },
      { id: "a1", sender: "assistant", text: "Updating the slide now." },
    ]);
  });

  it("ignores malformed message payloads", () => {
    expect(mapMessages([
      { nope: true },
      { info: { id: "a1", role: "assistant", time: { created: 2 } }, parts: [{ type: "text", id: "m1", text: "Done" }] },
    ])).toMatchObject([
      { id: "m1", sender: "assistant", text: "Done" },
    ]);
  });
});
