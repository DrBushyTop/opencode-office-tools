import { describe, expect, it } from "vitest";
import { normalizeOpencodeEvent } from "./opencode-events";

describe("opencode events", () => {
  it("emits full-text reasoning updates before later deltas", () => {
    const parts = new Map<string, { type: string; text: string; pending: string }>();

    expect(normalizeOpencodeEvent({
      type: "message.part.updated",
      properties: {
        part: {
          id: "r1",
          type: "reasoning",
          text: "Good first render! I can see a few issues to fix:",
        },
      },
    }, parts)).toEqual([
      {
        type: "assistant.reasoning_update",
        id: "r1",
        data: { content: "Good first render! I can see a few issues to fix:" },
      },
    ]);

    expect(normalizeOpencodeEvent({
      type: "message.part.delta",
      properties: {
        partID: "r1",
        field: "text",
        delta: " Let me run the visual QA.",
      },
    }, parts)).toEqual([
      {
        type: "assistant.reasoning_delta",
        id: "r1",
        data: { deltaContent: " Let me run the visual QA." },
      },
    ]);
  });

  it("resyncs assistant text from cumulative part updates", () => {
    const parts = new Map<string, { type: string; text: string; pending: string }>([
      ["a1", { type: "text", text: "Partial", pending: "" }],
    ]);

    expect(normalizeOpencodeEvent({
      type: "message.part.updated",
      properties: {
        part: {
          id: "a1",
          type: "text",
          text: "Partial answer with the missing prefix restored.",
        },
      },
    }, parts)).toEqual([
      {
        type: "assistant.message_update",
        id: "a1",
        data: { content: "Partial answer with the missing prefix restored." },
      },
    ]);
  });

  it("ignores non-text deltas", () => {
    const parts = new Map<string, { type: string; text: string; pending: string }>([
      ["t1", { type: "tool", text: "", pending: "" }],
    ]);

    expect(normalizeOpencodeEvent({
      type: "message.part.delta",
      properties: {
        partID: "t1",
        field: "arguments",
        delta: "{}",
      },
    }, parts)).toEqual([]);
  });

  it("replays buffered deltas once the part type is known", () => {
    const parts = new Map<string, { type: string; text: string; pending: string }>();

    expect(normalizeOpencodeEvent({
      type: "message.part.delta",
      properties: {
        partID: "r2",
        field: "text",
        delta: "Missing prefix.",
      },
    }, parts)).toEqual([]);

    expect(normalizeOpencodeEvent({
      type: "message.part.updated",
      properties: {
        part: {
          id: "r2",
          type: "reasoning",
        },
      },
    }, parts)).toEqual([
      {
        type: "assistant.reasoning_update",
        id: "r2",
        data: { content: "Missing prefix." },
      },
    ]);
  });
});
