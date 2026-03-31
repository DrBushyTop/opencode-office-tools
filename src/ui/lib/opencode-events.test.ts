import { describe, expect, it } from "vitest";
import { normalizeOpencodeEvent } from "./opencode-events";

describe("opencode events", () => {
  it("emits full-text reasoning updates before later deltas", () => {
    const parts = new Map<string, { type: string; text: string; pending: string }>();
    const roles = new Map<string, string>([["a1", "assistant"]]);

    expect(normalizeOpencodeEvent({
      type: "message.part.updated",
      properties: {
        part: {
          id: "r1",
          messageID: "a1",
          type: "reasoning",
          text: "Good first render! I can see a few issues to fix:",
        },
      },
    }, parts, roles)).toEqual([
      {
        type: "assistant.reasoning_update",
        id: "r1",
        data: { content: "Good first render! I can see a few issues to fix:" },
      },
    ]);

    expect(normalizeOpencodeEvent({
      type: "message.part.delta",
      properties: {
        messageID: "a1",
        partID: "r1",
        field: "text",
        delta: " Let me run the visual QA.",
      },
    }, parts, roles)).toEqual([
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
    const roles = new Map<string, string>([["m1", "assistant"]]);

    expect(normalizeOpencodeEvent({
      type: "message.part.updated",
      properties: {
        part: {
          id: "a1",
          messageID: "m1",
          type: "text",
          text: "Partial answer with the missing prefix restored.",
        },
      },
    }, parts, roles)).toEqual([
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
    const roles = new Map<string, string>([["m1", "assistant"]]);

    expect(normalizeOpencodeEvent({
      type: "message.part.delta",
      properties: {
        messageID: "m1",
        partID: "t1",
        field: "arguments",
        delta: "{}",
      },
    }, parts, roles)).toEqual([]);
  });

  it("replays buffered deltas once the part type is known", () => {
    const parts = new Map<string, { type: string; text: string; pending: string }>();
    const roles = new Map<string, string>([["a1", "assistant"]]);

    expect(normalizeOpencodeEvent({
      type: "message.part.delta",
      properties: {
        messageID: "a1",
        partID: "r2",
        field: "text",
        delta: "Missing prefix.",
      },
    }, parts, roles)).toEqual([]);

    expect(normalizeOpencodeEvent({
      type: "message.part.updated",
      properties: {
        part: {
          id: "r2",
          messageID: "a1",
          type: "reasoning",
        },
      },
    }, parts, roles)).toEqual([
      {
        type: "assistant.reasoning_update",
        id: "r2",
        data: { content: "Missing prefix." },
      },
    ]);
  });

  it("ignores user text parts so sent prompts do not echo as assistant output", () => {
    const parts = new Map<string, { type: string; text: string; pending: string }>();
    const roles = new Map<string, string>();

    expect(normalizeOpencodeEvent({
      type: "message.updated",
      properties: {
        info: {
          id: "u1",
          role: "user",
          time: { created: 1 },
        },
      },
    }, parts, roles)).toEqual([]);

    expect(normalizeOpencodeEvent({
      type: "message.part.updated",
      properties: {
        part: {
          id: "p1",
          messageID: "u1",
          type: "text",
          text: "No, use edit_slide_from_code instead",
        },
      },
    }, parts, roles)).toEqual([]);
  });
});
