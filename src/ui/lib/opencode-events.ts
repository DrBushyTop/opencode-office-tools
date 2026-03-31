import { z } from "zod";
import { getLatestSessionUsage } from "./opencode-usage";
import { jsonObjectSchema, opencodeMessagePartSchema, type OpencodeMessage, opencodeMessageSchema } from "./opencode-schemas";

export const uiEventSchema = z.object({
  type: z.enum([
    "assistant.message_delta",
    "assistant.message_update",
    "assistant.message",
    "assistant.reasoning_delta",
    "assistant.reasoning_update",
    "tool.execution_start",
    "tool.execution_complete",
    "assistant.turn_start",
    "assistant.turn_end",
    "session.error",
  ]),
  id: z.string().optional(),
  timestamp: z.string().optional(),
  data: jsonObjectSchema.optional(),
});

export type UiEvent = z.infer<typeof uiEventSchema>;

const sessionErrorEventSchema = z.object({
  type: z.literal("session.error"),
  properties: z.object({
    error: z.object({
      message: z.string().optional(),
      name: z.string().optional(),
    }).passthrough().optional(),
  }).passthrough().optional(),
}).passthrough();

const sessionStatusEventSchema = z.object({
  type: z.literal("session.status"),
  properties: z.object({
    status: z.object({
      type: z.string().optional(),
    }).passthrough().optional(),
  }).passthrough().optional(),
}).passthrough();

const messagePartDeltaEventSchema = z.object({
  type: z.literal("message.part.delta"),
  properties: z.object({
    messageID: z.string().optional(),
    partID: z.string().optional(),
    field: z.string().optional(),
    delta: z.string().optional(),
  }).passthrough(),
}).passthrough();

const messagePartUpdatedEventSchema = z.object({
  type: z.literal("message.part.updated"),
  properties: z.object({
    part: opencodeMessagePartSchema.optional(),
  }).passthrough(),
}).passthrough();

const messageUpdatedEventSchema = z.object({
  type: z.literal("message.updated"),
  properties: z.object({
    info: z.object({
      id: z.string(),
      role: z.string().optional(),
      time: z.object({
        created: z.union([z.string(), z.number(), z.date()]).optional(),
        completed: z.union([z.string(), z.number(), z.date()]).optional(),
      }).passthrough().optional(),
    }).passthrough().optional(),
  }).passthrough(),
}).passthrough();

export const trafficStats = {
  bytesIn: 0,
  bytesOut: 0,
  reset() {
    this.bytesIn = 0;
    this.bytesOut = 0;
  },
};

type PartState = {
  type: string;
  text: string;
  pending: string;
};

function partEvent(type: string, mode: "delta" | "update") {
  if (type === "reasoning") return mode === "delta" ? "assistant.reasoning_delta" : "assistant.reasoning_update";
  if (type === "text") return mode === "delta" ? "assistant.message_delta" : "assistant.message_update";
  return null;
}

function getAssistantText(message: OpencodeMessage): string {
  return (message.parts || [])
    .filter((part) => part.type === "text" && !part.synthetic)
    .map((part: any) => part.text || "")
    .join("\n\n")
    .trim();
}

function getAssistantParts(message: OpencodeMessage) {
  return Array.isArray(message.parts) ? message.parts : [];
}

function getErrorMessage(event: z.infer<typeof sessionErrorEventSchema>): string {
  return event.properties?.error?.message || event.properties?.error?.name || "Unknown session error";
}

export function normalizeOpencodeEvent(event: unknown, parts: Map<string, PartState>, roles?: Map<string, string>): UiEvent[] {
  const sessionErrorEvent = sessionErrorEventSchema.safeParse(event);
  if (sessionErrorEvent.success) {
    return [{ type: "session.error", data: { message: getErrorMessage(sessionErrorEvent.data) } }];
  }

  const sessionStatusEvent = sessionStatusEventSchema.safeParse(event);
  if (sessionStatusEvent.success) {
    if (sessionStatusEvent.data.properties?.status?.type === "busy") {
      return [{ type: "assistant.turn_start", data: {} }];
    }
    if (sessionStatusEvent.data.properties?.status?.type === "idle") {
      return [{ type: "assistant.turn_end", data: {} }];
    }
  }

  const messagePartDeltaEvent = messagePartDeltaEventSchema.safeParse(event);
  if (messagePartDeltaEvent.success) {
    const messageId = messagePartDeltaEvent.data.properties.messageID;
    const partId = messagePartDeltaEvent.data.properties.partID;
    if (!partId) return [];
    if (roles && roles.get(messageId || "") !== "assistant") return [];
    if (messagePartDeltaEvent.data.properties.field && messagePartDeltaEvent.data.properties.field !== "text") return [];
    const delta = messagePartDeltaEvent.data.properties.delta || "";
    const part = parts.get(partId);
    if (!delta) return [];
    if (!part) {
      parts.set(partId, { type: "", text: "", pending: delta });
      return [];
    }
    const type = partEvent(part.type, "delta");
    if (!type) {
      parts.set(partId, { ...part, pending: `${part.pending}${delta}` });
      return [];
    }
    parts.set(partId, { ...part, text: `${part.text}${delta}` });
    return [{
      type,
      id: partId,
      data: { deltaContent: delta },
    }];
  }

  const messagePartUpdatedEvent = messagePartUpdatedEventSchema.safeParse(event);
  if (messagePartUpdatedEvent.success) {
    const part = messagePartUpdatedEvent.data.properties.part;
    if (!part?.id) return [];
    if (roles && roles.get(part.messageID || "") !== "assistant") return [];

    if (part.type === "tool") {
      if (part.state?.status === "running") {
        return [
          {
            type: "tool.execution_start",
            id: part.id,
            data: {
              toolName: part.tool,
              arguments: part.state.input || {},
              metadata: part.state.metadata || {},
            },
          },
        ];
      }

      if (part.state?.status === "completed" || part.state?.status === "error") {
        return [
          {
            type: "tool.execution_complete",
            id: part.id,
            data: {
              toolName: part.tool,
              result: part.state.output,
              error: part.state.error,
              metadata: part.state.metadata || {},
            },
          },
        ];
      }
    }

    const prev = parts.get(part.id);
    const pending = prev?.pending || "";
    const text = typeof part.text === "string" && part.text.length > 0
      ? (part.text.startsWith(pending) ? part.text : `${pending}${part.text}`)
      : pending || prev?.text || "";
    parts.set(part.id, { type: part.type, text, pending: "" });

    const type = partEvent(part.type, "update");
    if (!type || !text || text === prev?.text) return [];

    return [{
      type,
      id: part.id,
      data: { content: text },
    }];
  }

  const messageUpdatedEvent = messageUpdatedEventSchema.safeParse(event);
  if (messageUpdatedEvent.success) {
    const info = messageUpdatedEvent.data.properties.info;
    if (roles && info?.id && info.role) roles.set(info.id, info.role);
    if (info?.role === "assistant" && info?.time?.completed) {
      return [
        {
          type: "assistant.message",
          id: info.id,
          timestamp: new Date(info.time.completed).toISOString(),
          data: {
            content: "",
          },
        },
      ];
    }
    return [];
  }

  return [];
}

export function getLatestAssistantMessage(messages: unknown[]): UiEvent | null {
  const parsedMessages = messages
    .map((message) => opencodeMessageSchema.safeParse(message))
    .filter((result) => result.success)
    .map((result) => result.data);

  const latest = [...parsedMessages].reverse().find((message) => message.info?.role === "assistant");
  if (!latest?.info) return null;

  const content = getAssistantText(latest);
  const parts = getAssistantParts(latest);
  const usage = getLatestSessionUsage([latest]);
  if (!content && parts.length === 0) return null;

  return {
    type: "assistant.message",
    id: latest.info.id,
    timestamp: new Date(latest.info.time.completed || latest.info.time.created || Date.now()).toISOString(),
    data: { content, parts, usage: usage || undefined },
  };
}
