import { z } from "zod";
import { getLatestSessionUsage } from "./opencode-usage";
import { jsonObjectSchema, opencodeMessagePartSchema, type OpencodeMessage, opencodeMessageSchema } from "./opencode-schemas";

export const uiEventSchema = z.object({
  type: z.enum([
    "assistant.message_delta",
    "assistant.message",
    "assistant.reasoning_delta",
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
    partID: z.string().optional(),
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

export function normalizeOpencodeEvent(event: unknown, partTypes: Map<string, string>): UiEvent[] {
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
    const partId = messagePartDeltaEvent.data.properties.partID;
    if (!partId) return [];
    const type = partTypes.get(partId);
    if (type === "reasoning") {
      return [{
        type: "assistant.reasoning_delta",
        id: partId,
        data: { deltaContent: messagePartDeltaEvent.data.properties.delta || "" },
      }];
    }
    if (type === "text") {
      return [{
        type: "assistant.message_delta",
        id: partId,
        data: { deltaContent: messagePartDeltaEvent.data.properties.delta || "" },
      }];
    }
    return [];
  }

  const messagePartUpdatedEvent = messagePartUpdatedEventSchema.safeParse(event);
  if (messagePartUpdatedEvent.success) {
    const part = messagePartUpdatedEvent.data.properties.part;
    if (!part?.id) return [];
    partTypes.set(part.id, part.type);

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
  }

  const messageUpdatedEvent = messageUpdatedEventSchema.safeParse(event);
  if (messageUpdatedEvent.success) {
    const info = messageUpdatedEvent.data.properties.info;
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
