import type { Message } from "../components/MessageList";
import type { ModelType } from "../components/HeaderBar";
import { savedSessionSchema, type OfficeHost, type SavedSession } from "../sessionStorage";
import { getOfficeHostLabel } from "./officeHost";
import { getLatestSessionUsage } from "./opencode-usage";
import { opencodeMessagePartSchema, opencodeMessageSchema, opencodeSessionInfoSchema, type OpencodeMessagePart } from "./opencode-schemas";
import { z } from "zod";

export type OpencodeSessionInfo = z.infer<typeof opencodeSessionInfoSchema>;

const opencodeMessagesSchema = z.array(opencodeMessageSchema);
const opencodeSessionListSchema = z.array(opencodeSessionInfoSchema);
const opencodeSessionSchema = opencodeSessionInfoSchema.pick({ id: true, title: true, directory: true, time: true });

function hostLabel(host: OfficeHost) {
  return getOfficeHostLabel(host);
}

export function makeSessionTitle(host: OfficeHost, text: string) {
  const prefix = `${hostLabel(host)}: `;
  const value = text.trim();
  if (!value) return `${prefix}New chat`;
  if (value.length <= 50) return `${prefix}${value}`;
  return `${prefix}${value.slice(0, 47)}...`;
}

function parseParts(parts: unknown): OpencodeMessagePart[] {
  const parsed = z.array(opencodeMessagePartSchema).safeParse(parts);
  return parsed.success ? parsed.data : [];
}

function text(parts: unknown = []) {
  return parseParts(parts)
    .filter((part) => part.type === "text" && !part.synthetic)
    .map((part) => part.text || "")
    .join("\n\n")
    .trim();
}

function images(parts: unknown = []) {
  return parseParts(parts)
    .filter((part) => part.type === "file" && String(part.mime || "").startsWith("image/"))
    .map((part) => ({ dataUrl: part.url || "", name: part.filename || "image" }));
}

export const sessionHistoryInternals = {
  text,
  images,
  carry,
  coalesceTranscriptMessages,
};

export function carry(live: Message[], next: Message[]) {
  const ids = new Set(next.map((item) => item.id));
  return live.filter((item) => item.sender !== "assistant" && !ids.has(item.id));
}

export function coalesceTranscriptMessages(history: Message[], live: Message[]) {
  const merged = [...history];
  const indexes = new Map(merged.map((item, index) => [item.id, index]));

  for (const item of live) {
    const index = indexes.get(item.id);
    if (index === undefined) {
      indexes.set(item.id, merged.length);
      merged.push(item);
      continue;
    }

    merged[index] = { ...merged[index], ...item };
  }

  return merged;
}

export function mapAssistantParts(parts: unknown = [], fallbackTime?: number): Message[] {
  return parseParts(parts).flatMap((part, index): Message[] => {
    const id = String(part.id || `part-${index}`);
    const time = new Date(part.state?.time?.start || part.time?.start || fallbackTime || Date.now());

    if (part.type === "tool") {
      return [{
        id,
        text: JSON.stringify(part.state?.input || {}, null, 2),
        sender: "tool" as const,
        startedAt: typeof part.state?.time?.start === "number" ? new Date(part.state.time.start) : undefined,
        finishedAt: typeof part.state?.time?.end === "number" ? new Date(part.state.time.end) : undefined,
        toolName: part.tool,
        toolArgs: part.state?.input || {},
        toolResult: part.state?.output,
        toolError: typeof part.state?.error === "string" ? part.state.error : undefined,
        toolMeta: part.state?.metadata || {},
        toolStatus: part.state?.status === "error" || part.state?.error ? "error" : part.state?.status === "pending" || part.state?.status === "running" ? "running" : "completed",
        timestamp: time,
      }];
    }

    if (part.type === "reasoning" && part.text) {
      return [{
        id,
        text: part.text,
        sender: "thinking" as const,
        timestamp: time,
      }];
    }

    if (part.type === "text" && !part.synthetic && part.text) {
      return [{
        id,
        text: part.text,
        sender: "assistant" as const,
        timestamp: time,
      }];
    }

    return [];
  });
}

export function mapMessages(items: unknown[]): Message[] {
  return items.flatMap((item) => {
    const parsed = opencodeMessageSchema.safeParse(item);
    if (!parsed.success) return [];

    const { info, parts } = parsed.data;
    if (info?.role === "user") {
      if (!info) return [];
      const body = text(parts);
      const files = images(parts);
      return [{
        id: info.id,
        text: body || (files.length ? `Sent ${files.length} image${files.length === 1 ? "" : "s"}` : ""),
        sender: "user" as const,
        timestamp: new Date(info.time.created),
        images: files.length ? files : undefined,
      }];
    }

    if (info?.role === "assistant") {
      return mapAssistantParts(parts, Number(info.time.created));
    }

    return [];
  });
}

export async function listSessions(host: OfficeHost, shared: boolean, directory?: string) {
  const query = new URLSearchParams({
    host,
    shared: shared ? "1" : "0",
  });
  if (directory) query.set("directory", directory);
  const response = await fetch(`/api/opencode/sessions?${query.toString()}`);
  if (!response.ok) throw new Error((await response.text()) || "Failed to load sessions");
  return opencodeSessionListSchema.parse(await response.json());
}

export async function deleteSession(id: string, directory?: string) {
  const response = await fetch(
    directory
      ? `/api/opencode/session/${id}?directory=${encodeURIComponent(directory)}`
      : `/api/opencode/session/${id}`,
    { method: "DELETE" },
  );
  if (!response.ok) throw new Error((await response.text()) || "Failed to delete session");
}

export async function updateSessionTitle(id: string, title: string, directory?: string) {
  const response = await fetch(
    directory
      ? `/api/opencode/session/${id}?directory=${encodeURIComponent(directory)}`
      : `/api/opencode/session/${id}`,
    {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ title }),
    },
  );
  if (!response.ok) throw new Error((await response.text()) || "Failed to update session title");
}

export async function restoreSession(id: string, model: ModelType, directory?: string): Promise<SavedSession> {
  const query = directory ? `?directory=${encodeURIComponent(directory)}` : "";
  const [sessionResponse, messagesResponse] = await Promise.all([
    fetch(`/api/opencode/session/${id}${query}`),
    fetch(`/api/opencode/session/${id}/messages${query}`),
  ]);

  if (!sessionResponse.ok) throw new Error((await sessionResponse.text()) || "Failed to load session");
  if (!messagesResponse.ok) throw new Error((await messagesResponse.text()) || "Failed to load messages");

  const session = opencodeSessionSchema.parse(await sessionResponse.json());
  const messages = opencodeMessagesSchema.parse(await messagesResponse.json());

  return savedSessionSchema.parse({
    id: session.id,
    title: session.title,
    directory: session.directory,
    model,
    messages: mapMessages(messages),
    usage: getLatestSessionUsage(messages),
    createdAt: new Date(session.time.created).toISOString(),
    updatedAt: new Date(session.time.updated).toISOString(),
  });
}
