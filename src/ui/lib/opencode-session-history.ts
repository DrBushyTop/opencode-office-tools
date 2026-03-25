import type { Message } from "../components/MessageList";
import type { ModelType } from "../components/HeaderBar";
import type { OfficeHost, SavedSession } from "../sessionStorage";

export interface OpencodeSessionInfo {
  id: string;
  title: string;
  directory: string;
  time: {
    created: number;
    updated: number;
  };
}

function hostLabel(host: OfficeHost) {
  if (host === "powerpoint") return "PowerPoint";
  if (host === "excel") return "Excel";
  return "Word";
}

export function makeSessionTitle(host: OfficeHost, text: string) {
  const prefix = `${hostLabel(host)}: `;
  const value = text.trim();
  if (!value) return `${prefix}New chat`;
  if (value.length <= 50) return `${prefix}${value}`;
  return `${prefix}${value.slice(0, 47)}...`;
}

function text(parts: any[] = []) {
  return parts
    .filter((part) => part.type === "text" && !part.synthetic)
    .map((part) => part.text || "")
    .join("\n\n")
    .trim();
}

function images(parts: any[] = []) {
  return parts
    .filter((part) => part.type === "file" && String(part.mime || "").startsWith("image/"))
    .map((part) => ({ dataUrl: part.url || "", name: part.filename || "image" }));
}

export const sessionHistoryInternals = {
  text,
  images,
};

export function mapMessages(items: any[]): Message[] {
  return items.flatMap((item) => {
    if (item.info?.role === "user") {
      const body = text(item.parts);
      const files = images(item.parts);
      return [{
        id: item.info.id,
        text: body || (files.length ? `Sent ${files.length} image${files.length === 1 ? "" : "s"}` : ""),
        sender: "user" as const,
        timestamp: new Date(item.info.time.created),
        images: files.length ? files : undefined,
      }];
    }

    if (item.info?.role === "assistant") {
      const body = text(item.parts);
      const tools = (item.parts || [])
        .filter((part: any) => part.type === "tool")
        .map((part: any) => ({
          id: part.id,
          text: JSON.stringify(part.state?.input || {}, null, 2),
          sender: "tool" as const,
          toolName: part.tool,
          toolArgs: part.state?.input || {},
          timestamp: new Date(part.state?.time?.start || item.info.time.created),
        }));
      const assistant = body
        ? [{
            id: item.info.id,
            text: body,
            sender: "assistant" as const,
            timestamp: new Date(item.info.time.completed || item.info.time.created),
          }]
        : [];
      return [...tools, ...assistant];
    }

    return [];
  });
}

export async function listSessions(host: OfficeHost, shared: boolean) {
  const response = await fetch(`/api/opencode/sessions?host=${encodeURIComponent(host)}&shared=${shared ? "1" : "0"}`);
  if (!response.ok) throw new Error((await response.text()) || "Failed to load sessions");
  return response.json() as Promise<OpencodeSessionInfo[]>;
}

export async function deleteSession(id: string) {
  const response = await fetch(`/api/opencode/session/${id}`, { method: "DELETE" });
  if (!response.ok) throw new Error((await response.text()) || "Failed to delete session");
}

export async function updateSessionTitle(id: string, title: string) {
  const response = await fetch(`/api/opencode/session/${id}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ title }),
  });
  if (!response.ok) throw new Error((await response.text()) || "Failed to update session title");
}

export async function restoreSession(id: string, model: ModelType): Promise<SavedSession> {
  const [sessionResponse, messagesResponse] = await Promise.all([
    fetch(`/api/opencode/session/${id}`),
    fetch(`/api/opencode/session/${id}/messages`),
  ]);

  if (!sessionResponse.ok) throw new Error((await sessionResponse.text()) || "Failed to load session");
  if (!messagesResponse.ok) throw new Error((await messagesResponse.text()) || "Failed to load messages");

  const session = await sessionResponse.json();
  const messages = await messagesResponse.json();

  return {
    id: session.id,
    title: session.title,
    model,
    messages: mapMessages(messages),
    createdAt: new Date(session.time.created).toISOString(),
    updatedAt: new Date(session.time.updated).toISOString(),
  };
}
