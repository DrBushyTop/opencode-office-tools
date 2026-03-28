import { getLatestAssistantMessage, normalizeOpencodeEvent, trafficStats, type UiEvent } from "./opencode-events";

export interface ModelInfo {
  key: string;
  label: string;
  providerID: string;
  modelID: string;
}

export interface SessionInfo {
  id: string;
  title?: string | null;
  parentID?: string | null;
}

interface PromptPart {
  type: "text" | "file";
  text?: string;
  mime?: string;
  url?: string;
  filename?: string;
}

interface PromptInput {
  model: { providerID: string; modelID: string };
  agent?: string;
  system: string;
  parts: PromptPart[];
  tools?: Record<string, boolean>;
}

async function readJson<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(path, init);
  if (!response.ok) {
    throw new Error((await response.text()) || `Request failed: ${response.status}`);
  }
  return response.json();
}

export class OpencodeClient {
  async getStatus() {
    return readJson<{ mode: string; baseUrl: string; directory: string; models: ModelInfo[] }>("/api/opencode/status");
  }

  async listModels() {
    const data = await readJson<{ models: ModelInfo[] }>("/api/models");
    return data.models;
  }

  async createSession(input: { title?: string } = {}) {
    return readJson<{ id: string; title: string }>("/api/opencode/session", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(input),
    });
  }

  async getMessages(sessionId: string) {
    return readJson<any[]>(`/api/opencode/session/${sessionId}/messages`);
  }

  async getSession(sessionId: string) {
    return readJson<SessionInfo>(`/api/opencode/session/${sessionId}`);
  }

  async getSessionChildren(sessionId: string) {
    return readJson<SessionInfo[]>(`/api/opencode/session/${sessionId}/children`);
  }

  async abortSession(sessionId: string) {
    return readJson<unknown>(`/api/opencode/session/${sessionId}/abort`, {
      method: "POST",
    });
  }

  async *query(sessionId: string, input: PromptInput, opts: { signal?: AbortSignal } = {}): AsyncGenerator<UiEvent, void, undefined> {
    const partTypes = new Map<string, string>();
    const queue: UiEvent[] = [];
    let done = false;
    let wake: (() => void) | null = null;
    let lastAssistantId = "";

    const eventSource = new EventSource(`/api/opencode/events?sessionId=${encodeURIComponent(sessionId)}`);
    const push = (event: UiEvent) => {
      queue.push(event);
      wake?.();
    };
    const stop = () => {
      done = true;
      eventSource.close();
      wake?.();
    };

    eventSource.onmessage = (message) => {
      trafficStats.bytesIn += message.data.length;
      try {
        const event = JSON.parse(message.data);
        for (const normalized of normalizeOpencodeEvent(event, partTypes)) {
          push(normalized);
        }
      } catch {
        push({ type: "session.error", data: { message: "Received malformed event data from OpenCode" } });
        done = true;
      }
    };

    eventSource.onerror = () => {
      if (opts.signal?.aborted || done) return;
      if (!done) {
        push({ type: "session.error", data: { message: "OpenCode event stream disconnected" } });
        done = true;
      }
    };

    opts.signal?.addEventListener("abort", stop, { once: true });

    trafficStats.bytesOut += JSON.stringify(input).length;
    const send = fetch(`/api/opencode/session/${sessionId}/message`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(input),
    }).then(async (response) => {
      if (!response.ok) {
        throw new Error((await response.text()) || "Failed to send prompt");
      }
    });

    send.catch((error: Error) => {
      push({ type: "session.error", data: { message: error.message } });
      done = true;
    });

    try {
      while (!done || queue.length > 0) {
        if (queue.length === 0) {
          await new Promise<void>((resolve) => {
            wake = resolve;
          });
          wake = null;
        }

        while (queue.length > 0) {
          const event = queue.shift()!;

          if (event.type === "assistant.message") {
            const messages = await this.getMessages(sessionId);
            const latest = getLatestAssistantMessage(messages);
            if (latest && latest.id !== lastAssistantId) {
              lastAssistantId = String(latest.id || "");
              yield latest;
            }
            continue;
          }

          if (event.type === "assistant.turn_end" || event.type === "session.error") {
            done = true;
          }

          yield event;
        }
      }
    } finally {
      opts.signal?.removeEventListener("abort", stop);
      eventSource.close();
    }
  }
}

export function createOpencodeClient() {
  return new OpencodeClient();
}
