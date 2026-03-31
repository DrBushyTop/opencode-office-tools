import { getLatestAssistantMessage, normalizeOpencodeEvent, trafficStats, type UiEvent } from "./opencode-events";
import { modelInfoSchema, opencodeConfigSchema, opencodeMessageSchema, sessionInfoSchema } from "./opencode-schemas";
import { z } from "zod";

export type ModelInfo = z.infer<typeof modelInfoSchema>;

export type SessionInfo = z.infer<typeof sessionInfoSchema>;

export type OpencodeConfig = z.infer<typeof opencodeConfigSchema>;

const promptPartSchema = z.object({
  type: z.enum(["text", "file"]),
  text: z.string().optional(),
  mime: z.string().optional(),
  url: z.string().optional(),
  filename: z.string().optional(),
});

const promptInputSchema = z.object({
  model: z.object({
    providerID: z.string(),
    modelID: z.string(),
  }),
  agent: z.string().optional(),
  system: z.string(),
  parts: z.array(promptPartSchema),
  tools: z.record(z.string(), z.boolean()).optional(),
});

type PromptInput = z.infer<typeof promptInputSchema>;

type SessionEventHandlers = {
  onEvent: (event: UiEvent) => void;
};

const statusSchema = z.object({
  mode: z.string(),
  baseUrl: z.string(),
  directory: z.string(),
  models: z.array(modelInfoSchema),
});

const modelsResponseSchema = z.object({
  models: z.array(modelInfoSchema),
});

const createSessionResponseSchema = z.object({
  id: z.string(),
  title: z.string(),
});

const opencodeMessagesSchema = z.array(opencodeMessageSchema);

async function readJson<T>(path: string, schema: z.ZodType<T>, init?: RequestInit): Promise<T> {
  const response = await fetch(path, init);
  if (!response.ok) {
    throw new Error((await response.text()) || `Request failed: ${response.status}`);
  }
  return schema.parse(await response.json());
}

export class OpencodeClient {
  async getStatus() {
    return readJson("/api/opencode/status", statusSchema);
  }

  async listModels() {
    const data = await readJson("/api/models", modelsResponseSchema);
    return data.models;
  }

  async createSession(input: { title?: string } = {}) {
    return readJson("/api/opencode/session", createSessionResponseSchema, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(input),
    });
  }

  async getMessages(sessionId: string) {
    return readJson(`/api/opencode/session/${sessionId}/messages`, opencodeMessagesSchema);
  }

  async getSession(sessionId: string) {
    return readJson(`/api/opencode/session/${sessionId}`, sessionInfoSchema);
  }

  async getSessionChildren(sessionId: string) {
    return readJson(`/api/opencode/session/${sessionId}/children`, z.array(sessionInfoSchema));
  }

  async abortSession(sessionId: string) {
    return readJson(`/api/opencode/session/${sessionId}/abort`, z.unknown(), {
      method: "POST",
    });
  }

  async getConfig() {
    return readJson("/api/opencode/config", opencodeConfigSchema);
  }

  async updateConfig(config: OpencodeConfig) {
    return readJson("/api/opencode/config", opencodeConfigSchema, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(config),
    });
  }

  async sendMessage(sessionId: string, input: PromptInput) {
    const payload = promptInputSchema.parse(input);
    trafficStats.bytesOut += JSON.stringify(payload).length;

    const response = await fetch(`/api/opencode/session/${sessionId}/message`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      throw new Error((await response.text()) || "Failed to send prompt");
    }
  }

  subscribe(sessionId: string, handlers: SessionEventHandlers, opts: { signal?: AbortSignal } = {}) {
    const parts = new Map<string, { type: string; text: string; pending: string }>();
    const eventSource = new EventSource(`/api/opencode/events?sessionId=${encodeURIComponent(sessionId)}`);
    let lastAssistantId = "";
    let closed = false;
    let failed = false;
    let readyTimer = 0;
    let resolveReady: (() => void) | null = null;
    let rejectReady: ((reason?: unknown) => void) | null = null;

    const ready = new Promise<void>((resolve, reject) => {
      resolveReady = resolve;
      rejectReady = reject;
      readyTimer = window.setTimeout(() => {
        resolveReady?.();
        resolveReady = null;
        rejectReady = null;
      }, 2000);
    });

    const settleReady = (error?: Error) => {
      if (readyTimer) {
        window.clearTimeout(readyTimer);
        readyTimer = 0;
      }

      if (error) {
        rejectReady?.(error);
        resolveReady = null;
        rejectReady = null;
        return;
      }

      resolveReady?.();
      resolveReady = null;
      rejectReady = null;
    };

    const close = () => {
      if (closed) return;
      closed = true;
      settleReady();
      eventSource.close();
      opts.signal?.removeEventListener("abort", close);
    };

    eventSource.onopen = () => {
      settleReady();
    };

    eventSource.onmessage = (message) => {
      trafficStats.bytesIn += message.data.length;
      settleReady();

      void (async () => {
        try {
          const event = JSON.parse(message.data);
          for (const normalized of normalizeOpencodeEvent(event, parts)) {
            if (normalized.type !== "assistant.message") {
              handlers.onEvent(normalized);
              continue;
            }

            const messages = await this.getMessages(sessionId);
            const latest = getLatestAssistantMessage(messages);
            if (!latest || latest.id === lastAssistantId) continue;
            lastAssistantId = String(latest.id || "");
            handlers.onEvent(latest);
          }
        } catch (error) {
          const message = error instanceof Error ? error.message : "Received malformed event data from OpenCode";
          handlers.onEvent({ type: "session.error", data: { message } });
        }
      })();
    };

    eventSource.onerror = () => {
      if (opts.signal?.aborted || closed || failed) return;
      failed = true;
      const error = new Error("OpenCode event stream disconnected");
      settleReady(error);
      handlers.onEvent({ type: "session.error", data: { message: error.message } });
    };

    opts.signal?.addEventListener("abort", close, { once: true });

    return {
      ready,
      close,
    };
  }

  async *query(sessionId: string, input: PromptInput, opts: { signal?: AbortSignal } = {}): AsyncGenerator<UiEvent, void, undefined> {
    const queue: UiEvent[] = [];
    let done = false;
    let wake: (() => void) | null = null;
    const push = (event: UiEvent) => {
      queue.push(event);
      wake?.();
    };

    const ctl = new AbortController();
    const stop = () => {
      done = true;
      ctl.abort();
      wake?.();
    };

    opts.signal?.addEventListener("abort", stop, { once: true });

    const subscription = this.subscribe(sessionId, { onEvent: push }, { signal: ctl.signal });
    const send = subscription.ready.then(() => this.sendMessage(sessionId, input));

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

          if (event.type === "assistant.turn_end" || event.type === "session.error") {
            done = true;
          }

          yield event;
        }
      }
    } finally {
      opts.signal?.removeEventListener("abort", stop);
      subscription.close();
    }
  }
}

export function createOpencodeClient() {
  return new OpencodeClient();
}
