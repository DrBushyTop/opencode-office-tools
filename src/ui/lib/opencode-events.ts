export interface UiEvent {
  type:
    | "assistant.message_delta"
    | "assistant.message"
    | "assistant.reasoning_delta"
    | "tool.execution_start"
    | "tool.execution_complete"
    | "assistant.turn_start"
    | "assistant.turn_end"
    | "session.error";
  id?: string;
  timestamp?: string;
  data?: Record<string, unknown>;
}

export const trafficStats = {
  bytesIn: 0,
  bytesOut: 0,
  reset() {
    this.bytesIn = 0;
    this.bytesOut = 0;
  },
};

function getAssistantText(message: any): string {
  return (message.parts || [])
    .filter((part: any) => part.type === "text" && !part.synthetic)
    .map((part: any) => part.text || "")
    .join("\n\n")
    .trim();
}

function getAssistantParts(message: any): any[] {
  return Array.isArray(message.parts) ? message.parts : [];
}

function getErrorMessage(event: any): string {
  return event.properties?.error?.message || event.properties?.error?.name || "Unknown session error";
}

export function normalizeOpencodeEvent(event: any, partTypes: Map<string, string>): UiEvent[] {
  if (event.type === "session.error") {
    return [{ type: "session.error", data: { message: getErrorMessage(event) } }];
  }

  if (event.type === "session.status") {
    if (event.properties?.status?.type === "busy") {
      return [{ type: "assistant.turn_start", data: {} }];
    }
    if (event.properties?.status?.type === "idle") {
      return [{ type: "assistant.turn_end", data: {} }];
    }
  }

  if (event.type === "message.part.delta") {
    const type = partTypes.get(event.properties?.partID);
    if (type === "reasoning") {
      return [{
        type: "assistant.reasoning_delta",
        id: event.properties?.partID,
        data: { deltaContent: event.properties?.delta || "" },
      }];
    }
    if (type === "text") {
      return [{
        type: "assistant.message_delta",
        id: event.properties?.partID,
        data: { deltaContent: event.properties?.delta || "" },
      }];
    }
    return [];
  }

  if (event.type === "message.part.updated") {
    const part = event.properties?.part;
    if (!part) return [];
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
            },
          },
        ];
      }
    }
  }

  if (event.type === "message.updated") {
    const info = event.properties?.info;
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

export function getLatestAssistantMessage(messages: any[]): UiEvent | null {
  const latest = [...messages].reverse().find((message: any) => message.info?.role === "assistant");
  if (!latest) return null;

  const content = getAssistantText(latest);
  const parts = getAssistantParts(latest);
  if (!content && parts.length === 0) return null;

  return {
    type: "assistant.message",
    id: latest.info.id,
    timestamp: new Date(latest.info.time.completed || latest.info.time.created || Date.now()).toISOString(),
    data: { content, parts },
  };
}
