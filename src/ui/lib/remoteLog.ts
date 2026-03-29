import { z } from "zod";

type RemoteLogLevel = "info" | "warn" | "error";

const remoteLogEntrySchema = z.object({
  level: z.enum(["info", "warn", "error"]),
  tag: z.string(),
  message: z.string(),
  detail: z.unknown().optional(),
});

type RemoteLogEntry = z.infer<typeof remoteLogEntrySchema>;

declare global {
  interface Window {
    __opencodeRemoteLogs?: RemoteLogEntry[];
  }
}

function normalizeDetail(detail: unknown): unknown {
  if (detail instanceof Error) {
    return {
      name: detail.name,
      message: detail.message,
      stack: detail.stack,
    };
  }

  if (typeof detail === "object" && detail !== null) {
    try {
      return JSON.parse(JSON.stringify(detail));
    } catch {
      return String(detail);
    }
  }

  return detail;
}

export function remoteLog(tag: string, message: string, detail?: unknown, level: RemoteLogLevel = "error") {
  const body = remoteLogEntrySchema.parse({ level, tag, message, detail: normalizeDetail(detail) });
  window.__opencodeRemoteLogs = window.__opencodeRemoteLogs || [];
  window.__opencodeRemoteLogs.push(body);
  fetch("/api/log", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  }).catch(() => {});
}
