export function remoteLog(tag: string, message: string, detail?: unknown) {
  const body = { level: "error", tag, message, detail: detail instanceof Error ? detail.message : detail };
  fetch("/api/log", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  }).catch(() => {});
}
