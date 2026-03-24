import type { OfficeHost } from "../sessionStorage";

export interface OfficeToolExecutor {
  execute(toolName: string, args: Record<string, unknown>): Promise<unknown>;
}

export function createOfficeToolBridge(host: OfficeHost, executor: OfficeToolExecutor) {
  let stopped = false;
  let timer = 0;

  const tick = async () => {
    if (stopped) return;

    try {
      await fetch("/api/office-tools/register", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ host }),
      });

      const response = await fetch(`/api/office-tools/poll?host=${encodeURIComponent(host)}`);
      const data = await response.json();

      if (data.request) {
        try {
          const result = await executor.execute(data.request.toolName, data.request.args || {});
          await fetch(`/api/office-tools/respond/${data.request.id}`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ result }),
          });
        } catch (error) {
          await fetch(`/api/office-tools/respond/${data.request.id}`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ error: error instanceof Error ? error.message : String(error) }),
          });
        }
      }
    } catch {}

    timer = window.setTimeout(tick, 750);
  };

  void tick();

  return {
    stop: async () => {
      stopped = true;
      window.clearTimeout(timer);
      await fetch(`/api/office-tools/register/${host}`, { method: "DELETE" }).catch(() => undefined);
    },
  };
}
