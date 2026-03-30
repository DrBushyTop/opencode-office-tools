import type { OfficeHost } from "../sessionStorage";
import type { PowerPointContextSnapshot } from "../tools/powerpointContext";
import { z } from "zod";

const officeBridgeSessionSchema = z.object({
  sessionToken: z.string().min(1),
});

const officeExecutorSchema = z.object({
  executorId: z.string().min(1),
});

const officeToolRequestSchema = z.object({
  id: z.string(),
  toolName: z.string(),
  args: z.record(z.string(), z.unknown()).optional(),
});

const officeToolPollSchema = z.object({
  request: officeToolRequestSchema.optional(),
}).passthrough();

export interface OfficeToolExecutor {
  execute(toolName: string, args: Record<string, unknown>): Promise<unknown>;
}

export async function readPowerPointContextSnapshot(): Promise<PowerPointContextSnapshot | null> {
  if (Office.context.host !== Office.HostType.PowerPoint) return null;
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) return null;

  return await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");

    const selectedSlides = context.presentation.getSelectedSlides();
    const selectedShapes = context.presentation.getSelectedShapes();
    selectedSlides.load("items/id");
    selectedShapes.load("items/id");
    await context.sync();

    const selectedSlideIds = selectedSlides.items.map((slide) => slide.id || "").filter(Boolean);
    const selectedShapeIds = selectedShapes.items.map((shape) => shape.id || "").filter(Boolean);
    const activeSlideId = selectedSlideIds[0];
    const activeSlideIndex = activeSlideId ? slides.items.findIndex((slide) => slide.id === activeSlideId) : undefined;

    return {
      selectedSlideIds,
      selectedShapeIds,
      activeSlideId,
      activeSlideIndex: activeSlideIndex !== undefined && activeSlideIndex >= 0 ? activeSlideIndex : undefined,
    };
  }).catch(() => null);
}

export function createOfficeToolBridge(host: OfficeHost, executor: OfficeToolExecutor) {
  let stopped = false;
  let timer = 0;
  let sessionToken = "";
  let executorId = "";

  const bridgeFetch = (input: string, init?: RequestInit) => fetch(input, {
    cache: "no-store",
    ...init,
  });

  const headers = () => ({
    "Content-Type": "application/json",
    "x-office-bridge-session": sessionToken,
    "x-office-executor-id": executorId,
  });

  const ensureSession = async () => {
    if (sessionToken) return;
    const response = await bridgeFetch("/api/office-tools/session");
    if (!response.ok) {
      throw new Error(`Failed to create Office bridge session: ${response.statusText}`);
    }
    const data = officeBridgeSessionSchema.parse(await response.json());
    sessionToken = data.sessionToken;
  };

  const ensureExecutor = async () => {
    if (executorId) return;
    await ensureSession();
    const response = await bridgeFetch("/api/office-tools/register", {
      method: "POST",
      headers: headers(),
      body: JSON.stringify({ host }),
    });
    if (!response.ok) {
      throw new Error(`Failed to register Office executor: ${response.statusText}`);
    }
    const data = officeExecutorSchema.parse(await response.json());
    executorId = data.executorId;
  };

  const ensureOk = async (response: Response) => {
    if (response.ok) return response;
    throw new Error(await response.text() || response.statusText);
  };

  const tick = async () => {
    if (stopped) return;

    try {
      await ensureExecutor();

      const response = await ensureOk(await bridgeFetch(`/api/office-tools/poll?executorId=${encodeURIComponent(executorId)}`, {
        headers: headers(),
      }));
      const data = officeToolPollSchema.parse(await response.json());

      if (data.request) {
        try {
          const result = await executor.execute(data.request.toolName, data.request.args || {});
          await ensureOk(await bridgeFetch(`/api/office-tools/respond/${data.request.id}`, {
            method: "POST",
            headers: headers(),
            body: JSON.stringify({ result }),
          }));
        } catch (error) {
          await ensureOk(await bridgeFetch(`/api/office-tools/respond/${data.request.id}`, {
            method: "POST",
            headers: headers(),
            body: JSON.stringify({ error: error instanceof Error ? error.message : String(error) }),
          }));
        }
      }
    } catch {
      executorId = "";
      sessionToken = "";
    }

    timer = window.setTimeout(tick, 750);
  };

  void tick();

  return {
    stop: async () => {
      stopped = true;
      window.clearTimeout(timer);
      if (executorId && sessionToken) {
        await bridgeFetch(`/api/office-tools/register/${executorId}`, {
          method: "DELETE",
          headers: headers(),
        }).catch(() => undefined);
      }
    },
  };
}
