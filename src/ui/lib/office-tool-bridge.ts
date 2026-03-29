import type { OfficeHost } from "../sessionStorage";
import type { PowerPointContextSnapshot } from "../tools/powerpointContext";

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

  const headers = () => ({
    "Content-Type": "application/json",
    "x-office-bridge-session": sessionToken,
    "x-office-executor-id": executorId,
  });

  const ensureSession = async () => {
    if (sessionToken) return;
    const response = await fetch("/api/office-tools/session");
    if (!response.ok) {
      throw new Error(`Failed to create Office bridge session: ${response.statusText}`);
    }
    const data = await response.json();
    sessionToken = String(data.sessionToken || "");
    if (!sessionToken) {
      throw new Error("Office bridge session token missing");
    }
  };

  const ensureExecutor = async () => {
    if (executorId) return;
    await ensureSession();
    const response = await fetch("/api/office-tools/register", {
      method: "POST",
      headers: headers(),
      body: JSON.stringify({ host }),
    });
    if (!response.ok) {
      throw new Error(`Failed to register Office executor: ${response.statusText}`);
    }
    const data = await response.json();
    executorId = String(data.executorId || "");
    if (!executorId) {
      throw new Error("Office executor id missing");
    }
  };

  const ensureOk = async (response: Response) => {
    if (response.ok) return response;
    throw new Error(await response.text() || response.statusText);
  };

  const tick = async () => {
    if (stopped) return;

    try {
      await ensureExecutor();

      const response = await ensureOk(await fetch(`/api/office-tools/poll?executorId=${encodeURIComponent(executorId)}`, {
        headers: headers(),
      }));
      const data = await response.json();

      if (data.request) {
        try {
          const result = await executor.execute(data.request.toolName, data.request.args || {});
          await ensureOk(await fetch(`/api/office-tools/respond/${data.request.id}`, {
            method: "POST",
            headers: headers(),
            body: JSON.stringify({ result }),
          }));
        } catch (error) {
          await ensureOk(await fetch(`/api/office-tools/respond/${data.request.id}`, {
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
        await fetch(`/api/office-tools/register/${executorId}`, {
          method: "DELETE",
          headers: headers(),
        }).catch(() => undefined);
      }
    },
  };
}
