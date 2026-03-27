import { describe, expect, it, vi } from "vitest";
import { manageSlide } from "./manageSlide";

describe("manageSlide", () => {
  it("rejects create on hosts without PowerPointApi 1.3", async () => {
    const contextStub = {
      presentation: { slides: { load: vi.fn(), items: [] } },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    const requirementsStub = { isSetSupported: vi.fn().mockReturnValue(false) };
    const runStub = vi.fn(async (callback: (context: typeof contextStub) => Promise<unknown>) => callback(contextStub));

    vi.stubGlobal("Office", { context: { requirements: requirementsStub } });
    vi.stubGlobal("PowerPoint", { run: runStub });

    await expect(manageSlide.handler({ action: "create" })).resolves.toMatchObject({
      resultType: "failure",
      error: "Creating slides requires PowerPointApi 1.3.",
    });

    expect(requirementsStub.isSetSupported).toHaveBeenCalledWith("PowerPointApi", "1.3");
  });

  it("rejects clear on hosts without PowerPointApi 1.3", async () => {
    const contextStub = {
      presentation: { slides: { load: vi.fn(), items: [] } },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    const requirementsStub = { isSetSupported: vi.fn().mockReturnValue(false) };
    const runStub = vi.fn(async (callback: (context: typeof contextStub) => Promise<unknown>) => callback(contextStub));

    vi.stubGlobal("Office", { context: { requirements: requirementsStub } });
    vi.stubGlobal("PowerPoint", { run: runStub });

    await expect(manageSlide.handler({ action: "clear", slideIndex: 0 })).resolves.toMatchObject({
      resultType: "failure",
      error: "Clearing slide shapes requires PowerPointApi 1.3.",
    });

    expect(requirementsStub.isSetSupported).toHaveBeenCalledWith("PowerPointApi", "1.3");
  });
});
