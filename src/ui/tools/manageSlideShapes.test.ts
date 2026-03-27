import { describe, expect, it, vi } from "vitest";
import { manageSlideShapes } from "./manageSlideShapes";

describe("manageSlideShapes", () => {
  it("rejects create on hosts without PowerPointApi 1.4", async () => {
    const slide = {};
    const contextStub = {
      presentation: { slides: { load: vi.fn(), items: [slide] } },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    const requirementsStub = {
      isSetSupported: vi.fn((setName: string, version: string) => setName === "PowerPointApi" && version !== "1.4"),
    };
    const runStub = vi.fn(async (callback: (context: typeof contextStub) => Promise<unknown>) => callback(contextStub));

    vi.stubGlobal("Office", { context: { requirements: requirementsStub } });
    vi.stubGlobal("PowerPoint", { run: runStub });

    await expect(manageSlideShapes.handler({ action: "create", slideIndex: 0, shapeType: "textBox", text: "Hello" })).resolves.toMatchObject({
      resultType: "failure",
      error: "Creating shapes requires PowerPointApi 1.4.",
    });

    expect(requirementsStub.isSetSupported).toHaveBeenCalledWith("PowerPointApi", "1.4");
  });

  it("rejects geometric shape create cleanly before touching host enums on unsupported hosts", async () => {
    const slide = {};
    const contextStub = {
      presentation: { slides: { load: vi.fn(), items: [slide] } },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    const requirementsStub = { isSetSupported: vi.fn().mockReturnValue(false) };
    const runStub = vi.fn(async (callback: (context: typeof contextStub) => Promise<unknown>) => callback(contextStub));

    vi.stubGlobal("Office", { context: { requirements: requirementsStub } });
    vi.stubGlobal("PowerPoint", { run: runStub });

    await expect(manageSlideShapes.handler({
      action: "create",
      slideIndex: 0,
      shapeType: "geometricShape",
      geometricShapeType: "Rectangle",
    })).resolves.toMatchObject({
      resultType: "failure",
      error: "Creating shapes requires PowerPointApi 1.4.",
    });
  });

  it("rejects update on hosts without PowerPointApi 1.3", async () => {
    const requirementsStub = { isSetSupported: vi.fn().mockReturnValue(false) };

    vi.stubGlobal("Office", { context: { requirements: requirementsStub } });

    await expect(manageSlideShapes.handler({
      action: "update",
      slideIndex: 0,
      shapeId: "shape-1",
      left: 10,
    })).resolves.toMatchObject({
      resultType: "failure",
      error: "Updating or deleting shapes requires PowerPointApi 1.3.",
    });

    expect(requirementsStub.isSetSupported).toHaveBeenCalledWith("PowerPointApi", "1.3");
  });
});
