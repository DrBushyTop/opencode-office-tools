import { describe, expect, it, vi } from "vitest";
import { editSlideWithCode, normalizeEditSlideCode, runEditSlideCode } from "./editSlideWithCode";
import { setPowerPointContextSnapshot } from "./powerpointContext";

describe("editSlideWithCode helpers", () => {
  it("removes code fences around live edit code", () => {
    const code = normalizeEditSlideCode(["```js", "targetShape.left = 24;", "```"].join("\n"));

    expect(code).toBe("targetShape.left = 24;");
  });

  it("runs async code with injected live slide bindings", async () => {
    const context = { sync: vi.fn().mockResolvedValue(undefined) } as unknown as PowerPoint.RequestContext;
    const targetShape = { left: 0 } as unknown as PowerPoint.Shape & { left: number };
    const slide = {} as PowerPoint.Slide;
    const shapes = {} as PowerPoint.ShapeCollection;
    const powerPointApi = { sentinel: true } as unknown as typeof PowerPoint;

    vi.stubGlobal("PowerPoint", powerPointApi);

    const result = await runEditSlideCode(
      "targetShape.left = 24; return `updated ${slideIndex}`;",
      {
        context,
        slide,
        shapes,
        targetShape,
        targetShapeId: "shape-1",
        targetShapeIndex: 0,
        slideIndex: 2,
      },
    );

    expect(targetShape.left).toBe(24);
    expect(result).toBe("updated 2");
  });
});

describe("editSlideWithCode", () => {
  it("uses the active slide and selected shape context when omitted", async () => {
    const shape = { id: "shape-1", name: "Title", left: 0 } as unknown as PowerPoint.Shape & { id: string; name: string; left: number };
    const slide = {
      shapes: {
        items: [shape],
        load: vi.fn(),
      },
    };
    const contextStub = {
      presentation: { slides: { load: vi.fn(), items: [slide] } },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    const requirementsStub = {
      isSetSupported: vi.fn((setName: string, version: string) => setName === "PowerPointApi" && version === "1.3"),
    };
    const runStub = vi.fn(async (callback: (context: typeof contextStub) => Promise<unknown>) => callback(contextStub));

    vi.stubGlobal("Office", { context: { requirements: requirementsStub } });
    vi.stubGlobal("PowerPoint", { run: runStub });
    setPowerPointContextSnapshot({ selectedSlideIds: ["slide-1"], selectedShapeIds: ["shape-1"], activeSlideId: "slide-1", activeSlideIndex: 0 });

    await expect(editSlideWithCode.handler({ code: "targetShape.left = 24;" })).resolves.toMatchObject({
      resultType: "success",
      slideIndex: 0,
      targetShapeId: "shape-1",
    });
    expect(shape.left).toBe(24);

    setPowerPointContextSnapshot(null);
  });

  it("rejects execution when no slide can be inferred", async () => {
    const contextStub = {
      presentation: { slides: { load: vi.fn(), items: [{}] } },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    const requirementsStub = {
      isSetSupported: vi.fn((setName: string, version: string) => setName === "PowerPointApi" && version === "1.3"),
    };
    const runStub = vi.fn(async (callback: (context: typeof contextStub) => Promise<unknown>) => callback(contextStub));

    vi.stubGlobal("Office", { context: { requirements: requirementsStub } });
    vi.stubGlobal("PowerPoint", { run: runStub });
    setPowerPointContextSnapshot(null);

    await expect(editSlideWithCode.handler({ code: "return 'noop';" })).resolves.toMatchObject({
      resultType: "failure",
      error: "slideIndex is required when no active slide can be inferred from the current PowerPoint context.",
    });
  });
});
