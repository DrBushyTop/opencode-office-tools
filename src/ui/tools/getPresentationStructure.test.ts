import { describe, expect, it, vi, beforeEach } from "vitest";

vi.mock("./powerpointShared", () => ({
  isPowerPointRequirementSetSupported: vi.fn((version: string) => version === "1.5" || version === "1.10"),
  supportsPowerPointPlaceholders: vi.fn(() => true),
  loadThemeColors: vi.fn(async () => ({ Accent1: "#123456", Accent2: "#abcdef" })),
  loadShapeSummaries: vi.fn(async (_context, shapes: Array<{ id: string; name: string; placeholderType?: string; placeholderContainedType?: string; left?: number; top?: number; width?: number; height?: number; text?: string }>) =>
    shapes.map((shape, index) => ({
      index,
      id: shape.id,
      name: shape.name,
      type: "textBox",
      text: shape.text || shape.name,
      placeholderType: shape.placeholderType,
      placeholderContainedType: shape.placeholderContainedType,
      left: shape.left ?? 0,
      top: shape.top ?? 0,
      width: shape.width ?? 100,
      height: shape.height ?? 20,
    }))),
  parseColor: vi.fn((value: string) => value),
  readOfficeValue: vi.fn((reader: () => unknown, fallback: unknown) => {
    try {
      const value = reader();
      return value === undefined ? fallback : value;
    } catch {
      return fallback;
    }
  }),
  summarizePlainText: vi.fn((text: string) => text),
  toolFailure: vi.fn((error: unknown) => ({ resultType: "failure", error: error instanceof Error ? error.message : String(error) })),
}));

import { getPresentationStructure } from "./getPresentationStructure";

describe("getPresentationStructure", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  it("returns structured template metadata in both mode", async () => {
    const layout = {
      id: "layout-1",
      name: "Timeline",
      type: "TitleAndContent",
      shapes: {
        items: [{ id: "shape-layout-title", name: "Layout Title", placeholderType: "Title", text: "Layout Title" }],
        load: vi.fn(),
      },
      background: {
        load: vi.fn(),
        isMasterBackgroundFollowed: true,
        areBackgroundGraphicsHidden: false,
      },
    };
    const master = {
      id: "master-1",
      name: "Corporate",
      layouts: { items: [layout] },
      shapes: {
        items: [{ id: "shape-master-title", name: "Master Title", placeholderType: "Title", text: "Master Title" }],
        load: vi.fn(),
      },
      background: { fill: { load: vi.fn(), type: "Solid" } },
      themeColorScheme: { getThemeColor: vi.fn(() => ({ value: "#123456" })) },
      load: vi.fn(),
    };
    const selectedSlides = { items: [{ id: "slide-1" }], load: vi.fn() };
    const selectedShapes = { items: [{ id: "shape-1" }], load: vi.fn() };
    const context = {
      presentation: {
        slideMasters: { items: [master], load: vi.fn() },
        slides: { items: [{ id: "slide-1" }, { id: "slide-2" }], load: vi.fn() },
        pageSetup: { slideWidth: 960, slideHeight: 540, load: vi.fn() },
        getSelectedSlides: vi.fn(() => selectedSlides),
        getSelectedShapes: vi.fn(() => selectedShapes),
      },
      sync: vi.fn().mockResolvedValue(undefined),
    } as unknown as PowerPoint.RequestContext;

    vi.stubGlobal("PowerPoint", {
      run: vi.fn(async (callback: (ctx: PowerPoint.RequestContext) => Promise<unknown>) => callback(context)),
    });

    const result = await getPresentationStructure.handler({ format: "both" });
    expect(result).toMatchObject({
      resultType: "success",
      summary: expect.stringContaining("Slides: 2"),
      structure: {
        slideCount: 2,
        activeSlideId: "slide-1",
        activeSlideIndex: 0,
        selectedSlideIds: ["slide-1"],
        selectedShapeIds: ["shape-1"],
        masters: [
          {
            id: "master-1",
            name: "Corporate",
            layouts: [{ id: "layout-1", name: "Timeline" }],
          },
        ],
      },
    });
  });
});
