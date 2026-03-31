import { beforeEach, describe, expect, it, vi } from "vitest";

const { supportsPowerPointPlaceholdersMock } = vi.hoisted(() => ({
  supportsPowerPointPlaceholdersMock: vi.fn(() => true),
}));

vi.mock("./powerpointShared", () => ({
  loadShapeSummaries: vi.fn(async (_context, shapes: Array<{ id: string; name: string; placeholderType?: string; placeholderContainedType?: string; left?: number; top?: number; width?: number; height?: number; text?: string }>) =>
    shapes.map((shape, index) => ({
      index,
      id: shape.id,
      name: shape.name,
      type: "placeholder",
      text: shape.text || "",
      placeholderType: shape.placeholderType,
      placeholderContainedType: shape.placeholderContainedType,
      left: shape.left ?? 0,
      top: shape.top ?? 0,
      width: shape.width ?? 100,
      height: shape.height ?? 20,
    }))),
  readOfficeValue: vi.fn((reader: () => unknown, fallback: unknown) => {
    try {
      const value = reader();
      return value === undefined ? fallback : value;
    } catch {
      return fallback;
    }
  }),
  supportsPowerPointPlaceholders: supportsPowerPointPlaceholdersMock,
  toolFailure: vi.fn((error: unknown) => ({ resultType: "failure", error: error instanceof Error ? error.message : String(error) })),
}));

import { listSlideLayouts } from "./listSlideLayouts";

describe("listSlideLayouts", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    supportsPowerPointPlaceholdersMock.mockReturnValue(true);
  });

  it("returns a flat layout catalog with placeholder inventory", async () => {
    const layout = {
      id: "layout-1",
      name: "Title and Content",
      type: "TitleAndContent",
      shapes: {
        items: [
          { id: "shape-title", name: "Title 1", placeholderType: "Title", text: "Click to add title" },
          { id: "shape-body", name: "Content Placeholder 2", placeholderType: "Body", placeholderContainedType: "Text", text: "Click to add text" },
        ],
        load: vi.fn(),
      },
    };
    const master = {
      id: "master-1",
      name: "Corporate",
      layouts: { items: [layout] },
      load: vi.fn(),
    };
    const context = {
      presentation: {
        slideMasters: { items: [master], load: vi.fn() },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    } as unknown as PowerPoint.RequestContext;

    vi.stubGlobal("PowerPoint", {
      run: vi.fn(async (callback: (ctx: PowerPoint.RequestContext) => Promise<unknown>) => callback(context)),
    });

    const result = await listSlideLayouts.handler();

    expect(result).toMatchObject({
      resultType: "success",
      slideMasterCount: 1,
      layoutCount: 1,
      placeholderMetadataSupported: true,
      layouts: [
        {
          slideMasterId: "master-1",
          slideMasterName: "Corporate",
          layoutId: "layout-1",
          layoutName: "Title and Content",
          layoutType: "TitleAndContent",
          placeholders: [
            {
              shapeId: "shape-title",
              placeholderName: "Title 1",
              placeholderType: "Title",
            },
            {
              shapeId: "shape-body",
              placeholderName: "Content Placeholder 2",
              placeholderType: "Body",
              placeholderContainedType: "Text",
            },
          ],
        },
      ],
    });
    expect(result).toMatchObject({
      textResultForLlm: expect.stringContaining("Found 1 layout across 1 slide master."),
    });
  });

  it("distinguishes unsupported placeholder metadata from missing placeholders", async () => {
    supportsPowerPointPlaceholdersMock.mockReturnValue(false);

    const layout = {
      id: "layout-1",
      name: "Title and Content",
      type: "TitleAndContent",
      shapes: {
        items: [{ id: "shape-title", name: "Title 1", placeholderType: "Title", text: "Click to add title" }],
        load: vi.fn(),
      },
    };
    const master = {
      id: "master-1",
      name: "Corporate",
      layouts: { items: [layout] },
      load: vi.fn(),
    };
    const context = {
      presentation: {
        slideMasters: { items: [master], load: vi.fn() },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    } as unknown as PowerPoint.RequestContext;

    vi.stubGlobal("PowerPoint", {
      run: vi.fn(async (callback: (ctx: PowerPoint.RequestContext) => Promise<unknown>) => callback(context)),
    });

    const result = await listSlideLayouts.handler();

    expect(result).toMatchObject({
      resultType: "success",
      placeholderMetadataSupported: false,
      layouts: [{ placeholders: [] }],
      textResultForLlm: expect.stringContaining("Placeholder metadata is unavailable on this host."),
    });
  });
});
