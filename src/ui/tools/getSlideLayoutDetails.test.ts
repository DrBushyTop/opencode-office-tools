import { beforeEach, describe, expect, it, vi } from "vitest";

const {
  loadPresentationLayoutCatalogFromDocumentMock,
  lookupPresentationLayoutMetadataMock,
  supportsPowerPointPlaceholdersMock,
} = vi.hoisted(() => ({
  loadPresentationLayoutCatalogFromDocumentMock: vi.fn<() => Promise<unknown | null>>(async () => null),
  lookupPresentationLayoutMetadataMock: vi.fn<(catalog: unknown, options: { layoutId?: string; slideMasterId?: string }) => unknown | null>(() => null),
  supportsPowerPointPlaceholdersMock: vi.fn(() => true),
}));

vi.mock("./powerpointShared", () => ({
  isPowerPointRequirementSetSupported: vi.fn(() => true),
  loadShapeSummaries: vi.fn(async (_context, shapes: Array<{ id: string; placeholderType?: string; placeholderContainedType?: string | null; left?: number; top?: number; width?: number; height?: number }>) =>
    shapes.map((shape, index) => ({
      index,
      id: shape.id,
      name: shape.id,
      type: "placeholder",
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
      return value ?? fallback;
    } catch {
      return fallback;
    }
  }),
  supportsPowerPointPlaceholders: supportsPowerPointPlaceholdersMock,
  toolFailure: vi.fn((error: unknown) => ({ resultType: "failure", error: error instanceof Error ? error.message : String(error) })),
}));

vi.mock("./powerpointLayoutCatalog", () => ({
  loadPresentationLayoutCatalogFromDocument: loadPresentationLayoutCatalogFromDocumentMock,
  lookupPresentationLayoutMetadata: lookupPresentationLayoutMetadataMock,
  resolveSlideLayoutMetadata: vi.fn((officeLayoutName: string, officeLayoutType: string, fallback?: { layoutName?: string; layoutType?: string } | null) => ({
    layoutName: officeLayoutName || fallback?.layoutName || "",
    layoutType: officeLayoutType || fallback?.layoutType || "Unknown",
  })),
}));

import { getSlideLayoutDetails } from "./getSlideLayoutDetails";

describe("getSlideLayoutDetails", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    loadPresentationLayoutCatalogFromDocumentMock.mockResolvedValue(null);
    lookupPresentationLayoutMetadataMock.mockReturnValue(null);
    supportsPowerPointPlaceholdersMock.mockReturnValue(true);
  });

  it("returns placeholder geometry for a single layout", async () => {
    const layout = {
      id: "layout-1",
      name: "Title and Content",
      type: "TitleAndContent",
      load: vi.fn(),
      shapes: {
        items: [
          { id: "shape-title", placeholderType: "Title", left: 10, top: 20, width: 300, height: 40 },
          { id: "shape-body", placeholderType: "Content", placeholderContainedType: "Text", left: 10, top: 80, width: 300, height: 200 },
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

    const result = await getSlideLayoutDetails.handler({ layoutId: "layout-1" });

    expect(result).toMatchObject({
      resultType: "success",
      slideMasterId: "master-1",
      slideMasterName: "Corporate",
      layoutId: "layout-1",
      layoutName: "Title and Content",
      layoutType: "TitleAndContent",
      placeholderMetadataSupported: true,
      placeholderCount: 2,
      placeholders: [
        { shapeId: "shape-title", placeholderType: "Title", left: 10, top: 20, width: 300, height: 40 },
        { shapeId: "shape-body", placeholderType: "Content", placeholderContainedType: "Text", left: 10, top: 80, width: 300, height: 200 },
      ],
    });
    expect(result).not.toHaveProperty("textResultForLlm");
  });

  it("requires slideMasterId when the same layout id is ambiguous", async () => {
    const makeLayout = () => ({
      id: "layout-1",
      name: "Title and Content",
      type: "TitleAndContent",
      load: vi.fn(),
      shapes: { items: [], load: vi.fn() },
    });
    const context = {
      presentation: {
        slideMasters: {
          items: [
            { id: "master-1", name: "Corporate", layouts: { items: [makeLayout()] }, load: vi.fn() },
            { id: "master-2", name: "Alt", layouts: { items: [makeLayout()] }, load: vi.fn() },
          ],
          load: vi.fn(),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    } as unknown as PowerPoint.RequestContext;

    vi.stubGlobal("PowerPoint", {
      run: vi.fn(async (callback: (ctx: PowerPoint.RequestContext) => Promise<unknown>) => callback(context)),
    });

    const result = await getSlideLayoutDetails.handler({ layoutId: "layout-1" });

    expect(result).toMatchObject({
      resultType: "failure",
      error: expect.stringContaining("ambiguous"),
    });
  });
});
