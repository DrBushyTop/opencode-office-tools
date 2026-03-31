import { beforeEach, describe, expect, it, vi } from "vitest";

const { loadPresentationLayoutCatalogFromDocumentMock, lookupPresentationLayoutMetadataMock } = vi.hoisted(() => ({
  loadPresentationLayoutCatalogFromDocumentMock: vi.fn<() => Promise<unknown | null>>(async () => null),
  lookupPresentationLayoutMetadataMock: vi.fn<(catalog: unknown, options: { layoutId?: string; slideMasterId?: string }) => unknown | null>(() => null),
}));

vi.mock("./powerpointShared", () => ({
  isPowerPointRequirementSetSupported: vi.fn(() => true),
  readOfficeValue: vi.fn((reader: () => unknown, fallback: unknown) => {
    try {
      const value = reader();
      return value ?? fallback;
    } catch {
      return fallback;
    }
  }),
  toolFailure: vi.fn((error: unknown) => ({ resultType: "failure", error: error instanceof Error ? error.message : String(error) })),
}));

vi.mock("./powerpointLayoutCatalog", () => ({
  loadPresentationLayoutCatalogFromDocument: loadPresentationLayoutCatalogFromDocumentMock,
  lookupPresentationLayoutMetadata: lookupPresentationLayoutMetadataMock,
  resolveSlideLayoutMetadata: vi.fn((officeLayoutName: string, officeLayoutType: string, fallback?: { layoutName?: string; layoutType?: string } | null) => ({
    layoutName: officeLayoutName || fallback?.layoutName || (officeLayoutType === "TitleAndContent" ? "Title and Content" : officeLayoutType || fallback?.layoutType || ""),
    layoutType: officeLayoutType || fallback?.layoutType || "Unknown",
  })),
}));

import { listSlideLayouts } from "./listSlideLayouts";

describe("listSlideLayouts", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    loadPresentationLayoutCatalogFromDocumentMock.mockResolvedValue(null);
    lookupPresentationLayoutMetadataMock.mockReturnValue(null);
  });

  it("returns a concise grouped overview", async () => {
    const masterA = {
      id: "master-1",
      name: "Corporate",
      layouts: {
        items: [
          { id: "layout-1", name: "Title and Content", type: "TitleAndContent", load: vi.fn() },
          { id: "layout-2", name: "Section Header", type: "SectionHeader", load: vi.fn() },
        ],
      },
      load: vi.fn(),
    };
    const masterB = {
      id: "master-2",
      name: "Alt",
      layouts: {
        items: [{ id: "layout-3", name: "Blank", type: "Blank", load: vi.fn() }],
      },
      load: vi.fn(),
    };
    const context = {
      presentation: {
        slideMasters: { items: [masterA, masterB], load: vi.fn() },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    } as unknown as PowerPoint.RequestContext;

    vi.stubGlobal("PowerPoint", {
      run: vi.fn(async (callback: (ctx: PowerPoint.RequestContext) => Promise<unknown>) => callback(context)),
    });

    const result = await listSlideLayouts.handler();

    expect(result).toMatchObject({
      resultType: "success",
      slideMasterCount: 2,
      layoutCount: 3,
      slideMasters: [
        {
          slideMasterId: "master-1",
          slideMasterName: "Corporate",
          layoutCount: 2,
          layouts: [
            { layoutId: "layout-1", layoutName: "Title and Content", layoutType: "TitleAndContent" },
            { layoutId: "layout-2", layoutName: "Section Header", layoutType: "SectionHeader" },
          ],
        },
        {
          slideMasterId: "master-2",
          slideMasterName: "Alt",
          layoutCount: 1,
          layouts: [
            { layoutId: "layout-3", layoutName: "Blank", layoutType: "Blank" },
          ],
        },
      ],
      layouts: [
        { layoutId: "layout-1", layoutName: "Title and Content", layoutType: "TitleAndContent" },
        { layoutId: "layout-2", layoutName: "Section Header", layoutType: "SectionHeader" },
        { layoutId: "layout-3", layoutName: "Blank", layoutType: "Blank" },
      ],
    });
    expect(result).toMatchObject({
      textResultForLlm: expect.stringContaining("Found 3 layouts across 2 slide masters."),
    });
  });

  it("falls back to Open XML layout metadata when Office returns blanks", async () => {
    loadPresentationLayoutCatalogFromDocumentMock.mockResolvedValue({ slideMasters: [] });
    lookupPresentationLayoutMetadataMock.mockImplementation((_catalog: unknown, options: { layoutId?: string; slideMasterId?: string }) => {
      if (!options.layoutId) return { slideMasterName: "Corporate XML" };
      return {
        slideMasterName: "Corporate XML",
        layoutName: "Light: Title and Subtitle",
        layoutType: "Title",
      };
    });

    const layout = {
      id: "2147483700#2384306777",
      name: "",
      load: vi.fn(),
      get type() {
        throw new Error("type not available");
      },
    };
    const master = {
      id: "2147483698#626277182",
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
      slideMasters: [
        {
          layouts: [
            {
              layoutId: "2147483700#2384306777",
              layoutName: "Light: Title and Subtitle",
              layoutType: "Title",
            },
          ],
        },
      ],
    });
  });

  it("loads layout metadata directly so host type values are available", async () => {
    const layout = {
      id: "layout-1",
      name: "Overview",
      type: "Custom",
      load: vi.fn(),
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

    expect(layout.load).toHaveBeenCalled();
    expect(result).toMatchObject({
      layouts: [{ layoutId: "layout-1", layoutName: "Overview", layoutType: "Custom" }],
    });
  });
});
