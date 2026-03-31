import { beforeEach, describe, expect, it, vi } from "vitest";

const {
  loadTextFramesMock,
  fetchImageUrlAsBase64Mock,
  getShapeBoundsMock,
  createImageRectangleMock,
  toPowerPointTableValuesMock,
} = vi.hoisted(() => ({
  loadTextFramesMock: vi.fn(),
  fetchImageUrlAsBase64Mock: vi.fn(),
  getShapeBoundsMock: vi.fn(),
  createImageRectangleMock: vi.fn(),
  toPowerPointTableValuesMock: vi.fn((values: Array<Array<string | number | boolean>>) => values.map((row) => row.map((cell) => String(cell)))),
}));

vi.mock("./powerpointShared", () => ({
  isPowerPointRequirementSetSupported: vi.fn((version: string) => version === "1.3" || version === "1.8"),
  readOfficeValue: vi.fn((reader: () => unknown, fallback: unknown) => {
    try {
      const value = reader();
      return value === undefined ? fallback : value;
    } catch {
      return fallback;
    }
  }),
  toolFailure: vi.fn((error: unknown) => ({ resultType: "failure", error: error instanceof Error ? error.message : String(error) })),
}));

vi.mock("./powerpointText", () => ({
  loadTextFrames: loadTextFramesMock,
}));

vi.mock("./powerpointNativeContent", () => ({
  createImageRectangle: createImageRectangleMock,
  fetchImageUrlAsBase64: fetchImageUrlAsBase64Mock,
  getShapeBounds: getShapeBoundsMock,
  toPowerPointTableValues: toPowerPointTableValuesMock,
}));

import { createSlideFromLayout } from "./createSlideFromLayout";

function createTextFrame() {
  return {
    isNullObject: false,
    hasText: false,
    load: vi.fn(),
    textRange: {
      text: "",
      load: vi.fn(),
    },
  };
}

describe("createSlideFromLayout", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    loadTextFramesMock.mockReset();
    fetchImageUrlAsBase64Mock.mockReset();
    getShapeBoundsMock.mockReset();
    createImageRectangleMock.mockReset();
    toPowerPointTableValuesMock.mockClear();
  });

  it("creates a slide from a layout and writes text bindings into placeholders", async () => {
    const titleFrame = createTextFrame();
    const titlePlaceholder = {
      id: "ph-title",
      name: "Title 1",
      type: "placeholder",
      placeholderFormat: { type: "Title", containedType: null, load: vi.fn() },
      getTextFrameOrNullObject: vi.fn(() => titleFrame),
      load: vi.fn(),
    };
    const createdSlide = {
      id: "slide-created",
      load: vi.fn(),
      delete: vi.fn(),
      moveTo: vi.fn(),
      shapes: {
        items: [titlePlaceholder],
        load: vi.fn(),
        addTable: vi.fn(),
      },
    };
    let slideWasCreated = false;
    const addSlide = vi.fn(() => {
      slideWasCreated = true;
      return createdSlide;
    });
    const slides = {
      items: [{ id: "slide-0" }],
      load: vi.fn(),
      add: addSlide,
    };
    const context = {
      presentation: { slides },
      sync: vi.fn(async () => {
        if (slideWasCreated && slides.items.length === 1) {
          slides.items.push(createdSlide);
        }
      }),
    } as unknown as PowerPoint.RequestContext;

    loadTextFramesMock.mockResolvedValue([titleFrame]);

    vi.stubGlobal("Office", { context: { requirements: { isSetSupported: vi.fn(() => true) } } });
    vi.stubGlobal("PowerPoint", {
      ShapeType: { placeholder: "placeholder" },
      run: vi.fn(async (callback: (ctx: PowerPoint.RequestContext) => Promise<unknown>) => callback(context)),
    });

    const result = await createSlideFromLayout.handler({
      layoutId: "layout-1",
      targetIndex: 0,
      bindings: [{ placeholderType: "Title", text: "Q4 plan" }],
    });

    expect(addSlide).toHaveBeenCalledWith({ layoutId: "layout-1" });
    expect(createdSlide.moveTo).toHaveBeenCalledWith(0);
    expect(titleFrame.textRange.text).toBe("Q4 plan");
    expect(result).toMatchObject({
      resultType: "success",
      slideId: "slide-created",
      slideIndex: 0,
      layoutId: "layout-1",
      appliedBindings: [
        {
          bindingType: "text",
          placeholderType: "Title",
          placeholderName: "Title 1",
          text: "Q4 plan",
        },
      ],
    });
  });

  it("maps image and table bindings into placeholder bounds", async () => {
    const imagePlaceholder = {
      id: "ph-image",
      name: "Picture Placeholder 2",
      type: "placeholder",
      placeholderFormat: { type: "Picture", containedType: null, load: vi.fn() },
      delete: vi.fn(),
      getTextFrameOrNullObject: vi.fn(),
      load: vi.fn(),
    };
    const tablePlaceholder = {
      id: "ph-table",
      name: "Content Placeholder 3",
      type: "placeholder",
      placeholderFormat: { type: "Body", containedType: "Table", load: vi.fn() },
      delete: vi.fn(),
      getTextFrameOrNullObject: vi.fn(),
      load: vi.fn(),
    };
    const createdImageShape = { id: "image-shape", name: "Picture Placeholder 2", load: vi.fn() };
    const createdTableShape = { id: "table-shape", name: "Content Placeholder 3", load: vi.fn() };
    const addTable = vi.fn(() => createdTableShape);
    const createdSlide = {
      id: "slide-created",
      load: vi.fn(),
      delete: vi.fn(),
      moveTo: vi.fn(),
      shapes: {
        items: [imagePlaceholder, tablePlaceholder],
        load: vi.fn(),
        addTable,
      },
    };
    let slideWasCreated = false;
    const addSlide = vi.fn(() => {
      slideWasCreated = true;
      return createdSlide;
    });
    const slides = {
      items: [] as Array<{ id: string } | typeof createdSlide>,
      load: vi.fn(),
      add: addSlide,
    };
    const context = {
      presentation: { slides },
      sync: vi.fn(async () => {
        if (slideWasCreated && slides.items.length === 0) {
          slides.items.push(createdSlide);
        }
      }),
    } as unknown as PowerPoint.RequestContext;

    fetchImageUrlAsBase64Mock.mockResolvedValue("IMAGE64");
    getShapeBoundsMock.mockImplementation(async (shape: { id: string; name: string }) => shape.id === "ph-image"
      ? { left: 10, top: 20, width: 300, height: 180, name: shape.name, id: shape.id }
      : { left: 40, top: 220, width: 400, height: 120, name: shape.name, id: shape.id });
    createImageRectangleMock.mockReturnValue(createdImageShape);

    vi.stubGlobal("Office", { context: { requirements: { isSetSupported: vi.fn(() => true) } } });
    vi.stubGlobal("PowerPoint", {
      ShapeType: { placeholder: "placeholder" },
      run: vi.fn(async (callback: (ctx: PowerPoint.RequestContext) => Promise<unknown>) => callback(context)),
    });

    const result = await createSlideFromLayout.handler({
      layoutId: "layout-visual",
      bindings: [
        { placeholderName: "Picture Placeholder 2", imageUrl: "https://example.com/photo.png" },
        { placeholderName: "Content Placeholder 3", tableData: [["Region", "Revenue"], ["North", 42]] },
      ],
    });

    expect(fetchImageUrlAsBase64Mock).toHaveBeenCalledWith("https://example.com/photo.png");
    expect(createImageRectangleMock).toHaveBeenCalledWith(createdSlide, {
      left: 10,
      top: 20,
      width: 300,
      height: 180,
      name: "Picture Placeholder 2",
      imageBase64: "IMAGE64",
    });
    expect(addTable).toHaveBeenCalledWith(2, 2, {
      values: [["Region", "Revenue"], ["North", "42"]],
      left: 40,
      top: 220,
      width: 400,
      height: 120,
    });
    expect(imagePlaceholder.delete).toHaveBeenCalled();
    expect(tablePlaceholder.delete).toHaveBeenCalled();
    expect(result).toMatchObject({
      resultType: "success",
      slideIndex: 0,
      appliedBindings: [
        { bindingType: "image", shapeId: "image-shape", placeholderName: "Picture Placeholder 2" },
        { bindingType: "table", shapeId: "table-shape", placeholderName: "Content Placeholder 3" },
      ],
    });
  });

  it("consumes matching placeholders so repeated placeholderType bindings target distinct shapes", async () => {
    const bodyFrameA = createTextFrame();
    const bodyFrameB = createTextFrame();
    const bodyPlaceholderA = {
      id: "ph-body-a",
      name: "Body 1",
      type: "placeholder",
      placeholderFormat: { type: "Body", containedType: "Text", load: vi.fn() },
      getTextFrameOrNullObject: vi.fn(() => bodyFrameA),
      load: vi.fn(),
    };
    const bodyPlaceholderB = {
      id: "ph-body-b",
      name: "Body 2",
      type: "placeholder",
      placeholderFormat: { type: "Body", containedType: "Text", load: vi.fn() },
      getTextFrameOrNullObject: vi.fn(() => bodyFrameB),
      load: vi.fn(),
    };
    const createdSlide = {
      id: "slide-created",
      load: vi.fn(),
      delete: vi.fn(),
      moveTo: vi.fn(),
      shapes: {
        items: [bodyPlaceholderA, bodyPlaceholderB],
        load: vi.fn(),
        addTable: vi.fn(),
      },
    };
    let slideWasCreated = false;
    const slides = {
      items: [] as Array<{ id: string } | typeof createdSlide>,
      load: vi.fn(),
      add: vi.fn(() => {
        slideWasCreated = true;
        return createdSlide;
      }),
    };
    const context = {
      presentation: { slides },
      sync: vi.fn(async () => {
        if (slideWasCreated && slides.items.length === 0) slides.items.push(createdSlide);
      }),
    } as unknown as PowerPoint.RequestContext;

    loadTextFramesMock.mockImplementation(async (_context, shapes: Array<{ id: string }>) => shapes[0]?.id === "ph-body-a" ? [bodyFrameA] : [bodyFrameB]);

    vi.stubGlobal("Office", { context: { requirements: { isSetSupported: vi.fn(() => true) } } });
    vi.stubGlobal("PowerPoint", {
      ShapeType: { placeholder: "placeholder" },
      run: vi.fn(async (callback: (ctx: PowerPoint.RequestContext) => Promise<unknown>) => callback(context)),
    });

    const result = await createSlideFromLayout.handler({
      layoutId: "layout-1",
      bindings: [
        { placeholderType: "Body", text: "Alpha" },
        { placeholderType: "Body", text: "Beta" },
      ],
    });

    expect(bodyFrameA.textRange.text).toBe("Alpha");
    expect(bodyFrameB.textRange.text).toBe("Beta");
    expect(result).toMatchObject({
      resultType: "success",
      appliedBindings: [
        { bindingType: "text", shapeId: "ph-body-a", text: "Alpha" },
        { bindingType: "text", shapeId: "ph-body-b", text: "Beta" },
      ],
    });
  });

  it("rolls back the created slide when binding resolution fails", async () => {
    const titleFrame = createTextFrame();
    const titlePlaceholder = {
      id: "ph-title",
      name: "Title 1",
      type: "placeholder",
      placeholderFormat: { type: "Title", containedType: null, load: vi.fn() },
      getTextFrameOrNullObject: vi.fn(() => titleFrame),
      load: vi.fn(),
    };
    const createdSlide = {
      id: "slide-created",
      load: vi.fn(),
      delete: vi.fn(),
      moveTo: vi.fn(),
      shapes: {
        items: [titlePlaceholder],
        load: vi.fn(),
        addTable: vi.fn(),
      },
    };
    let slideWasCreated = false;
    const slides = {
      items: [] as Array<{ id: string } | typeof createdSlide>,
      load: vi.fn(),
      add: vi.fn(() => {
        slideWasCreated = true;
        return createdSlide;
      }),
    };
    const context = {
      presentation: { slides },
      sync: vi.fn(async () => {
        if (slideWasCreated && slides.items.length === 0) slides.items.push(createdSlide);
      }),
    } as unknown as PowerPoint.RequestContext;

    loadTextFramesMock.mockResolvedValue([titleFrame]);

    vi.stubGlobal("Office", { context: { requirements: { isSetSupported: vi.fn(() => true) } } });
    vi.stubGlobal("PowerPoint", {
      ShapeType: { placeholder: "placeholder" },
      run: vi.fn(async (callback: (ctx: PowerPoint.RequestContext) => Promise<unknown>) => callback(context)),
    });

    const result = await createSlideFromLayout.handler({
      layoutId: "layout-1",
      bindings: [
        { placeholderType: "Title", text: "Keep me?" },
        { placeholderType: "Body", text: "Missing placeholder" },
      ],
    });

    expect(createdSlide.delete).toHaveBeenCalled();
    expect(result).toMatchObject({
      resultType: "failure",
      error: expect.stringContaining("Created slide was rolled back."),
    });
  });

  it("rejects ragged tableData before mutating the deck", async () => {
    const result = await createSlideFromLayout.handler({
      layoutId: "layout-1",
      bindings: [{ placeholderName: "Table", tableData: [["A", "B"], ["Only one cell"]] }],
    });

    expect(result).toMatchObject({
      resultType: "failure",
      error: "tableData must be a non-empty rectangular 2D array.",
    });
  });
});
