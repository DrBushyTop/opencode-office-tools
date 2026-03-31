import { beforeEach, describe, expect, it, vi } from "vitest";

const {
  requirementSupportMock,
  loadThemeColorsMock,
  loadTextFramesMock,
} = vi.hoisted(() => ({
  requirementSupportMock: vi.fn((_version: string) => true),
  loadThemeColorsMock: vi.fn(),
  loadTextFramesMock: vi.fn(),
}));

vi.mock("./powerpointShared", () => ({
  isPowerPointRequirementSetSupported: requirementSupportMock,
  loadThemeColors: loadThemeColorsMock,
  normalizeHexColor: vi.fn((value: string) => {
    const trimmed = value.trim();
    return trimmed.startsWith("#") ? trimmed : /^[0-9A-Fa-f]{6}$/.test(trimmed) ? `#${trimmed}` : trimmed;
  }),
  parseColor: vi.fn((value: string | null | undefined) => {
    if (!value) return "(none)";
    return value.startsWith("#") ? value : /^[0-9A-Fa-f]{6}$/.test(value) ? `#${value}` : value;
  }),
  readOfficeValue: vi.fn((reader: () => unknown, fallback: unknown) => {
    try {
      const value = reader();
      return value === undefined ? fallback : value;
    } catch {
      return fallback;
    }
  }),
  toolFailure: vi.fn((error: unknown, hint?: string) => {
    const message = error instanceof Error ? error.message : String(error);
    const fullMessage = hint ? `${message} ${hint}` : message;
    return { resultType: "failure", error: fullMessage, textResultForLlm: fullMessage, toolTelemetry: {} };
  }),
}));

vi.mock("./powerpointText", () => ({
  loadTextFrames: loadTextFramesMock,
}));

import { editSlideMaster } from "./editSlideMaster";

function createMaster(overrides: Record<string, unknown> = {}) {
  return {
    id: "master-1",
    name: "Corporate",
    load: vi.fn(),
    themeColorScheme: {
      setThemeColor: vi.fn(),
    },
    shapes: {
      items: [],
      load: vi.fn(),
      addTextBox: vi.fn(),
      addGeometricShape: vi.fn(),
    },
    ...overrides,
  };
}

describe("editSlideMaster", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    requirementSupportMock.mockReset();
    requirementSupportMock.mockReturnValue(true);
    loadThemeColorsMock.mockReset();
    loadTextFramesMock.mockReset();
  });

  it("updates theme colors on the default slide master and returns structured changes", async () => {
    const master = createMaster();
    const context = {
      presentation: {
        slideMasters: { items: [master], load: vi.fn() },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    } as unknown as PowerPoint.RequestContext;

    loadThemeColorsMock
      .mockResolvedValueOnce({ Accent1: "#112233", Hyperlink: "#445566" })
      .mockResolvedValueOnce({ Accent1: "#223344", Hyperlink: "#556677" });

    vi.stubGlobal("PowerPoint", {
      run: vi.fn(async (callback: (ctx: PowerPoint.RequestContext) => Promise<unknown>) => callback(context)),
    });

    const result = await editSlideMaster.handler({
      themeColors: {
        Accent1: "223344",
        Hyperlink: "#556677",
      },
    });

    expect(master.themeColorScheme.setThemeColor).toHaveBeenCalledWith("Accent1", "#223344");
    expect(master.themeColorScheme.setThemeColor).toHaveBeenCalledWith("Hyperlink", "#556677");
    expect(result).toMatchObject({
      resultType: "success",
      slideMasterId: "master-1",
      slideMasterName: "Corporate",
      changedThemeColors: [
        { color: "Accent1", previous: "#112233", next: "#223344" },
        { color: "Hyperlink", previous: "#445566", next: "#556677" },
      ],
      decorativeShapeResults: [],
    });
  });

  it("creates a decorative text box on the requested slide master", async () => {
    const createdShape = {
      id: "shape-new",
      name: "Created shape",
      type: "TextBox",
      left: 10,
      top: 12,
      width: 120,
      height: 28,
      visible: true,
      load: vi.fn(),
      fill: { setSolidColor: vi.fn() },
      lineFormat: { color: "", weight: 0 },
    } as unknown as PowerPoint.Shape;
    const textFrame = {
      isNullObject: false,
      hasText: true,
      textRange: {
        text: "Brand bar",
        font: { color: "", size: 0 },
      },
    } as unknown as PowerPoint.TextFrame;
    const targetMaster = createMaster({
      id: "master-2",
      name: "Alt",
      shapes: {
        items: [],
        load: vi.fn(),
        addTextBox: vi.fn(() => createdShape),
        addGeometricShape: vi.fn(),
      },
    });
    const context = {
      presentation: {
        slideMasters: { items: [createMaster(), targetMaster], load: vi.fn() },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    } as unknown as PowerPoint.RequestContext;

    loadThemeColorsMock.mockResolvedValue({ Accent1: "#112233" });
    loadTextFramesMock.mockResolvedValue([textFrame]);

    vi.stubGlobal("PowerPoint", {
      run: vi.fn(async (callback: (ctx: PowerPoint.RequestContext) => Promise<unknown>) => callback(context)),
    });

    const result = await editSlideMaster.handler({
      slideMasterId: "master-2",
      decorativeShapes: [{
        action: "create",
        shapeType: "textBox",
        text: "Brand bar",
        name: "Header ribbon",
        left: 10,
        top: 12,
        width: 120,
        height: 28,
        fillColor: "123456",
        fontColor: "ABCDEF",
        fontSize: 20,
      }],
    });

    expect(targetMaster.shapes.addTextBox).toHaveBeenCalledWith("Brand bar", {
      left: 10,
      top: 12,
      width: 120,
      height: 28,
    });
    expect(createdShape.fill.setSolidColor).toHaveBeenCalledWith("#123456");
    expect((textFrame.textRange.font as { color: string }).color).toBe("#ABCDEF");
    expect((textFrame.textRange.font as { size: number }).size).toBe(20);
    expect((createdShape as { name: string }).name).toBe("Header ribbon");
    expect(result).toMatchObject({
      resultType: "success",
      slideMasterId: "master-2",
      decorativeShapeResults: [
        {
          action: "create",
          shapeId: "shape-new",
          name: "Header ribbon",
        },
      ],
    });
  });

  it("fails clearly when the requested slide master does not exist", async () => {
    const context = {
      presentation: {
        slideMasters: { items: [createMaster(), createMaster({ id: "master-2", name: "Alt" })], load: vi.fn() },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    } as unknown as PowerPoint.RequestContext;

    vi.stubGlobal("PowerPoint", {
      run: vi.fn(async (callback: (ctx: PowerPoint.RequestContext) => Promise<unknown>) => callback(context)),
    });

    const result = await editSlideMaster.handler({
      slideMasterId: "missing-master",
      themeColors: { Accent1: "#123456" },
    });

    expect(result).toMatchObject({
      resultType: "failure",
      error: expect.stringContaining("Slide master \"missing-master\" was not found"),
    });
    expect(result).toMatchObject({
      error: expect.stringContaining("master-1"),
    });
  });

  it("fails clearly for decorative shape editing when the host lacks shape creation support", async () => {
    requirementSupportMock.mockImplementation((version: string) => version !== "1.4");

    const context = {
      presentation: {
        slideMasters: { items: [createMaster()], load: vi.fn() },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    } as unknown as PowerPoint.RequestContext;

    vi.stubGlobal("PowerPoint", {
      run: vi.fn(async (callback: (ctx: PowerPoint.RequestContext) => Promise<unknown>) => callback(context)),
    });

    const result = await editSlideMaster.handler({
      decorativeShapes: [{
        action: "create",
        shapeType: "textBox",
        text: "Brand bar",
      }],
    });

    expect(result).toMatchObject({
      resultType: "failure",
      error: "Creating decorative master shapes requires PowerPointApi 1.4.",
    });
  });

  it("does not apply theme changes when preflight shape validation fails later in the request", async () => {
    const master = createMaster();
    const context = {
      presentation: {
        slideMasters: { items: [master], load: vi.fn() },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    } as unknown as PowerPoint.RequestContext;

    vi.stubGlobal("PowerPoint", {
      run: vi.fn(async (callback: (ctx: PowerPoint.RequestContext) => Promise<unknown>) => callback(context)),
    });

    const result = await editSlideMaster.handler({
      themeColors: { Accent1: "#123456" },
      decorativeShapes: [{
        action: "update",
        shapeId: "missing-shape",
        name: "Updated",
      }],
    });

    expect(result).toMatchObject({
      resultType: "failure",
      error: expect.stringContaining("missing-shape"),
    });
    expect(master.themeColorScheme.setThemeColor).not.toHaveBeenCalled();
  });

  it("rejects invalid geometric shape types before calling the host API", async () => {
    const master = createMaster();
    const context = {
      presentation: {
        slideMasters: { items: [master], load: vi.fn() },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    } as unknown as PowerPoint.RequestContext;

    vi.stubGlobal("PowerPoint", {
      GeometricShapeType: {
        Rectangle: "Rectangle",
        Ellipse: "Ellipse",
      },
      run: vi.fn(async (callback: (ctx: PowerPoint.RequestContext) => Promise<unknown>) => callback(context)),
    });

    const result = await editSlideMaster.handler({
      decorativeShapes: [{
        action: "create",
        shapeType: "geometricShape",
        geometricShapeType: "NotARealShape",
      }],
    });

    expect(result).toMatchObject({
      resultType: "failure",
      error: expect.stringContaining("geometricShapeType must be a valid PowerPoint.GeometricShapeType value"),
    });
    expect(master.shapes.addGeometricShape).not.toHaveBeenCalled();
  });

  it("rejects blank theme color values instead of silently ignoring them", async () => {
    const result = await editSlideMaster.handler({
      themeColors: { Accent1: "   " },
    });

    expect(result).toMatchObject({
      resultType: "failure",
      error: "Color must be a 6-digit hex value like #123ABC.",
    });
  });
});
