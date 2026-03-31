import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from "@xmldom/xmldom";
import { strToU8, zipSync } from "fflate";
import { describe, expect, it, vi } from "vitest";
import { manageSlideShapes } from "./manageSlideShapes";
import { setPowerPointContextSnapshot } from "./powerpointContext";

function createPresentationBase64(entries: Record<string, string>) {
  let binary = "";
  zipSync(Object.fromEntries(
    Object.entries(entries).map(([path, contents]) => [path, strToU8(contents)]),
  )).forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary);
}

if (typeof DOMParser === "undefined") {
  vi.stubGlobal("DOMParser", XmldomParser);
}

if (typeof XMLSerializer === "undefined") {
  vi.stubGlobal("XMLSerializer", XmldomSerializer);
}

describe("manageSlideShapes", () => {
  it("uses the active slide and selected shape context when slideIndex and shapeId are omitted", async () => {
    const shape = { id: "shape-1", name: "Title", left: 0 } as unknown as PowerPoint.Shape & { left: number };
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

    await expect(manageSlideShapes.handler({ action: "update", left: 24 })).resolves.toBe("Updated shape on slide 1.");
    expect(shape.left).toBe(24);
    setPowerPointContextSnapshot(null);
  });

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

  it("adds a sparse-update hint for ambiguous InvalidArgument failures", async () => {
    const requirementsStub = {
      isSetSupported: vi.fn((setName: string, version: string) => setName === "PowerPointApi" && version === "1.3"),
    };
    const error = Object.assign(new Error("InvalidArgument"), { code: "InvalidArgument" });
    const runStub = vi.fn(async () => {
      throw error;
    });

    vi.stubGlobal("Office", { context: { requirements: requirementsStub } });
    vi.stubGlobal("PowerPoint", { run: runStub });

    const result = await manageSlideShapes.handler({
      action: "update",
      slideIndex: 0,
      shapeId: "shape-1",
      text: "Click to edit text",
      fontName: "",
      fontColor: "",
    });

    expect(result).toMatchObject({
      resultType: "failure",
    });
    if (result && typeof result === "object" && "error" in result) {
      expect(String(result.error)).toContain("for update, pass only the target");
      expect(String(result.error)).toContain("empty strings");
    }
  });

  it("updates a shape when shapeId matches the exported XML cNvPr id after slide replacement", async () => {
    const shape = { id: "office-body", name: "Body", left: 0 } as unknown as PowerPoint.Shape & { left: number };
    const xmlBase64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cSld>
          <p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr/>
            <p:sp><p:nvSpPr><p:cNvPr id="10" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>
            <p:sp><p:nvSpPr><p:cNvPr id="11" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>
          </p:spTree>
        </p:cSld>
      </p:sld>`,
    });
    const slide = {
      shapes: {
        items: [{ id: "office-title", name: "Title" }, shape],
        load: vi.fn(),
      },
      exportAsBase64: vi.fn(() => ({ value: xmlBase64 })),
    };
    const contextStub = {
      presentation: { slides: { load: vi.fn(), items: [slide] } },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    const requirementsStub = {
      isSetSupported: vi.fn((setName: string, version: string) => setName === "PowerPointApi" && ["1.3", "1.8"].includes(version)),
    };
    const runStub = vi.fn(async (callback: (context: typeof contextStub) => Promise<unknown>) => callback(contextStub));

    vi.stubGlobal("Office", { context: { requirements: requirementsStub } });
    vi.stubGlobal("PowerPoint", { run: runStub });

    await expect(manageSlideShapes.handler({
      action: "update",
      slideIndex: 0,
      shapeId: "11",
      left: 10,
    })).resolves.toBe("Updated shape on slide 1.");

    expect(shape.left).toBe(10);
  });

  it("accepts fontSize zero when forwarding text formatting updates", async () => {
    const font = { size: 12 };
    const frame = {
      isNullObject: false,
      hasText: true,
      textRange: {
        text: "Hello",
        load: vi.fn(),
        font,
        paragraphFormat: {},
      },
      load: vi.fn(),
    };
    const shape = {
      getTextFrameOrNullObject: vi.fn(() => frame),
    } as unknown as PowerPoint.Shape;
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

    await expect(manageSlideShapes.handler({
      action: "update",
      slideIndex: 0,
      shapeIndex: 0,
      fontSize: 0,
    })).resolves.toBe("Updated shape on slide 1.");

    expect(font.size).toBe(0);
  });

  it("sanitizes dense echo-style update payloads before applying the patch", async () => {
    const fill = {
      clear: vi.fn(),
      setSolidColor: vi.fn(),
      transparency: 0.4,
    };
    const lineFormat = {
      color: "#333333",
      weight: 2,
      transparency: 0.25,
      visible: false,
    };
    const font = {
      name: "Aptos",
      size: 18,
      color: "#111111",
      bold: true,
      italic: true,
      underline: "Single",
      strikethrough: true,
      allCaps: true,
      smallCaps: true,
      subscript: true,
      superscript: true,
      doubleStrikethrough: true,
    };
    const paragraphFormat = {
      horizontalAlignment: "Center",
      bulletFormat: { visible: true },
      indentLevel: 2,
    };
    const frame = {
      isNullObject: false,
      hasText: true,
      textRange: {
        text: "Old title",
        load: vi.fn(),
        font,
        paragraphFormat,
      },
      autoSizeSetting: "AutoSizeTextToFitShape",
      wordWrap: false,
      verticalAlignment: "Middle",
      leftMargin: 8,
      rightMargin: 8,
      topMargin: 8,
      bottomMargin: 8,
      load: vi.fn(),
    };
    const shape = {
      name: "Existing title",
      left: 40,
      top: 50,
      width: 320,
      height: 90,
      rotation: 12,
      visible: false,
      altTextTitle: "Existing alt title",
      altTextDescription: "Existing alt description",
      fill,
      lineFormat,
      getTextFrameOrNullObject: vi.fn(() => frame),
    } as unknown as PowerPoint.Shape & {
      name: string;
      left: number;
      top: number;
      width: number;
      height: number;
      rotation: number;
      visible: boolean;
      altTextTitle: string;
      altTextDescription: string;
    };
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

    await expect(manageSlideShapes.handler({
      action: "update",
      slideIndex: 0,
      shapeIndex: 0,
      shapeIds: [],
      placeholderType: "",
      shapeType: "textBox",
      geometricShapeType: "",
      connectorType: "Straight",
      text: "Project schedule",
      name: "",
      left: 0,
      top: 0,
      width: 0,
      height: 0,
      rotation: 0,
      visible: true,
      altTextTitle: "",
      altTextDescription: "",
      fillColor: "",
      fillTransparency: 0,
      clearFill: false,
      lineColor: "",
      lineWeight: 0,
      lineTransparency: 0,
      lineVisible: true,
      fontName: "",
      fontSize: 0,
      fontColor: "",
      bold: false,
      italic: false,
      underline: "None",
      strikethrough: false,
      allCaps: false,
      smallCaps: false,
      subscript: false,
      superscript: false,
      doubleStrikethrough: false,
      paragraphAlignment: "Left",
      bulletVisible: false,
      indentLevel: 0,
      textAutoSize: "AutoSizeNone",
      wordWrap: true,
      verticalAlignment: "Top",
      marginLeft: 0,
      marginRight: 0,
      marginTop: 0,
      marginBottom: 0,
    })).resolves.toBe("Updated shape on slide 1.");

    expect(frame.textRange.text).toBe("Project schedule");
    expect(shape.left).toBe(40);
    expect(shape.width).toBe(320);
    expect(shape.visible).toBe(false);
    expect(fill.setSolidColor).not.toHaveBeenCalled();
    expect(fill.clear).not.toHaveBeenCalled();
    expect(lineFormat.weight).toBe(2);
    expect(font.size).toBe(18);
    expect(paragraphFormat.horizontalAlignment).toBe("Center");
    expect(frame.wordWrap).toBe(false);
  });

  it("sanitizes noisy create payloads before applying shape formatting", async () => {
    const fill = {
      clear: vi.fn(),
      setSolidColor: vi.fn(),
      transparency: 0.4,
    };
    const lineFormat = {
      color: "#333333",
      weight: 2,
      transparency: 0.25,
      visible: true,
    };
    const font = {
      name: "Calibri",
      size: 12,
      color: "#000000",
      bold: false,
      italic: false,
      underline: "None",
      strikethrough: false,
      allCaps: false,
      smallCaps: false,
      subscript: false,
      superscript: false,
      doubleStrikethrough: false,
    };
    const paragraphFormat = {
      horizontalAlignment: "Left",
      bulletFormat: { visible: false },
      indentLevel: 0,
    };
    const frame = {
      isNullObject: false,
      hasText: true,
      textRange: {
        text: "",
        load: vi.fn(),
        font,
        paragraphFormat,
      },
      autoSizeSetting: "AutoSizeNone",
      wordWrap: false,
      verticalAlignment: "Top",
      leftMargin: 5,
      rightMargin: 5,
      topMargin: 5,
      bottomMargin: 5,
      load: vi.fn(),
    };
    const createdShape = {
      id: "shape-created",
      name: "",
      visible: true,
      altTextTitle: "",
      altTextDescription: "",
      fill,
      lineFormat,
      textFrame: frame,
      getTextFrameOrNullObject: vi.fn(() => frame),
      delete: vi.fn(),
      load: vi.fn(),
    } as unknown as PowerPoint.Shape & {
      id: string;
      name: string;
      visible: boolean;
      altTextTitle: string;
      altTextDescription: string;
    };
    const addTextBox = vi.fn(() => createdShape);
    const slide = {
      shapes: {
        addTextBox,
      },
    };
    const contextStub = {
      presentation: { slides: { load: vi.fn(), items: [slide] } },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    const requirementsStub = {
      isSetSupported: vi.fn((setName: string, version: string) => setName === "PowerPointApi" && ["1.3", "1.4"].includes(version)),
    };
    const runStub = vi.fn(async (callback: (context: typeof contextStub) => Promise<unknown>) => callback(contextStub));

    vi.stubGlobal("Office", { context: { requirements: requirementsStub } });
    vi.stubGlobal("PowerPoint", { run: runStub });

    await expect(manageSlideShapes.handler({
      action: "create",
      slideIndex: 0,
      shapeId: "",
      shapeIndex: 0,
      shapeIds: [],
      placeholderType: "",
      shapeType: "textBox",
      geometricShapeType: "",
      connectorType: "Straight",
      text: "Test / go-live",
      name: "Phase Label Launch",
      left: 56,
      top: 312,
      width: 104,
      height: 24,
      rotation: 0,
      visible: true,
      altTextTitle: "",
      altTextDescription: "",
      fillColor: "",
      fillTransparency: 0,
      clearFill: true,
      lineColor: "",
      lineWeight: 0,
      lineTransparency: 0,
      lineVisible: false,
      fontName: "Aptos",
      fontSize: 13,
      fontColor: "#FFFFFF",
      bold: false,
      italic: false,
      underline: "None",
      strikethrough: false,
      allCaps: false,
      smallCaps: false,
      subscript: false,
      superscript: false,
      doubleStrikethrough: false,
      paragraphAlignment: "Left",
      bulletVisible: false,
      indentLevel: 0,
      textAutoSize: "AutoSizeNone",
      wordWrap: true,
      verticalAlignment: "Middle",
      marginLeft: 0,
      marginRight: 0,
      marginTop: 0,
      marginBottom: 0,
    })).resolves.toBe("Created textBox shape-created on slide 1.");

    expect(addTextBox).toHaveBeenCalledWith("Test / go-live", { left: 56, top: 312, width: 104, height: 24 });
    expect(fill.clear).toHaveBeenCalled();
    expect(fill.setSolidColor).not.toHaveBeenCalled();
    expect(lineFormat.visible).toBe(false);
    expect(lineFormat.weight).toBe(2);
    expect(font.name).toBe("Aptos");
    expect(font.size).toBe(13);
    expect(font.color).toBe("#FFFFFF");
    expect(frame.verticalAlignment).toBe("Middle");
    expect(createdShape.name).toBe("Phase Label Launch");
  });
});
