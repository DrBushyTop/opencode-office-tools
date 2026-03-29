import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from "@xmldom/xmldom";
import { strToU8, zipSync } from "fflate";
import { describe, expect, it, vi } from "vitest";
import { manageSlideShapes } from "./manageSlideShapes";
import { setPowerPointContextSnapshot } from "./powerpointContext";

function createPresentationBase64(entries: Record<string, string>) {
  return Buffer.from(zipSync(Object.fromEntries(
    Object.entries(entries).map(([path, contents]) => [path, strToU8(contents)]),
  ))).toString("base64");
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
});
