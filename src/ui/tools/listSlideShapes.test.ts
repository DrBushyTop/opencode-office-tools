import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from "@xmldom/xmldom";
import { strToU8, zipSync } from "fflate";
import { afterEach, describe, expect, it, vi } from "vitest";
import { listSlideShapes } from "./listSlideShapes";
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

function createSingleShapeSlideBase64(options: { xmlShapeId: string; name: string; text: string; left?: number; top?: number; width?: number; height?: number }) {
  const { xmlShapeId, name, text, left = 10, top = 20, width = 30, height = 40 } = options;
  const emu = (points: number) => String(points * 12700);
  return createPresentationBase64({
    "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
    "ppt/slides/slide1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cSld>
          <p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="${xmlShapeId}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="${emu(left)}" y="${emu(top)}"/><a:ext cx="${emu(width)}" cy="${emu(height)}"/></a:xfrm></p:spPr>
              <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>
            </p:sp>
          </p:spTree>
        </p:cSld>
        <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
      </p:sld>`,
  });
}

if (typeof DOMParser === "undefined") {
  vi.stubGlobal("DOMParser", XmldomParser);
}

if (typeof XMLSerializer === "undefined") {
  vi.stubGlobal("XMLSerializer", XmldomSerializer);
}

afterEach(() => {
  setPowerPointContextSnapshot(null);
});

describe("listSlideShapes", () => {
  it("fails clearly when slideIndex is omitted and no active slide can be inferred", async () => {
    const slides = {
      items: [{ id: "slide-1" }, { id: "slide-2" }],
      load: vi.fn(),
    };
    const contextStub = {
      presentation: { slides },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    const runStub = vi.fn(async (callback: (context: typeof contextStub) => Promise<unknown>) => callback(contextStub));

    vi.stubGlobal("Office", { context: { requirements: { isSetSupported: vi.fn().mockReturnValue(true) } } });
    vi.stubGlobal("PowerPoint", { run: runStub, ShapeType: { table: "Table", placeholder: "Placeholder" } });

    await expect(listSlideShapes.handler()).resolves.toMatchObject({
      resultType: "failure",
      error: "slideIndex is required when no active slide can be inferred from the current PowerPoint context.",
    });
  });

  it("uses the active slide context and returns stable refs", async () => {
    const frame = {
      isNullObject: false,
      hasText: true,
      load: vi.fn(),
      textRange: {
        text: "Hello shape-ref world",
        load: vi.fn(),
      },
    };
    const shape = {
      id: "office-title",
      name: "Title",
      type: "TextBox",
      left: 10,
      top: 20,
      width: 30,
      height: 40,
      rotation: 0,
      zOrderPosition: 0,
      visible: true,
      altTextTitle: "",
      altTextDescription: "",
      load: vi.fn(),
      getTextFrameOrNullObject: vi.fn(() => frame),
    };
    const slide = {
      id: "slide-active",
      load: vi.fn(),
      exportAsBase64: vi.fn(() => ({ value: createSingleShapeSlideBase64({ xmlShapeId: "10", name: "Title", text: "Hello shape-ref world" }) })),
      shapes: {
        items: [shape],
        load: vi.fn(),
      },
    };
    const contextStub = {
      presentation: { slides: { items: [slide], load: vi.fn() } },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    const runStub = vi.fn(async (callback: (context: typeof contextStub) => Promise<unknown>) => callback(contextStub));

    vi.stubGlobal("Office", { context: { requirements: { isSetSupported: vi.fn().mockReturnValue(true) } } });
    vi.stubGlobal("PowerPoint", { run: runStub, ShapeType: { table: "Table", placeholder: "Placeholder" } });
    setPowerPointContextSnapshot({ selectedSlideIds: ["slide-active"], selectedShapeIds: [], activeSlideId: "slide-active", activeSlideIndex: 0 });

    const result = await listSlideShapes.handler({ detail: false });

    expect(result).toMatchObject({
      resultType: "success",
      slideIndex: 0,
      slideId: "slide-active",
      detail: false,
      shapes: [
        {
          index: 0,
          ref: "slide-id:slide-active/shape:10",
          slideId: "slide-active",
          xmlShapeId: "10",
          name: "Title",
          type: "TextBox",
          xmlType: "sp",
          box: { left: 10, top: 20, width: 30, height: 40 },
          hasText: true,
          textPreview: "Hello shape-ref world",
        },
      ],
    });
  });

  it("keeps blank text placeholders aligned with XML text bodies", async () => {
    const frame = {
      isNullObject: false,
      hasText: false,
      load: vi.fn(),
      textRange: {
        text: "",
        load: vi.fn(),
      },
    };
    const shape = {
      id: "office-title",
      name: "Title",
      type: "TextBox",
      left: 10,
      top: 20,
      width: 30,
      height: 40,
      rotation: 0,
      zOrderPosition: 0,
      visible: true,
      altTextTitle: "",
      altTextDescription: "",
      load: vi.fn(),
      getTextFrameOrNullObject: vi.fn(() => frame),
    };
    const slide = {
      id: "slide-active",
      load: vi.fn(),
      exportAsBase64: vi.fn(() => ({ value: createSingleShapeSlideBase64({ xmlShapeId: "10", name: "Title", text: "" }) })),
      shapes: {
        items: [shape],
        load: vi.fn(),
      },
    };
    const contextStub = {
      presentation: { slides: { items: [slide], load: vi.fn() } },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    const runStub = vi.fn(async (callback: (context: typeof contextStub) => Promise<unknown>) => callback(contextStub));

    vi.stubGlobal("Office", { context: { requirements: { isSetSupported: vi.fn().mockReturnValue(true) } } });
    vi.stubGlobal("PowerPoint", { run: runStub, ShapeType: { table: "Table", placeholder: "Placeholder" } });
    setPowerPointContextSnapshot({ selectedSlideIds: ["slide-active"], selectedShapeIds: [], activeSlideId: "slide-active", activeSlideIndex: 0 });

    const result = await listSlideShapes.handler({ detail: false });

    expect(result).toMatchObject({
      resultType: "success",
      shapes: [
        {
          ref: "slide-id:slide-active/shape:10",
          hasText: true,
          textPreview: "(empty)",
        },
      ],
    });
  });
});
