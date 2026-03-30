import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from "@xmldom/xmldom";
import { strToU8, zipSync } from "fflate";
import { afterEach, describe, expect, it, vi } from "vitest";
import { OpenXmlPackage, parseXml } from "./openXmlPackage";
import { addSlideAnimation } from "./addSlideAnimation";
import { clearSlideExportCache } from "./powerpointOpenXml";

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

function ensureXmlGlobals() {
  if (typeof DOMParser === "undefined") {
    vi.stubGlobal("DOMParser", XmldomParser);
  }
  if (typeof XMLSerializer === "undefined") {
    vi.stubGlobal("XMLSerializer", XmldomSerializer);
  }
}

function baseSlideXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <p:cSld>
      <p:spTree>
        <p:nvGrpSpPr>
          <p:cNvPr id="1" name=""/>
          <p:cNvGrpSpPr/>
          <p:nvPr/>
        </p:nvGrpSpPr>
        <p:grpSpPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="0" cy="0"/>
            <a:chOff x="0" y="0"/>
            <a:chExt cx="0" cy="0"/>
          </a:xfrm>
        </p:grpSpPr>
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="10" name="Title"/>
            <p:cNvSpPr/>
            <p:nvPr/>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm>
              <a:off x="127000" y="127000"/>
              <a:ext cx="1270000" cy="508000"/>
            </a:xfrm>
          </p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>One</a:t></a:r></a:p></p:txBody>
        </p:sp>
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="11" name="Body"/>
            <p:cNvSpPr/>
            <p:nvPr/>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm>
              <a:off x="254000" y="254000"/>
              <a:ext cx="1270000" cy="508000"/>
            </a:xfrm>
          </p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Two</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:spTree>
    </p:cSld>
    <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
  </p:sld>`;
}

afterEach(() => {
  clearSlideExportCache();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("addSlideAnimation", () => {
  it("falls back from stale shapeId values to exported XML shape ids", async () => {
    ensureXmlGlobals();
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
    });
    let insertedBase64 = "";

    const sourceSlide = {
      id: "slide-1",
      load: vi.fn(),
      exportAsBase64: vi.fn(() => ({ value: base64 })),
      delete: vi.fn(),
      shapes: {
        items: [
          { id: "office-shape-100", name: "Title" },
          { id: "office-shape-200", name: "Body" },
        ],
        load: vi.fn(),
      },
    };
    const insertedSlide = { id: "slide-new" };
    const slides = {
      items: [sourceSlide],
      load: vi.fn(),
      getItemAt: vi.fn((index: number) => [sourceSlide][index]),
    } as any;
    // After insert + delete batch, only the replacement slide remains.
    const finalSlides = {
      items: [insertedSlide],
      load: vi.fn(),
    } as any;
    const presentation = {
      slides,
      insertSlidesFromBase64: vi.fn((mutated: string) => {
        insertedBase64 = mutated;
        presentation.slides = finalSlides;
      }),
    } as any;
    const context = {
      presentation,
      sync: vi.fn().mockResolvedValue(undefined),
    } as any;

    vi.stubGlobal("Office", {
      context: {
        requirements: {
          isSetSupported: vi.fn((setName: string, version: string) => setName === "PowerPointApi" && version === "1.8"),
        },
      },
    });
    vi.stubGlobal("PowerPoint", {
      run: vi.fn(async (callback: (requestContext: typeof context) => Promise<unknown>) => callback(context)),
      InsertSlideFormatting: {
        keepSourceFormatting: "KeepSourceFormatting",
      },
    });

    const result = await addSlideAnimation.handler({
      slideIndex: 0,
      shapeId: "11",
      type: "scale",
      start: "withPrevious",
      scaleXPercent: 90,
      scaleYPercent: 90,
    });

    const slideDoc = parseXml(new OpenXmlPackage(insertedBase64).readText("ppt/slides/slide1.xml"));
    const target = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "spTgt")[0];

    expect(target?.getAttribute("spid")).toBe("11");
    expect(result).toMatchObject({
      resultType: "success",
      slideIndex: 0,
      slideId: "slide-new",
      shapeId: "office-shape-200",
      refreshedShapeId: "office-shape-200",
      textResultForLlm: "Added a scale animation to slide 1 targeting shape office-shape-200.",
    });
  });

  it("matches the XML target by shape metadata when Office shape order differs", async () => {
    ensureXmlGlobals();
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
    });
    let insertedBase64 = "";

    const sourceSlide = {
      id: "slide-1",
      load: vi.fn(),
      exportAsBase64: vi.fn(() => ({ value: base64 })),
      delete: vi.fn(),
      shapes: {
        items: [
          { id: "office-body", name: "Body", left: 20, top: 20, width: 100, height: 40 },
          { id: "office-title", name: "Title", left: 10, top: 10, width: 100, height: 40 },
        ],
        load: vi.fn(),
      },
    };
    const insertedSlide = { id: "slide-new" };
    const slides = {
      items: [sourceSlide],
      load: vi.fn(),
      getItemAt: vi.fn((index: number) => [sourceSlide][index]),
    } as any;
    const finalSlides = {
      items: [insertedSlide],
      load: vi.fn(),
    } as any;
    const presentation = {
      slides,
      insertSlidesFromBase64: vi.fn((mutated: string) => {
        insertedBase64 = mutated;
        presentation.slides = finalSlides;
      }),
    } as any;
    const context = {
      presentation,
      sync: vi.fn().mockResolvedValue(undefined),
    } as any;

    vi.stubGlobal("Office", {
      context: {
        requirements: {
          isSetSupported: vi.fn((setName: string, version: string) => setName === "PowerPointApi" && version === "1.8"),
        },
      },
    });
    vi.stubGlobal("PowerPoint", {
      run: vi.fn(async (callback: (requestContext: typeof context) => Promise<unknown>) => callback(context)),
      InsertSlideFormatting: {
        keepSourceFormatting: "KeepSourceFormatting",
      },
    });

    const result = await addSlideAnimation.handler({
      slideIndex: 0,
      shapeId: "office-body",
      type: "fade",
      start: "afterPrevious",
      durationMs: 350,
      delayMs: 150,
    });

    expect(result).toMatchObject({
      resultType: "success",
      slideIndex: 0,
      slideId: "slide-new",
      shapeId: "office-body",
      refreshedShapeId: "office-body",
      textResultForLlm: "Added a fade animation to slide 1 targeting shape office-body.",
    });
    const slideDoc = parseXml(new OpenXmlPackage(insertedBase64).readText("ppt/slides/slide1.xml"));
    const target = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "spTgt")[0];

    expect(target?.getAttribute("spid")).toBe("11");
  });
});
