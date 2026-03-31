import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from "@xmldom/xmldom";
import { strToU8, zipSync } from "fflate";
import { describe, expect, it, vi } from "vitest";
import {
  getShapeTextAutoSizeSetting,
  reapplyShapeTextAutoSizeSetting,
  resolveSlideShapeByIdWithXmlFallback,
} from "./powerpointShapeTarget";

function createPresentationBase64(entries: Record<string, string>) {
  let binary = "";
  zipSync(Object.fromEntries(
    Object.entries(entries).map(([path, contents]) => [path, strToU8(contents)]),
  )).forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary);
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
          <p:spPr/>
        </p:sp>
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="11" name="Body"/>
            <p:cNvSpPr/>
            <p:nvPr/>
          </p:nvSpPr>
          <p:spPr/>
        </p:sp>
      </p:spTree>
    </p:cSld>
    <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
  </p:sld>`;
}

if (typeof DOMParser === "undefined") {
  vi.stubGlobal("DOMParser", XmldomParser);
}

if (typeof XMLSerializer === "undefined") {
  vi.stubGlobal("XMLSerializer", XmldomSerializer);
}

describe("resolveSlideShapeByIdWithXmlFallback", () => {
  it("returns a direct Office shape.id match without exporting the slide", async () => {
    const context = { sync: vi.fn().mockResolvedValue(undefined) } as unknown as PowerPoint.RequestContext;
    const shapes = [{ id: "office-title", name: "Title" }, { id: "office-body", name: "Body" }] as unknown as PowerPoint.Shape[];
    const exportAsBase64 = vi.fn(() => ({ value: "unused" }));
    const slide = {
      shapes: { items: shapes, load: vi.fn() },
      exportAsBase64,
    } as unknown as PowerPoint.Slide;

    vi.stubGlobal("Office", { context: { requirements: { isSetSupported: vi.fn().mockReturnValue(true) } } });

    await expect(resolveSlideShapeByIdWithXmlFallback(context, slide, 0, "office-body")).resolves.toMatchObject({
      shape: shapes[1],
      shapeId: "office-body",
      shapeIndex: 1,
    });
    expect(exportAsBase64).not.toHaveBeenCalled();
  });

  it("falls back from exported XML cNvPr ids to the current Office shape.id", async () => {
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
    });
    const context = { sync: vi.fn().mockResolvedValue(undefined) } as unknown as PowerPoint.RequestContext;
    const shapes = [{ id: "office-title", name: "Title" }, { id: "office-body", name: "Body" }] as unknown as PowerPoint.Shape[];
    const exportAsBase64 = vi.fn(() => ({ value: base64 }));
    const slide = {
      shapes: { items: shapes, load: vi.fn() },
      exportAsBase64,
    } as unknown as PowerPoint.Slide;

    vi.stubGlobal("Office", { context: { requirements: { isSetSupported: vi.fn((setName: string, version: string) => setName === "PowerPointApi" && version === "1.8") } } });

    await expect(resolveSlideShapeByIdWithXmlFallback(context, slide, 0, 11)).resolves.toMatchObject({
      shape: shapes[1],
      shapeId: "office-body",
      shapeIndex: 1,
    });
    expect(exportAsBase64).toHaveBeenCalledTimes(1);
  });
});

describe("shape text auto-size helpers", () => {
  it("reads a text frame auto-size setting through XML id fallback", async () => {
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
    });
    const frame = {
      isNullObject: false,
      autoSizeSetting: "AutoSizeShapeToFitText",
      load: vi.fn(),
    };
    const shape = {
      id: "office-body",
      name: "Body",
      getTextFrameOrNullObject: vi.fn(() => frame),
    } as unknown as PowerPoint.Shape;
    const context = { sync: vi.fn().mockResolvedValue(undefined) } as unknown as PowerPoint.RequestContext;
    const slide = {
      shapes: { items: [{ id: "office-title", name: "Title" }, shape], load: vi.fn() },
      exportAsBase64: vi.fn(() => ({ value: base64 })),
    } as unknown as PowerPoint.Slide;

    vi.stubGlobal("Office", { context: { requirements: { isSetSupported: vi.fn((setName: string, version: string) => setName === "PowerPointApi" && version === "1.8") } } });

    await expect(getShapeTextAutoSizeSetting(context, slide, 0, "11")).resolves.toBe("AutoSizeShapeToFitText");
    expect(frame.load).toHaveBeenCalledWith(["isNullObject", "autoSizeSetting"]);
  });

  it("reapplies the requested text auto-size setting after a round-trip when the replacement shape is still text-capable", async () => {
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
    });
    const frame = {
      isNullObject: false,
      autoSizeSetting: "AutoSizeNone",
      load: vi.fn(),
    };
    const shape = {
      id: "office-body-new",
      name: "Body",
      getTextFrameOrNullObject: vi.fn(() => frame),
    } as unknown as PowerPoint.Shape;
    const context = { sync: vi.fn().mockResolvedValue(undefined) } as unknown as PowerPoint.RequestContext;
    const slide = {
      shapes: { items: [{ id: "office-title-new", name: "Title" }, shape], load: vi.fn() },
      exportAsBase64: vi.fn(() => ({ value: base64 })),
    } as unknown as PowerPoint.Slide;

    vi.stubGlobal("Office", { context: { requirements: { isSetSupported: vi.fn((setName: string, version: string) => setName === "PowerPointApi" && version === "1.8") } } });

    await expect(reapplyShapeTextAutoSizeSetting(context, slide, 0, "11", "AutoSizeTextToFitShape")).resolves.toBe(true);
    expect(frame.autoSizeSetting).toBe("AutoSizeTextToFitShape");
  });
});
