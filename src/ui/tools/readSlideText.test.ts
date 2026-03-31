import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from "@xmldom/xmldom";
import { strToU8, zipSync } from "fflate";
import { afterEach, describe, expect, it, vi } from "vitest";
import { readSlideText } from "./readSlideText";

function createPresentationBase64(entries: Record<string, string>) {
  let binary = "";
  zipSync(Object.fromEntries(
    Object.entries(entries).map(([path, contents]) => [path, strToU8(contents)]),
  )).forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary);
}

function createSlideBase64() {
  return createPresentationBase64({
    "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
    "ppt/slides/slide1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cSld>
          <p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="10" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr/>
              <p:txBody>
                <a:bodyPr/>
                <a:lstStyle/>
                <a:p><a:r><a:t>One</a:t></a:r></a:p>
                <a:p><a:r><a:t>Two</a:t></a:r></a:p>
              </p:txBody>
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
});

describe("readSlideText", () => {
  it("validates arguments", async () => {
    await expect(readSlideText.handler({})).resolves.toMatchObject({
      resultType: "failure",
    });
  });

  it("returns raw paragraph XML for the targeted shape ref", async () => {
    const slide = {
      id: "slide-1",
      load: vi.fn(),
      exportAsBase64: vi.fn(() => ({ value: createSlideBase64() })),
    };
    const contextStub = {
      presentation: { slides: { items: [slide], load: vi.fn() } },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    const runStub = vi.fn(async (callback: (context: typeof contextStub) => Promise<unknown>) => callback(contextStub));

    vi.stubGlobal("Office", { context: { requirements: { isSetSupported: vi.fn().mockReturnValue(true) } } });
    vi.stubGlobal("PowerPoint", { run: runStub });

    const result = await readSlideText.handler({ ref: "slide-id:slide-1/shape:10" });

    expect(result).toMatchObject({
      resultType: "success",
      ref: "slide-id:slide-1/shape:10",
      slideId: "slide-1",
      xmlShapeId: "10",
      paragraphsXml: [
        expect.stringContaining("<a:t>One</a:t>"),
        expect.stringContaining("<a:t>Two</a:t>"),
      ],
    });
  });

  it("adds a round-trip refresh hint when the slide id in the ref is stale", async () => {
    const slide = {
      id: "slide-new",
      load: vi.fn(),
      exportAsBase64: vi.fn(() => ({ value: createSlideBase64() })),
    };
    const contextStub = {
      presentation: { slides: { items: [slide], load: vi.fn() } },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    const runStub = vi.fn(async (callback: (context: typeof contextStub) => Promise<unknown>) => callback(contextStub));

    vi.stubGlobal("Office", { context: { requirements: { isSetSupported: vi.fn().mockReturnValue(true) } } });
    vi.stubGlobal("PowerPoint", { run: runStub });

    const result = await readSlideText.handler({ ref: "slide-id:slide-old/shape:10" });

    expect(result).toMatchObject({ resultType: "failure" });
    if (result && typeof result === "object" && "error" in result) {
      expect(String(result.error)).toContain('Slide "slide-old" was not found in the current presentation.');
      expect(String(result.error)).toContain("Re-run get_presentation_overview to refresh current slideIndex values");
    }
  });

  it("adds a round-trip refresh hint when the shape ref is stale on the current slide", async () => {
    const slide = {
      id: "slide-1",
      load: vi.fn(),
      exportAsBase64: vi.fn(() => ({ value: createPresentationBase64({
        "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
        "ppt/slides/slide1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld>
              <p:spTree>
                <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
                <p:sp><p:nvSpPr><p:cNvPr id="99" name="Other"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>
              </p:spTree>
            </p:cSld>
            <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
          </p:sld>`,
      }) })),
    };
    const contextStub = {
      presentation: { slides: { items: [slide], load: vi.fn() } },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    const runStub = vi.fn(async (callback: (context: typeof contextStub) => Promise<unknown>) => callback(contextStub));

    vi.stubGlobal("Office", { context: { requirements: { isSetSupported: vi.fn().mockReturnValue(true) } } });
    vi.stubGlobal("PowerPoint", { run: runStub });

    const result = await readSlideText.handler({ ref: "slide-id:slide-1/shape:10" });

    expect(result).toMatchObject({ resultType: "failure" });
    if (result && typeof result === "object" && "error" in result) {
      expect(String(result.error)).toContain('Could not find shape ref "slide-id:slide-1/shape:10" on exported slide "slide-1".');
      expect(String(result.error)).toContain("Re-run get_presentation_overview to refresh current slideIndex values");
    }
  });
});
