import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from "@xmldom/xmldom";
import { strToU8, zipSync } from "fflate";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { editSlideText } from "./editSlideText";
import { clearSlideExportCache } from "./powerpointOpenXml";
import { inspectSlideXmlFromBase64Presentation } from "./powerpointSlideXml";

function createPresentationBase64(entries: Record<string, string>) {
  let binary = "";
  zipSync(Object.fromEntries(
    Object.entries(entries).map(([path, contents]) => [path, strToU8(contents)]),
  )).forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary);
}

function createSlideBase64(titleText: string) {
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
              <p:spPr><a:xfrm><a:off x="127000" y="254000"/><a:ext cx="381000" cy="508000"/></a:xfrm></p:spPr>
              <p:txBody>
                <a:bodyPr wrap="square"/>
                <a:lstStyle/>
                <a:p><a:r><a:t>${titleText}</a:t></a:r></a:p>
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

beforeEach(() => {
  clearSlideExportCache();
});

afterEach(() => {
  clearSlideExportCache();
});

describe("editSlideText", () => {
  it("validates arguments", async () => {
    await expect(editSlideText.handler({ ref: "slide-id:slide-1/shape:10" })).resolves.toMatchObject({
      resultType: "failure",
    });
  });

  it("replaces one shape's paragraph XML and refreshes the returned ref after the slide round-trip", async () => {
    const sourceBase64 = createSlideBase64("Original");
    const replacementBase64Holder = { value: sourceBase64 };

    const sourceFrame = {
      isNullObject: false,
      autoSizeSetting: "AutoSizeShapeToFitText",
      load: vi.fn(),
    };
    const replacementFrame = {
      isNullObject: false,
      autoSizeSetting: "AutoSizeNone",
      load: vi.fn(),
    };
    const sourceShape = {
      id: "office-title",
      name: "Title",
      getTextFrameOrNullObject: vi.fn(() => sourceFrame),
    };
    const replacementShape = {
      id: "office-title-new",
      name: "Title",
      getTextFrameOrNullObject: vi.fn(() => replacementFrame),
    };
    const sourceSlide = {
      id: "slide-old",
      load: vi.fn(),
      exportAsBase64: vi.fn(() => ({ value: sourceBase64 })),
      delete: vi.fn(),
      shapes: { items: [sourceShape], load: vi.fn() },
    };
    const replacementSlide = {
      id: "slide-new",
      load: vi.fn(),
      exportAsBase64: vi.fn(() => ({ value: replacementBase64Holder.value })),
      shapes: { items: [replacementShape], load: vi.fn() },
    };
    const slides = {
      items: [sourceSlide],
      load: vi.fn(),
      getItemAt: vi.fn((_index: number) => sourceSlide),
    } as any;
    const finalSlides = {
      items: [replacementSlide],
      load: vi.fn(),
    } as any;
    const presentation = {
      slides,
      insertSlidesFromBase64: vi.fn((mutated: string) => {
        replacementBase64Holder.value = mutated;
        presentation.slides = finalSlides;
      }),
    } as any;
    const contextStub = {
      presentation,
      sync: vi.fn().mockResolvedValue(undefined),
    } as any;
    const runStub = vi.fn(async (callback: (context: typeof contextStub) => Promise<unknown>) => callback(contextStub));

    vi.stubGlobal("Office", {
      context: {
        requirements: {
          isSetSupported: vi.fn((setName: string, version: string) => setName === "PowerPointApi" && version === "1.8"),
        },
      },
    });
    vi.stubGlobal("PowerPoint", {
      run: runStub,
      InsertSlideFormatting: { keepSourceFormatting: "KeepSourceFormatting" },
    });

    const result = await editSlideText.handler({
      ref: "slide-id:slide-old/shape:10",
      paragraphsXml: ['<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Updated</a:t></a:r></a:p>'],
    });

    expect(result).toMatchObject({
      resultType: "success",
      ref: "slide-id:slide-new/shape:10",
      slideId: "slide-new",
      xmlShapeId: "10",
      slideIndex: 0,
    });
    expect(replacementFrame.autoSizeSetting).toBe("AutoSizeShapeToFitText");
    expect(presentation.insertSlidesFromBase64).toHaveBeenCalledTimes(1);

    const inspection = inspectSlideXmlFromBase64Presentation(replacementBase64Holder.value, { slideId: "slide-new" });
    expect(inspection.shapes[0]?.xmlShapeId).toBe("10");
    expect(inspection.shapes[0]?.textBody?.textContent || "").toContain("Updated");
  });
});
