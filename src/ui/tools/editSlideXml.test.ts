import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from "@xmldom/xmldom";
import { strToU8, zipSync } from "fflate";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { editSlideXml } from "./editSlideXml";
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

function createSlideBase64(titleText: string, bodyText: string) {
  return createPresentationBase64({
    "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
    "ppt/slides/slide1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cSld>
          <p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
            <p:sp><p:nvSpPr><p:cNvPr id="10" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${titleText}</a:t></a:r></a:p></p:txBody></p:sp>
            <p:sp><p:nvSpPr><p:cNvPr id="11" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${bodyText}</a:t></a:r></a:p></p:txBody></p:sp>
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

describe("editSlideXml", () => {
  it("rejects mixed-slide batches before touching PowerPoint", async () => {
    await expect(editSlideXml.handler({
      replacements: [
        { ref: "slide-id:slide-1/shape:10", paragraphsXml: ["<a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>"] },
        { ref: "slide-id:slide-2/shape:11", paragraphsXml: ["<a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>"] },
      ],
    })).resolves.toMatchObject({
      resultType: "failure",
      error: "All replacements must target the same slide.",
    });
  });

  it("batch edits multiple shapes on one slide with a single round-trip and refreshes refs", async () => {
    const sourceBase64 = createSlideBase64("Before title", "Before body");
    const replacementBase64Holder = { value: sourceBase64 };

    const sourceTitleShape = { id: "office-title", name: "Title", getTextFrameOrNullObject: vi.fn(() => ({ isNullObject: false, autoSizeSetting: "AutoSizeNone", load: vi.fn() })) };
    const sourceBodyShape = { id: "office-body", name: "Body", getTextFrameOrNullObject: vi.fn(() => ({ isNullObject: false, autoSizeSetting: "AutoSizeNone", load: vi.fn() })) };
    const replacementTitleShape = { id: "office-title-new", name: "Title", getTextFrameOrNullObject: vi.fn(() => ({ isNullObject: false, autoSizeSetting: "AutoSizeNone", load: vi.fn() })) };
    const replacementBodyShape = { id: "office-body-new", name: "Body", getTextFrameOrNullObject: vi.fn(() => ({ isNullObject: false, autoSizeSetting: "AutoSizeNone", load: vi.fn() })) };
    const sourceSlide = {
      id: "slide-old",
      load: vi.fn(),
      exportAsBase64: vi.fn(() => ({ value: sourceBase64 })),
      delete: vi.fn(),
      shapes: { items: [sourceTitleShape, sourceBodyShape], load: vi.fn() },
    };
    const replacementSlide = {
      id: "slide-new",
      load: vi.fn(),
      exportAsBase64: vi.fn(() => ({ value: replacementBase64Holder.value })),
      shapes: { items: [replacementTitleShape, replacementBodyShape], load: vi.fn() },
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
    let syncInFlight = false;
    const contextStub = {
      presentation,
      sync: vi.fn(async () => {
        if (syncInFlight) {
          throw new Error("Concurrent context.sync() is not allowed in this test.");
        }
        syncInFlight = true;
        await Promise.resolve();
        syncInFlight = false;
      }),
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

    const result = await editSlideXml.handler({
      replacements: [
        {
          ref: "slide-id:slide-old/shape:10",
          paragraphsXml: ['<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>After title</a:t></a:r></a:p>'],
        },
        {
          ref: "slide-id:slide-old/shape:11",
          paragraphsXml: ['<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>After body</a:t></a:r></a:p>'],
        },
      ],
    });

    expect(result).toMatchObject({
      resultType: "success",
      slideId: "slide-new",
      slideIndex: 0,
      replacements: [
        { ref: "slide-id:slide-new/shape:10", slideId: "slide-new", xmlShapeId: "10" },
        { ref: "slide-id:slide-new/shape:11", slideId: "slide-new", xmlShapeId: "11" },
      ],
    });
    expect(presentation.insertSlidesFromBase64).toHaveBeenCalledTimes(1);

    const inspection = inspectSlideXmlFromBase64Presentation(replacementBase64Holder.value, { slideId: "slide-new" });
    expect(inspection.shapes[0]?.textBody?.textContent || "").toContain("After title");
    expect(inspection.shapes[1]?.textBody?.textContent || "").toContain("After body");
  });
});
