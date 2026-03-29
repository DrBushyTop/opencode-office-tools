import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from "@xmldom/xmldom";
import { strToU8, zipSync } from "fflate";
import { afterEach, describe, expect, it, vi } from "vitest";
import { extractSlideTransitionFromBase64Presentation } from "./powerpointOpenXml";
import { setSlideTransition } from "./setSlideTransition";

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
      </p:spTree>
    </p:cSld>
    <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
  </p:sld>`;
}

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("setSlideTransition", () => {
  it("applies the same transition to multiple slides when slideIndex is an array", async () => {
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
    });
    const insertedBase64: string[] = [];

    const slidesState: any[] = [];
    const slides = {
      items: slidesState,
      load: vi.fn(),
    } as any;

    const makeSlide = (id: string) => ({
      id,
      load: vi.fn(),
      exportAsBase64: vi.fn(() => ({ value: base64 })),
      delete: vi.fn(() => {
        const index = slidesState.findIndex((slide) => slide.id === id);
        if (index >= 0) {
          slidesState.splice(index, 1);
        }
      }),
    });

    slidesState.push(makeSlide("slide-1"), makeSlide("slide-2"), makeSlide("slide-3"));

    let insertedCount = 0;
    const presentation = {
      slides,
      insertSlidesFromBase64: vi.fn((mutated: string, options?: { targetSlideId?: string }) => {
        insertedBase64.push(mutated);
        insertedCount += 1;
        const insertedSlide = makeSlide(`slide-new-${insertedCount}`);
        if (options?.targetSlideId) {
          const targetIndex = slidesState.findIndex((slide) => slide.id === options.targetSlideId);
          slidesState.splice(targetIndex + 1, 0, insertedSlide);
        } else {
          slidesState.unshift(insertedSlide);
        }
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

    const result = await setSlideTransition.handler({
      slideIndex: [0, 1],
      effect: "fade",
      durationMs: 600,
    });

    expect(insertedBase64).toHaveLength(2);
    expect(extractSlideTransitionFromBase64Presentation(insertedBase64[0])).toMatchObject({ effect: "fade", durationMs: 600 });
    expect(extractSlideTransitionFromBase64Presentation(insertedBase64[1])).toMatchObject({ effect: "fade", durationMs: 600 });
    expect(result).toMatchObject({
      resultType: "success",
      slideIndexes: [0, 1],
      slideIds: ["slide-new-1", "slide-new-2"],
      textResultForLlm: "Set the fade transition on slides 1, 2.",
    });
  });
});
