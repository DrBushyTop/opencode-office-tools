import { strToU8, zipSync } from "fflate";
import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from "@xmldom/xmldom";
import { describe, expect, it, vi } from "vitest";
import { OpenXmlPackage, parseXml } from "./openXmlPackage";
import {
  addSlideAnimationInBase64Presentation,
  extractSlideTransitionFromBase64Presentation,
  replaceSlideWithMutatedOpenXml,
  setSlideTransitionInBase64Presentation,
  setSpeakerNotesInBase64Presentation,
} from "./powerpointOpenXml";

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
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="10" name="Title"/>
            <p:cNvSpPr/>
            <p:nvPr/>
          </p:nvSpPr>
          <p:spPr/>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>One</a:t></a:r></a:p></p:txBody>
        </p:sp>
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="11" name="Body"/>
            <p:cNvSpPr/>
            <p:nvPr/>
          </p:nvSpPr>
          <p:spPr/>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Two</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:spTree>
    </p:cSld>
    <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
  </p:sld>`;
}

describe("replaceSlideWithMutatedOpenXml", () => {
  it("inserts relative to the source slide and deletes the original by id", async () => {
    const sourceSlide = {
      id: "slide-2",
      load: vi.fn(),
      exportAsBase64: vi.fn(() => ({ value: "BASE64" })),
    };
    const slideA = { id: "slide-1", load: vi.fn() };
    const insertedSlide = { id: "slide-new" };
    const originalSlide = { id: "slide-2", delete: vi.fn() };
    const slides = {
      items: [slideA, sourceSlide],
      load: vi.fn(),
    } as any;

    const updatedSlides = {
      items: [slideA, insertedSlide, originalSlide],
      load: vi.fn(),
    } as any;

    const presentation = {
      slides,
      insertSlidesFromBase64: vi.fn(() => {
        presentation.slides = updatedSlides;
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
      InsertSlideFormatting: {
        keepSourceFormatting: "KeepSourceFormatting",
      },
    });

    await replaceSlideWithMutatedOpenXml(context, 1, (value) => `${value}-mutated`);

    expect(presentation.insertSlidesFromBase64).toHaveBeenCalledWith("BASE64-mutated", {
      formatting: "KeepSourceFormatting",
      targetSlideId: "slide-2",
    });
    expect(originalSlide.delete).toHaveBeenCalledTimes(1);
  });

  it("uses the source slide id for first-slide replacement and deletes the original by id", async () => {
    const sourceSlide = {
      id: "slide-1",
      load: vi.fn(),
      exportAsBase64: vi.fn(() => ({ value: "BASE64" })),
    };
    const insertedSlide = { id: "slide-new" };
    const originalSlide = { id: "slide-1", delete: vi.fn() };
    const slides = {
      items: [sourceSlide],
      load: vi.fn(),
    } as any;

    const updatedSlides = {
      items: [insertedSlide, originalSlide],
      load: vi.fn(),
    } as any;

    const presentation = {
      slides,
      insertSlidesFromBase64: vi.fn(() => {
        presentation.slides = updatedSlides;
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
      InsertSlideFormatting: {
        keepSourceFormatting: "KeepSourceFormatting",
      },
    });

    await replaceSlideWithMutatedOpenXml(context, 0, (value) => `${value}-mutated`);

    expect(presentation.insertSlidesFromBase64).toHaveBeenCalledWith("BASE64-mutated", {
      formatting: "KeepSourceFormatting",
      targetSlideId: "slide-1",
    });
    expect(originalSlide.delete).toHaveBeenCalledTimes(1);
  });

  it("deletes the original slide by id even if the inserted slide is not adjacent", async () => {
    const sourceSlide = {
      id: "slide-2",
      load: vi.fn(),
      exportAsBase64: vi.fn(() => ({ value: "BASE64" })),
    };
    const slideA = { id: "slide-1", load: vi.fn() };
    const insertedSlide = { id: "slide-new" };
    const originalSlide = { id: "slide-2", delete: vi.fn() };
    const slides = {
      items: [slideA, sourceSlide],
      load: vi.fn(),
    } as any;

    const updatedSlides = {
      items: [slideA, originalSlide, insertedSlide],
      load: vi.fn(),
    } as any;

    const presentation = {
      slides,
      insertSlidesFromBase64: vi.fn(() => {
        presentation.slides = updatedSlides;
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
      InsertSlideFormatting: {
        keepSourceFormatting: "KeepSourceFormatting",
      },
    });

    await replaceSlideWithMutatedOpenXml(context, 1, (value) => `${value}-mutated`);

    expect(originalSlide.delete).toHaveBeenCalledTimes(1);
  });

  it("round-trips transition duration metadata", () => {
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
    });

    const mutated = setSlideTransitionInBase64Presentation(base64, {
      effect: "fade",
      durationMs: 700,
      advanceOnClick: true,
    });

    expect(extractSlideTransitionFromBase64Presentation(mutated)).toMatchObject({
      effect: "fade",
      durationMs: 700,
      advanceOnClick: true,
    });
  });

  it("maps animation targets to exported XML shape ids", () => {
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
    });

    const mutated = addSlideAnimationInBase64Presentation(base64, {
      shapeId: "office-shape-id",
      type: "scale",
      start: "withPrevious",
      scaleXPercent: 90,
      scaleYPercent: 90,
    }, 1);
    const slideXml = new OpenXmlPackage(mutated).readText("ppt/slides/slide1.xml");
    const slideDoc = parseXml(slideXml);
    const target = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "spTgt")[0];

    expect(target?.getAttribute("spid")).toBe("11");
  });

  it("creates notes slide relationships back to the slide and notes master", () => {
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
      "ppt/notesMasters/notesMaster1.xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
    });

    const mutated = setSpeakerNotesInBase64Presentation(base64, "Hello notes");
    const pkg = new OpenXmlPackage(mutated);
    const slideRels = pkg.readText("ppt/slides/_rels/slide1.xml.rels");
    const notesRels = pkg.readText("ppt/notesSlides/_rels/notesSlide1.xml.rels");
    const contentTypes = pkg.readText("[Content_Types].xml");

    expect(slideRels).toContain("relationships/notesSlide");
    expect(notesRels).toContain("relationships/notesMaster");
    expect(notesRels).toContain("relationships/slide");
    expect(notesRels).toContain("../slides/slide1.xml");
    expect(contentTypes).toContain("/ppt/notesSlides/notesSlide1.xml");
  });
});
