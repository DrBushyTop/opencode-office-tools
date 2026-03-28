import { strToU8, zipSync } from "fflate";
import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from "@xmldom/xmldom";
import { describe, expect, it, vi } from "vitest";
import { OpenXmlPackage, parseXml } from "./openXmlPackage";
import {
  addSlideAnimationInBase64Presentation,
  extractSpeakerNotesFromBase64Presentation,
  extractSlideTransitionFromBase64Presentation,
  findSlideShapeIndexByXmlShapeIdInBase64Presentation,
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
  it("inserts before the source slide position and deletes the original by id", async () => {
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
      targetSlideId: "slide-1",
    });
    expect(originalSlide.delete).toHaveBeenCalledTimes(1);
  });

  it("omits targetSlideId for first-slide replacement and deletes the original by id", async () => {
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

  it("finds shape indexes by exported XML shape id", () => {
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
    });

    expect(findSlideShapeIndexByXmlShapeIdInBase64Presentation(base64, "10")).toBe(0);
    expect(findSlideShapeIndexByXmlShapeIdInBase64Presentation(base64, "11")).toBe(1);
    expect(findSlideShapeIndexByXmlShapeIdInBase64Presentation(base64, "999")).toBe(-1);
  });

  it("places withPrevious animations in the same timing group", () => {
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
    });

    const withShortAnimation = addSlideAnimationInBase64Presentation(base64, {
      shapeId: "shape-1",
      type: "scale",
      start: "withPrevious",
      durationMs: 500,
      scaleXPercent: 90,
      scaleYPercent: 90,
    }, 0);
    const withLongAnimation = addSlideAnimationInBase64Presentation(withShortAnimation, {
      shapeId: "shape-2",
      type: "rotate",
      start: "withPrevious",
      durationMs: 1600,
      angleDegrees: 180,
    }, 1);

    const slideDoc = parseXml(new OpenXmlPackage(withLongAnimation).readText("ppt/slides/slide1.xml"));
    // Both animations should be separate per-shape p:par nodes inside the same timing group
    const animScales = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "animScale");
    const animRots = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "animRot");
    expect(animScales.length).toBe(1);
    expect(animRots.length).toBe(1);

    // Both should have their own per-shape p:par with nodeType="withEffect"
    const cTns = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "cTn");
    const withEffectCTns = Array.from(cTns).filter((n) => n.getAttribute("nodeType") === "withEffect");
    expect(withEffectCTns.length).toBe(2);

    // Each should have its own duration
    expect(withEffectCTns[0].getAttribute("dur")).toBe("500");
    expect(withEffectCTns[1].getAttribute("dur")).toBe("1600");
  });

  it("serializes delay for afterPrevious animations", () => {
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
    });

    const mutated = addSlideAnimationInBase64Presentation(base64, {
      shapeId: "shape-1",
      type: "scale",
      start: "afterPrevious",
      delayMs: 250,
      durationMs: 900,
      scaleXPercent: 110,
      scaleYPercent: 110,
    }, 0);

    const slideDoc = parseXml(new OpenXmlPackage(mutated).readText("ppt/slides/slide1.xml"));
    // The delay should be on the per-shape animation cTn's stCondLst, which has nodeType="afterEffect"
    const cTns = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "cTn");
    const afterEffectCTn = Array.from(cTns).find((n) => n.getAttribute("nodeType") === "afterEffect");
    expect(afterEffectCTn).toBeDefined();
    const stCondLst = afterEffectCTn?.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "stCondLst")[0];
    const cond = stCondLst?.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "cond")[0];
    expect(cond?.getAttribute("delay")).toBe("250");
    expect(cond?.hasAttribute("evt")).toBe(false);
  });

  it("creates an appear entrance animation with visibility set", () => {
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
    });

    const mutated = addSlideAnimationInBase64Presentation(base64, {
      shapeId: "shape-1",
      type: "appear",
      start: "onClick",
    }, 0);

    const slideDoc = parseXml(new OpenXmlPackage(mutated).readText("ppt/slides/slide1.xml"));
    const setNodes = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "set");
    expect(setNodes.length).toBe(1);
    const attrName = setNodes[0].getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "attrName")[0];
    expect(attrName?.textContent).toBe("style.visibility");
    const strVal = setNodes[0].getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "strVal")[0];
    expect(strVal?.getAttribute("val")).toBe("visible");
    const cTns = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "cTn");
    const entrCtn = Array.from(cTns).find((n) => n.getAttribute("presetClass") === "entr");
    expect(entrCtn).toBeDefined();
    expect(entrCtn?.getAttribute("presetID")).toBe("1");
  });

  it("creates a fade entrance animation with animEffect and visibility set", () => {
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
    });

    const mutated = addSlideAnimationInBase64Presentation(base64, {
      shapeId: "shape-1",
      type: "fade",
      start: "afterPrevious",
      durationMs: 500,
    }, 0);

    const slideDoc = parseXml(new OpenXmlPackage(mutated).readText("ppt/slides/slide1.xml"));
    const setNodes = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "set");
    expect(setNodes.length).toBe(1);
    const animEffects = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "animEffect");
    expect(animEffects.length).toBe(1);
    expect(animEffects[0].getAttribute("transition")).toBe("in");
    expect(animEffects[0].getAttribute("filter")).toBe("fade");
    const cTns = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "cTn");
    const entrCtn = Array.from(cTns).find((n) => n.getAttribute("presetClass") === "entr");
    expect(entrCtn?.getAttribute("presetID")).toBe("10");
  });

  it("creates a flyIn entrance animation with motion path and direction subtype", () => {
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
    });

    const mutated = addSlideAnimationInBase64Presentation(base64, {
      shapeId: "shape-1",
      type: "flyIn",
      start: "withPrevious",
      direction: "left",
      durationMs: 700,
    }, 0);

    const slideDoc = parseXml(new OpenXmlPackage(mutated).readText("ppt/slides/slide1.xml"));
    const animMotions = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "animMotion");
    expect(animMotions.length).toBe(1);
    expect(animMotions[0].getAttribute("path")).toBe("M -1 0 L 0 0 E");
    expect(animMotions[0].getAttribute("origin")).toBe("layout");
    const cTns = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "cTn");
    const entrCtn = Array.from(cTns).find((n) => n.getAttribute("presetClass") === "entr");
    expect(entrCtn?.getAttribute("presetID")).toBe("2");
    expect(entrCtn?.getAttribute("presetSubtype")).toBe("4");
  });

  it("creates a wipe entrance animation with animEffect filter", () => {
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
    });

    const mutated = addSlideAnimationInBase64Presentation(base64, {
      shapeId: "shape-1",
      type: "wipe",
      start: "onClick",
      direction: "right",
    }, 0);

    const slideDoc = parseXml(new OpenXmlPackage(mutated).readText("ppt/slides/slide1.xml"));
    const animEffects = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "animEffect");
    expect(animEffects.length).toBe(1);
    expect(animEffects[0].getAttribute("filter")).toBe("wipe(right)");
    const cTns = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "cTn");
    const entrCtn = Array.from(cTns).find((n) => n.getAttribute("presetClass") === "entr");
    expect(entrCtn?.getAttribute("presetID")).toBe("22");
    expect(entrCtn?.getAttribute("presetSubtype")).toBe("4");
  });

  it("creates a zoomIn entrance animation with animScale from/to", () => {
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
    });

    const mutated = addSlideAnimationInBase64Presentation(base64, {
      shapeId: "shape-1",
      type: "zoomIn",
      start: "afterPrevious",
      durationMs: 400,
    }, 0);

    const slideDoc = parseXml(new OpenXmlPackage(mutated).readText("ppt/slides/slide1.xml"));
    const animScales = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "animScale");
    expect(animScales.length).toBe(1);
    const from = animScales[0].getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "from")[0];
    const to = animScales[0].getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "to")[0];
    expect(from?.getAttribute("x")).toBe("0");
    expect(from?.getAttribute("y")).toBe("0");
    expect(to?.getAttribute("x")).toBe("100000");
    expect(to?.getAttribute("y")).toBe("100000");
    const cTns = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "cTn");
    const entrCtn = Array.from(cTns).find((n) => n.getAttribute("presetClass") === "entr");
    expect(entrCtn?.getAttribute("presetID")).toBe("23");
  });

  it("supports staggered entrance animations with afterPrevious and delayMs", () => {
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
    });

    const first = addSlideAnimationInBase64Presentation(base64, {
      shapeId: "shape-1",
      type: "fade",
      start: "onClick",
      durationMs: 300,
    }, 0);
    const second = addSlideAnimationInBase64Presentation(first, {
      shapeId: "shape-2",
      type: "fade",
      start: "afterPrevious",
      delayMs: 200,
      durationMs: 300,
    }, 1);

    const slideDoc = parseXml(new OpenXmlPackage(second).readText("ppt/slides/slide1.xml"));
    const conds = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "cond");
    const delayCond = Array.from(conds).find((c) => c.getAttribute("delay") === "200" && !c.hasAttribute("evt"));
    expect(delayCond).toBeDefined();
    const setNodes = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "set");
    expect(setNodes.length).toBe(2);
    const animEffects = slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "animEffect");
    expect(animEffects.length).toBe(2);
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

  it("reads the canonical speaker notes shape written by setSpeakerNotesInBase64Presentation", () => {
    const base64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
      "ppt/notesMasters/notesMaster1.xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
      "ppt/slides/_rels/slide1.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>',
      "ppt/notesSlides/notesSlide1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
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
                <p:cNvPr id="2" name="Other body"/>
                <p:cNvSpPr/>
                <p:nvPr><p:ph type="body" idx="9"/></p:nvPr>
              </p:nvSpPr>
              <p:spPr/>
              <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Legacy body text</a:t></a:r></a:p></p:txBody>
            </p:sp>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="3" name="General text"/>
                <p:cNvSpPr/>
                <p:nvPr/>
              </p:nvSpPr>
              <p:spPr/>
              <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>General note text</a:t></a:r></a:p></p:txBody>
            </p:sp>
          </p:spTree>
        </p:cSld>
        <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
      </p:notes>`,
    });

    const mutated = setSpeakerNotesInBase64Presentation(base64, "Canonical notes");

    expect(extractSpeakerNotesFromBase64Presentation(mutated)).toBe("Canonical notes");
  });
});
