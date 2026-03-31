import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from "@xmldom/xmldom";
import { strToU8, zipSync } from "fflate";
import { describe, expect, it, vi } from "vitest";
import { OpenXmlPackage, parseXml } from "./openXmlPackage";
import {
  getShapeParagraphXmlByTarget,
  inspectSlideXmlFromBase64Presentation,
  replaceShapeParagraphXmlInSlideDocument,
  replaceShapeParagraphXmlInSlideInspection,
  replaceShapeParagraphXmlInBase64Presentation,
  resolveSlideXmlShapeTarget,
} from "./powerpointSlideXml";
import { buildPowerPointShapeRef } from "./powerpointShapeRefs";

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
          <p:spPr>
            <a:xfrm>
              <a:off x="12700" y="25400"/>
              <a:ext cx="381000" cy="508000"/>
            </a:xfrm>
          </p:spPr>
          <p:txBody>
            <a:bodyPr wrap="square" rtlCol="0"/>
            <a:lstStyle><a:lvl1pPr marL="123"/></a:lstStyle>
            <a:p><a:r><a:t>One</a:t></a:r></a:p>
            <a:p><a:r><a:t>Two</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="11" name="Body"/>
            <p:cNvSpPr/>
            <p:nvPr/>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm>
              <a:off x="50800" y="76200"/>
              <a:ext cx="762000" cy="889000"/>
            </a:xfrm>
          </p:spPr>
          <p:txBody>
            <a:bodyPr anchor="ctr"/>
            <a:lstStyle><a:lvl2pPr marL="456"/></a:lstStyle>
            <a:p><a:r><a:t>Body text</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
        <p:graphicFrame>
          <p:nvGraphicFramePr>
            <p:cNvPr id="12" name="Chart 1"/>
            <p:cNvGraphicFramePr/>
            <p:nvPr/>
          </p:nvGraphicFramePr>
          <p:xfrm>
            <a:off x="88900" y="101600"/>
            <a:ext cx="1016000" cy="1143000"/>
          </p:xfrm>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/>
          </a:graphic>
        </p:graphicFrame>
        <p:pic>
          <p:nvPicPr>
            <p:cNvPr id="13" name="Photo"/>
            <p:cNvPicPr/>
            <p:nvPr/>
          </p:nvPicPr>
          <p:blipFill/>
          <p:spPr>
            <a:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="127000" cy="127000"/>
            </a:xfrm>
          </p:spPr>
        </p:pic>
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

function createBase64() {
  return createPresentationBase64({
    "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
    "ppt/slides/slide1.xml": baseSlideXml(),
  });
}

describe("powerpointSlideXml", () => {
  it("inspects slide shapes in XML order and exposes paragraph XML", () => {
    const inspection = inspectSlideXmlFromBase64Presentation(createBase64(), { slideId: "slide-stable-1" });

    expect(inspection.shapes.map((shape) => ({
      index: shape.index,
      xmlShapeId: shape.xmlShapeId,
      name: shape.name,
      type: shape.type,
      hasTextBody: !!shape.textBody,
    }))).toEqual([
      { index: 0, xmlShapeId: "10", name: "Title", type: "sp", hasTextBody: true },
      { index: 1, xmlShapeId: "11", name: "Body", type: "sp", hasTextBody: true },
      { index: 2, xmlShapeId: "12", name: "Chart 1", type: "graphicFrame", hasTextBody: false },
      { index: 3, xmlShapeId: "13", name: "Photo", type: "pic", hasTextBody: false },
    ]);

    const paragraphs = getShapeParagraphXmlByTarget(resolveSlideXmlShapeTarget(inspection, {
      slideId: "slide-stable-1",
      xmlShapeId: "10",
      ref: buildPowerPointShapeRef("slide-stable-1", "10"),
    }));
    expect(paragraphs).toHaveLength(2);
    expect(paragraphs[0]).toContain("<a:t>One</a:t>");
    expect(paragraphs[1]).toContain("<a:t>Two</a:t>");
  });

  it("extracts geometry for graphicFrame shapes", () => {
    const inspection = inspectSlideXmlFromBase64Presentation(createBase64(), { slideId: "slide-stable-1" });
    expect(inspection.shapes[2]).toMatchObject({
      xmlShapeId: "12",
      type: "graphicFrame",
      box: {
        left: 7,
        top: 8,
        width: 80,
        height: 90,
      },
    });
  });

  it("replaces one shape's paragraphs while preserving bodyPr and lstStyle", () => {
    const mutated = replaceShapeParagraphXmlInBase64Presentation(createBase64(), [
      {
        target: {
          slideId: "slide-stable-1",
          xmlShapeId: "10",
          ref: buildPowerPointShapeRef("slide-stable-1", "10"),
        },
        paragraphsXml: [
          '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Updated title</a:t></a:r></a:p>',
        ],
      },
    ], { slideId: "slide-stable-1" });

    const slideDoc = parseXml(new OpenXmlPackage(mutated).readText("ppt/slides/slide1.xml"));
    const titleShape = inspectSlideXmlFromBase64Presentation(mutated, { slideId: "slide-stable-1" }).shapes[0];
    const titleTextBody = titleShape.textBody;

    expect(titleTextBody?.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "bodyPr")[0]?.getAttribute("wrap")).toBe("square");
    expect(titleTextBody?.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "bodyPr")[0]?.getAttribute("rtlCol")).toBe("0");
    expect(titleTextBody?.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "lvl1pPr")[0]?.getAttribute("marL")).toBe("123");

    const paragraphs = getShapeParagraphXmlByTarget(resolveSlideXmlShapeTarget(
      inspectSlideXmlFromBase64Presentation(mutated, { slideId: "slide-stable-1" }),
      { slideId: "slide-stable-1", xmlShapeId: "10", ref: buildPowerPointShapeRef("slide-stable-1", "10") },
    ));
    expect(paragraphs).toHaveLength(1);
    expect(paragraphs[0]).toContain("<a:t>Updated title</a:t>");
    expect(slideDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "t")[0]?.textContent).toBe("Updated title");
  });

  it("replaces multiple shapes in one slide mutation", () => {
    const mutated = replaceShapeParagraphXmlInBase64Presentation(createBase64(), [
      {
        target: {
          slideId: "slide-stable-1",
          xmlShapeId: "10",
          ref: buildPowerPointShapeRef("slide-stable-1", "10"),
        },
        paragraphsXml: ['<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>First</a:t></a:r></a:p>'],
      },
      {
        target: {
          slideId: "slide-stable-1",
          xmlShapeId: "11",
          ref: buildPowerPointShapeRef("slide-stable-1", "11"),
        },
        paragraphsXml: ['<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Second</a:t></a:r></a:p>'],
      },
    ], { slideId: "slide-stable-1" });

    const inspection = inspectSlideXmlFromBase64Presentation(mutated, { slideId: "slide-stable-1" });
    expect(getShapeParagraphXmlByTarget(resolveSlideXmlShapeTarget(inspection, {
      slideId: "slide-stable-1",
      xmlShapeId: "10",
      ref: buildPowerPointShapeRef("slide-stable-1", "10"),
    }))[0]).toContain("<a:t>First</a:t>");
    expect(getShapeParagraphXmlByTarget(resolveSlideXmlShapeTarget(inspection, {
      slideId: "slide-stable-1",
      xmlShapeId: "11",
      ref: buildPowerPointShapeRef("slide-stable-1", "11"),
    }))[0]).toContain("<a:t>Second</a:t>");
  });

  it("rejects malformed paragraph XML before mutating the slide document", () => {
    const inspection = inspectSlideXmlFromBase64Presentation(createBase64(), { slideId: "slide-stable-1" });
    const before = getShapeParagraphXmlByTarget(resolveSlideXmlShapeTarget(inspection, {
      slideId: "slide-stable-1",
      xmlShapeId: "10",
      ref: buildPowerPointShapeRef("slide-stable-1", "10"),
    }));

    expect(() => replaceShapeParagraphXmlInSlideInspection(inspection, [
      {
        target: {
          slideId: "slide-stable-1",
          xmlShapeId: "10",
          ref: buildPowerPointShapeRef("slide-stable-1", "10"),
        },
        paragraphsXml: ['<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Still original</a:t></a:r></a:p>'],
      },
      {
        target: {
          slideId: "slide-stable-1",
          xmlShapeId: "11",
          ref: buildPowerPointShapeRef("slide-stable-1", "11"),
        },
        paragraphsXml: ['<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Broken</a:t></a:r>'],
      },
    ])).toThrow(/Invalid paragraph XML for shape 11 at paragraph 0/i);

    expect(getShapeParagraphXmlByTarget(resolveSlideXmlShapeTarget(inspection, {
      slideId: "slide-stable-1",
      xmlShapeId: "10",
      ref: buildPowerPointShapeRef("slide-stable-1", "10"),
    }))).toEqual(before);
  });

  it("rejects duplicate xml shape ids in document-level batch edits", () => {
    const inspection = inspectSlideXmlFromBase64Presentation(createBase64(), { slideId: "slide-stable-1" });
    const before = getShapeParagraphXmlByTarget(resolveSlideXmlShapeTarget(inspection, {
      slideId: "slide-stable-1",
      xmlShapeId: "10",
      ref: buildPowerPointShapeRef("slide-stable-1", "10"),
    }));

    expect(() => replaceShapeParagraphXmlInSlideDocument(inspection.slideDoc, [
      {
        xmlShapeId: "10",
        paragraphsXml: ['<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>One</a:t></a:r></a:p>'],
      },
      {
        xmlShapeId: "10",
        paragraphsXml: ['<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Two</a:t></a:r></a:p>'],
      },
    ])).toThrow(/Duplicate shape target/i);

    expect(getShapeParagraphXmlByTarget(resolveSlideXmlShapeTarget(inspection, {
      slideId: "slide-stable-1",
      xmlShapeId: "10",
      ref: buildPowerPointShapeRef("slide-stable-1", "10"),
    }))).toEqual(before);
  });

  it("accepts extension and foreign markup that should round-trip safely", () => {
    const paragraphXml = '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:foo="urn:test"><a:r><a:rPr><a:extLst><a:ext uri="demo"><foo:marker foo:flag="yes"/></a:ext></a:extLst></a:rPr><a:t>With extension</a:t></a:r></a:p>';
    const mutated = replaceShapeParagraphXmlInBase64Presentation(createBase64(), [
      {
        target: {
          slideId: "slide-stable-1",
          xmlShapeId: "10",
          ref: buildPowerPointShapeRef("slide-stable-1", "10"),
        },
        paragraphsXml: [paragraphXml],
      },
    ], { slideId: "slide-stable-1" });

    const inspection = inspectSlideXmlFromBase64Presentation(mutated, { slideId: "slide-stable-1" });
    const roundTripped = getShapeParagraphXmlByTarget(resolveSlideXmlShapeTarget(inspection, {
      slideId: "slide-stable-1",
      xmlShapeId: "10",
      ref: buildPowerPointShapeRef("slide-stable-1", "10"),
    }))[0];

    expect(roundTripped).toContain("<foo:marker foo:flag=\"yes\"/>");
    expect(roundTripped).toContain("<a:t>With extension</a:t>");
  });

  it("rejects nested unsupported node types like CDATA", () => {
    expect(() => replaceShapeParagraphXmlInBase64Presentation(createBase64(), [
      {
        target: {
          slideId: "slide-stable-1",
          xmlShapeId: "10",
          ref: buildPowerPointShapeRef("slide-stable-1", "10"),
        },
        paragraphsXml: ['<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t><![CDATA[Bad]]></a:t></a:r></a:p>'],
      },
    ], { slideId: "slide-stable-1" })).toThrow(/unsupported XML node type/i);
  });

  it("rejects duplicate targets in one batch before mutating the slide", () => {
    const inspection = inspectSlideXmlFromBase64Presentation(createBase64(), { slideId: "slide-stable-1" });
    const titleRef = buildPowerPointShapeRef("slide-stable-1", "10");
    const before = getShapeParagraphXmlByTarget(resolveSlideXmlShapeTarget(inspection, {
      slideId: "slide-stable-1",
      xmlShapeId: "10",
      ref: titleRef,
    }));

    expect(() => replaceShapeParagraphXmlInSlideInspection(inspection, [
      {
        target: { slideId: "slide-stable-1", xmlShapeId: "10", ref: titleRef },
        paragraphsXml: ['<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>One</a:t></a:r></a:p>'],
      },
      {
        target: { slideId: "slide-stable-1", xmlShapeId: "10", ref: titleRef },
        paragraphsXml: ['<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Two</a:t></a:r></a:p>'],
      },
    ])).toThrow(/Duplicate shape target/i);

    expect(getShapeParagraphXmlByTarget(resolveSlideXmlShapeTarget(inspection, {
      slideId: "slide-stable-1",
      xmlShapeId: "10",
      ref: titleRef,
    }))).toEqual(before);
  });

  it("rejects multi-slide packages so callers must target an explicit exported slide", () => {
    const multiSlideBase64 = createPresentationBase64({
      "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
      "ppt/slides/slide1.xml": baseSlideXml(),
      "ppt/slides/slide2.xml": baseSlideXml().replace('id="10" name="Title"', 'id="20" name="Second slide"'),
    });

    expect(() => inspectSlideXmlFromBase64Presentation(multiSlideBase64)).toThrow(/single-slide PowerPoint export/i);
  });
});
