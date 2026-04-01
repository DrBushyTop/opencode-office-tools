import { strToU8, zipSync } from "fflate";
import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from "@xmldom/xmldom";
import { describe, expect, it, vi } from "vitest";
import { OpenXmlPackage, parseXml } from "./openXmlPackage";
import {
  createChartInBase64Presentation,
  deleteChartInBase64Presentation,
  updateChartInBase64Presentation,
} from "./powerpointChartXml";

if (typeof DOMParser === "undefined") {
  vi.stubGlobal("DOMParser", XmldomParser);
}

if (typeof XMLSerializer === "undefined") {
  vi.stubGlobal("XMLSerializer", XmldomSerializer);
}

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function createPresentationBase64(entries: Record<string, string>) {
  let binary = "";
  zipSync(Object.fromEntries(
    Object.entries(entries).map(([path, contents]) => [path, strToU8(contents)]),
  )).forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary);
}

function baseContentTypesXml() {
  return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"/ppt/slides/slide1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/></Types>";
}

function baseSlideXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <p:sld xmlns:a="${NS_A}" xmlns:p="${NS_P}" xmlns:r="${NS_R}">
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
              <a:off x="381000" y="190500"/>
              <a:ext cx="4572000" cy="685800"/>
            </a:xfrm>
          </p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:spTree>
    </p:cSld>
    <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
  </p:sld>`;
}

function basePresentationBase64() {
  return createPresentationBase64({
    "[Content_Types].xml": baseContentTypesXml(),
    "ppt/slides/slide1.xml": baseSlideXml(),
  });
}

describe("powerpointChartXml", () => {
  it("creates a real chart part, relationship, content-type override, and graphic frame", () => {
    const created = createChartInBase64Presentation(basePresentationBase64(), {
      chartType: "column",
      title: "Revenue",
      categories: ["Q1", "Q2"],
      series: [
        { name: "North", values: [10, 20] },
        { name: "South", values: [15, 25] },
      ],
      stacked: true,
      left: 72,
      top: 108,
      width: 360,
      height: 240,
    });

    expect(created.chartPartPath).toBe("ppt/charts/chart1.xml");
    expect(created.xmlShapeId).toBe("11");

    const pkg = new OpenXmlPackage(created.base64);
    expect(pkg.has("ppt/charts/chart1.xml")).toBe(true);

    const contentTypesDoc = parseXml(pkg.readText("[Content_Types].xml"));
    const overrides = Array.from(contentTypesDoc.getElementsByTagName("Override"));
    const chartOverride = overrides.find((node) => node.getAttribute("PartName") === "/ppt/charts/chart1.xml");
    expect(chartOverride?.getAttribute("ContentType")).toBe("application/vnd.openxmlformats-officedocument.drawingml.chart+xml");

    const relsDoc = parseXml(pkg.readText("ppt/slides/_rels/slide1.xml.rels"));
    const chartRel = Array.from(relsDoc.getElementsByTagName("Relationship")).find((node) => node.getAttribute("Id") === created.relationshipId);
    expect(chartRel?.getAttribute("Type")).toBe("http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart");
    expect(chartRel?.getAttribute("Target")).toBe("../charts/chart1.xml");

    const slideDoc = parseXml(pkg.readText("ppt/slides/slide1.xml"));
    const frame = Array.from(slideDoc.getElementsByTagNameNS(NS_P, "graphicFrame"))
      .find((node) => node.getElementsByTagNameNS(NS_P, "cNvPr")[0]?.getAttribute("id") === created.xmlShapeId);
    expect(frame).toBeTruthy();
    expect(frame?.getElementsByTagNameNS(NS_C, "chart")[0]?.getAttributeNS(NS_R, "id") || frame?.getElementsByTagNameNS(NS_C, "chart")[0]?.getAttribute("r:id")).toBe(created.relationshipId);

    const chartDoc = parseXml(pkg.readText(created.chartPartPath));
    expect(chartDoc.getElementsByTagNameNS(NS_C, "style")[0]?.getAttribute("val")).toBe("2");
    expect(chartDoc.getElementsByTagNameNS(NS_C, "title").length).toBe(1);
    expect(chartDoc.getElementsByTagNameNS(NS_C, "legendPos")[0]?.getAttribute("val")).toBe("t");
    expect(chartDoc.getElementsByTagNameNS(NS_C, "barChart").length).toBe(1);
    expect(chartDoc.getElementsByTagNameNS(NS_C, "barDir")[0]?.getAttribute("val")).toBe("col");
    expect(chartDoc.getElementsByTagNameNS(NS_C, "grouping")[0]?.getAttribute("val")).toBe("stacked");
    expect(chartDoc.getElementsByTagNameNS(NS_C, "overlap")[0]?.getAttribute("val")).toBe("100");
    expect(chartDoc.getElementsByTagNameNS(NS_C, "ser").length).toBe(2);
    expect(chartDoc.getElementsByTagNameNS(NS_C, "dLbls").length).toBe(2);
  });

  it("updates an existing chart in place and rewrites both frame geometry and chart XML", () => {
    const created = createChartInBase64Presentation(basePresentationBase64(), {
      chartType: "column",
      title: "Revenue",
      categories: ["Q1", "Q2"],
      series: [{ name: "North", values: [10, 20] }],
      left: 72,
      top: 108,
      width: 360,
      height: 240,
    });

    const updated = updateChartInBase64Presentation(created.base64, created.xmlShapeId, {
      chartType: "doughnut",
      title: "Mix",
      categories: ["A", "B", "C"],
      series: [{ name: "Share", values: [40, 35, 25] }],
      left: 96,
      top: 144,
      width: 320,
      height: 220,
    });

    expect(updated.xmlShapeId).toBe(created.xmlShapeId);
    expect(updated.chartPartPath).toBe(created.chartPartPath);
    expect(updated.relationshipId).toBe(created.relationshipId);

    const pkg = new OpenXmlPackage(updated.base64);
    expect(pkg.listPaths().filter((path) => /^ppt\/charts\/chart\d+\.xml$/.test(path))).toEqual(["ppt/charts/chart1.xml"]);

    const slideDoc = parseXml(pkg.readText("ppt/slides/slide1.xml"));
    const frame = Array.from(slideDoc.getElementsByTagNameNS(NS_P, "graphicFrame"))
      .find((node) => node.getElementsByTagNameNS(NS_P, "cNvPr")[0]?.getAttribute("id") === created.xmlShapeId);
    const xfrm = frame?.getElementsByTagNameNS(NS_P, "xfrm")[0];
    expect(xfrm?.getElementsByTagNameNS(NS_A, "off")[0]?.getAttribute("x")).toBe(String(96 * 12700));
    expect(xfrm?.getElementsByTagNameNS(NS_A, "off")[0]?.getAttribute("y")).toBe(String(144 * 12700));

    const chartDoc = parseXml(pkg.readText("ppt/charts/chart1.xml"));
    expect(chartDoc.getElementsByTagNameNS(NS_C, "doughnutChart").length).toBe(1);
    expect(chartDoc.getElementsByTagNameNS(NS_C, "holeSize")[0]?.getAttribute("val")).toBe("50");
    expect(Array.from(chartDoc.getElementsByTagNameNS(NS_A, "t")).some((node) => node.textContent === "Mix")).toBe(true);
  });

  it("deletes a chart frame, its relationship, the chart part, and its content-type override", () => {
    const created = createChartInBase64Presentation(basePresentationBase64(), {
      chartType: "line",
      title: "Trend",
      categories: ["Jan", "Feb"],
      series: [{ name: "North", values: [5, 8] }],
    });

    const deleted = deleteChartInBase64Presentation(created.base64, created.xmlShapeId);

    expect(deleted.xmlShapeId).toBe(created.xmlShapeId);
    expect(deleted.chartPartPath).toBe(created.chartPartPath);

    const pkg = new OpenXmlPackage(deleted.base64);
    expect(pkg.has(created.chartPartPath)).toBe(false);

    const contentTypesDoc = parseXml(pkg.readText("[Content_Types].xml"));
    const overrides = Array.from(contentTypesDoc.getElementsByTagName("Override"));
    expect(overrides.some((node) => node.getAttribute("PartName") === "/ppt/charts/chart1.xml")).toBe(false);

    const relsDoc = parseXml(pkg.readText("ppt/slides/_rels/slide1.xml.rels"));
    expect(Array.from(relsDoc.getElementsByTagName("Relationship")).some((node) => node.getAttribute("Id") === created.relationshipId)).toBe(false);

    const slideDoc = parseXml(pkg.readText("ppt/slides/slide1.xml"));
    expect(Array.from(slideDoc.getElementsByTagNameNS(NS_P, "graphicFrame"))
      .some((node) => node.getElementsByTagNameNS(NS_P, "cNvPr")[0]?.getAttribute("id") === created.xmlShapeId)).toBe(false);
  });

  it("rejects unsupported pie and scatter definitions early", () => {
    expect(() => createChartInBase64Presentation(basePresentationBase64(), {
      chartType: "pie",
      categories: ["A", "B"],
      series: [
        { name: "One", values: [1, 2] },
        { name: "Two", values: [3, 4] },
      ],
    })).toThrow(/pie charts require exactly one series/i);

    expect(() => createChartInBase64Presentation(basePresentationBase64(), {
      chartType: "scatter",
      categories: ["A", "B"],
      series: [{ name: "One", values: [1, 2] }],
    })).toThrow(/scatter categories\[0\] must be numeric strings/i);
  });

  it("rejects update and delete when an existing chart has unmanaged dependent relationships", () => {
    const created = createChartInBase64Presentation(basePresentationBase64(), {
      chartType: "column",
      title: "Revenue",
      categories: ["Q1", "Q2"],
      series: [{ name: "North", values: [10, 20] }],
    });
    const pkg = new OpenXmlPackage(created.base64);
    pkg.writeText("ppt/charts/_rels/chart1.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/package\" Target=\"../embeddings/Microsoft_Excel_Sheet1.xlsx\"/></Relationships>");
    const withDependencies = pkg.toBase64();

    expect(() => updateChartInBase64Presentation(withDependencies, created.xmlShapeId, {
      chartType: "line",
      categories: ["Q1", "Q2"],
      series: [{ name: "North", values: [10, 20] }],
    })).toThrow(/dependent chart relationships/i);

    expect(() => deleteChartInBase64Presentation(withDependencies, created.xmlShapeId)).toThrow(/dependent chart relationships/i);
  });

  it("applies fontColor to title, axes, legend, and data labels", () => {
    const created = createChartInBase64Presentation(basePresentationBase64(), {
      chartType: "column",
      title: "Dark Slide Chart",
      categories: ["Q1", "Q2"],
      series: [{ name: "Revenue", values: [10, 20] }],
      fontColor: "#FFFFFF",
    });

    const pkg = new OpenXmlPackage(created.base64);
    const chartXml = pkg.readText(created.chartPartPath);
    const chartDoc = parseXml(chartXml);

    // Title should have explicit white fill
    const titleRichText = chartDoc.getElementsByTagNameNS(NS_C, "title")[0];
    const titleFills = titleRichText?.getElementsByTagNameNS(NS_A, "srgbClr");
    expect(Array.from(titleFills || []).some((node) => node.getAttribute("val") === "FFFFFF")).toBe(true);

    // Axes should have txPr with white fill
    const catAx = chartDoc.getElementsByTagNameNS(NS_C, "catAx")[0];
    const catAxTxPr = catAx?.getElementsByTagNameNS(NS_C, "txPr")[0];
    expect(catAxTxPr).toBeTruthy();
    const catAxFill = catAxTxPr?.getElementsByTagNameNS(NS_A, "srgbClr")[0];
    expect(catAxFill?.getAttribute("val")).toBe("FFFFFF");

    const valAx = chartDoc.getElementsByTagNameNS(NS_C, "valAx")[0];
    const valAxTxPr = valAx?.getElementsByTagNameNS(NS_C, "txPr")[0];
    expect(valAxTxPr).toBeTruthy();
    const valAxFill = valAxTxPr?.getElementsByTagNameNS(NS_A, "srgbClr")[0];
    expect(valAxFill?.getAttribute("val")).toBe("FFFFFF");

    // Legend should have txPr with white fill
    const legend = chartDoc.getElementsByTagNameNS(NS_C, "legend")[0];
    const legendTxPr = legend?.getElementsByTagNameNS(NS_C, "txPr")[0];
    expect(legendTxPr).toBeTruthy();
    const legendFill = legendTxPr?.getElementsByTagNameNS(NS_A, "srgbClr")[0];
    expect(legendFill?.getAttribute("val")).toBe("FFFFFF");

    // Data labels should have txPr with white fill
    const dLbls = chartDoc.getElementsByTagNameNS(NS_C, "dLbls")[0];
    const dLblsTxPr = dLbls?.getElementsByTagNameNS(NS_C, "txPr")[0];
    expect(dLblsTxPr).toBeTruthy();
    const dLblsFill = dLblsTxPr?.getElementsByTagNameNS(NS_A, "srgbClr")[0];
    expect(dLblsFill?.getAttribute("val")).toBe("FFFFFF");
  });

  it("does not emit txPr elements when fontColor is omitted", () => {
    const created = createChartInBase64Presentation(basePresentationBase64(), {
      chartType: "column",
      title: "Default Chart",
      categories: ["Q1", "Q2"],
      series: [{ name: "Revenue", values: [10, 20] }],
    });

    const pkg = new OpenXmlPackage(created.base64);
    const chartDoc = parseXml(pkg.readText(created.chartPartPath));

    // Legend should not have txPr when no fontColor
    const legend = chartDoc.getElementsByTagNameNS(NS_C, "legend")[0];
    expect(legend?.getElementsByTagNameNS(NS_C, "txPr").length).toBe(0);

    // Axes should not have txPr when no fontColor
    const catAx = chartDoc.getElementsByTagNameNS(NS_C, "catAx")[0];
    expect(catAx?.getElementsByTagNameNS(NS_C, "txPr").length).toBe(0);
  });

  it("hides data labels when showDataLabels is false", () => {
    const created = createChartInBase64Presentation(basePresentationBase64(), {
      chartType: "column",
      title: "Clean Chart",
      categories: ["Q1", "Q2"],
      series: [{ name: "Revenue", values: [10, 20] }],
      showDataLabels: false,
    });

    const pkg = new OpenXmlPackage(created.base64);
    const chartDoc = parseXml(pkg.readText(created.chartPartPath));

    const dLbls = chartDoc.getElementsByTagNameNS(NS_C, "dLbls")[0];
    const showVal = dLbls?.getElementsByTagNameNS(NS_C, "showVal")[0];
    expect(showVal?.getAttribute("val")).toBe("0");
  });

  it("hides legend when showLegend is false", () => {
    const created = createChartInBase64Presentation(basePresentationBase64(), {
      chartType: "column",
      title: "No Legend",
      categories: ["Q1", "Q2"],
      series: [{ name: "Revenue", values: [10, 20] }],
      showLegend: false,
    });

    const pkg = new OpenXmlPackage(created.base64);
    const chartDoc = parseXml(pkg.readText(created.chartPartPath));

    expect(chartDoc.getElementsByTagNameNS(NS_C, "legend").length).toBe(0);
  });

  it("sets legend position from legendPosition parameter", () => {
    const created = createChartInBase64Presentation(basePresentationBase64(), {
      chartType: "column",
      title: "Bottom Legend",
      categories: ["Q1", "Q2"],
      series: [{ name: "Revenue", values: [10, 20] }],
      legendPosition: "bottom",
    });

    const pkg = new OpenXmlPackage(created.base64);
    const chartDoc = parseXml(pkg.readText(created.chartPartPath));

    const legendPos = chartDoc.getElementsByTagNameNS(NS_C, "legendPos")[0];
    expect(legendPos?.getAttribute("val")).toBe("b");
  });

  it("accepts fontColor without hash prefix", () => {
    const created = createChartInBase64Presentation(basePresentationBase64(), {
      chartType: "line",
      title: "No Hash",
      categories: ["A", "B"],
      series: [{ name: "S1", values: [1, 2] }],
      fontColor: "00FF00",
    });

    const pkg = new OpenXmlPackage(created.base64);
    const chartDoc = parseXml(pkg.readText(created.chartPartPath));

    const srgbClrs = chartDoc.getElementsByTagNameNS(NS_A, "srgbClr");
    expect(Array.from(srgbClrs).some((node) => node.getAttribute("val") === "00FF00")).toBe(true);
  });
});
