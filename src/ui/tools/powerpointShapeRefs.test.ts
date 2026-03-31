import { describe, expect, it } from "vitest";
import { buildPowerPointShapeRef, enrichShapeSummariesWithRefs, parsePowerPointShapeRef } from "./powerpointShapeRefs";

describe("powerpointShapeRefs", () => {
  it("builds and parses stable shape refs", () => {
    const ref = buildPowerPointShapeRef("slide/alpha:#1", "42");

    expect(ref).toBe("slide-id:slide%2Falpha%3A%231/shape:42");
    expect(parsePowerPointShapeRef(ref)).toEqual({ slideId: "slide/alpha:#1", xmlShapeId: "42", ref });
  });

  it("rejects malformed refs clearly", () => {
    expect(() => parsePowerPointShapeRef("shape:42")).toThrow(/Invalid PowerPoint shape ref/i);
    expect(() => parsePowerPointShapeRef("slide-id:/shape:42")).toThrow(/Invalid PowerPoint shape ref/i);
    expect(() => parsePowerPointShapeRef("slide-id:slide-1/shape:")).toThrow(/Invalid PowerPoint shape ref/i);
    expect(() => parsePowerPointShapeRef("slide-id:slide-1/shape:abc")).toThrow(/Invalid PowerPoint shape ref/i);
    expect(() => parsePowerPointShapeRef("slide-id:%E0%A4%A/shape:42")).toThrow(/Invalid PowerPoint shape ref/i);
    expect(() => buildPowerPointShapeRef("   ", "42")).toThrow(/Invalid PowerPoint shape ref/i);
  });

  it("maps Office shape summaries to XML-first refs by index", () => {
    const enriched = enrichShapeSummariesWithRefs("slide-stable-99", [
      {
        index: 0,
        id: "office-1",
        name: "Duplicate",
        type: "TextBox",
        left: 10,
        top: 20,
        width: 30,
        height: 40,
        text: "A",
      },
      {
        index: 1,
        id: "office-2",
        name: "Duplicate",
        type: "TextBox",
        left: 11,
        top: 21,
        width: 31,
        height: 41,
        text: "B",
      },
    ], [
      {
        index: 0,
        xmlShapeId: "10",
        name: "Duplicate",
        type: "sp",
        shapeElement: {} as Element,
        textBody: {} as Element,
        box: { left: 10, top: 20, width: 30, height: 40 },
      },
      {
        index: 1,
        xmlShapeId: "11",
        name: "Duplicate",
        type: "sp",
        shapeElement: {} as Element,
        textBody: {} as Element,
        box: { left: 11, top: 21, width: 31, height: 41 },
      },
    ]);

    expect(enriched.map((shape: { id: string; xmlShapeId: string; ref: string }) => ({ id: shape.id, xmlShapeId: shape.xmlShapeId, ref: shape.ref }))).toEqual([
      { id: "office-1", xmlShapeId: "10", ref: "slide-id:slide-stable-99/shape:10" },
      { id: "office-2", xmlShapeId: "11", ref: "slide-id:slide-stable-99/shape:11" },
    ]);
  });

  it("rejects Office/XML count mismatches before assigning refs", () => {
    expect(() => enrichShapeSummariesWithRefs("slide-stable-99", [
      {
        index: 0,
        id: "office-1",
        name: "Only one",
        type: "TextBox",
        left: 10,
        top: 20,
        width: 30,
        height: 40,
      },
    ], [])).toThrow(/Shape count mismatch/i);
  });

  it("rejects Office/XML order mismatches before assigning refs", () => {
    expect(() => enrichShapeSummariesWithRefs("slide-stable-99", [
      {
        index: 0,
        id: "office-1",
        name: "First",
        type: "TextBox",
        left: 10,
        top: 20,
        width: 30,
        height: 40,
      },
      {
        index: 1,
        id: "office-2",
        name: "Second",
        type: "TextBox",
        left: 50,
        top: 60,
        width: 70,
        height: 80,
      },
    ], [
      {
        index: 0,
        xmlShapeId: "10",
        name: "First",
        type: "sp",
        shapeElement: {} as Element,
        textBody: {} as Element,
        box: { left: 50, top: 60, width: 70, height: 80 },
      },
      {
        index: 1,
        xmlShapeId: "11",
        name: "Second",
        type: "sp",
        shapeElement: {} as Element,
        textBody: {} as Element,
        box: { left: 10, top: 20, width: 30, height: 40 },
      },
    ])).toThrow(/Shape order mismatch/i);
  });

  it("can skip text-presence alignment when text was not loaded", () => {
    const enriched = enrichShapeSummariesWithRefs("slide-stable-99", [
      {
        index: 0,
        id: "office-1",
        name: "Title",
        type: "TextBox",
        left: 10,
        top: 20,
        width: 30,
        height: 40,
      },
    ], [
      {
        index: 0,
        xmlShapeId: "10",
        name: "Title",
        type: "sp",
        shapeElement: {} as Element,
        textBody: {} as Element,
        box: { left: 10, top: 20, width: 30, height: 40 },
      },
    ], { enforceTextPresence: false });

    expect(enriched[0]?.ref).toBe("slide-id:slide-stable-99/shape:10");
  });
});
