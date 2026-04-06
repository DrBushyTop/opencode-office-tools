import { describe, expect, it, vi } from "vitest";
import { buildGeneratedSlide, normalizeAddSlideCode } from "./addSlideFromCode";

describe("addSlideFromCode helpers", () => {
  it("removes common bootstrap boilerplate", () => {
    const code = normalizeAddSlideCode(`
const pptxgen = require('pptxgenjs');
const pptx = new pptxgen();
const slide = pptx.addSlide();
slide.addText('Hello', { x: 1, y: 1, w: 1, h: 1 });
`);

    expect(code).toBe("slide.addText('Hello', { x: 1, y: 1, w: 1, h: 1 });");
  });

  it("maps generated slide aliases to the injected slide", () => {
    const addText = vi.fn();
    const slide = { addText };
    const pptx = {
      ShapeType: { rect: "rect" },
      AlignH: {},
      AlignV: {},
    };

    buildGeneratedSlide(
      "let s = pptx.addSlide(); s.addText('Hello', { x: 1, y: 1, w: 1, h: 1 });",
      slide,
      pptx,
    );

    expect(addText).toHaveBeenCalledWith("Hello", { x: 1, y: 1, w: 1, h: 1 });
  });

  it("keeps direct slide-only recipes intact", () => {
    const code = normalizeAddSlideCode(`slide.addShape("rect", { x: 1, y: 1, w: 2, h: 1 });`);

    expect(code).toBe(`slide.addShape("rect", { x: 1, y: 1, w: 2, h: 1 });`);
  });

  it("exposes pptxgen helpers through the runtime scope", () => {
    const addShape = vi.fn();
    const slide = { addShape };
    const pptx = {
      ShapeType: { rect: "rect" },
      AlignH: { center: "center" },
      AlignV: { mid: "mid" },
    };

    buildGeneratedSlide(
      `slide.addShape(ShapeType.rect, { x: 1, y: 2, w: 3, h: 1, align: AlignH.center, valign: AlignV.mid });`,
      slide,
      pptx,
    );

    expect(addShape).toHaveBeenCalledWith("rect", { x: 1, y: 2, w: 3, h: 1, align: "center", valign: "mid" });
  });
});
