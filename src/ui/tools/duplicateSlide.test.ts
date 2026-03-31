import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("./powerpointShared", () => ({
  isPowerPointRequirementSetSupported: vi.fn((version: string) => version === "1.8"),
  toolFailure: vi.fn((error: unknown) => ({ resultType: "failure", error: error instanceof Error ? error.message : String(error) })),
}));

vi.mock("./powerpointContext", () => ({
  getPowerPointContextSnapshot: vi.fn(() => ({ activeSlideIndex: 1 })),
}));

import { duplicateSlide } from "./duplicateSlide";

describe("duplicateSlide", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  it("duplicates the active slide immediately after itself by default", async () => {
    const slide0 = { id: "slide-0" };
    const slide1 = {
      id: "slide-1",
      load: vi.fn(),
      exportAsBase64: vi.fn(() => ({ value: "BASE64" })),
    };
    const slide2 = { id: "slide-2", load: vi.fn() };
    const duplicatedSlide = { id: "slide-1-copy", load: vi.fn() };
    const slides = {
      items: [slide0, slide1, slide2],
      load: vi.fn(),
    };
    let wasDuplicated = false;
    const presentation = {
      slides,
      insertSlidesFromBase64: vi.fn(() => {
        wasDuplicated = true;
      }),
    };
    const context = {
      presentation,
      sync: vi.fn(async () => {
        if (wasDuplicated && slides.items.length === 3) {
          slides.items.splice(2, 0, duplicatedSlide);
        }
      }),
    } as unknown as PowerPoint.RequestContext;

    vi.stubGlobal("Office", { context: { requirements: { isSetSupported: vi.fn(() => true) } } });
    vi.stubGlobal("PowerPoint", {
      run: vi.fn(async (callback: (ctx: PowerPoint.RequestContext) => Promise<unknown>) => callback(context)),
    });

    const result = await duplicateSlide.handler();

    expect(presentation.insertSlidesFromBase64).toHaveBeenCalledWith("BASE64", {
      formatting: "KeepSourceFormatting",
      targetSlideId: "slide-1",
    });
    expect(result).toMatchObject({
      resultType: "success",
      sourceIndex: 1,
      targetIndex: 2,
      duplicatedSlideId: "slide-1-copy",
      formatting: "KeepSourceFormatting",
      textResultForLlm: "Duplicated slide 2 to position 3.",
    });
  });

  it("moves the duplicated slide to index 0 when targetIndex is 0", async () => {
    const slide0 = { id: "slide-0", load: vi.fn() };
    const slide1 = {
      id: "slide-1",
      load: vi.fn(),
      exportAsBase64: vi.fn(() => ({ value: "BASE64" })),
    };
    const duplicatedSlide = { id: "slide-1-copy", load: vi.fn(), moveTo: vi.fn() };
    const slides = {
      items: [slide0, slide1],
      load: vi.fn(),
    };
    let wasDuplicated = false;
    const presentation = {
      slides,
      insertSlidesFromBase64: vi.fn(() => {
        wasDuplicated = true;
      }),
    };
    const context = {
      presentation,
      sync: vi.fn(async () => {
        if (wasDuplicated && slides.items.length === 2) {
          slides.items.push(duplicatedSlide);
        }
      }),
    } as unknown as PowerPoint.RequestContext;

    vi.stubGlobal("Office", { context: { requirements: { isSetSupported: vi.fn(() => true) } } });
    vi.stubGlobal("PowerPoint", {
      run: vi.fn(async (callback: (ctx: PowerPoint.RequestContext) => Promise<unknown>) => callback(context)),
    });

    const result = await duplicateSlide.handler({ slideIndex: 1, targetIndex: 0 });

    expect(duplicatedSlide.moveTo).toHaveBeenCalledWith(0);
    expect(result).toMatchObject({
      resultType: "success",
      sourceIndex: 1,
      targetIndex: 0,
      duplicatedSlideId: "slide-1-copy",
      textResultForLlm: "Duplicated slide 2 to position 1.",
    });
  });
});
