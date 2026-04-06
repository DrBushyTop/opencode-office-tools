import { afterEach, describe, expect, it, vi } from "vitest";
import { getSlideImage } from "./getSlideImage";

describe("getSlideImage", () => {
  afterEach(() => {
    vi.restoreAllMocks();
    delete (globalThis as { PowerPoint?: unknown }).PowerPoint;
  });

  it("rejects invalid slide indexes before calling PowerPoint", async () => {
    const result = await getSlideImage.handler({ slideIndex: -1 });

    expect(result).toMatchObject({
      resultType: "failure",
      error: "slideIndex must be a non-negative integer.",
    });
  });

  it("returns an out-of-range error when the requested slide does not exist", async () => {
    const context = {
      presentation: {
        slides: {
          items: [{ getImageAsBase64: vi.fn() }, { getImageAsBase64: vi.fn() }],
          load: vi.fn(),
        },
      },
      sync: vi.fn(),
    };
    (globalThis as { PowerPoint?: unknown }).PowerPoint = {
      run: vi.fn(async (callback: (context: any) => Promise<unknown>) => callback(context)),
    };

    const result = await getSlideImage.handler({ slideIndex: 5, width: 640 });

    expect(result).toMatchObject({
      resultType: "failure",
      error: "Invalid slideIndex",
      textResultForLlm: "Invalid slideIndex 5. Must be 0-1 (current slide count: 2)",
    });
  });

  it("returns a PNG payload for a valid slide", async () => {
    const imageResult = { value: "abc123" };
    const slide = { getImageAsBase64: vi.fn(() => imageResult) };
    const context = {
      presentation: {
        slides: {
          items: [slide],
          load: vi.fn(),
        },
      },
      sync: vi.fn(),
    };
    (globalThis as { PowerPoint?: unknown }).PowerPoint = {
      run: vi.fn(async (callback: (context: any) => Promise<unknown>) => callback(context)),
    };

    const result = await getSlideImage.handler({ slideIndex: 0, width: 1024 });

    expect(slide.getImageAsBase64).toHaveBeenCalledWith({ width: 1024 });
    expect(result).toMatchObject({
      resultType: "success",
      textResultForLlm: "Rendered slide 1 of 1 as a 1024px PNG snapshot.",
      binaryResultsForLlm: [
        {
          data: "abc123",
          mimeType: "image/png",
          type: "image",
          description: "Slide 1 of 1",
        },
      ],
    });
  });

  it("surfaces the host-version fallback when slide capture is unavailable", async () => {
    const error = Object.assign(new Error("getImageAsBase64 is unavailable"), { code: "InvalidOperation" });
    (globalThis as { PowerPoint?: unknown }).PowerPoint = {
      run: vi.fn(async () => {
        throw error;
      }),
    };

    const result = await getSlideImage.handler({ slideIndex: 0 });

    expect(result).toMatchObject({
      resultType: "failure",
      error: "API not available",
      textResultForLlm: expect.stringContaining("cannot export slide images"),
    });
  });
});
