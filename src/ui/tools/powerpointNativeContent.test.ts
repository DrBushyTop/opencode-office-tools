import { describe, expect, it, vi, afterEach } from "vitest";
import { fetchImageUrlAsBase64 } from "./powerpointNativeContent";

describe("powerpointNativeContent", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("rejects invalid image urls before fetching", async () => {
    const fetchMock = vi.fn();
    vi.stubGlobal("fetch", fetchMock);

    await expect(fetchImageUrlAsBase64("not-a-url")).rejects.toThrow(/valid HTTPS URL/i);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("rejects non-https image urls before fetching", async () => {
    const fetchMock = vi.fn();
    vi.stubGlobal("fetch", fetchMock);

    await expect(fetchImageUrlAsBase64("http://example.com/a.png")).rejects.toThrow(/must use HTTPS/i);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("rejects non-image responses", async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      headers: {
        get: vi.fn((name: string) => (name.toLowerCase() === "content-type" ? "text/html" : null)),
      },
    });
    vi.stubGlobal("fetch", fetchMock);

    await expect(fetchImageUrlAsBase64("https://example.com/a.png")).rejects.toThrow(/image content type/i);
  });

  it("rejects oversized image responses by content-length", async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      headers: {
        get: vi.fn((name: string) => {
          if (name.toLowerCase() === "content-type") return "image/png";
          if (name.toLowerCase() === "content-length") return String(10 * 1024 * 1024 + 1);
          return null;
        }),
      },
    });
    vi.stubGlobal("fetch", fetchMock);

    await expect(fetchImageUrlAsBase64("https://example.com/a.png")).rejects.toThrow(/too large/i);
  });
});
