import { afterEach, describe, expect, it, vi } from "vitest";

vi.mock("./excelShared", async () => {
  const actual = await vi.importActual<any>("./excelShared");
  return {
    ...actual,
    getWorksheet: vi.fn(),
    isExcelRequirementSetSupported: vi.fn(),
  };
});

describe("getRangeImage", () => {
  afterEach(() => {
    vi.restoreAllMocks();
    vi.resetModules();
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  it("fails when the host does not support range image export", async () => {
    const excelShared = await import("./excelShared");
    vi.mocked(excelShared.isExcelRequirementSetSupported).mockReturnValue(false);

    const { getRangeImage } = await import("./getRangeImage");
    const result = await getRangeImage.handler({ range: "A1:C3" });

    expect(result).toMatchObject({
      resultType: "failure",
      error: "This Excel host cannot export range images. Use a host with ExcelApi 1.7+ and try again.",
    });
  });

  it("returns a PNG payload for a valid range", async () => {
    const excelShared = await import("./excelShared");
    vi.mocked(excelShared.isExcelRequirementSetSupported).mockReturnValue(true);

    const image = { value: "abc123" };
    const range = {
      address: "Sheet1!A1:F7",
      load: vi.fn(),
      getImage: vi.fn(() => image),
    };
    const worksheet = {
      name: "Sheet1",
      getRange: vi.fn(() => range),
    };
    vi.mocked(excelShared.getWorksheet).mockResolvedValue(worksheet as never);

    (globalThis as { Excel?: unknown }).Excel = {
      run: vi.fn(async (callback: (context: { sync: ReturnType<typeof vi.fn> }) => Promise<unknown>) => callback({ sync: vi.fn() })),
    };

    const { getRangeImage } = await import("./getRangeImage");
    const result = await getRangeImage.handler({ range: "A1:F7", sheetName: "Sheet1" });

    expect(worksheet.getRange).toHaveBeenCalledWith("A1:F7");
    expect(range.getImage).toHaveBeenCalled();
    expect(result).toMatchObject({
      resultType: "success",
      textResultForLlm: "Rendered Sheet1!A1:F7 in Sheet1 as a PNG snapshot.",
      binaryResultsForLlm: [
        {
          data: "abc123",
          mimeType: "image/png",
          type: "image",
          description: "Sheet1 Sheet1!A1:F7",
        },
      ],
    });
  });

  it("surfaces Excel debug information on export failures", async () => {
    const excelShared = await import("./excelShared");
    vi.mocked(excelShared.isExcelRequirementSetSupported).mockReturnValue(true);

    const excelError = Object.assign(new Error("The requested operation is invalid."), {
      code: "InvalidOperation",
      debugInfo: {
        message: "Range.getImage is unavailable in this host.",
        errorLocation: "Range.getImage",
      },
    });

    (globalThis as { Excel?: unknown }).Excel = {
      run: vi.fn(async () => {
        throw excelError;
      }),
    };

    const { getRangeImage } = await import("./getRangeImage");
    const result = await getRangeImage.handler({ range: "A1" });

    expect(result).toMatchObject({
      resultType: "failure",
      error: "The requested operation is invalid.: Range.getImage is unavailable in this host. (at Range.getImage) [InvalidOperation]",
    });
  });
});
