import { beforeEach, describe, expect, it, vi } from "vitest";
import { z } from "zod";

vi.mock("./powerpointOpenXml", () => ({
  replaceSlideWithMutatedOpenXml: vi.fn(),
}));

vi.mock("./powerpointChartXml", () => ({
  createChartInBase64Presentation: vi.fn(),
  updateChartInBase64Presentation: vi.fn(),
  deleteChartInBase64Presentation: vi.fn(),
  slideChartTypeSchema: z.enum(["column", "bar", "line", "pie", "doughnut", "area", "scatter"]),
  slideChartLegendPositionSchema: z.enum(["top", "bottom", "left", "right"]),
  slideChartDefinitionSchema: {
    parse: (value: unknown) => value,
  },
}));

vi.mock("./powerpointNativeContent", () => ({
  getSlideById: vi.fn(),
}));

import { editSlideChart } from "./editSlideChart";
import { createChartInBase64Presentation, deleteChartInBase64Presentation, updateChartInBase64Presentation } from "./powerpointChartXml";
import { replaceSlideWithMutatedOpenXml } from "./powerpointOpenXml";
import { getSlideById } from "./powerpointNativeContent";

describe("editSlideChart", () => {
  beforeEach(() => {
    vi.resetAllMocks();
    vi.stubGlobal("PowerPoint", {
      run: async (callback: (context: unknown) => Promise<unknown>) => callback({ presentation: { slides: {} } }),
    });
  });

  it("returns refreshed chart refs after create", async () => {
    vi.mocked(createChartInBase64Presentation).mockReturnValue({
      base64: "mutated",
      xmlShapeId: "77",
      chartPartPath: "ppt/charts/chart1.xml",
      relationshipId: "rId5",
    });
    vi.mocked(replaceSlideWithMutatedOpenXml).mockImplementation(async (_context, slideIndex, mutate) => {
      expect(slideIndex).toBe(2);
      expect(mutate("source", {} as never)).toBe("mutated");
      return {
        originalSlideId: "slide-old",
        replacementSlideId: "slide-new",
        finalSlideIndex: 4,
      };
    });

    const result = await editSlideChart.handler({
      action: "create",
      slideIndex: 2,
      chartType: "column",
      title: "Revenue",
      series: [{ name: "North", values: [1, 2] }],
    });

    expect(result).toMatchObject({
      resultType: "success",
      slideId: "slide-new",
      slideIndex: 4,
      xmlShapeId: "77",
      ref: "slide-id:slide-new/shape:77",
    });
  });

  it("resolves the slide from ref for update and returns the refreshed ref", async () => {
    vi.mocked(getSlideById).mockResolvedValue({ slide: { id: "slide-old" } as never, slideIndex: 3 });
    vi.mocked(updateChartInBase64Presentation).mockReturnValue({
      base64: "mutated",
      xmlShapeId: "42",
      chartPartPath: "ppt/charts/chart2.xml",
      relationshipId: "rId9",
    });
    vi.mocked(replaceSlideWithMutatedOpenXml).mockImplementation(async (_context, slideIndex, mutate) => {
      expect(slideIndex).toBe(3);
      expect(mutate("source", {} as never)).toBe("mutated");
      return {
        originalSlideId: "slide-old",
        replacementSlideId: "slide-newer",
        finalSlideIndex: 5,
      };
    });

    const result = await editSlideChart.handler({
      action: "update",
      ref: "slide-id:slide-old/shape:42",
      chartType: "pie",
      series: [{ name: "North", values: [1, 2] }],
    });

    expect(result).toMatchObject({
      resultType: "success",
      slideId: "slide-newer",
      slideIndex: 5,
      xmlShapeId: "42",
      ref: "slide-id:slide-newer/shape:42",
    });
  });

  it("returns deleted target info and replacement slide info for delete", async () => {
    vi.mocked(getSlideById).mockResolvedValue({ slide: { id: "slide-old" } as never, slideIndex: 1 });
    vi.mocked(deleteChartInBase64Presentation).mockReturnValue({
      base64: "mutated",
      xmlShapeId: "42",
      chartPartPath: "ppt/charts/chart2.xml",
      relationshipId: "rId9",
    });
    vi.mocked(replaceSlideWithMutatedOpenXml).mockImplementation(async (_context, slideIndex, mutate) => {
      expect(slideIndex).toBe(1);
      expect(mutate("source", {} as never)).toBe("mutated");
      return {
        originalSlideId: "slide-old",
        replacementSlideId: "slide-after-delete",
        finalSlideIndex: 2,
      };
    });

    const result = await editSlideChart.handler({
      action: "delete",
      ref: "slide-id:slide-old/shape:42",
    });

    expect(result).toMatchObject({
      resultType: "success",
      slideId: "slide-after-delete",
      slideIndex: 2,
      deletedTarget: {
        ref: "slide-id:slide-old/shape:42",
        slideId: "slide-old",
        xmlShapeId: "42",
      },
    });
  });
});
