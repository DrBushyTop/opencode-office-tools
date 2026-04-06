import { afterEach, describe, expect, it } from "vitest";

describe("sessionStorage", () => {
  afterEach(() => {
    delete (globalThis as { Office?: unknown }).Office;
  });

  it("normalizes Office host names into storage host ids", async () => {
    (globalThis as { Office?: unknown }).Office = {
      HostType: {
        PowerPoint: "PowerPoint",
        Word: "Word",
        Excel: "Excel",
        OneNote: "OneNote",
      },
    };

    const { getHostFromOfficeHost } = await import("./sessionStorage");

    expect(getHostFromOfficeHost("PowerPoint" as never)).toBe("powerpoint");
    expect(getHostFromOfficeHost("Word" as never)).toBe("word");
    expect(getHostFromOfficeHost("Excel" as never)).toBe("excel");
    expect(getHostFromOfficeHost("OneNote" as never)).toBe("onenote");
    expect(getHostFromOfficeHost("Unknown" as never)).toBe("word");
  });
});
