import { afterEach, describe, expect, it } from "vitest";
import { getOfficeHostLabel, normalizeOfficeHost } from "./officeHost";

describe("officeHost", () => {
  afterEach(() => {
    delete (globalThis as { Office?: unknown }).Office;
  });

  it("normalizes Office host names into session host ids", () => {
    (globalThis as { Office?: unknown }).Office = {
      HostType: {
        PowerPoint: "PowerPoint",
        Word: "Word",
        Excel: "Excel",
      },
    };

    expect(normalizeOfficeHost("PowerPoint")).toBe("powerpoint");
    expect(normalizeOfficeHost("Word")).toBe("word");
    expect(normalizeOfficeHost("Excel")).toBe("excel");
    expect(normalizeOfficeHost("unknown")).toBe("word");
  });

  it("returns human-readable host labels", () => {
    expect(getOfficeHostLabel("powerpoint")).toBe("PowerPoint");
    expect(getOfficeHostLabel("word")).toBe("Word");
    expect(getOfficeHostLabel("excel")).toBe("Excel");
  });
});
