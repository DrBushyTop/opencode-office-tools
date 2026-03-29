import { describe, expect, it, vi } from "vitest";

describe("wordShared", () => {
  it("can be imported without a Word global at module load time", async () => {
    const globalWithWord = globalThis as typeof globalThis & { Word?: unknown };
    const previousWord = globalWithWord.Word;
    // @ts-expect-error test intentionally removes host global
    delete globalThis.Word;

    await expect(import("./wordShared")).resolves.toBeTruthy();

    if (previousWord !== undefined) {
      globalWithWord.Word = previousWord;
    }
  });

  it("parses generic document range addresses", async () => {
    const { DocumentRangeAddressSchema, parseDocumentRangeAddress } = await import("./wordShared");

    expect(parseDocumentRangeAddress("selection")).toEqual({ kind: "selection" });
    expect(parseDocumentRangeAddress("bookmark[Clause A]")).toEqual({ kind: "bookmark", name: "Clause A" });
    expect(parseDocumentRangeAddress("content_control[id=12]")).toEqual({ kind: "contentControl", by: "id", value: 12 });
    expect(parseDocumentRangeAddress("content_control[index=3]")).toEqual({ kind: "contentControl", by: "index", value: 3 });
    expect(parseDocumentRangeAddress("table[2].cell[4,5]")).toEqual({ kind: "table", tableIndex: 2, rowIndex: 4, cellIndex: 5 });
    expect(parseDocumentRangeAddress("table[2]")).toEqual({ kind: "table", tableIndex: 2 });
    expect(parseDocumentRangeAddress("content_control[id=0]")).toBeNull();
    expect(parseDocumentRangeAddress("table[1].cell[2]")).toBeNull();
    expect(DocumentRangeAddressSchema.safeParse({ kind: "table", tableIndex: 1, rowIndex: 1 }).success).toBe(false);
  });

  it("parses structural document part addresses through schema-backed helpers", async () => {
    const { DocumentPartAddressSchema, parseDocumentPartAddress } = await import("./wordShared");

    expect(parseDocumentPartAddress("headers_footers")).toEqual({ kind: "headersFootersOverview" });
    expect(parseDocumentPartAddress("table_of_contents")).toEqual({ kind: "tableOfContents" });
    expect(parseDocumentPartAddress("section[*].header.primary")).toEqual({
      kind: "section",
      section: "*",
      target: "header",
      type: "primary",
    });
    expect(DocumentPartAddressSchema.safeParse({ kind: "section", section: 0 }).success).toBe(false);
  });

  it("returns full OOXML payloads without trimming the package markup", async () => {
    const { readResolvedWordTarget } = await import("./wordShared");
    const fullOoxml = '<?xml version="1.0"?><pkg:package><pkg:part><w:document><w:body><w:p/></w:body></w:document></pkg:part></pkg:package>';
    const context = { sync: async () => undefined } as Word.RequestContext;
    const resolved = {
      kind: "range" as const,
      label: "selection",
      target: {
        getOoxml: () => ({ value: fullOoxml }),
      } as unknown as Word.Range,
    };

    await expect(readResolvedWordTarget(context, resolved, "ooxml")).resolves.toBe(fullOoxml);
  });

  it("rejects clearing whole-table targets to avoid deleting the table", async () => {
    const { writeResolvedWordTarget } = await import("./wordShared");
    const deleteMock = vi.fn();

    expect(() => {
      writeResolvedWordTarget(
        {
          kind: "range",
          label: "table[1]",
          clearBehavior: "reject",
          target: { delete: deleteMock } as unknown as Word.Range,
        },
        "clear",
        "html",
        undefined,
        "replace",
      );
    }).toThrow(/would remove the entire table/i);

    expect(deleteMock).not.toHaveBeenCalled();
  });

  it("deletes range-backed cell targets when clearing", async () => {
    const { writeResolvedWordTarget } = await import("./wordShared");
    const deleteMock = vi.fn();

    writeResolvedWordTarget(
      {
        kind: "range",
        label: "selection",
        target: { delete: deleteMock } as unknown as Word.Range,
      },
      "clear",
      "html",
      undefined,
      "replace",
    );

    expect(deleteMock).toHaveBeenCalledTimes(1);
  });

  it("rejects table cell targets on hosts without WordApi 1.3", async () => {
    const { resolveDocumentRangeTarget } = await import("./wordShared");
    const originalOffice = globalThis.Office;
    const loadMock = vi.fn();
    const syncMock = vi.fn().mockResolvedValue(undefined);

    globalThis.Office = {
      context: {
        requirements: {
          isSetSupported: vi.fn((setName: string, version: string) => !(setName === "WordApi" && version === "1.3")),
        },
      },
    } as unknown as typeof Office;

    const context = {
      document: {
        body: {
          tables: {
            items: [{ getCellOrNullObject: vi.fn() }],
            load: loadMock,
          },
        },
      },
      sync: syncMock,
    } as unknown as Word.RequestContext;

    await expect(
      resolveDocumentRangeTarget(context, { kind: "table", tableIndex: 1, rowIndex: 1, cellIndex: 1 }),
    ).rejects.toThrow("Table targets require WordApi 1.3.");

    expect(loadMock).not.toHaveBeenCalled();
    expect(syncMock).not.toHaveBeenCalled();

    globalThis.Office = originalOffice;
  });

  it("rejects table targets on hosts without WordApi 1.3 before accessing tables", async () => {
    const { resolveDocumentRangeTarget } = await import("./wordShared");
    const originalOffice = globalThis.Office;

    globalThis.Office = {
      context: {
        requirements: {
          isSetSupported: vi.fn((setName: string, version: string) => !(setName === "WordApi" && version === "1.3")),
        },
      },
    } as unknown as typeof Office;

    const bodyAccessError = new Error("tables getter should not run");
    const document = {};
    Object.defineProperty(document, "body", {
      get() {
        throw bodyAccessError;
      },
    });

    const context = {
      document,
      sync: vi.fn().mockResolvedValue(undefined),
    } as unknown as Word.RequestContext;

    await expect(resolveDocumentRangeTarget(context, { kind: "table", tableIndex: 1 })).rejects.toThrow(
      "Table targets require WordApi 1.3.",
    );

    globalThis.Office = originalOffice;
  });

  it("uses WordApi 1.4 for bookmark support checks", async () => {
    const { resolveDocumentRangeTarget } = await import("./wordShared");
    const originalOffice = globalThis.Office;

    globalThis.Office = {
      context: {
        requirements: {
          isSetSupported: vi.fn((setName: string, version: string) => setName === "WordApi" && version === "1.4"),
        },
      },
    } as unknown as typeof Office;

    const range = {
      load: vi.fn(),
      isNullObject: false,
    };
    const syncMock = vi.fn().mockResolvedValue(undefined);
    const context = {
      document: {
        getBookmarkRangeOrNullObject: vi.fn(() => range),
      },
      sync: syncMock,
    } as unknown as Word.RequestContext;

    await expect(resolveDocumentRangeTarget(context, { kind: "bookmark", name: "Clause A" })).resolves.toMatchObject({
      kind: "range",
      label: "bookmark[Clause A]",
      target: range,
    });

    expect(syncMock).toHaveBeenCalledTimes(1);

    globalThis.Office = originalOffice;
  });
});
