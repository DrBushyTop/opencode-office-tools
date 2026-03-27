import { describe, expect, it } from "vitest";

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
    const { parseDocumentRangeAddress } = await import("./wordShared");

    expect(parseDocumentRangeAddress("selection")).toEqual({ kind: "selection" });
    expect(parseDocumentRangeAddress("bookmark[Clause A]")).toEqual({ kind: "bookmark", name: "Clause A" });
    expect(parseDocumentRangeAddress("content_control[id=12]")).toEqual({ kind: "contentControl", by: "id", value: 12 });
    expect(parseDocumentRangeAddress("content_control[index=3]")).toEqual({ kind: "contentControl", by: "index", value: 3 });
    expect(parseDocumentRangeAddress("table[2].cell[4,5]")).toEqual({ kind: "table", tableIndex: 2, rowIndex: 4, cellIndex: 5 });
    expect(parseDocumentRangeAddress("table[2]")).toEqual({ kind: "table", tableIndex: 2 });
    expect(parseDocumentRangeAddress("content_control[id=0]")).toBeNull();
    expect(parseDocumentRangeAddress("table[1].cell[2]")).toBeNull();
  });
});
