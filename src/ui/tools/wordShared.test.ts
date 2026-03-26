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
});
