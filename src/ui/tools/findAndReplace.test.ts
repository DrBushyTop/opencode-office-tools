import { describe, expect, it } from "vitest";
import { findAndReplace } from "./findAndReplace";
import { findDocumentText } from "./findDocumentText";

describe("findAndReplace", () => {
  it("rejects whitespace-only search text", async () => {
    const result = await findAndReplace.handler({ find: "   ", replace: "x" });

    expect(result).toMatchObject({
      resultType: "failure",
      error: "Search text cannot be empty.",
    });
  });

  it("mentions table cell scopes in unsupported-address guidance", async () => {
    const replaceResult = await findAndReplace.handler({ find: "x", replace: "y", address: "bad" });
    const findResult = await findDocumentText.handler({ find: "x", address: "bad" });

    expect(replaceResult).toMatchObject({
      resultType: "failure",
      error: expect.stringContaining("table[1].cell[2,3]"),
    });
    expect(findResult).toMatchObject({
      resultType: "failure",
      error: expect.stringContaining("table[1].cell[2,3]"),
    });
  });

  it("validates numeric search options with zod", async () => {
    const result = await findDocumentText.handler({ find: "x", maxResults: 0 });

    expect(result).toMatchObject({
      resultType: "failure",
      error: expect.stringContaining("maxResults"),
    });
  });
});
