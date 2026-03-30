import { describe, expect, it, vi } from "vitest";
import { applyStyleToSelection } from "./applyStyleToSelection";

describe("applyStyleToSelection", () => {
  it("accepts fontSize zero when forwarding selection formatting", async () => {
    const font = { size: 12 };
    const selection = {
      text: "Hello",
      font,
      load: vi.fn(),
    };
    const contextStub = {
      document: {
        getSelection: vi.fn(() => selection),
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    vi.stubGlobal("Word", {
      run: async (callback: (context: typeof contextStub) => Promise<unknown>) => callback(contextStub),
    });

    await expect(applyStyleToSelection.handler({ fontSize: 0 })).resolves.toBe("Applied formatting: 0pt.");
    expect(font.size).toBe(0);
  });
});
