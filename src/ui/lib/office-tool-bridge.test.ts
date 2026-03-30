import { describe, expect, it, vi } from "vitest";

import { createOfficeToolBridge } from "./office-tool-bridge";

describe("office tool bridge", () => {
  it("disables caching for bridge requests", async () => {
    const fetchMock = vi.fn(async (input: string, init?: RequestInit) => {
      if (input === "/api/office-tools/session") {
        return {
          ok: true,
          json: async () => ({ sessionToken: "session-token" }),
        } satisfies Partial<Response> as Response;
      }

      if (input === "/api/office-tools/register") {
        return {
          ok: true,
          json: async () => ({ executorId: "executor-id" }),
        } satisfies Partial<Response> as Response;
      }

      if (input === "/api/office-tools/poll?executorId=executor-id") {
        return {
          ok: true,
          json: async () => ({}),
        } satisfies Partial<Response> as Response;
      }

      if (input === "/api/office-tools/register/executor-id") {
        return {
          ok: true,
          json: async () => ({ ok: true }),
        } satisfies Partial<Response> as Response;
      }

      throw new Error(`Unexpected fetch: ${input} ${JSON.stringify(init)}`);
    });

    const originalFetch = globalThis.fetch;
    const originalWindow = globalThis.window;
    const fakeWindow = {
      setTimeout: vi.fn(() => 1),
      clearTimeout: vi.fn(() => undefined),
    } as unknown as Window & typeof globalThis;

    vi.stubGlobal("window", fakeWindow);
    globalThis.fetch = fetchMock as typeof fetch;

    try {
      const bridge = createOfficeToolBridge("word", {
        execute: vi.fn(),
      });

      await vi.waitFor(() => expect(fetchMock).toHaveBeenCalledTimes(3));

      expect(fetchMock).toHaveBeenNthCalledWith(1, "/api/office-tools/session", expect.objectContaining({ cache: "no-store" }));
      expect(fetchMock).toHaveBeenNthCalledWith(2, "/api/office-tools/register", expect.objectContaining({ cache: "no-store" }));
      expect(fetchMock).toHaveBeenNthCalledWith(3, "/api/office-tools/poll?executorId=executor-id", expect.objectContaining({ cache: "no-store" }));

      await bridge.stop();

      expect(fetchMock).toHaveBeenNthCalledWith(4, "/api/office-tools/register/executor-id", expect.objectContaining({
        cache: "no-store",
        method: "DELETE",
      }));
      expect(fakeWindow.setTimeout).toHaveBeenCalled();
      expect(fakeWindow.clearTimeout).toHaveBeenCalledWith(1);
    } finally {
      globalThis.fetch = originalFetch;
      if (originalWindow === undefined) {
        vi.unstubAllGlobals();
      } else {
        globalThis.window = originalWindow;
      }
    }
  });
});
