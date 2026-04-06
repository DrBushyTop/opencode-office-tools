import { afterEach, describe, expect, it, vi } from "vitest";
import { z } from "zod";

describe("useLocalStorage", () => {
  afterEach(() => {
    vi.resetModules();
    vi.restoreAllMocks();
    delete (globalThis as { localStorage?: unknown; window?: unknown }).localStorage;
    delete (globalThis as { window?: unknown }).window;
  });

  it("reads a stored value and persists updates through the storage adapter", async () => {
    const useSyncExternalStore = vi.fn((subscribe: (onStoreChange: () => void) => () => void, getSnapshot: () => unknown) => {
      const unsubscribe = subscribe(vi.fn());
      unsubscribe();
      return getSnapshot();
    });
    vi.doMock("react", () => ({ useSyncExternalStore }));

    const localStorage = {
      getItem: vi.fn(() => JSON.stringify({ enabled: true })),
      setItem: vi.fn(),
    };
    (globalThis as { localStorage?: unknown; window?: unknown }).localStorage = localStorage;
    (globalThis as { window?: unknown }).window = {
      addEventListener: vi.fn(),
      removeEventListener: vi.fn(),
    };

    const { useLocalStorage } = await import("./useLocalStorage");
    const [value, setter] = useLocalStorage("feature", { enabled: false }, z.object({ enabled: z.boolean() }));

    expect(value).toEqual({ enabled: true });
    expect(localStorage.getItem).toHaveBeenCalledWith("feature");

    setter({ enabled: false });
    expect(localStorage.setItem).toHaveBeenCalledWith("feature", JSON.stringify({ enabled: false }));
  });

  it("falls back to the default value when JSON parsing fails", async () => {
    const useSyncExternalStore = vi.fn((_: unknown, getSnapshot: () => unknown) => getSnapshot());
    vi.doMock("react", () => ({ useSyncExternalStore }));

    (globalThis as { localStorage?: unknown; window?: unknown }).localStorage = {
      getItem: vi.fn(() => "not-json"),
      setItem: vi.fn(),
    };
    (globalThis as { window?: unknown }).window = {
      addEventListener: vi.fn(),
      removeEventListener: vi.fn(),
    };

    const { useLocalStorage } = await import("./useLocalStorage");
    const [value] = useLocalStorage("broken", { enabled: false });

    expect(value).toEqual({ enabled: false });
  });

  it("falls back to the default value when schema validation fails", async () => {
    const useSyncExternalStore = vi.fn((_: unknown, getSnapshot: () => unknown) => getSnapshot());
    vi.doMock("react", () => ({ useSyncExternalStore }));

    (globalThis as { localStorage?: unknown; window?: unknown }).localStorage = {
      getItem: vi.fn(() => JSON.stringify({ enabled: "yes" })),
      setItem: vi.fn(),
    };
    (globalThis as { window?: unknown }).window = {
      addEventListener: vi.fn(),
      removeEventListener: vi.fn(),
    };

    const { useLocalStorage } = await import("./useLocalStorage");
    const [value] = useLocalStorage("feature", { enabled: false }, z.object({ enabled: z.boolean() }));

    expect(value).toEqual({ enabled: false });
  });
});
