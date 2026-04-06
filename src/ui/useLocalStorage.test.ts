import { afterEach, describe, expect, it, vi } from "vitest";
import { z } from "zod";

describe("useLocalStorage", () => {
  afterEach(() => {
    vi.resetModules();
    vi.restoreAllMocks();
    delete (globalThis as { localStorage?: unknown }).localStorage;
  });

  it("reads a stored value through the lazy state initializer and persists it", async () => {
    const setValue = vi.fn();
    const useState = vi.fn((value: unknown | (() => unknown)) => [typeof value === "function" ? value() : value, setValue]);
    const useEffect = vi.fn((effect: () => void | (() => void)) => effect());
    vi.doMock("react", () => ({ useState, useEffect }));

    const localStorage = {
      getItem: vi.fn(() => JSON.stringify({ enabled: true })),
      setItem: vi.fn(),
    };
    (globalThis as { localStorage?: unknown }).localStorage = localStorage;

    const { useLocalStorage } = await import("./useLocalStorage");
    const [value, setter] = useLocalStorage("feature", { enabled: false }, z.object({ enabled: z.boolean() }));

    expect(value).toEqual({ enabled: true });
    expect(setter).toBe(setValue);
    expect(localStorage.getItem).toHaveBeenCalledWith("feature");
    expect(localStorage.setItem).toHaveBeenCalledWith("feature", JSON.stringify({ enabled: true }));
  });

  it("falls back to the default value when JSON parsing fails", async () => {
    const useState = vi.fn((value: unknown | (() => unknown)) => [typeof value === "function" ? value() : value, vi.fn()]);
    const useEffect = vi.fn();
    vi.doMock("react", () => ({ useState, useEffect }));

    (globalThis as { localStorage?: unknown }).localStorage = {
      getItem: vi.fn(() => "not-json"),
      setItem: vi.fn(),
    };

    const { useLocalStorage } = await import("./useLocalStorage");
    const [value] = useLocalStorage("broken", { enabled: false });

    expect(value).toEqual({ enabled: false });
  });

  it("falls back to the default value when schema validation fails", async () => {
    const useState = vi.fn((value: unknown | (() => unknown)) => [typeof value === "function" ? value() : value, vi.fn()]);
    const useEffect = vi.fn();
    vi.doMock("react", () => ({ useState, useEffect }));

    (globalThis as { localStorage?: unknown }).localStorage = {
      getItem: vi.fn(() => JSON.stringify({ enabled: "yes" })),
      setItem: vi.fn(),
    };

    const { useLocalStorage } = await import("./useLocalStorage");
    const [value] = useLocalStorage("feature", { enabled: false }, z.object({ enabled: z.boolean() }));

    expect(value).toEqual({ enabled: false });
  });
});
