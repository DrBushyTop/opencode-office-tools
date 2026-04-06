import { afterEach, describe, expect, it, vi } from "vitest";

describe("useIsDarkMode", () => {
  afterEach(() => {
    vi.resetModules();
    vi.restoreAllMocks();
    delete (globalThis as { window?: unknown }).window;
  });

  it("uses the current dark-mode media query state and subscribes to changes", async () => {
    const setIsDarkMode = vi.fn();
    const useState = vi.fn((value: boolean | (() => boolean)) => [typeof value === "function" ? value() : value, setIsDarkMode]);
    let cleanup: (() => void) | undefined;
    const useEffect = vi.fn((effect: () => void | (() => void)) => {
      const value = effect();
      cleanup = typeof value === "function" ? value : undefined;
    });

    vi.doMock("react", () => ({ useState, useEffect }));

    const addEventListener = vi.fn();
    const removeEventListener = vi.fn();
    const mediaQuery = {
      matches: true,
      addEventListener,
      removeEventListener,
    };
    (globalThis as { window?: unknown }).window = {
      matchMedia: vi.fn(() => mediaQuery),
    };

    const { useIsDarkMode } = await import("./useIsDarkMode");
    const value = useIsDarkMode();

    expect(value).toBe(true);
    expect(addEventListener).toHaveBeenCalledWith("change", expect.any(Function));

    const changeHandler = addEventListener.mock.calls[0]?.[1] as (event: { matches: boolean }) => void;
    changeHandler({ matches: false });
    expect(setIsDarkMode).toHaveBeenCalledWith(false);

    cleanup?.();
    expect(removeEventListener).toHaveBeenCalledWith("change", changeHandler);
  });
});
