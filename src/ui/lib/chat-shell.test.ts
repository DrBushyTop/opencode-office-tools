import { describe, expect, it } from "vitest";
import type { OfficeHost } from "../sessionStorage";
import type { PowerPointContextSnapshot } from "../tools/powerpointContext";
import { buildHeaderSubtitle, deriveConnectionIndicator } from "./chat-shell";

describe("chat shell helpers", () => {
  describe("buildHeaderSubtitle", () => {
    it("builds the base host/runtime/tool-count subtitle for non-PowerPoint hosts", () => {
      expect(buildHeaderSubtitle({
        host: "word" satisfies OfficeHost,
        runtimeMode: "local",
        enabledToolCount: 14,
      })).toBe("Word • local • 14 tools");

      expect(buildHeaderSubtitle({
        host: "excel" satisfies OfficeHost,
        runtimeMode: "",
        enabledToolCount: 9,
      })).toBe("Excel • 9 tools");
    });

    it("preserves PowerPoint context composition after the base subtitle", () => {
      const powerpointContext: PowerPointContextSnapshot = {
        selectedSlideIds: ["slide-4"],
        selectedShapeIds: ["shape-1", "shape-2"],
        activeSlideId: "slide-4",
        activeSlideIndex: 3,
      };

      expect(buildHeaderSubtitle({
        host: "powerpoint",
        runtimeMode: "remote",
        enabledToolCount: 18,
        powerpointContext,
      })).toBe("PowerPoint • remote • 18 tools • Slide 4 • 2 shapes selected");
    });

    it("keeps empty PowerPoint fallback context when no slide or shapes are active", () => {
      expect(buildHeaderSubtitle({
        host: "powerpoint",
        runtimeMode: "",
        enabledToolCount: 18,
        powerpointContext: null,
      })).toBe("PowerPoint • 18 tools • No active slide • No shapes selected");
    });
  });

  describe("deriveConnectionIndicator", () => {
    it("returns a connecting indicator while the shell status is still loading", () => {
      expect(deriveConnectionIndicator({
        isLoading: true,
        hasLoaded: false,
        hasFailed: false,
      })).toEqual({
        state: "connecting",
        label: "Connecting…",
      });
    });

    it("returns a connected indicator after a successful load", () => {
      expect(deriveConnectionIndicator({
        isLoading: false,
        hasLoaded: true,
        hasFailed: false,
      })).toEqual({
        state: "connected",
        label: "Connected",
      });
    });

    it("returns a disconnected indicator after a failed load", () => {
      expect(deriveConnectionIndicator({
        isLoading: false,
        hasLoaded: false,
        hasFailed: true,
      })).toEqual({
        state: "disconnected",
        label: "Offline",
      });
    });
  });
});
