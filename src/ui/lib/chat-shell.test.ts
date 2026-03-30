import { describe, expect, it } from "vitest";
import type { OfficeHost } from "../sessionStorage";
import type { PowerPointContextSnapshot } from "../tools/powerpointContext";
import { buildHeaderSubtitle, buildPowerPointContextLabel, deriveConnectionIndicator } from "./chat-shell";

describe("chat shell helpers", () => {
  describe("buildHeaderSubtitle", () => {
    it("builds the base host/tool-count subtitle for non-PowerPoint hosts", () => {
      expect(buildHeaderSubtitle({
        host: "word" satisfies OfficeHost,
        enabledToolCount: 14,
      })).toBe("Word • 14 tools");

      expect(buildHeaderSubtitle({
        host: "excel" satisfies OfficeHost,
        enabledToolCount: 9,
      })).toBe("Excel • 9 tools");
    });

    it("keeps the PowerPoint subtitle focused on host and tool count", () => {
      const powerpointContext: PowerPointContextSnapshot = {
        selectedSlideIds: ["slide-4"],
        selectedShapeIds: ["shape-1", "shape-2"],
        activeSlideId: "slide-4",
        activeSlideIndex: 3,
      };

      expect(buildHeaderSubtitle({
        host: "powerpoint",
        enabledToolCount: 18,
      })).toBe("PowerPoint • 18 tools");

      expect(buildPowerPointContextLabel(powerpointContext)).toBe("Slide 4 • 2 shapes selected");
    });

    it("keeps empty PowerPoint fallback context when no slide or shapes are active", () => {
      expect(buildPowerPointContextLabel(null)).toBe("No active slide • No shapes selected");
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
