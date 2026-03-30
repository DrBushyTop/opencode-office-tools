import type { OfficeHost } from "../sessionStorage";
import type { PowerPointContextSnapshot } from "../tools/powerpointContext";

type ConnectionInput = {
  isLoading: boolean;
  hasLoaded: boolean;
  hasFailed: boolean;
};

function hostLabel(host: OfficeHost) {
  if (host === "powerpoint") return "PowerPoint";
  if (host === "excel") return "Excel";
  if (host === "onenote") return "OneNote";
  return "Word";
}

function powerpointLabel(context: PowerPointContextSnapshot | null | undefined) {
  return [
    context?.activeSlideIndex !== undefined ? `Slide ${context.activeSlideIndex + 1}` : "No active slide",
    context?.selectedShapeIds.length ? `${context.selectedShapeIds.length} shape${context.selectedShapeIds.length === 1 ? "" : "s"} selected` : "No shapes selected",
  ];
}

export function buildPowerPointContextLabel(context: PowerPointContextSnapshot | null | undefined) {
  return powerpointLabel(context).join(" • ");
}

export function buildHeaderSubtitle(input: {
  host: OfficeHost;
  enabledToolCount: number;
}) {
  const bits = [
    hostLabel(input.host),
    `${input.enabledToolCount} tools`,
  ];

  return bits.filter(Boolean).join(" • ");
}

export function deriveConnectionIndicator(input: ConnectionInput) {
  if (input.isLoading && !input.hasLoaded) {
    return { state: "connecting", label: "Connecting…" } as const;
  }

  if (input.hasFailed && !input.hasLoaded) {
    return { state: "disconnected", label: "Offline" } as const;
  }

  return { state: "connected", label: "Connected" } as const;
}
