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

export function buildHeaderSubtitle(input: {
  host: OfficeHost;
  runtimeMode: string;
  enabledToolCount: number;
  powerpointContext?: PowerPointContextSnapshot | null;
}) {
  const bits = [
    hostLabel(input.host),
    input.runtimeMode || undefined,
    `${input.enabledToolCount} tools`,
    ...(input.host === "powerpoint" ? powerpointLabel(input.powerpointContext) : []),
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
