import type { OfficeHost } from "../sessionStorage";

const hostLabels: Record<OfficeHost, string> = {
  powerpoint: "PowerPoint",
  word: "Word",
  excel: "Excel",
  onenote: "OneNote",
};

export function normalizeOfficeHost(host: unknown): OfficeHost {
  if (host === Office.HostType.PowerPoint) return "powerpoint";
  if (host === Office.HostType.Excel) return "excel";
  if (host === Office.HostType.OneNote) return "onenote";
  return "word";
}

export function getOfficeHostLabel(host: OfficeHost) {
  return hostLabels[host];
}
