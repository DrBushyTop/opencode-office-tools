import type { ToolResultFailure } from "./types";

export type HeaderFooterTarget = "header" | "footer";
export type HeaderFooterTypeName = "primary" | "firstPage" | "evenPages";

export const headerFooterTypeMap: Record<HeaderFooterTypeName, Word.HeaderFooterType> = {
  primary: Word.HeaderFooterType.primary,
  firstPage: Word.HeaderFooterType.firstPage,
  evenPages: Word.HeaderFooterType.evenPages,
};

export function getHeaderFooterBody(section: Word.Section, target: HeaderFooterTarget, type: HeaderFooterTypeName) {
  return target === "header"
    ? section.getHeader(headerFooterTypeMap[type])
    : section.getFooter(headerFooterTypeMap[type]);
}

export function isWordDesktopRequirementSetSupported(version: string) {
  return Office.context.requirements.isSetSupported("WordApiDesktop", version);
}

export function isWordRequirementSetSupported(version: string) {
  return Office.context.requirements.isSetSupported("WordApi", version);
}

export function toolFailure(error: unknown): ToolResultFailure {
  const message = error instanceof Error ? error.message : String(error);
  return { textResultForLlm: message, resultType: "failure", error: message, toolTelemetry: {} };
}

export function summarizePlainText(text: string, limit = 80) {
  const normalized = text.replace(/\s+/g, " ").trim();
  if (!normalized) return "(empty)";
  return normalized.length > limit ? `${normalized.slice(0, limit - 3)}...` : normalized;
}
