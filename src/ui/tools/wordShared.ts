import type { ToolResultFailure } from "./types";

export type HeaderFooterTarget = "header" | "footer";
export type HeaderFooterTypeName = "primary" | "firstPage" | "evenPages";
export type SectionSelector = number | "*";

export type DocumentPartAddress =
  | { kind: "headersFootersOverview" }
  | { kind: "tableOfContents" }
  | { kind: "section"; section: SectionSelector; target?: HeaderFooterTarget; type?: HeaderFooterTypeName };

function getHeaderFooterType(type: HeaderFooterTypeName): Word.HeaderFooterType {
  switch (type) {
    case "firstPage":
      return Word.HeaderFooterType.firstPage;
    case "evenPages":
      return Word.HeaderFooterType.evenPages;
    case "primary":
    default:
      return Word.HeaderFooterType.primary;
  }
}

export function getHeaderFooterBody(section: Word.Section, target: HeaderFooterTarget, type: HeaderFooterTypeName) {
  return target === "header"
    ? section.getHeader(getHeaderFooterType(type))
    : section.getFooter(getHeaderFooterType(type));
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

export function parseDocumentPartAddress(address: string): DocumentPartAddress | null {
  const normalized = String(address || "").trim();
  if (!normalized) return null;
  if (normalized === "headers_footers") {
    return { kind: "headersFootersOverview" };
  }
  if (normalized === "table_of_contents") {
    return { kind: "tableOfContents" };
  }

  const match = normalized.match(/^section\[(\*|\d+)\](?:\.(header|footer)\.(primary|firstPage|evenPages))?$/);
  if (!match) return null;

  return {
    kind: "section",
    section: match[1] === "*" ? "*" : Number(match[1]),
    target: match[2] as HeaderFooterTarget | undefined,
    type: match[3] as HeaderFooterTypeName | undefined,
  };
}

export function isValidSectionSelector(section: SectionSelector) {
  return section === "*" || (Number.isInteger(section) && section > 0);
}
