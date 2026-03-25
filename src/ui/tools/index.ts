import { getDocumentContent } from "./getDocumentContent";
import { setDocumentContent } from "./setDocumentContent";
import { getSelection } from "./getSelection";
import { getPresentationContent } from "./getPresentationContent";
import { getPresentationOverview } from "./getPresentationOverview";
import { getSlideImage } from "./getSlideImage";
import { setPresentationContent } from "./setPresentationContent";
import { addSlideFromCode } from "./addSlideFromCode";
import { clearSlide } from "./clearSlide";
import { updateSlideShape } from "./updateSlideShape";
import { getWorkbookContent } from "./getWorkbookContent";
import { setWorkbookContent } from "./setWorkbookContent";
import { getSelectedRange } from "./getSelectedRange";
import { setSelectedRange } from "./setSelectedRange";
import { getWorkbookInfo } from "./getWorkbookInfo";

// New Word tools
import { getDocumentOverview } from "./getDocumentOverview";
import { getSelectionText } from "./getSelectionText";
import { insertContentAtSelection } from "./insertContentAtSelection";
import { findAndReplace } from "./findAndReplace";
import { getDocumentSection } from "./getDocumentSection";
import { insertTable } from "./insertTable";
import { applyStyleToSelection } from "./applyStyleToSelection";

// New PowerPoint tools
import { getSlideNotes } from "./getSlideNotes";
import { setSlideNotes } from "./setSlideNotes";
import { duplicateSlide } from "./duplicateSlide";

// New Excel tools
import { getWorkbookOverview } from "./getWorkbookOverview";
import { findAndReplaceCells } from "./findAndReplaceCells";
import { insertChart } from "./insertChart";
import { applyCellFormatting } from "./applyCellFormatting";
import { createNamedRange } from "./createNamedRange";
import { getOfficeToolNames } from "../../shared/office-tool-definitions";

export const wordTools = [
  getDocumentOverview,
  getDocumentContent,
  getDocumentSection,
  setDocumentContent,
  getSelection,
  getSelectionText,
  insertContentAtSelection,
  findAndReplace,
  insertTable,
  applyStyleToSelection,
];

export const powerpointTools = [
  getPresentationOverview,
  getPresentationContent,
  getSlideImage,
  getSlideNotes,
  setPresentationContent,
  addSlideFromCode,
  clearSlide,
  updateSlideShape,
  setSlideNotes,
  duplicateSlide,
];

export const excelTools = [
  getWorkbookOverview,
  getWorkbookInfo,
  getWorkbookContent,
  setWorkbookContent,
  getSelectedRange,
  setSelectedRange,
  findAndReplaceCells,
  insertChart,
  applyCellFormatting,
  createNamedRange,
];

export function getToolsForHost(host: typeof Office.HostType[keyof typeof Office.HostType]) {
  switch (host) {
    case Office.HostType.Word:
      return wordTools;
    case Office.HostType.PowerPoint:
      return powerpointTools;
    case Office.HostType.Excel:
      return excelTools;
    default:
      return [];
  }
}

export function getOfficeToolExecutor(host: typeof Office.HostType[keyof typeof Office.HostType]) {
  const tools = getToolsForHost(host);
  const map = new Map(tools.map((tool) => [tool.name, tool.handler]));

  return {
    async execute(toolName: string, args: Record<string, unknown>) {
      const handler = map.get(toolName);
      if (!handler) {
        throw new Error(`Tool '${toolName}' is not available for this host`);
      }

      const result = await handler(args as never, {
        sessionId: "office-bridge",
        toolCallId: crypto.randomUUID(),
        toolName,
        arguments: args,
      } as never);

      if (typeof result === "string") return result;
      if ((result as any).resultType === "failure") {
        throw new Error((result as any).error || (result as any).textResultForLlm || "Tool execution failed");
      }

      return (result as any).textResultForLlm || result;
    },
  };
}

export function getToolNamesForHost(host: "word" | "powerpoint" | "excel") {
  return getOfficeToolNames(host);
}
