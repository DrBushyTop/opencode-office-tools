import { getDocumentContent } from "./getDocumentContent";
import { setDocumentContent } from "./setDocumentContent";
import { getSelection } from "./getSelection";
import { getPresentationContent } from "./getPresentationContent";
import { getPresentationOverview } from "./getPresentationOverview";
import { getSlideImage } from "./getSlideImage";
import { addSlideFromCode } from "./addSlideFromCode";
import { getWorkbookContent } from "./getWorkbookContent";
import { setWorkbookContent } from "./setWorkbookContent";
import { getSelectedRange } from "./getSelectedRange";
import { setSelectedRange } from "./setSelectedRange";
import { getWorkbookInfo } from "./getWorkbookInfo";

// New Word tools
import { getDocumentOverview } from "./getDocumentOverview";
import { getSelectionText } from "./getSelectionText";
import { getSelectionHtml } from "./getSelectionHtml";
import { insertContentAtSelection } from "./insertContentAtSelection";
import { findAndReplace } from "./findAndReplace";
import { getDocumentSection } from "./getDocumentSection";
import { insertTable } from "./insertTable";
import { applyStyleToSelection } from "./applyStyleToSelection";
import { getDocumentPart } from "./getDocumentPart";
import { setDocumentPart } from "./setDocumentPart";
import { getDocumentRange } from "./getDocumentRange";
import { setDocumentRange } from "./setDocumentRange";
import { findDocumentText } from "./findDocumentText";
import { getDocumentTargets } from "./getDocumentTargets";

// New PowerPoint tools
import { addSlideAnimation } from "./addSlideAnimation";
import { clearSlideAnimations } from "./clearSlideAnimations";
import { getSlideNotes } from "./getSlideNotes";
import { getSlideTransition } from "./getSlideTransition";
import { setSlideNotes } from "./setSlideNotes";
import { setSlideTransition } from "./setSlideTransition";
import { getPresentationStructure } from "./getPresentationStructure";
import { getSlideShapes } from "./getSlideShapes";
import { manageSlide } from "./manageSlide";
import { manageSlideShapes } from "./manageSlideShapes";

// New Excel tools
import { getWorkbookOverview } from "./getWorkbookOverview";
import { findAndReplaceCells } from "./findAndReplaceCells";
import { applyCellFormatting } from "./applyCellFormatting";
import { manageChart } from "./manageChart";
import { manageNamedRange } from "./manageNamedRange";
import { manageRange } from "./manageRange";
import { manageWorksheet } from "./manageWorksheet";
import { manageTable } from "./manageTable";
import { getOfficeToolNames } from "../../shared/office-tool-definitions";

const officeToolHandlers = {
  [getDocumentOverview.name]: getDocumentOverview,
  [getDocumentContent.name]: getDocumentContent,
  [getDocumentPart.name]: getDocumentPart,
  [getDocumentSection.name]: getDocumentSection,
  [setDocumentContent.name]: setDocumentContent,
  [setDocumentPart.name]: setDocumentPart,
  [getSelection.name]: getSelection,
  [getSelectionText.name]: getSelectionText,
  [getSelectionHtml.name]: getSelectionHtml,
  [getDocumentRange.name]: getDocumentRange,
  [setDocumentRange.name]: setDocumentRange,
  [findDocumentText.name]: findDocumentText,
  [getDocumentTargets.name]: getDocumentTargets,
  [insertContentAtSelection.name]: insertContentAtSelection,
  [findAndReplace.name]: findAndReplace,
  [insertTable.name]: insertTable,
  [applyStyleToSelection.name]: applyStyleToSelection,
  [getWorkbookOverview.name]: getWorkbookOverview,
  [getWorkbookInfo.name]: getWorkbookInfo,
  [getWorkbookContent.name]: getWorkbookContent,
  [setWorkbookContent.name]: setWorkbookContent,
  [getSelectedRange.name]: getSelectedRange,
  [setSelectedRange.name]: setSelectedRange,
  [findAndReplaceCells.name]: findAndReplaceCells,
  [applyCellFormatting.name]: applyCellFormatting,
  [manageChart.name]: manageChart,
  [manageNamedRange.name]: manageNamedRange,
  [manageRange.name]: manageRange,
  [manageWorksheet.name]: manageWorksheet,
  [manageTable.name]: manageTable,
  [getPresentationOverview.name]: getPresentationOverview,
  [getPresentationStructure.name]: getPresentationStructure,
  [getPresentationContent.name]: getPresentationContent,
  [getSlideImage.name]: getSlideImage,
  [getSlideShapes.name]: getSlideShapes,
  [addSlideAnimation.name]: addSlideAnimation,
  [clearSlideAnimations.name]: clearSlideAnimations,
  [getSlideNotes.name]: getSlideNotes,
  [getSlideTransition.name]: getSlideTransition,
  [manageSlide.name]: manageSlide,
  [manageSlideShapes.name]: manageSlideShapes,
  [addSlideFromCode.name]: addSlideFromCode,
  [setSlideNotes.name]: setSlideNotes,
  [setSlideTransition.name]: setSlideTransition,
};

export const allOfficeTools = Object.values(officeToolHandlers);

export const wordTools = [
  getDocumentOverview,
  getDocumentContent,
  getDocumentPart,
  getDocumentSection,
  setDocumentContent,
  setDocumentPart,
  getSelection,
  getSelectionText,
  getSelectionHtml,
  getDocumentRange,
  setDocumentRange,
  findDocumentText,
  getDocumentTargets,
  insertContentAtSelection,
  findAndReplace,
  insertTable,
  applyStyleToSelection,
];

export const powerpointTools = [
  getPresentationOverview,
  getPresentationStructure,
  getPresentationContent,
  getSlideImage,
  getSlideShapes,
  addSlideAnimation,
  clearSlideAnimations,
  getSlideNotes,
  getSlideTransition,
  manageSlide,
  manageSlideShapes,
  addSlideFromCode,
  setSlideNotes,
  setSlideTransition,
];

export const excelTools = [
  getWorkbookOverview,
  getWorkbookInfo,
  getWorkbookContent,
  setWorkbookContent,
  getSelectedRange,
  setSelectedRange,
  findAndReplaceCells,
  applyCellFormatting,
  manageChart,
  manageNamedRange,
  manageRange,
  manageWorksheet,
  manageTable,
];

export function getToolsForHost(host: typeof Office.HostType[keyof typeof Office.HostType]) {
  const registryHost = host === Office.HostType.Word
    ? "word"
    : host === Office.HostType.PowerPoint
      ? "powerpoint"
      : host === Office.HostType.Excel
        ? "excel"
        : null;

  if (!registryHost) return [];

  return getOfficeToolNames(registryHost).map((name) => {
    const tool = officeToolHandlers[name as keyof typeof officeToolHandlers];
    if (!tool) {
      throw new Error(`Missing Office tool handler for '${name}'`);
    }
    return tool;
  });
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
