import type { Tool, ToolArguments, ToolHandlerResult } from "./types";
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
import { getSlideAnimations } from "./getSlideAnimations";
import { getSlideNotes } from "./getSlideNotes";
import { getSlideTransition } from "./getSlideTransition";
import { setSlideNotes } from "./setSlideNotes";
import { setSlideTransition } from "./setSlideTransition";
import { getPresentationStructure } from "./getPresentationStructure";
import { getSlideShapes } from "./getSlideShapes";
import { manageSlide } from "./manageSlide";
import { manageSlideShapes } from "./manageSlideShapes";
import { manageSlideMedia } from "./manageSlideMedia";
import { manageSlideTable } from "./manageSlideTable";
import { manageSlideChart } from "./manageSlideChart";
import { insertBusinessLayout } from "./insertBusinessLayout";
import { createSlideFromTemplate } from "./createSlideFromTemplate";

// New Excel tools
import { getWorkbookOverview } from "./getWorkbookOverview";
import { findAndReplaceCells } from "./findAndReplaceCells";
import { applyCellFormatting } from "./applyCellFormatting";
import { manageChart } from "./manageChart";
import { manageNamedRange } from "./manageNamedRange";
import { manageRange } from "./manageRange";
import { manageWorksheet } from "./manageWorksheet";
import { manageTable } from "./manageTable";
import { getNotebookOverview } from "./getNotebookOverview";
import { getPageContent } from "./getPageContent";
import { getNoteSelection } from "./getNoteSelection";
import { setNoteSelection } from "./setNoteSelection";
import { createPage } from "./createPage";
import { setPageTitle } from "./setPageTitle";
import { appendPageContent } from "./appendPageContent";
import { navigateToPage } from "./navigateToPage";
import { getOfficeToolNames } from "../../shared/office-tool-definitions";
import { isToolResultFailure } from "./toolShared";

type OfficeHost = typeof Office.HostType[keyof typeof Office.HostType];
type RegistryHost = "word" | "powerpoint" | "excel" | "onenote";

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
  [getNotebookOverview.name]: getNotebookOverview,
  [getPageContent.name]: getPageContent,
  [getNoteSelection.name]: getNoteSelection,
  [setNoteSelection.name]: setNoteSelection,
  [createPage.name]: createPage,
  [setPageTitle.name]: setPageTitle,
  [appendPageContent.name]: appendPageContent,
  [navigateToPage.name]: navigateToPage,
  [getPresentationOverview.name]: getPresentationOverview,
  [getPresentationStructure.name]: getPresentationStructure,
  [getPresentationContent.name]: getPresentationContent,
  [getSlideImage.name]: getSlideImage,
  [getSlideShapes.name]: getSlideShapes,
  [addSlideAnimation.name]: addSlideAnimation,
  [clearSlideAnimations.name]: clearSlideAnimations,
  [getSlideAnimations.name]: getSlideAnimations,
  [getSlideNotes.name]: getSlideNotes,
  [getSlideTransition.name]: getSlideTransition,
  [manageSlide.name]: manageSlide,
  [manageSlideShapes.name]: manageSlideShapes,
  [manageSlideMedia.name]: manageSlideMedia,
  [manageSlideTable.name]: manageSlideTable,
  [manageSlideChart.name]: manageSlideChart,
  [insertBusinessLayout.name]: insertBusinessLayout,
  [createSlideFromTemplate.name]: createSlideFromTemplate,
  [addSlideFromCode.name]: addSlideFromCode,
  [setSlideNotes.name]: setSlideNotes,
  [setSlideTransition.name]: setSlideTransition,
} satisfies Record<string, Tool>;

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
  getSlideAnimations,
  getSlideNotes,
  getSlideTransition,
  manageSlide,
  manageSlideShapes,
  manageSlideMedia,
  manageSlideTable,
  manageSlideChart,
  insertBusinessLayout,
  createSlideFromTemplate,
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

export const onenoteTools = [
  getNotebookOverview,
  getPageContent,
  getNoteSelection,
  setNoteSelection,
  createPage,
  setPageTitle,
  appendPageContent,
  navigateToPage,
];

function resolveRegistryHost(host: OfficeHost): RegistryHost | null {
  return host === Office.HostType.Word
    ? "word"
    : host === Office.HostType.PowerPoint
      ? "powerpoint"
      : host === Office.HostType.Excel
        ? "excel"
        : host === Office.HostType.OneNote
          ? "onenote"
          : null;
}

export function getToolsForHost(host: OfficeHost): Tool[] {
  const registryHost = resolveRegistryHost(host);

  if (!registryHost) return [];

  return getOfficeToolNames(registryHost).map((name) => {
    const tool = officeToolHandlers[name as keyof typeof officeToolHandlers];
    if (!tool) {
      throw new Error(`Missing Office tool handler for '${name}'`);
    }
    return tool;
  });
}

function createToolExecutionContext(toolName: string, args: ToolArguments) {
  return {
    sessionId: "office-bridge",
    toolCallId: crypto.randomUUID(),
    toolName,
    arguments: args,
  };
}

export function normalizeToolExecutionResult(result: ToolHandlerResult) {
  if (typeof result === "string") return result;
  if (isToolResultFailure(result)) {
    throw new Error(result.error || result.textResultForLlm || "Tool execution failed");
  }
  return result;
}

export function getOfficeToolExecutor(host: OfficeHost) {
  const tools = getToolsForHost(host);
  const map = new Map(tools.map((tool) => [tool.name, tool.handler]));

  return {
    async execute(toolName: string, args: ToolArguments) {
      const handler = map.get(toolName);
      if (!handler) {
        throw new Error(`Tool '${toolName}' is not available for this host`);
      }

      const result = await handler(args as never, createToolExecutionContext(toolName, args));
      return normalizeToolExecutionResult(result);
    },
  };
}

export function getToolNamesForHost(host: "word" | "powerpoint" | "excel" | "onenote") {
  return getOfficeToolNames(host);
}
