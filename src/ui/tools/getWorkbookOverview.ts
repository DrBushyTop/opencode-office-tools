import type { Tool } from "./types";
import { isExcelRequirementSetSupported, toolFailure } from "./excelShared";

type WorksheetOverview = {
  id: string;
  name: string;
  position: number;
  visibility: string;
  protection: string;
  usedRangeText: string;
  usedCellCount: number;
  chartCount: number;
  tableCount: number;
  pivotCount: number;
  autoFilterText?: string;
  frozenPaneText?: string;
  tableDetails?: string;
  worksheetNames?: string;
};

function summarizeWorksheet(sheet: WorksheetOverview, activeSheetName: string) {
  const lines = [
    `${sheet.position + 1}. ${sheet.name}${sheet.name === activeSheetName ? " (active)" : ""}`,
    `   identity: id=${sheet.id}, visibility=${sheet.visibility}, protection=${sheet.protection}`,
    `   footprint: ${sheet.usedRangeText}`,
    `   objects: charts=${sheet.chartCount}, tables=${sheet.tableCount}, pivotTables=${sheet.pivotCount}`,
  ];

  if (sheet.autoFilterText) lines.push(`   filtering: ${sheet.autoFilterText}`);
  if (sheet.frozenPaneText) lines.push(`   frozen panes: ${sheet.frozenPaneText}`);
  if (sheet.tableDetails) lines.push(`   table catalog: ${sheet.tableDetails}`);
  if (sheet.worksheetNames) lines.push(`   local names: ${sheet.worksheetNames}`);
  return lines;
}

function summarizeWorkbookTotals(worksheets: WorksheetOverview[]) {
  return {
    cells: worksheets.reduce((sum, sheet) => sum + sheet.usedCellCount, 0),
    charts: worksheets.reduce((sum, sheet) => sum + sheet.chartCount, 0),
    tables: worksheets.reduce((sum, sheet) => sum + sheet.tableCount, 0),
    pivotTables: worksheets.reduce((sum, sheet) => sum + sheet.pivotCount, 0),
  };
}

async function collectWorksheetOverviews(context: Excel.RequestContext, supportsFreeze: boolean, supportsAutoFilter: boolean, supportsPivotTables: boolean) {
  const workbook = context.workbook;
  const sheets = workbook.worksheets;
  const names = workbook.names;

  sheets.load("items");
  names.load("items");
  await context.sync();

  const activeSheet = workbook.worksheets.getActiveWorksheet();
  activeSheet.load("name");

  const usedRanges = sheets.items.map((sheet) => sheet.getUsedRangeOrNullObject());
  const chartCounts = sheets.items.map((sheet) => sheet.charts.getCount());
  const tableCounts = sheets.items.map((sheet) => sheet.tables.getCount());
  const pivotCounts = supportsPivotTables ? sheets.items.map((sheet) => sheet.pivotTables.getCount()) : [];
  const tableCollections = sheets.items.map((sheet) => sheet.tables);
  const sheetNames = sheets.items.map((sheet) => sheet.names);
  const protections = sheets.items.map((sheet) => sheet.protection);
  const freezeRanges = supportsFreeze ? sheets.items.map((sheet) => sheet.freezePanes.getLocationOrNullObject()) : [];
  const autoFilters = supportsAutoFilter ? sheets.items.map((sheet) => sheet.autoFilter) : [];

  for (const [index, sheet] of sheets.items.entries()) {
    sheet.load(["id", "name", "position", "visibility"]);
    usedRanges[index].load(["isNullObject", "address", "rowCount", "columnCount"]);
    tableCollections[index].load("items/name,items/style,items/showTotals");
    sheetNames[index].load("items/name");
    protections[index].load("protected");
    if (supportsFreeze) {
      freezeRanges[index].load(["isNullObject", "address"]);
    }
    if (supportsAutoFilter) {
      autoFilters[index].load(["enabled", "isDataFiltered"]);
    }
  }

  for (const name of names.items) {
    name.load(["name", "value", "scope"]);
  }

  await context.sync();

  const worksheetSummaries: WorksheetOverview[] = sheets.items.map((sheet, index) => {
    const usedRange = usedRanges[index];
    const chartCount = chartCounts[index].value;
    const tableCount = tableCounts[index].value;
    const pivotCount = supportsPivotTables ? pivotCounts[index].value : 0;
    const usedRangeText = usedRange.isNullObject
      ? "empty"
      : `${usedRange.address} (${usedRange.rowCount} rows x ${usedRange.columnCount} cols)`;

    return {
      id: sheet.id,
      name: sheet.name,
      position: sheet.position,
      visibility: sheet.visibility || "Visible",
      protection: protections[index].protected ? "protected" : "unprotected",
      usedRangeText,
      usedCellCount: usedRange.isNullObject ? 0 : usedRange.rowCount * usedRange.columnCount,
      chartCount,
      tableCount,
      pivotCount,
      autoFilterText: supportsAutoFilter
        ? `${autoFilters[index].enabled ? "enabled" : "disabled"}; filtered=${autoFilters[index].isDataFiltered ? "yes" : "no"}`
        : undefined,
      frozenPaneText: supportsFreeze
        ? (freezeRanges[index].isNullObject ? "none" : freezeRanges[index].address)
        : undefined,
      tableDetails: tableCollections[index].items.length
        ? tableCollections[index].items.map((table) => `${table.name} (${table.style}${table.showTotals ? ", totals" : ""})`).join(", ")
        : undefined,
      worksheetNames: sheetNames[index].items.length
        ? sheetNames[index].items.map((name) => name.name).join(", ")
        : undefined,
    };
  });

  return {
    activeSheetName: activeSheet.name,
    worksheetSummaries,
    workbookNames: names.items.map((name) => ({ name: name.name, value: name.value, scope: name.scope })),
  };
}

function renderWorkbookOverview(data: Awaited<ReturnType<typeof collectWorksheetOverviews>>) {
  const totals = summarizeWorkbookTotals(data.worksheetSummaries);
  const lines = [
    "Workbook inventory",
    `${"━".repeat(40)}`,
    `Sheet count: ${data.worksheetSummaries.length}`,
    `Active sheet: ${data.activeSheetName}`,
    "",
    "Sheet inventory:",
  ];

  for (const sheet of data.worksheetSummaries) {
    lines.push(...summarizeWorksheet(sheet, data.activeSheetName), "");
  }

  lines.push(
    "Workbook totals:",
    `- cells in used ranges: ${totals.cells.toLocaleString()}`,
    `- charts: ${totals.charts}`,
    `- tables: ${totals.tables}`,
    `- pivotTables: ${totals.pivotTables}`,
  );

  if (data.workbookNames.length) {
    lines.push("", `Workbook names (${data.workbookNames.length}):`);
    for (const name of data.workbookNames) {
      lines.push(`- ${name.name} [${name.scope}]: ${name.value}`);
    }
  }

  return lines.join("\n");
}

export const getWorkbookOverview: Tool = {
  name: "get_workbook_overview",
  description: "Inspect the workbook and return a structured inventory of sheets, named ranges, object counts, filters, and workbook-wide totals.",
  parameters: {
    type: "object",
    properties: {},
  },
  handler: async () => {
    try {
      return await Excel.run(async (context) => {
        const supportsFreeze = isExcelRequirementSetSupported("1.7");
        const supportsAutoFilter = isExcelRequirementSetSupported("1.9");
        const supportsPivotTables = isExcelRequirementSetSupported("1.3");
        const overview = await collectWorksheetOverviews(context, supportsFreeze, supportsAutoFilter, supportsPivotTables);
        return renderWorkbookOverview(overview);
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
