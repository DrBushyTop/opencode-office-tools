import type { ToolResultFailure } from "./types";

export function toolFailure(error: unknown): ToolResultFailure {
  const message = error instanceof Error ? error.message : String(error);
  return { textResultForLlm: message, resultType: "failure", error: message, toolTelemetry: {} };
}

export function isExcelRequirementSetSupported(version: string) {
  return Office.context.requirements.isSetSupported("ExcelApi", version);
}

export function normalizeExcelColor(value: string | null | undefined) {
  if (!value) return "(none)";
  if (value.startsWith("#")) return value;
  return /^[0-9A-Fa-f]{6}$/.test(value) ? `#${value}` : value;
}

export function cellToString(value: unknown) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  return JSON.stringify(value);
}

export function splitSheetQualifiedRange(input: string) {
  const bangIndex = input.lastIndexOf("!");
  if (bangIndex === -1) return null;

  const rawSheet = input.slice(0, bangIndex);
  const rangeAddress = input.slice(bangIndex + 1);
  const normalizedSheet = rawSheet.startsWith("'") && rawSheet.endsWith("'")
    ? rawSheet.slice(1, -1).replace(/''/g, "'")
    : rawSheet;

  return { sheetName: normalizedSheet, rangeAddress };
}

const a1ReferencePattern = /^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/;

function looksLikeA1Reference(reference: string) {
  return a1ReferencePattern.test(reference.trim());
}

export async function qualifyNamedRangeReference(context: Excel.RequestContext, reference: string) {
  const trimmedReference = reference.trim();
  if (trimmedReference.startsWith("=")) return trimmedReference;
  if (trimmedReference.includes("!")) return `=${trimmedReference}`;
  if (!looksLikeA1Reference(trimmedReference)) return `=${trimmedReference}`;

  const activeSheet = context.workbook.worksheets.getActiveWorksheet();
  activeSheet.load("name");
  await context.sync();
  const escapedSheetName = activeSheet.name.replace(/'/g, "''");
  return `='${escapedSheetName}'!${trimmedReference}`;
}

export function rangeGridToLines(
  values: unknown[][],
  formulas: unknown[][],
  texts?: string[][],
  numberFormats?: string[][],
) {
  const rows: string[] = [];
  for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
    const rowData: string[] = [];
    for (let columnIndex = 0; columnIndex < values[rowIndex].length; columnIndex += 1) {
      const value = values[rowIndex][columnIndex];
      const formula = formulas[rowIndex]?.[columnIndex];
      const text = texts?.[rowIndex]?.[columnIndex];
      const numberFormat = numberFormats?.[rowIndex]?.[columnIndex];
      const valueText = cellToString(value);
      const formulaText = cellToString(formula);
      const displayText = text !== undefined ? ` text=${JSON.stringify(text)}` : "";
      const formatText = numberFormat ? ` format=${JSON.stringify(numberFormat)}` : "";

      if (formulaText && formulaText !== valueText) {
        rowData.push(`${formulaText} (= ${valueText})${displayText}${formatText}`.trim());
      } else {
        rowData.push(`${valueText}${displayText}${formatText}`.trim());
      }
    }
    rows.push(rowData.join(" | "));
  }
  return rows;
}

export async function getWorksheet(context: Excel.RequestContext, sheetName?: string) {
  if (sheetName) {
    const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
    sheet.load(["isNullObject", "name"]);
    await context.sync();
    if ((sheet as Excel.Worksheet & { isNullObject?: boolean }).isNullObject) {
      throw new Error(`Worksheet ${sheetName} was not found.`);
    }
    return sheet as Excel.Worksheet;
  }

  const activeSheet = context.workbook.worksheets.getActiveWorksheet();
  activeSheet.load("name");
  await context.sync();
  return activeSheet;
}

export async function describeRange(
  context: Excel.RequestContext,
  range: Excel.Range,
  sheetName: string,
  options: {
    detail?: boolean;
    includeNumberFormats?: boolean;
    includeTables?: boolean;
    includeValidation?: boolean;
    includeMergedAreas?: boolean;
  } = {},
) {
  const {
    detail = false,
    includeNumberFormats = detail,
    includeTables = detail,
    includeValidation = detail,
    includeMergedAreas = detail,
  } = options;

  const supportsValidation = includeValidation && isExcelRequirementSetSupported("1.8");
  const supportsTables = includeTables && isExcelRequirementSetSupported("1.9");
  const supportsPivotTables = includeTables && isExcelRequirementSetSupported("1.12");
  const supportsMergedAreas = includeMergedAreas && isExcelRequirementSetSupported("1.13");
  const supportsConditionalFormats = isExcelRequirementSetSupported("1.6");

  range.load(["address", "rowCount", "columnCount", "values", "formulas", "text"]);
  if (includeNumberFormats) {
    range.load("numberFormat");
  }

  const conditionalFormatCount = supportsConditionalFormats ? range.conditionalFormats.getCount() : null;

  const dataValidation = supportsValidation ? range.dataValidation : null;
  if (dataValidation) {
    dataValidation.load(["type", "valid", "ignoreBlanks", "rule"]);
  }

  const tables = supportsTables ? range.getTables(false) : null;
  if (tables) {
    tables.load("items/name,items/style,items/showHeaders,items/showTotals");
  }

  const pivotTables = supportsPivotTables ? range.getPivotTables(false) : null;
  if (pivotTables) {
    pivotTables.load("items/name");
  }

  const mergedAreas = supportsMergedAreas ? range.getMergedAreasOrNullObject() : null;
  if (mergedAreas) {
    mergedAreas.load(["isNullObject", "address"]);
  }

  await context.sync();

  const lines = [
    `Worksheet: ${sheetName}`,
    `Range: ${range.address}`,
    `Dimensions: ${range.rowCount} rows x ${range.columnCount} columns`,
  ];

  if (tables) {
    const tableSummary = tables.items.length
      ? tables.items.map((table) => `${table.name} (${table.style}, headers=${table.showHeaders ? "on" : "off"}, totals=${table.showTotals ? "on" : "off"})`).join(", ")
      : "(none)";
    lines.push(`Tables: ${tableSummary}`);
  }

  if (pivotTables) {
    lines.push(`PivotTables: ${pivotTables.items.length ? pivotTables.items.map((pivot) => pivot.name).join(", ") : "(none)"}`);
  }

  lines.push(`Conditional formats: ${conditionalFormatCount ? conditionalFormatCount.value : "unavailable on this host (requires ExcelApi 1.6)"}`);

  if (dataValidation) {
    lines.push(`Data validation: type=${dataValidation.type}, valid=${String(dataValidation.valid)}, ignoreBlanks=${dataValidation.ignoreBlanks}`);
    if (dataValidation.type !== Excel.DataValidationType.none && dataValidation.rule) {
      lines.push(`Validation rule: ${JSON.stringify(dataValidation.rule)}`);
    }
  } else if (includeValidation) {
    lines.push("Data validation: unavailable on this host (requires ExcelApi 1.8)");
  }

  if (mergedAreas) {
    lines.push(`Merged areas: ${mergedAreas.isNullObject ? "(none)" : mergedAreas.address}`);
  } else if (includeMergedAreas) {
    lines.push("Merged areas: unavailable on this host (requires ExcelApi 1.13)");
  }

  lines.push("");
  lines.push(...rangeGridToLines(range.values, range.formulas, detail ? range.text : undefined, includeNumberFormats ? range.numberFormat : undefined));

  return lines.join("\n");
}
