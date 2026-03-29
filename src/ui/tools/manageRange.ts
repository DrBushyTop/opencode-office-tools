import { z } from "zod";
import type { Tool } from "./types";
import { getWorksheet, isExcelRequirementSetSupported, nonNegativeIntegerSchema, parseToolArgs, splitSheetQualifiedRange, toolFailure } from "./excelShared";

type RangeAction = "clear" | "insert" | "delete" | "copy" | "fill" | "sort" | "filter";

type FilterOperation = "apply" | "clearAll" | "remove" | "reapply";

type FilterArgs = {
  filterOn?: Excel.FilterOn | "BottomItems" | "BottomPercent" | "CellColor" | "Dynamic" | "FontColor" | "Values" | "TopItems" | "TopPercent" | "Icon" | "Custom";
  criterion1?: string;
  criterion2?: string;
  filterValues?: string[];
  dynamicCriteria?: string;
  filterColor?: string;
};

const manageRangeArgsSchema = z.object({
  action: z.enum(["clear", "insert", "delete", "copy", "fill", "sort", "filter"]),
  sheetName: z.string().optional(),
  range: z.string(),
  clearMode: z.enum(["All", "Formats", "Contents", "Hyperlinks", "RemoveHyperlinks", "ResetContents"]).default("Contents"),
  insertShift: z.enum(["Down", "Right"]).default("Down"),
  deleteShift: z.enum(["Up", "Left"]).default("Up"),
  sourceRange: z.string().optional(),
  copyType: z.enum(["All", "Formulas", "Values", "Formats", "Link"]).default("All"),
  skipBlanks: z.boolean().default(false),
  transpose: z.boolean().default(false),
  destinationRange: z.string().optional(),
  fillType: z.enum(["FillDefault", "FillCopy", "FillSeries", "FillFormats", "FillValues", "FillDays", "FillWeekdays", "FillMonths", "FillYears", "LinearTrend", "GrowthTrend", "FlashFill"]).default("FillDefault"),
  sortKey: nonNegativeIntegerSchema("sortKey and columnIndex must be non-negative integers when provided.").optional(),
  sortAscending: z.boolean().default(true),
  hasHeaders: z.boolean().optional(),
  matchCase: z.boolean().default(false),
  sortOrientation: z.enum(["Rows", "Columns"]).optional(),
  sortMethod: z.enum(["PinYin", "StrokeCount"]).optional(),
  filterOperation: z.enum(["apply", "clearAll", "remove", "reapply"]).default("apply"),
  columnIndex: nonNegativeIntegerSchema("sortKey and columnIndex must be non-negative integers when provided.").optional(),
  filterOn: z.enum(["BottomItems", "BottomPercent", "CellColor", "Dynamic", "FontColor", "Values", "TopItems", "TopPercent", "Icon", "Custom"]).optional(),
  criterion1: z.string().optional(),
  criterion2: z.string().optional(),
  filterOperator: z.enum(["And", "Or"]).optional(),
  filterValues: z.array(z.string()).optional(),
  dynamicCriteria: z.enum(["Unknown", "AboveAverage", "AllDatesInPeriodApril", "AllDatesInPeriodAugust", "AllDatesInPeriodDecember", "AllDatesInPeriodFebruray", "AllDatesInPeriodJanuary", "AllDatesInPeriodJuly", "AllDatesInPeriodJune", "AllDatesInPeriodMarch", "AllDatesInPeriodMay", "AllDatesInPeriodNovember", "AllDatesInPeriodOctober", "AllDatesInPeriodQuarter1", "AllDatesInPeriodQuarter2", "AllDatesInPeriodQuarter3", "AllDatesInPeriodQuarter4", "AllDatesInPeriodSeptember", "BelowAverage", "LastMonth", "LastQuarter", "LastWeek", "LastYear", "NextMonth", "NextQuarter", "NextWeek", "NextYear", "ThisMonth", "ThisQuarter", "ThisWeek", "ThisYear", "Today", "Tomorrow", "YearToDate", "Yesterday"]).optional(),
  filterColor: z.string().optional(),
});

function normalizeRangeTarget(range: string, sheetName?: string) {
  const qualified = splitSheetQualifiedRange(range);
  return {
    sheetName: qualified?.sheetName || sheetName,
    rangeAddress: qualified?.rangeAddress || range,
  };
}

export function hasFilterCriteria({ filterOn, criterion1, criterion2, filterValues, dynamicCriteria, filterColor }: FilterArgs) {
  return Boolean(
    filterOn
    || criterion1
    || criterion2
    || filterValues?.length
    || dynamicCriteria
    || filterColor,
  );
}

export function normalizeRangeAddressForComparison(address: string) {
  return address.replace(/\$/g, "").replace(/\s+/g, "").toUpperCase();
}

function filterRangeMismatchMessage(operation: Exclude<FilterOperation, "apply">, requestedAddress: string, actualAddress: string | null) {
  if (!actualAddress) {
    return `Cannot ${operation} filters for ${requestedAddress} because the worksheet has no active AutoFilter on that range.`;
  }

  return `Cannot ${operation} filters for ${requestedAddress} because the worksheet AutoFilter is currently scoped to ${actualAddress}.`;
}

export const manageRange: Tool = {
  name: "manage_range",
  description: "Perform generic Excel range operations such as clear, insert, delete, copy, fill, sort, and filter.",
  parameters: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: ["clear", "insert", "delete", "copy", "fill", "sort", "filter"],
        description: "Range operation to perform.",
      },
      sheetName: {
        type: "string",
        description: "Worksheet name for range values that are not already sheet-qualified.",
      },
      range: {
        type: "string",
        description: "Target range such as A1:D10 or Sheet1!A1:D10.",
      },
      clearMode: {
        type: "string",
        enum: ["All", "Formats", "Contents", "Hyperlinks", "RemoveHyperlinks", "ResetContents"],
        description: "What to clear for the clear action. Default is Contents.",
      },
      insertShift: {
        type: "string",
        enum: ["Down", "Right"],
        description: "Shift direction for insert.",
      },
      deleteShift: {
        type: "string",
        enum: ["Up", "Left"],
        description: "Shift direction for delete.",
      },
      sourceRange: {
        type: "string",
        description: "Source range for copy.",
      },
      copyType: {
        type: "string",
        enum: ["All", "Formulas", "Values", "Formats", "Link"],
        description: "Content to copy for copy. Default is All.",
      },
      skipBlanks: {
        type: "boolean",
        description: "Skip blank source cells during copy.",
      },
      transpose: {
        type: "boolean",
        description: "Transpose during copy.",
      },
      destinationRange: {
        type: "string",
        description: "Destination range for fill.",
      },
      fillType: {
        type: "string",
        enum: ["FillDefault", "FillCopy", "FillSeries", "FillFormats", "FillValues", "FillDays", "FillWeekdays", "FillMonths", "FillYears", "LinearTrend", "GrowthTrend", "FlashFill"],
        description: "AutoFill mode for fill. Default is FillDefault.",
      },
      sortKey: {
        type: "number",
        description: "Zero-based column or row offset within the range for sort.",
      },
      sortAscending: {
        type: "boolean",
        description: "Whether sort is ascending. Default true.",
      },
      hasHeaders: {
        type: "boolean",
        description: "Whether the sorted range has headers.",
      },
      matchCase: {
        type: "boolean",
        description: "Whether sort should be case-sensitive.",
      },
      sortOrientation: {
        type: "string",
        enum: ["Rows", "Columns"],
        description: "Whether to sort rows or columns.",
      },
      sortMethod: {
        type: "string",
        enum: ["PinYin", "StrokeCount"],
        description: "Chinese character ordering method for sort.",
      },
      filterOperation: {
        type: "string",
        enum: ["apply", "clearAll", "remove", "reapply"],
        description: "Filter lifecycle operation. Default is apply.",
      },
      columnIndex: {
        type: "number",
        description: "Zero-based column offset within the range for filter apply.",
      },
      filterOn: {
        type: "string",
        enum: ["BottomItems", "BottomPercent", "CellColor", "Dynamic", "FontColor", "Values", "TopItems", "TopPercent", "Icon", "Custom"],
        description: "How to filter the column when filterOperation is apply.",
      },
      criterion1: {
        type: "string",
        description: "Primary filter criterion, such as >50 or =North.",
      },
      criterion2: {
        type: "string",
        description: "Secondary filter criterion for custom filters.",
      },
      filterOperator: {
        type: "string",
        enum: ["And", "Or"],
        description: "How to combine custom criteria.",
      },
      filterValues: {
        type: "array",
        items: { type: "string" },
        description: "Explicit values to keep visible when filtering by values.",
      },
      dynamicCriteria: {
        type: "string",
        enum: ["Unknown", "AboveAverage", "AllDatesInPeriodApril", "AllDatesInPeriodAugust", "AllDatesInPeriodDecember", "AllDatesInPeriodFebruray", "AllDatesInPeriodJanuary", "AllDatesInPeriodJuly", "AllDatesInPeriodJune", "AllDatesInPeriodMarch", "AllDatesInPeriodMay", "AllDatesInPeriodNovember", "AllDatesInPeriodOctober", "AllDatesInPeriodQuarter1", "AllDatesInPeriodQuarter2", "AllDatesInPeriodQuarter3", "AllDatesInPeriodQuarter4", "AllDatesInPeriodSeptember", "BelowAverage", "LastMonth", "LastQuarter", "LastWeek", "LastYear", "NextMonth", "NextQuarter", "NextWeek", "NextYear", "ThisMonth", "ThisQuarter", "ThisWeek", "ThisYear", "Today", "Tomorrow", "YearToDate", "Yesterday"],
        description: "Dynamic filter criteria when filterOn is Dynamic.",
      },
      filterColor: {
        type: "string",
        description: "HTML color string for cellColor or fontColor filters.",
      },
    },
    required: ["action", "range"],
  },
  handler: async (args) => {
    const parsedArgs = parseToolArgs(manageRangeArgsSchema, args);
    if (!parsedArgs.success) return parsedArgs.failure;

    const {
      action,
      sheetName,
      range,
      clearMode,
      insertShift,
      deleteShift,
      sourceRange,
      copyType,
      skipBlanks,
      transpose,
      destinationRange,
      fillType,
      sortKey,
      sortAscending,
      hasHeaders,
      matchCase,
      sortOrientation,
      sortMethod,
      filterOperation,
      columnIndex,
      filterOn,
      criterion1,
      criterion2,
      filterOperator,
      filterValues,
      dynamicCriteria,
      filterColor,
    } = parsedArgs.data;

    if (action === "filter" && filterOperation === "apply" && hasFilterCriteria({
      filterOn,
      criterion1,
      criterion2,
      filterValues,
      dynamicCriteria,
      filterColor,
    }) && columnIndex === undefined) {
      return toolFailure("columnIndex is required when applying filter criteria.");
    }

    try {
      return await Excel.run(async (context) => {
        const target = normalizeRangeTarget(range, sheetName);
        const worksheet = await getWorksheet(context, target.sheetName);
        worksheet.load("name");
        const targetRange = worksheet.getRange(target.rangeAddress);
        targetRange.load("address");
        await context.sync();

        switch (action) {
          case "clear":
            targetRange.clear(clearMode);
            await context.sync();
            return `Cleared ${clearMode.toLowerCase()} in ${targetRange.address}.`;
          case "insert": {
            const insertedRange = targetRange.insert(insertShift);
            insertedRange.load("address");
            await context.sync();
            return `Inserted cells at ${insertedRange.address}, shifting ${insertShift.toLowerCase()}.`;
          }
          case "delete":
            targetRange.delete(deleteShift);
            await context.sync();
            return `Deleted cells at ${targetRange.address}, shifting ${deleteShift.toLowerCase()}.`;
          case "copy": {
            if (!isExcelRequirementSetSupported("1.9")) {
              return toolFailure("Copying ranges requires ExcelApi 1.9.");
            }
            if (!sourceRange) return toolFailure("sourceRange is required for copy.");
            const sourceTarget = normalizeRangeTarget(sourceRange, sheetName);
            const sourceSheet = await getWorksheet(context, sourceTarget.sheetName);
            const source = sourceSheet.getRange(sourceTarget.rangeAddress);
            source.load("address");
            targetRange.copyFrom(source, copyType, skipBlanks, transpose);
            await context.sync();
            return `Copied ${copyType.toLowerCase()} from ${source.address} to ${targetRange.address}.`;
          }
          case "fill":
            if (!isExcelRequirementSetSupported("1.9")) {
              return toolFailure("AutoFill requires ExcelApi 1.9.");
            }
            if (!destinationRange) return toolFailure("destinationRange is required for fill.");
            {
              const fillTarget = normalizeRangeTarget(destinationRange, sheetName);
              if (fillTarget.sheetName && fillTarget.sheetName !== worksheet.name) {
                return toolFailure("destinationRange for fill must be on the same worksheet as range.");
              }
              targetRange.autoFill(fillTarget.rangeAddress, fillType);
              await context.sync();
              return `Filled from ${targetRange.address} into ${fillTarget.rangeAddress} using ${fillType}.`;
            }
          case "sort":
            if (sortKey === undefined) return toolFailure("sortKey is required for sort.");
            targetRange.sort.apply([{ key: sortKey, ascending: sortAscending }], matchCase, hasHeaders, sortOrientation, sortMethod);
            await context.sync();
            return `Sorted ${targetRange.address} by offset ${sortKey} (${sortAscending ? "ascending" : "descending"}).`;
          case "filter":
            if (!isExcelRequirementSetSupported("1.9")) {
              return toolFailure("Filtering ranges requires ExcelApi 1.9.");
            }

            switch (filterOperation) {
              case "clearAll":
              case "remove":
              case "reapply": {
                const activeFilterRange = worksheet.autoFilter.getRangeOrNullObject();
                activeFilterRange.load(["isNullObject", "address"]);
                await context.sync();

                const actualFilterAddress = (activeFilterRange as Excel.Range & { isNullObject?: boolean }).isNullObject
                  ? null
                  : activeFilterRange.address;

                if (!actualFilterAddress || normalizeRangeAddressForComparison(actualFilterAddress) !== normalizeRangeAddressForComparison(targetRange.address)) {
                  return toolFailure(filterRangeMismatchMessage(filterOperation, targetRange.address, actualFilterAddress));
                }

                if (filterOperation === "clearAll") {
                  worksheet.autoFilter.clearCriteria();
                  await context.sync();
                  return `Cleared filter criteria for ${actualFilterAddress}.`;
                }

                if (filterOperation === "remove") {
                  worksheet.autoFilter.remove();
                  await context.sync();
                  return `Removed filters for ${actualFilterAddress}.`;
                }

                worksheet.autoFilter.reapply();
                await context.sync();
                return `Reapplied filters for ${actualFilterAddress}.`;
              }
              case "apply": {
                const effectiveFilterOn = filterOn
                  || (filterValues?.length ? "Values" : undefined)
                  || (dynamicCriteria ? "Dynamic" : undefined)
                  || (criterion1 || criterion2 ? "Custom" : undefined);
                const criteria = effectiveFilterOn
                  ? {
                    filterOn: effectiveFilterOn,
                    criterion1,
                    criterion2,
                    operator: filterOperator,
                    values: filterValues,
                    dynamicCriteria: dynamicCriteria as Excel.DynamicFilterCriteria | undefined,
                    color: filterColor,
                  }
                  : undefined;

                worksheet.autoFilter.apply(targetRange, columnIndex, criteria);
                await context.sync();
                return criteria
                  ? `Applied ${effectiveFilterOn} filter to ${targetRange.address}${columnIndex !== undefined ? ` at column offset ${columnIndex}` : ""}.`
                  : `Enabled filters for ${targetRange.address}.`;
              }
              default:
                return toolFailure(`Unsupported filterOperation ${filterOperation}.`);
            }
          default:
            return toolFailure(`Unsupported action ${action}.`);
        }
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
