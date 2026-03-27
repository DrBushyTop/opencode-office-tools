import type { Tool } from "./types";
import { getWorksheet, isExcelRequirementSetSupported, splitSheetQualifiedRange, toolFailure } from "./excelShared";

type RangeAction = "clear" | "insert" | "delete" | "copy" | "fill" | "sort" | "filter";

function normalizeRangeTarget(range: string, sheetName?: string) {
  const qualified = splitSheetQualifiedRange(range);
  return {
    sheetName: qualified?.sheetName || sheetName,
    rangeAddress: qualified?.rangeAddress || range,
  };
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
    const {
      action,
      sheetName,
      range,
      clearMode = "Contents",
      insertShift = "Down",
      deleteShift = "Up",
      sourceRange,
      copyType = "All",
      skipBlanks = false,
      transpose = false,
      destinationRange,
      fillType = "FillDefault",
      sortKey,
      sortAscending = true,
      hasHeaders,
      matchCase = false,
      sortOrientation,
      sortMethod,
      filterOperation = "apply",
      columnIndex,
      filterOn,
      criterion1,
      criterion2,
      filterOperator,
      filterValues,
      dynamicCriteria,
      filterColor,
    } = args as {
      action: RangeAction;
      sheetName?: string;
      range: string;
      clearMode?: "All" | "Formats" | "Contents" | "Hyperlinks" | "RemoveHyperlinks" | "ResetContents";
      insertShift?: "Down" | "Right";
      deleteShift?: "Up" | "Left";
      sourceRange?: string;
      copyType?: "All" | "Formulas" | "Values" | "Formats" | "Link";
      skipBlanks?: boolean;
      transpose?: boolean;
      destinationRange?: string;
      fillType?: "FillDefault" | "FillCopy" | "FillSeries" | "FillFormats" | "FillValues" | "FillDays" | "FillWeekdays" | "FillMonths" | "FillYears" | "LinearTrend" | "GrowthTrend" | "FlashFill";
      sortKey?: number;
      sortAscending?: boolean;
      hasHeaders?: boolean;
      matchCase?: boolean;
      sortOrientation?: "Rows" | "Columns";
      sortMethod?: "PinYin" | "StrokeCount";
      filterOperation?: "apply" | "clearAll" | "remove" | "reapply";
      columnIndex?: number;
      filterOn?: Excel.FilterOn | "BottomItems" | "BottomPercent" | "CellColor" | "Dynamic" | "FontColor" | "Values" | "TopItems" | "TopPercent" | "Icon" | "Custom";
      criterion1?: string;
      criterion2?: string;
      filterOperator?: "And" | "Or";
      filterValues?: string[];
      dynamicCriteria?: Excel.DynamicFilterCriteria;
      filterColor?: string;
    };

    if ((sortKey !== undefined || columnIndex !== undefined) && (!Number.isInteger(sortKey ?? columnIndex) || (sortKey ?? columnIndex)! < 0)) {
      return toolFailure("sortKey and columnIndex must be non-negative integers when provided.");
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
                worksheet.autoFilter.clearCriteria();
                await context.sync();
                return `Cleared filter criteria for ${targetRange.address}.`;
              case "remove":
                worksheet.autoFilter.remove();
                await context.sync();
                return `Removed filters for ${targetRange.address}.`;
              case "reapply":
                worksheet.autoFilter.reapply();
                await context.sync();
                return `Reapplied filters for ${targetRange.address}.`;
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
                    dynamicCriteria,
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
