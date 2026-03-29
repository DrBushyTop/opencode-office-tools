import type { ToolResultFailure } from "./types";
import { z } from "zod";
import { createToolFailure, summarizePlainText as summarizeSharedPlainText } from "./toolShared";

export const HeaderFooterTargetSchema = z.enum(["header", "footer"]);
export const HeaderFooterTypeNameSchema = z.enum(["primary", "firstPage", "evenPages"]);
export const SectionSelectorSchema = z.union([z.literal("*"), z.number().int().positive()]);
export const DocumentContentFormatSchema = z.enum(["summary", "text", "html", "ooxml"]);
export const DocumentWriteFormatSchema = z.enum(["html", "text", "ooxml"]);
export const DocumentWriteOperationSchema = z.enum(["replace", "insert", "clear"]);
export const DocumentWriteLocationSchema = z.enum(["replace", "before", "after", "start", "end"]);

export type HeaderFooterTarget = z.infer<typeof HeaderFooterTargetSchema>;
export type HeaderFooterTypeName = z.infer<typeof HeaderFooterTypeNameSchema>;
export type SectionSelector = z.infer<typeof SectionSelectorSchema>;
export type DocumentContentFormat = z.infer<typeof DocumentContentFormatSchema>;
export type DocumentWriteFormat = z.infer<typeof DocumentWriteFormatSchema>;
export type DocumentWriteOperation = z.infer<typeof DocumentWriteOperationSchema>;
export type DocumentWriteLocation = z.infer<typeof DocumentWriteLocationSchema>;
type InlineInsertLocation = Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end;

export const DocumentPartAddressSchema = z.union([
  z.object({ kind: z.literal("headersFootersOverview") }),
  z.object({ kind: z.literal("tableOfContents") }),
  z.object({
    kind: z.literal("section"),
    section: SectionSelectorSchema,
    target: HeaderFooterTargetSchema.optional(),
    type: HeaderFooterTypeNameSchema.optional(),
  }),
]);

export const DocumentRangeAddressSchema = z.discriminatedUnion("kind", [
  z.object({ kind: z.literal("body") }),
  z.object({ kind: z.literal("selection") }),
  z.object({ kind: z.literal("bookmark"), name: z.string().min(1) }),
  z.object({
    kind: z.literal("contentControl"),
    by: z.enum(["id", "index"]),
    value: z.number().int().positive(),
  }),
  z.object({
    kind: z.literal("table"),
    tableIndex: z.number().int().positive(),
    rowIndex: z.number().int().positive().optional(),
    cellIndex: z.number().int().positive().optional(),
  }),
]).refine(
  (address) => address.kind !== "table"
    || ((address.rowIndex === undefined && address.cellIndex === undefined)
      || (address.rowIndex !== undefined && address.cellIndex !== undefined)),
  { message: "Table cell addresses must include both rowIndex and cellIndex." },
);

export type DocumentPartAddress = z.infer<typeof DocumentPartAddressSchema>;
export type DocumentRangeAddress = z.infer<typeof DocumentRangeAddressSchema>;

export type ResolvedWordTarget =
  | { kind: "body"; label: string; target: Word.Body }
  | { kind: "range"; label: string; target: Word.Range; clearBehavior?: "delete" | "reject" }
  | { kind: "contentControl"; label: string; target: Word.ContentControl };

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
  return createToolFailure(error);
}

export function getZodErrorMessage(error: z.ZodError) {
  const issue = error.issues[0];
  if (!issue) return "Invalid arguments.";

  const path = issue.path.length > 0 ? `${issue.path.join(".")}: ` : "";
  return `${path}${issue.message}`;
}

export function summarizePlainText(text: string, limit = 80) {
  return summarizeSharedPlainText(text, limit);
}

export function resolveInsertLocation(location: DocumentWriteLocation): Word.InsertLocation {
  switch (location) {
    case "before":
      return Word.InsertLocation.before;
    case "after":
      return Word.InsertLocation.after;
    case "start":
      return Word.InsertLocation.start;
    case "end":
      return Word.InsertLocation.end;
    case "replace":
    default:
      return Word.InsertLocation.replace;
  }
}

function parsePositiveInteger(value: string) {
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed > 0 ? parsed : null;
}

export function parseDocumentRangeAddress(address: string): DocumentRangeAddress | null {
  const normalized = String(address || "").trim();
  if (!normalized) return null;

  if (/^(document|body)$/i.test(normalized)) {
    return DocumentRangeAddressSchema.parse({ kind: "body" });
  }

  if (/^selection$/i.test(normalized)) {
    return DocumentRangeAddressSchema.parse({ kind: "selection" });
  }

  const bookmarkMatch = normalized.match(/^bookmark\[(.+)\]$/i);
  if (bookmarkMatch) {
    const name = bookmarkMatch[1].trim();
    return name ? DocumentRangeAddressSchema.parse({ kind: "bookmark", name }) : null;
  }

  const contentControlMatch = normalized.match(/^content_control\[(id|index)=(\d+)\]$/i);
  if (contentControlMatch) {
    const value = parsePositiveInteger(contentControlMatch[2]);
    if (!value) return null;
    return DocumentRangeAddressSchema.parse({
      kind: "contentControl",
      by: contentControlMatch[1].toLowerCase(),
      value,
    });
  }

  const tableMatch = normalized.match(/^table\[(\d+)\](?:\.cell\[(\d+),(\d+)\])?$/i);
  if (tableMatch) {
    const tableIndex = parsePositiveInteger(tableMatch[1]);
    const rowIndex = tableMatch[2] ? parsePositiveInteger(tableMatch[2]) : null;
    const cellIndex = tableMatch[3] ? parsePositiveInteger(tableMatch[3]) : null;
    if (!tableIndex) return null;
    if ((rowIndex && !cellIndex) || (!rowIndex && cellIndex)) return null;

    return DocumentRangeAddressSchema.parse({
      kind: "table",
      tableIndex,
      ...(rowIndex && cellIndex ? { rowIndex, cellIndex } : {}),
    });
  }

  return null;
}

async function resolveContentControlTarget(
  context: Word.RequestContext,
  address: Extract<DocumentRangeAddress, { kind: "contentControl" }>,
): Promise<ResolvedWordTarget> {
  if (address.by === "id") {
    const contentControl = context.document.contentControls.getByIdOrNullObject(address.value);
    contentControl.load("isNullObject");
    await context.sync();
    if (contentControl.isNullObject) {
      throw new Error(`Content control ${address.value} does not exist.`);
    }

    return {
      kind: "contentControl",
      label: `content_control[id=${address.value}]`,
      target: contentControl,
    };
  }

  const collection = context.document.contentControls;
  collection.load("items");
  await context.sync();

  const contentControl = collection.items[address.value - 1];
  if (!contentControl) {
    throw new Error(`Content control index ${address.value} does not exist.`);
  }

  contentControl.load("id");
  await context.sync();

  return {
    kind: "contentControl",
    label: `content_control[id=${contentControl.id}]`,
    target: contentControl,
  };
}

export async function resolveDocumentRangeTarget(
  context: Word.RequestContext,
  address: DocumentRangeAddress,
): Promise<ResolvedWordTarget> {
  switch (address.kind) {
    case "body":
      return { kind: "body", label: "document", target: context.document.body };
    case "selection":
      return { kind: "range", label: "selection", target: context.document.getSelection() };
    case "bookmark": {
      if (!isWordRequirementSetSupported("1.4")) {
        throw new Error("Bookmark targets require WordApi 1.4.");
      }

      const range = context.document.getBookmarkRangeOrNullObject(address.name);
      range.load("isNullObject");
      await context.sync();

      if (range.isNullObject) {
        throw new Error(`Bookmark '${address.name}' does not exist.`);
      }

      return { kind: "range", label: `bookmark[${address.name}]`, target: range };
    }
    case "contentControl":
      return resolveContentControlTarget(context, address);
    case "table": {
      if (!isWordRequirementSetSupported("1.3")) {
        throw new Error("Table targets require WordApi 1.3.");
      }

      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      const table = tables.items[address.tableIndex - 1];
      if (!table) {
        throw new Error(`Table ${address.tableIndex} does not exist.`);
      }

      if (address.rowIndex && address.cellIndex) {
        const cell = table.getCellOrNullObject(address.rowIndex - 1, address.cellIndex - 1);
        cell.load("isNullObject");
        await context.sync();

        if (cell.isNullObject) {
          throw new Error(`Cell ${address.rowIndex},${address.cellIndex} does not exist in table ${address.tableIndex}.`);
        }

        return {
          kind: "body",
          label: `table[${address.tableIndex}].cell[${address.rowIndex},${address.cellIndex}]`,
          target: cell.body,
        };
      }

      return {
        kind: "range",
        label: `table[${address.tableIndex}]`,
        target: table.getRange(Word.RangeLocation.whole),
        clearBehavior: "reject",
      };
    }
    default:
      throw new Error("Unsupported target.");
  }
}

export async function readResolvedWordTarget(
  context: Word.RequestContext,
  resolved: ResolvedWordTarget,
  format: DocumentContentFormat,
) {
  if (format === "html") {
    const html = resolved.target.getHtml();
    await context.sync();
    return html.value || "(empty)";
  }

  if (format === "ooxml") {
    const ooxml = resolved.target.getOoxml();
    await context.sync();
    return ooxml.value || "(empty)";
  }

  resolved.target.load("text");
  await context.sync();

  if (format === "summary") {
    return summarizePlainText(resolved.target.text || "");
  }

  return resolved.target.text || "(empty)";
}

function insertIntoResolvedTarget(
  resolved: ResolvedWordTarget,
  format: DocumentWriteFormat,
  content: string,
  location: DocumentWriteLocation,
) {
  if ((resolved.kind === "body" || resolved.kind === "contentControl") && (location === "before" || location === "after")) {
    throw new Error(`The ${resolved.label} target supports replace, start, or end insertion only.`);
  }

  if (resolved.kind === "range") {
    const insertLocation = resolveInsertLocation(location);
    switch (format) {
      case "text":
        resolved.target.insertText(content, insertLocation);
        return;
      case "ooxml":
        resolved.target.insertOoxml(content, insertLocation);
        return;
      case "html":
      default:
        resolved.target.insertHtml(content, insertLocation);
        return;
    }
  }

  const insertLocation: InlineInsertLocation = location === "start"
    ? Word.InsertLocation.start
    : location === "end"
      ? Word.InsertLocation.end
      : Word.InsertLocation.replace;

  switch (format) {
    case "text":
      resolved.target.insertText(content, insertLocation);
      return;
    case "ooxml":
      resolved.target.insertOoxml(content, insertLocation);
      return;
    case "html":
    default:
      resolved.target.insertHtml(content, insertLocation);
      return;
  }
}

export function writeResolvedWordTarget(
  resolved: ResolvedWordTarget,
  operation: DocumentWriteOperation,
  format: DocumentWriteFormat,
  content: string | undefined,
  location: DocumentWriteLocation,
) {
  if (operation === "clear") {
    if (resolved.kind === "range") {
      if (resolved.clearBehavior === "reject") {
        throw new Error(`Clearing ${resolved.label} is not supported because it would remove the entire table. Target a cell such as ${resolved.label}.cell[1,1] or use replace instead.`);
      }
      resolved.target.delete();
    } else {
      resolved.target.clear();
    }
    return;
  }

  if (content === undefined) {
    throw new Error("content is required for replace or insert operations.");
  }

  insertIntoResolvedTarget(resolved, format, content, operation === "replace" ? "replace" : location);
}

export function parseDocumentPartAddress(address: string): DocumentPartAddress | null {
  const normalized = String(address || "").trim();
  if (!normalized) return null;
  if (normalized === "headers_footers") {
    return DocumentPartAddressSchema.parse({ kind: "headersFootersOverview" });
  }
  if (normalized === "table_of_contents") {
    return DocumentPartAddressSchema.parse({ kind: "tableOfContents" });
  }

  const match = normalized.match(/^section\[(\*|\d+)\](?:\.(header|footer)\.(primary|firstPage|evenPages))?$/);
  if (!match) return null;

  return DocumentPartAddressSchema.parse({
    kind: "section",
    section: match[1] === "*" ? "*" : Number(match[1]),
    target: match[2],
    type: match[3],
  });
}

export function isValidSectionSelector(section: SectionSelector) {
  return SectionSelectorSchema.safeParse(section).success;
}
