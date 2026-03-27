import type { ToolResultFailure } from "./types";

export type OneNoteSelectionFormat = "text" | "matrix";
export type OneNoteSelectionWriteFormat = "text" | "html" | "image";
export type OneNotePageContentFormat = "summary" | "text" | "json";
export type OneNotePagePlacement = "sectionEnd" | "before" | "after";

export interface OneNoteParagraphSummary {
  type: string;
  text?: string;
  html?: string;
  rowCount?: number;
  columnCount?: number;
  description?: string;
  width?: number;
  height?: number;
}

export interface OneNoteContentSummary {
  id: string;
  type: string;
  left?: number;
  top?: number;
  paragraphs?: OneNoteParagraphSummary[];
  description?: string;
}

export function toolFailure(error: unknown): ToolResultFailure {
  const message = error instanceof Error ? error.message : String(error);
  return { textResultForLlm: message, resultType: "failure", error: message, toolTelemetry: {} };
}

export function isOneNoteRequirementSetSupported(version: string) {
  return Office.context.requirements.isSetSupported("OneNoteApi", version);
}

export function summarizePlainText(text: string, limit = 90) {
  const normalized = String(text || "").replace(/\s+/g, " ").trim();
  if (!normalized) return "(empty)";
  return normalized.length > limit ? `${normalized.slice(0, limit - 3)}...` : normalized;
}

export function ensureNonEmptyHtml(html: string) {
  return String(html || "").trim();
}

export function normalizeImagePayload(content: string) {
  const trimmed = String(content || "").trim();
  const dataUrlMatch = trimmed.match(/^data:image\/[a-zA-Z0-9.+-]+;base64,(.+)$/);
  return dataUrlMatch ? dataUrlMatch[1] : trimmed;
}

export function parseSelectionFormat(value: unknown): OneNoteSelectionFormat {
  return value === "matrix" ? "matrix" : "text";
}

export function parseSelectionWriteFormat(value: unknown): OneNoteSelectionWriteFormat {
  return value === "html" || value === "image" ? value : "text";
}

export function parsePagePlacement(value: unknown): OneNotePagePlacement {
  return value === "before" || value === "after" ? value : "sectionEnd";
}

export function parsePageContentFormat(value: unknown): OneNotePageContentFormat {
  return value === "text" || value === "json" ? value : "summary";
}

export function formatPageText(contentItems: OneNoteContentSummary[]) {
  const blocks: string[] = [];

  for (const item of contentItems) {
    if (item.type === "Outline") {
      const paragraphText = (item.paragraphs || []).map((paragraph) => {
        if (paragraph.type === "RichText") return paragraph.text || "";
        if (paragraph.type === "Table") return `[Table ${paragraph.rowCount || 0}x${paragraph.columnCount || 0}]`;
        if (paragraph.type === "Image") return `[Image${paragraph.description ? `: ${paragraph.description}` : ""}]`;
        return `[${paragraph.type}]`;
      }).filter(Boolean);

      if (paragraphText.length) {
        blocks.push(paragraphText.join("\n"));
      }
      continue;
    }

    if (item.type === "Image") {
      blocks.push(`[Image${item.description ? `: ${item.description}` : ""}]`);
      continue;
    }

    blocks.push(`[${item.type}]`);
  }

  return blocks.join("\n\n").trim() || "(empty page)";
}

export function formatPageSummary(page: { title: string; id: string; pageLevel?: number }, contentItems: OneNoteContentSummary[]) {
  const outlineCount = contentItems.filter((item) => item.type === "Outline").length;
  const imageCount = contentItems.filter((item) => item.type === "Image").length;
  const otherCount = contentItems.filter((item) => item.type !== "Outline" && item.type !== "Image").length;
  const paragraphCount = contentItems.reduce((sum, item) => sum + (item.paragraphs?.length || 0), 0);
  const preview = summarizePlainText(formatPageText(contentItems), 240);

  return [
    `Page ${JSON.stringify(page.title || "Untitled")} (${page.id})`,
    `Level: ${page.pageLevel ?? 0}`,
    `Top-level content: outlines=${outlineCount}, images=${imageCount}, other=${otherCount}`,
    `Paragraph-like items: ${paragraphCount}`,
    `Preview: ${preview}`,
  ].join("\n");
}

export function asJsonString(value: unknown) {
  return JSON.stringify(value, null, 2);
}

export function loadActivePageOrThrow(context: OneNote.RequestContext) {
  return context.application.getActivePage();
}

export function loadActiveSectionOrThrow(context: OneNote.RequestContext) {
  return context.application.getActiveSection();
}

export async function getSelectedDataAsync<T>(coercionType: Office.CoercionType): Promise<T> {
  return await new Promise<T>((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(coercionType, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value as T);
        return;
      }

      reject(new Error(result.error?.message || "Failed to read OneNote selection."));
    });
  });
}

export async function setSelectedDataAsync(data: string, coercionType: Office.CoercionType) {
  await new Promise<void>((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(data, { coercionType }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
        return;
      }

      reject(new Error(result.error?.message || "Failed to update OneNote selection."));
    });
  });
}

export async function loadParagraphSummaries(
  context: OneNote.RequestContext,
  outline: OneNote.Outline,
  _format: OneNotePageContentFormat,
): Promise<OneNoteParagraphSummary[]> {
  const paragraphs = outline.paragraphs;
  paragraphs.load("items/id,type");
  await context.sync();

  const richTextResults: Array<{ summary: OneNoteParagraphSummary; richText: OneNote.RichText; html: OfficeExtension.ClientResult<string> }> = [];
  const imageParagraphs: Array<{ summary: OneNoteParagraphSummary; image: OneNote.Image }> = [];
  const tableParagraphs: Array<{ summary: OneNoteParagraphSummary; table: OneNote.Table }> = [];
  const summaries: OneNoteParagraphSummary[] = [];

  for (const paragraph of paragraphs.items) {
    const type = String(paragraph.type);
    if (type === "RichText") {
      const richText = paragraph.richText;
      richText.load("text");
      const html = richText.getHtml();
      const summary: OneNoteParagraphSummary = { type };
      summaries.push(summary);
      richTextResults.push({ summary, richText, html });
      continue;
    }

    if (type === "Image") {
      const image = paragraph.image;
      image.load(["description", "width", "height"]);
      const summary: OneNoteParagraphSummary = { type };
      summaries.push(summary);
      imageParagraphs.push({ summary, image });
      continue;
    }

    if (type === "Table") {
      const table = paragraph.table;
      table.load(["rowCount", "columnCount"]);
      const summary: OneNoteParagraphSummary = { type };
      summaries.push(summary);
      tableParagraphs.push({ summary, table });
      continue;
    }

    summaries.push({ type });
  }

  if (richTextResults.length || imageParagraphs.length || tableParagraphs.length) {
    await context.sync();
  }

  for (const item of richTextResults) {
    item.summary.text = item.richText.text || (item.html.value ? item.html.value.replace(/<[^>]+>/g, " ") : "");
    item.summary.html = item.html.value || undefined;
  }

  for (const item of imageParagraphs) {
    item.summary.description = item.image.description || "";
    item.summary.width = item.image.width;
    item.summary.height = item.image.height;
  }

  for (const item of tableParagraphs) {
    item.summary.rowCount = item.table.rowCount;
    item.summary.columnCount = item.table.columnCount;
  }

  return summaries;
}

export async function loadPageContentSummaries(
  context: OneNote.RequestContext,
  page: OneNote.Page,
  format: OneNotePageContentFormat,
): Promise<OneNoteContentSummary[]> {
  const contents = page.contents;
  contents.load("items/id,type,left,top");
  await context.sync();

  const outlineItems: Array<{ summary: OneNoteContentSummary; outline: OneNote.Outline }> = [];
  const imageItems: Array<{ summary: OneNoteContentSummary; image: OneNote.Image }> = [];
  const summaries: OneNoteContentSummary[] = [];

  for (const content of contents.items) {
    const type = String(content.type);
    const summary: OneNoteContentSummary = {
      id: content.id,
      type,
      left: content.left,
      top: content.top,
    };
    summaries.push(summary);

    if (type === "Outline") {
      outlineItems.push({ summary, outline: content.outline });
      continue;
    }

    if (type === "Image") {
      const image = content.image;
      image.load(["description", "width", "height"]);
      imageItems.push({ summary, image });
    }
  }

  if (imageItems.length) {
    await context.sync();
  }

  for (const item of imageItems) {
    item.summary.description = item.image.description || "";
  }

  for (const item of outlineItems) {
    item.summary.paragraphs = await loadParagraphSummaries(context, item.outline, format);
  }

  return summaries;
}

export async function findSectionById(
  context: OneNote.RequestContext,
  notebook: OneNote.Notebook,
  sectionId: string,
): Promise<OneNote.Section | null> {
  notebook.sections.load("items/id,name");
  notebook.sectionGroups.load("items/id,name");
  await context.sync();

  for (const section of notebook.sections.items) {
    if (section.id === sectionId) return section;
  }

  for (const group of notebook.sectionGroups.items) {
    const found = await findSectionByIdInGroup(context, group, sectionId);
    if (found) return found;
  }

  return null;
}

async function findSectionByIdInGroup(
  context: OneNote.RequestContext,
  group: OneNote.SectionGroup,
  sectionId: string,
): Promise<OneNote.Section | null> {
  group.sections.load("items/id,name");
  group.sectionGroups.load("items/id,name");
  await context.sync();

  for (const section of group.sections.items) {
    if (section.id === sectionId) return section;
  }

  for (const child of group.sectionGroups.items) {
    const found = await findSectionByIdInGroup(context, child, sectionId);
    if (found) return found;
  }

  return null;
}

export async function findPageById(
  context: OneNote.RequestContext,
  notebook: OneNote.Notebook,
  pageId: string,
): Promise<OneNote.Page | null> {
  notebook.sections.load("items/id,name");
  notebook.sectionGroups.load("items/id,name");
  await context.sync();

  for (const section of notebook.sections.items) {
    const found = await findPageByIdInSection(context, section, pageId);
    if (found) return found;
  }

  for (const group of notebook.sectionGroups.items) {
    const found = await findPageByIdInGroup(context, group, pageId);
    if (found) return found;
  }

  return null;
}

async function findPageByIdInSection(
  context: OneNote.RequestContext,
  section: OneNote.Section,
  pageId: string,
): Promise<OneNote.Page | null> {
  section.pages.load("items/id,title,pageLevel,clientUrl");
  await context.sync();
  return section.pages.items.find((page) => page.id === pageId) || null;
}

async function findPageByIdInGroup(
  context: OneNote.RequestContext,
  group: OneNote.SectionGroup,
  pageId: string,
): Promise<OneNote.Page | null> {
  group.sections.load("items/id,name");
  group.sectionGroups.load("items/id,name");
  await context.sync();

  for (const section of group.sections.items) {
    const found = await findPageByIdInSection(context, section, pageId);
    if (found) return found;
  }

  for (const child of group.sectionGroups.items) {
    const found = await findPageByIdInGroup(context, child, pageId);
    if (found) return found;
  }

  return null;
}

export async function appendSectionOverview(
  context: OneNote.RequestContext,
  section: OneNote.Section,
  lines: string[],
  indent: string,
  activeIds: { sectionId: string; pageId: string },
) {
  section.load(["id", "name"]);
  section.pages.load("items/id,title,pageLevel");
  await context.sync();

  lines.push(`${indent}- Section ${JSON.stringify(section.name)} (${section.id})${section.id === activeIds.sectionId ? " <- active" : ""}`);
  for (const page of section.pages.items) {
    lines.push(`${indent}  - Page ${JSON.stringify(page.title || "Untitled")} (${page.id})${page.id === activeIds.pageId ? " <- active" : ""}, level=${page.pageLevel}, clientUrl=${page.clientUrl || "(none)"}`);
  }
}

export async function appendSectionGroupOverview(
  context: OneNote.RequestContext,
  group: OneNote.SectionGroup,
  lines: string[],
  indent: string,
  activeIds: { sectionId: string; pageId: string },
) {
  group.load(["id", "name"]);
  group.sections.load("items/id,name");
  group.sectionGroups.load("items/id,name");
  await context.sync();

  lines.push(`${indent}- Section group ${JSON.stringify(group.name)} (${group.id})`);
  for (const section of group.sections.items) {
    await appendSectionOverview(context, section, lines, `${indent}  `, activeIds);
  }
  for (const child of group.sectionGroups.items) {
    await appendSectionGroupOverview(context, child, lines, `${indent}  `, activeIds);
  }
}
