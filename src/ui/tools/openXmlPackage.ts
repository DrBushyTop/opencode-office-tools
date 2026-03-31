import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

export type OpenXmlPartPath = string;
type ZipEntries = Record<OpenXmlPartPath, Uint8Array>;

function decodeBase64(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function encodeBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
  }
  return btoa(binary);
}

export class OpenXmlPackage {
  private readonly entries: ZipEntries;

  constructor(base64: string) {
    this.entries = unzipSync(decodeBase64(base64));
  }

  listPaths(): OpenXmlPartPath[] {
    return Object.keys(this.entries).sort();
  }

  has(path: OpenXmlPartPath): boolean {
    return this.entries[path] !== undefined;
  }

  readText(path: OpenXmlPartPath): string {
    const entry = this.entries[path];
    if (!entry) throw new Error(`Open XML part not found: ${path}`);
    return strFromU8(entry);
  }

  readBytes(path: OpenXmlPartPath): Uint8Array {
    const entry = this.entries[path];
    if (!entry) throw new Error(`Open XML part not found: ${path}`);
    return entry.slice();
  }

  writeText(path: OpenXmlPartPath, contents: string): void {
    this.entries[path] = strToU8(contents);
  }

  writeBytes(path: OpenXmlPartPath, contents: Uint8Array): void {
    this.entries[path] = contents.slice();
  }

  delete(path: OpenXmlPartPath): void {
    delete this.entries[path];
  }

  toBase64(): string {
    return encodeBase64(zipSync(this.entries, { level: 1 }));
  }
}

export function parseXml(xml: string): XMLDocument {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  const parseError = doc.getElementsByTagName("parsererror")[0];
  if (parseError) {
    throw new Error(`Invalid XML: ${parseError.textContent || "parser error"}`);
  }
  return doc;
}

export function serializeXml(doc: XMLDocument): string {
  return new XMLSerializer().serializeToString(doc);
}

export function relationshipPartPath(sourcePartPath: OpenXmlPartPath): OpenXmlPartPath {
  const slashIndex = sourcePartPath.lastIndexOf("/");
  const directory = slashIndex === -1 ? "" : sourcePartPath.slice(0, slashIndex);
  const filename = slashIndex === -1 ? sourcePartPath : sourcePartPath.slice(slashIndex + 1);
  return `${directory}/_rels/${filename}.rels`;
}

export function dirname(path: OpenXmlPartPath): string {
  const slashIndex = path.lastIndexOf("/");
  return slashIndex === -1 ? "" : path.slice(0, slashIndex);
}

export function resolveTargetPath(sourcePartPath: OpenXmlPartPath, target: string): OpenXmlPartPath {
  const baseSegments = dirname(sourcePartPath).split("/").filter(Boolean);
  const targetSegments = target.split("/");
  const segments = [...baseSegments];

  for (const segment of targetSegments) {
    if (!segment || segment === ".") continue;
    if (segment === "..") {
      segments.pop();
    } else {
      segments.push(segment);
    }
  }

  return segments.join("/");
}

export function nextRelationshipId(relationshipsDoc: XMLDocument): string {
  const existingIds = Array.from(relationshipsDoc.getElementsByTagName("Relationship"))
    .map((relationship) => relationship.getAttribute("Id") || "")
    .map((id) => /^rId(\d+)$/.exec(id)?.[1])
    .filter(Boolean)
    .map((value) => Number(value));
  const nextId = existingIds.length ? Math.max(...existingIds) + 1 : 1;
  return `rId${nextId}`;
}

export function createRelationshipsDocument(): XMLDocument {
  return parseXml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
}
