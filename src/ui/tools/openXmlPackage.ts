import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

type ZipEntries = Record<string, Uint8Array>;

function decodeBase64(base64: string) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function encodeBase64(bytes: Uint8Array) {
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

  listPaths() {
    return Object.keys(this.entries).sort();
  }

  has(path: string) {
    return this.entries[path] !== undefined;
  }

  readText(path: string) {
    const entry = this.entries[path];
    if (!entry) throw new Error(`Open XML part not found: ${path}`);
    return strFromU8(entry);
  }

  writeText(path: string, contents: string) {
    this.entries[path] = strToU8(contents);
  }

  toBase64() {
    return encodeBase64(zipSync(this.entries, { level: 6 }));
  }
}

export function parseXml(xml: string) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  const parseError = doc.getElementsByTagName("parsererror")[0];
  if (parseError) {
    throw new Error(`Invalid XML: ${parseError.textContent || "parser error"}`);
  }
  return doc;
}

export function serializeXml(doc: XMLDocument) {
  return new XMLSerializer().serializeToString(doc);
}

export function relationshipPartPath(sourcePartPath: string) {
  const slashIndex = sourcePartPath.lastIndexOf("/");
  const directory = slashIndex === -1 ? "" : sourcePartPath.slice(0, slashIndex);
  const filename = slashIndex === -1 ? sourcePartPath : sourcePartPath.slice(slashIndex + 1);
  return `${directory}/_rels/${filename}.rels`;
}

export function dirname(path: string) {
  const slashIndex = path.lastIndexOf("/");
  return slashIndex === -1 ? "" : path.slice(0, slashIndex);
}

export function resolveTargetPath(sourcePartPath: string, target: string) {
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

export function nextRelationshipId(relationshipsDoc: XMLDocument) {
  const existingIds = Array.from(relationshipsDoc.getElementsByTagName("Relationship"))
    .map((relationship) => relationship.getAttribute("Id") || "")
    .map((id) => /^rId(\d+)$/.exec(id)?.[1])
    .filter(Boolean)
    .map((value) => Number(value));
  const nextId = existingIds.length ? Math.max(...existingIds) + 1 : 1;
  return `rId${nextId}`;
}

export function createRelationshipsDocument() {
  return parseXml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
}
