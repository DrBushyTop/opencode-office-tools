import {
  OpenXmlPackage,
  createRelationshipsDocument,
  nextRelationshipId,
  parseXml,
  relationshipPartPath,
  resolveTargetPath,
  serializeXml,
} from "./openXmlPackage";
import { isPowerPointRequirementSetSupported } from "./powerpointShared";

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const RELATIONSHIP_TYPE_NOTES_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const RELATIONSHIP_TYPE_NOTES_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
const CONTENT_TYPE_NOTES_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";

const EXCLUDED_NOTE_PLACEHOLDER_TYPES = new Set(["sldImg", "hdr", "dt", "ftr", "sldNum"]);

function getFirstSlidePath(pkg: OpenXmlPackage) {
  const slides = pkg.listPaths().filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path));
  if (!slides.length) {
    throw new Error("The exported PowerPoint package does not contain a slide XML part.");
  }
  return slides.sort((left, right) => left.localeCompare(right, undefined, { numeric: true }))[0];
}

function getNotesMasterPath(pkg: OpenXmlPackage) {
  return pkg.listPaths().find((path) => /^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(path)) || null;
}

function getRelationshipTarget(relationshipsDoc: XMLDocument, type: string) {
  const relationships = Array.from(relationshipsDoc.getElementsByTagName("Relationship"));
  return relationships.find((relationship) => relationship.getAttribute("Type") === type) || null;
}

function getOrCreateRelationshipsDoc(pkg: OpenXmlPackage, partPath: string) {
  const relsPath = relationshipPartPath(partPath);
  const doc = pkg.has(relsPath) ? parseXml(pkg.readText(relsPath)) : createRelationshipsDocument();
  return { relsPath, doc };
}

function ensureContentTypeOverride(pkg: OpenXmlPackage, partPath: string, contentType: string) {
  const contentTypesDoc = parseXml(pkg.readText("[Content_Types].xml"));
  const overrides = Array.from(contentTypesDoc.getElementsByTagName("Override"));
  const partName = `/${partPath}`;
  if (!overrides.some((override) => override.getAttribute("PartName") === partName)) {
    const override = contentTypesDoc.createElementNS(contentTypesDoc.documentElement.namespaceURI, "Override");
    override.setAttribute("PartName", partName);
    override.setAttribute("ContentType", contentType);
    contentTypesDoc.documentElement.appendChild(override);
    pkg.writeText("[Content_Types].xml", serializeXml(contentTypesDoc));
  }
}

function getPlaceholderType(shape: Element) {
  return shape.getElementsByTagNameNS(NS_P, "ph")[0]?.getAttribute("type") || null;
}

function getTextBody(shape: Element) {
  return shape.getElementsByTagNameNS(NS_P, "txBody")[0] || null;
}

function extractTextBody(textBody: Element | null) {
  if (!textBody) return "";
  const paragraphs = Array.from(textBody.getElementsByTagNameNS(NS_A, "p"));
  return paragraphs
    .map((paragraph) => Array.from(paragraph.getElementsByTagNameNS(NS_A, "t")).map((node) => node.textContent || "").join(""))
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function buildParagraph(doc: XMLDocument, text: string) {
  const paragraph = doc.createElementNS(NS_A, "a:p");
  if (text) {
    const run = doc.createElementNS(NS_A, "a:r");
    const runProps = doc.createElementNS(NS_A, "a:rPr");
    runProps.setAttribute("lang", "en-US");
    run.appendChild(runProps);
    const textNode = doc.createElementNS(NS_A, "a:t");
    textNode.textContent = text;
    run.appendChild(textNode);
    paragraph.appendChild(run);
  }
  paragraph.appendChild(doc.createElementNS(NS_A, "a:endParaRPr"));
  return paragraph;
}

function writeTextBody(textBody: Element, text: string) {
  while (textBody.firstChild) {
    textBody.removeChild(textBody.firstChild);
  }

  textBody.appendChild(textBody.ownerDocument.createElementNS(NS_A, "a:bodyPr"));
  textBody.appendChild(textBody.ownerDocument.createElementNS(NS_A, "a:lstStyle"));

  const lines = text.replace(/\r\n/g, "\n").split("\n");
  const effectiveLines = lines.length ? lines : [""];
  for (const line of effectiveLines) {
    textBody.appendChild(buildParagraph(textBody.ownerDocument, line));
  }
}

function getSpeakerNotesShape(notesDoc: XMLDocument) {
  const shapes = Array.from(notesDoc.getElementsByTagNameNS(NS_P, "sp"));
  return shapes.find((shape) => {
    const placeholderType = getPlaceholderType(shape);
    return placeholderType === "body" || (!placeholderType && getTextBody(shape));
  }) || null;
}

function ensureSpeakerNotesShape(notesDoc: XMLDocument) {
  const existing = getSpeakerNotesShape(notesDoc);
  if (existing) return existing;

  const spTree = notesDoc.getElementsByTagNameNS(NS_P, "spTree")[0];
  if (!spTree) {
    throw new Error("The notes slide XML is missing its shape tree.");
  }

  const shapeIds = Array.from(notesDoc.getElementsByTagNameNS(NS_P, "cNvPr"))
    .map((node) => Number(node.getAttribute("id") || 0));
  const nextShapeId = shapeIds.length ? Math.max(...shapeIds) + 1 : 2;

  const shape = notesDoc.createElementNS(NS_P, "p:sp");
  const nvSpPr = notesDoc.createElementNS(NS_P, "p:nvSpPr");
  const cNvPr = notesDoc.createElementNS(NS_P, "p:cNvPr");
  cNvPr.setAttribute("id", String(nextShapeId));
  cNvPr.setAttribute("name", `Speaker Notes ${nextShapeId}`);
  const cNvSpPr = notesDoc.createElementNS(NS_P, "p:cNvSpPr");
  const spLocks = notesDoc.createElementNS(NS_A, "a:spLocks");
  spLocks.setAttribute("noGrp", "1");
  cNvSpPr.appendChild(spLocks);
  const nvPr = notesDoc.createElementNS(NS_P, "p:nvPr");
  const ph = notesDoc.createElementNS(NS_P, "p:ph");
  ph.setAttribute("type", "body");
  ph.setAttribute("idx", "1");
  nvPr.appendChild(ph);
  nvSpPr.append(cNvPr, cNvSpPr, nvPr);

  const spPr = notesDoc.createElementNS(NS_P, "p:spPr");
  const txBody = notesDoc.createElementNS(NS_P, "p:txBody");
  writeTextBody(txBody, "");

  shape.append(nvSpPr, spPr, txBody);
  spTree.appendChild(shape);
  return shape;
}

function buildNotesSlideXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="${NS_A}" xmlns:p="${NS_P}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Speaker Notes"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:endParaRPr/></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:notes>`;
}

function ensureNotesSlide(pkg: OpenXmlPackage, slidePath: string) {
  const notesMasterPath = getNotesMasterPath(pkg);
  if (!notesMasterPath) {
    throw new Error("The exported slide package does not contain a notes master. This PowerPoint host cannot round-trip speaker notes through the current Open XML fallback.");
  }

  const { relsPath: slideRelsPath, doc: slideRelsDoc } = getOrCreateRelationshipsDoc(pkg, slidePath);
  const existingNotesRelationship = getRelationshipTarget(slideRelsDoc, RELATIONSHIP_TYPE_NOTES_SLIDE);
  if (existingNotesRelationship) {
    const target = existingNotesRelationship.getAttribute("Target");
    if (!target) {
      throw new Error("The slide notes relationship is missing its target.");
    }
    const notesSlidePath = resolveTargetPath(slidePath, target);
    return notesSlidePath;
  }

  const slideNumber = /slide(\d+)\.xml$/.exec(slidePath)?.[1] || "1";
  const notesSlidePath = `ppt/notesSlides/notesSlide${slideNumber}.xml`;
  const notesSlideRelsPath = relationshipPartPath(notesSlidePath);
  const notesMasterFilename = notesMasterPath.split("/").pop();
  if (!notesMasterFilename) {
    throw new Error("The notes master path is invalid.");
  }
  const notesMasterRelativeTarget = `../notesMasters/${notesMasterFilename}`;

  pkg.writeText(notesSlidePath, buildNotesSlideXml());
  pkg.writeText(notesSlideRelsPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIP_TYPE_NOTES_MASTER}" Target="${notesMasterRelativeTarget}"/></Relationships>`);

  const relationship = slideRelsDoc.createElementNS(slideRelsDoc.documentElement.namespaceURI, "Relationship");
  relationship.setAttribute("Id", nextRelationshipId(slideRelsDoc));
  relationship.setAttribute("Type", RELATIONSHIP_TYPE_NOTES_SLIDE);
  relationship.setAttribute("Target", `../notesSlides/notesSlide${slideNumber}.xml`);
  slideRelsDoc.documentElement.appendChild(relationship);
  pkg.writeText(slideRelsPath, serializeXml(slideRelsDoc));

  ensureContentTypeOverride(pkg, notesSlidePath, CONTENT_TYPE_NOTES_SLIDE);
  return notesSlidePath;
}

export function extractSpeakerNotesFromBase64Presentation(base64: string) {
  const pkg = new OpenXmlPackage(base64);
  const slidePath = getFirstSlidePath(pkg);
  const slideRelsPath = relationshipPartPath(slidePath);
  if (!pkg.has(slideRelsPath)) {
    return "";
  }

  const slideRelsDoc = parseXml(pkg.readText(slideRelsPath));
  const notesRelationship = getRelationshipTarget(slideRelsDoc, RELATIONSHIP_TYPE_NOTES_SLIDE);
  if (!notesRelationship) {
    return "";
  }

  const target = notesRelationship.getAttribute("Target");
  if (!target) return "";
  const notesPath = resolveTargetPath(slidePath, target);
  if (!pkg.has(notesPath)) return "";

  const notesDoc = parseXml(pkg.readText(notesPath));
  const shapes = Array.from(notesDoc.getElementsByTagNameNS(NS_P, "sp"));
  const noteBlocks = shapes
    .filter((shape) => !EXCLUDED_NOTE_PLACEHOLDER_TYPES.has(getPlaceholderType(shape) || ""))
    .map((shape) => extractTextBody(getTextBody(shape)))
    .filter(Boolean);
  return noteBlocks.join("\n\n").trim();
}

export function setSpeakerNotesInBase64Presentation(base64: string, notes: string) {
  const pkg = new OpenXmlPackage(base64);
  const slidePath = getFirstSlidePath(pkg);
  const notesPath = ensureNotesSlide(pkg, slidePath);
  const notesDoc = parseXml(pkg.readText(notesPath));
  const notesShape = ensureSpeakerNotesShape(notesDoc);
  const textBody = getTextBody(notesShape);
  if (!textBody) {
    throw new Error("The notes slide does not contain a writable text body.");
  }

  writeTextBody(textBody, notes);
  pkg.writeText(notesPath, serializeXml(notesDoc));
  return pkg.toBase64();
}

export async function replaceSlideWithMutatedOpenXml(
  context: PowerPoint.RequestContext,
  slideIndex: number,
  mutate: (base64: string) => string,
) {
  if (!isPowerPointRequirementSetSupported("1.8")) {
    throw new Error("PowerPoint Open XML slide round-tripping requires PowerPointApi 1.8.");
  }

  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();

  if (slideIndex < 0 || slideIndex >= slides.items.length) {
    throw new Error(`Invalid slideIndex ${slideIndex}. Must be 0-${slides.items.length - 1}.`);
  }

  const sourceSlide = slides.items[slideIndex];
  sourceSlide.load("id");

  let previousSlideId: string | undefined;
  if (slideIndex > 0) {
    const previousSlide = slides.items[slideIndex - 1];
    previousSlide.load("id");
  }
  await context.sync();

  if (slideIndex > 0) {
    previousSlideId = slides.items[slideIndex - 1].id;
  }

  const exported = sourceSlide.exportAsBase64();
  await context.sync();
  const mutated = mutate(exported.value);

  context.presentation.insertSlidesFromBase64(mutated, {
    formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
    ...(previousSlideId ? { targetSlideId: previousSlideId } : {}),
  });
  await context.sync();

  slides.load("items");
  await context.sync();
  const originalSlideAfterInsert = slides.items[slideIndex + 1];
  if (!originalSlideAfterInsert) {
    throw new Error("Failed to locate the original slide after Open XML round-trip insertion.");
  }
  originalSlideAfterInsert.delete();
  await context.sync();
}

export async function exportSlideAsBase64(context: PowerPoint.RequestContext, slideIndex: number) {
  if (!isPowerPointRequirementSetSupported("1.8")) {
    throw new Error("PowerPoint Open XML slide export requires PowerPointApi 1.8.");
  }

  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  if (slideIndex < 0 || slideIndex >= slides.items.length) {
    throw new Error(`Invalid slideIndex ${slideIndex}. Must be 0-${slides.items.length - 1}.`);
  }

  const exported = slides.items[slideIndex].exportAsBase64();
  await context.sync();
  return exported.value;
}
