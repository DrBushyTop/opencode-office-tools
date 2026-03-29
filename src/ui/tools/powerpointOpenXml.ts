import {
  OpenXmlPackage,
  createRelationshipsDocument,
  nextRelationshipId,
  parseXml,
  relationshipPartPath,
  resolveTargetPath,
  serializeXml,
} from "./openXmlPackage";
import { invalidSlideIndexMessage, isPowerPointRequirementSetSupported } from "./powerpointShared";
import { z } from "zod";

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const RELATIONSHIP_TYPE_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const RELATIONSHIP_TYPE_NOTES_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const RELATIONSHIP_TYPE_NOTES_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
const CONTENT_TYPE_NOTES_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
const ANIMATABLE_SHAPE_LOCAL_NAMES = new Set(["sp", "cxnSp", "pic", "graphicFrame", "grpSp", "contentPart"]);
const ELEMENT_NODE = 1;

const EXCLUDED_NOTE_PLACEHOLDER_TYPES = new Set(["sldImg", "hdr", "dt", "ftr", "sldNum"]);

function getDirectChildByTagName(parent: Element, namespace: string, localName: string) {
  return Array.from(parent.childNodes).find(
    (node) => node.nodeType === ELEMENT_NODE && (node as Element).namespaceURI === namespace && (node as Element).localName === localName,
  ) as Element | undefined;
}

export const slideTransitionEffectSchema = z.enum(["none", "cut", "fade", "dissolve", "random", "randomBar", "push", "wipe", "split", "cover", "pull", "zoom"]);
export const slideTransitionSpeedSchema = z.enum(["slow", "medium", "fast"]);
export const slideTransitionDirectionSchema = z.enum(["left", "right", "up", "down", "horizontal", "vertical", "in", "out"]);
export const slideTransitionOrientationSchema = z.enum(["horizontal", "vertical"]);

export const slideTransitionDefinitionSchema = z.object({
  effect: slideTransitionEffectSchema,
  speed: slideTransitionSpeedSchema.optional(),
  advanceOnClick: z.boolean().optional(),
  advanceAfterMs: z.number().optional(),
  durationMs: z.number().optional(),
  direction: slideTransitionDirectionSchema.optional(),
  orientation: slideTransitionOrientationSchema.optional(),
  throughBlack: z.boolean().optional(),
});

export type SlideTransitionDefinition = z.infer<typeof slideTransitionDefinitionSchema>;

export const slideAnimationTypeSchema = z.enum([
  "appear", "fade", "flyIn", "wipe", "zoomIn", "floatIn", "riseUp", "peekIn", "growAndTurn",
  "complementaryColor", "changeFillColor", "changeLineColor",
  "motionPath", "scale", "rotate",
]);
export const slideAnimationStartSchema = z.enum(["onClick", "withPrevious", "afterPrevious"]);
export const slideAnimationPathOriginSchema = z.enum(["parent", "layout"]);
export const slideAnimationPathEditModeSchema = z.enum(["relative", "fixed"]);
export const slideAnimationDirectionSchema = z.enum(["left", "right", "up", "down"]);
export const slideAnimationColorSpaceSchema = z.enum(["hsl", "rgb"]);

export const slideAnimationDefinitionSchema = z.object({
  type: slideAnimationTypeSchema,
  start: slideAnimationStartSchema,
  durationMs: z.number().optional(),
  delayMs: z.number().optional(),
  repeatCount: z.number().optional(),
  shapeId: z.string(),
  path: z.string().optional(),
  pathOrigin: slideAnimationPathOriginSchema.optional(),
  pathEditMode: slideAnimationPathEditModeSchema.optional(),
  scaleXPercent: z.number().optional(),
  scaleYPercent: z.number().optional(),
  angleDegrees: z.number().optional(),
  direction: slideAnimationDirectionSchema.optional(),
  toColor: z.string().optional(),
  colorSpace: slideAnimationColorSpaceSchema.optional(),
});

export type AnimationType = z.infer<typeof slideAnimationTypeSchema>;
export type SlideAnimationDefinition = z.infer<typeof slideAnimationDefinitionSchema>;

const ENTRANCE_ANIMATION_TYPES = new Set(["appear", "fade", "flyIn", "wipe", "zoomIn", "floatIn", "riseUp", "peekIn", "growAndTurn"]);

const EMPHASIS_COLOR_TYPES = new Set(["complementaryColor", "changeFillColor", "changeLineColor"]);

const ENTRANCE_PRESET_IDS: Record<string, number> = {
  appear: 1,
  fade: 10,
  flyIn: 2,
  wipe: 22,
  zoomIn: 23,
  floatIn: 30,
  riseUp: 34,
  peekIn: 42,
  growAndTurn: 37,
};

const EMPHASIS_PRESET_IDS: Record<string, number> = {
  complementaryColor: 70,
  changeFillColor: 54,
  changeLineColor: 60,
};

const FLY_IN_DIRECTION_SUBTYPES: Record<string, number> = {
  left: 4,
  right: 2,
  up: 1,
  down: 3,
};

const WIPE_DIRECTION_SUBTYPES: Record<string, number> = {
  left: 2,
  right: 4,
  up: 10,
  down: 3,
};

interface SlideAnimationMutationDefinition extends Omit<SlideAnimationDefinition, "shapeId"> {
  targetXmlShapeId: string;
}

function isEmphasisColorAnimation(animation: SlideAnimationMutationDefinition) {
  return EMPHASIS_COLOR_TYPES.has(animation.type);
}

function getEmphasisPresetId(animation: SlideAnimationMutationDefinition) {
  return EMPHASIS_PRESET_IDS[animation.type];
}

function getFirstSlidePath(pkg: OpenXmlPackage) {
  const slides = pkg.listPaths().filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path));
  if (!slides.length) {
    throw new Error("The exported PowerPoint package does not contain a slide XML part.");
  }
  return slides.sort((left, right) => left.localeCompare(right, undefined, { numeric: true }))[0];
}

function getFirstSlideDocument(pkg: OpenXmlPackage) {
  const slidePath = getFirstSlidePath(pkg);
  return { slidePath, slideDoc: parseXml(pkg.readText(slidePath)) };
}

function getNotesMasterPath(pkg: OpenXmlPackage) {
  return pkg.listPaths().find((path) => /^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(path)) || null;
}

function getRelationshipTarget(relationshipsDoc: XMLDocument, type: string) {
  const relationships = Array.from(relationshipsDoc.getElementsByTagName("Relationship"));
  return relationships.find((relationship) => relationship.getAttribute("Type") === type) || null;
}

function getAlternateContentTransitionNodes(slideDoc: XMLDocument) {
  const alternateContents = Array.from(slideDoc.getElementsByTagNameNS(NS_MC, "AlternateContent"));
  return alternateContents.flatMap((alternateContent) => {
    const choices = Array.from(alternateContent.getElementsByTagNameNS(NS_MC, "Choice"));
    const fallbacks = Array.from(alternateContent.getElementsByTagNameNS(NS_MC, "Fallback"));
    return [...choices, ...fallbacks].flatMap((node) => Array.from(node.getElementsByTagNameNS(NS_P, "transition")));
  });
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

function getSlideShapeElementsInOrder(slideDoc: XMLDocument) {
  const spTree = slideDoc.getElementsByTagNameNS(NS_P, "spTree")[0];
  if (!spTree) {
    throw new Error("The slide XML is missing its shape tree.");
  }

  return Array.from(spTree.childNodes).filter(
    (node) => node.nodeType === ELEMENT_NODE
      && (node as Element).namespaceURI === NS_P
      && ANIMATABLE_SHAPE_LOCAL_NAMES.has((node as Element).localName),
  ) as Element[];
}

function getXmlShapeId(shape: Element, shapeIndex: number) {
  const cNvPr = shape.getElementsByTagNameNS(NS_P, "cNvPr")[0];
  const xmlShapeId = cNvPr?.getAttribute("id");
  if (!xmlShapeId) {
    throw new Error(`The exported slide XML is missing a non-visual shape id for shapeIndex ${shapeIndex}.`);
  }

  return xmlShapeId;
}

function resolveAnimationTargetXmlShapeId(slideDoc: XMLDocument, shapeIndex: number) {
  const shapes = getSlideShapeElementsInOrder(slideDoc);
  const shape = shapes[shapeIndex];
  if (!shape) {
    throw new Error(`The exported slide XML does not contain shapeIndex ${shapeIndex}. Available shape indexes: 0-${Math.max(shapes.length - 1, 0)}.`);
  }

  return getXmlShapeId(shape, shapeIndex);
}

export function findSlideShapeIndexByXmlShapeIdInBase64Presentation(base64: string, xmlShapeId: string) {
  const pkg = new OpenXmlPackage(base64);
  const { slideDoc } = getFirstSlideDocument(pkg);
  return getSlideShapeElementsInOrder(slideDoc).findIndex((shape, shapeIndex) => getXmlShapeId(shape, shapeIndex) === xmlShapeId);
}

export function listXmlShapeIdsInBase64Presentation(base64: string) {
  const pkg = new OpenXmlPackage(base64);
  const { slideDoc } = getFirstSlideDocument(pkg);
  return getSlideShapeElementsInOrder(slideDoc).map((shape, shapeIndex) => getXmlShapeId(shape, shapeIndex));
}

export const openXmlRoundTripResultSchema = z.object({
  originalSlideId: z.string(),
  replacementSlideId: z.string(),
  finalSlideIndex: z.number(),
});

export type OpenXmlRoundTripResult = z.infer<typeof openXmlRoundTripResultSchema>;

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
  nvSpPr.appendChild(cNvPr);
  nvSpPr.appendChild(cNvSpPr);
  nvSpPr.appendChild(nvPr);

  const spPr = notesDoc.createElementNS(NS_P, "p:spPr");
  const txBody = notesDoc.createElementNS(NS_P, "p:txBody");
  writeTextBody(txBody, "");

  shape.appendChild(nvSpPr);
  shape.appendChild(spPr);
  shape.appendChild(txBody);
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

function getTransitionEffectDetails(transition: Element) {
  const effect = Array.from(transition.childNodes).find((node) => node.nodeType === ELEMENT_NODE) as Element | undefined;
  if (!effect) return { effect: "none" as const };

  const localName = effect.localName;
  switch (localName) {
    case "cut":
      return {
        effect: "cut" as const,
        throughBlack: effect.getAttribute("thruBlk") === "1" || effect.getAttribute("thruBlk") === "true",
      };
    case "randomBar":
      return {
        effect: "randomBar" as const,
        direction: (effect.getAttribute("dir") === "vert" ? "vertical" : "horizontal") as SlideTransitionDefinition["direction"],
      };
    case "push":
    case "wipe":
    case "cover":
    case "pull": {
      const dir = effect.getAttribute("dir") || "";
      return {
        effect: localName as SlideTransitionDefinition["effect"],
        direction: ({ l: "left", r: "right", u: "up", d: "down" } as Record<string, SlideTransitionDefinition["direction"]>)[dir] || undefined,
      };
    }
    case "split":
      return {
        effect: "split" as const,
        direction: (effect.getAttribute("dir") === "out" ? "out" : "in") as SlideTransitionDefinition["direction"],
        orientation: (effect.getAttribute("orient") === "vert" ? "vertical" : "horizontal") as SlideTransitionDefinition["orientation"],
      };
    default:
      return { effect: localName as SlideTransitionDefinition["effect"] };
  }
}

function buildTransitionElement(doc: XMLDocument, definition: SlideTransitionDefinition) {
  const transition = doc.createElementNS(NS_P, "p:transition");
  if (definition.speed) {
    transition.setAttribute("spd", definition.speed === "medium" ? "med" : definition.speed);
  }
  if (definition.advanceOnClick !== undefined) {
    transition.setAttribute("advClick", definition.advanceOnClick ? "1" : "0");
  }
  if (definition.advanceAfterMs !== undefined) {
    transition.setAttribute("advTm", String(definition.advanceAfterMs));
  }

  if (definition.effect === "none") {
    return transition;
  }

  const effect = doc.createElementNS(NS_P, `p:${definition.effect}`);
  switch (definition.effect) {
    case "cut":
      if (definition.throughBlack) effect.setAttribute("thruBlk", "1");
      break;
    case "randomBar":
      effect.setAttribute("dir", definition.direction === "vertical" ? "vert" : "horz");
      break;
    case "push":
    case "wipe":
    case "cover":
    case "pull": {
      const dirMap: Record<string, string> = { left: "l", right: "r", up: "u", down: "d" };
      if (definition.direction && dirMap[definition.direction]) {
        effect.setAttribute("dir", dirMap[definition.direction]);
      }
      break;
    }
    case "split":
      effect.setAttribute("dir", definition.direction === "out" ? "out" : "in");
      effect.setAttribute("orient", definition.orientation === "vertical" ? "vert" : "horz");
      break;
    default:
      break;
  }

  transition.appendChild(effect);
  return transition;
}

function buildTransitionNode(doc: XMLDocument, definition: SlideTransitionDefinition) {
  if (definition.durationMs === undefined) {
    return buildTransitionElement(doc, definition);
  }

  const alternateContent = doc.createElementNS(NS_MC, "mc:AlternateContent");
  const choice = doc.createElementNS(NS_MC, "mc:Choice");
  choice.setAttribute("Requires", "p14");
  const fallback = doc.createElementNS(NS_MC, "mc:Fallback");

  const choiceTransition = buildTransitionElement(doc, definition);
  choiceTransition.setAttributeNS(NS_P14, "p14:dur", String(definition.durationMs));
  const fallbackTransition = buildTransitionElement(doc, { ...definition, durationMs: undefined });

  choice.appendChild(choiceTransition);
  fallback.appendChild(fallbackTransition);
  alternateContent.appendChild(choice);
  alternateContent.appendChild(fallback);
  return alternateContent;
}

function clearSlideTransitionNodes(slideDoc: XMLDocument) {
  const directTransitions = Array.from(slideDoc.documentElement.getElementsByTagNameNS(NS_P, "transition"))
    .filter((node) => node.parentNode === slideDoc.documentElement);
  for (const transition of directTransitions) {
    transition.parentNode?.removeChild(transition);
  }

  const alternateContents = Array.from(slideDoc.getElementsByTagNameNS(NS_MC, "AlternateContent"));
  for (const alternateContent of alternateContents) {
    if (alternateContent.getElementsByTagNameNS(NS_P, "transition").length) {
      alternateContent.parentNode?.removeChild(alternateContent);
    }
  }
}

function nextTimeNodeId(slideDoc: XMLDocument) {
  const ids = Array.from(slideDoc.getElementsByTagNameNS(NS_P, "cTn"))
    .map((node) => Number(node.getAttribute("id") || 0));
  return String((ids.length ? Math.max(...ids) : 0) + 1);
}

function createTimeNodeIdAllocator(slideDoc: XMLDocument) {
  let nextId = Number(nextTimeNodeId(slideDoc));
  return () => String(nextId++);
}

function getAnimationDurationMs(animation: SlideAnimationMutationDefinition) {
  return animation.durationMs ?? 1000;
}

function getOrCreateChild(parent: Element, namespace: string, qualifiedName: string) {
  const localName = qualifiedName.split(":").pop() || qualifiedName;
  const existing = Array.from(parent.childNodes).find(
    (node) => node.nodeType === ELEMENT_NODE && (node as Element).namespaceURI === namespace && (node as Element).localName === localName,
  ) as Element | undefined;
  if (existing) return existing;
  const created = parent.ownerDocument.createElementNS(namespace, qualifiedName);
  parent.appendChild(created);
  return created;
}

function buildTimingRoot(doc: XMLDocument) {
  const timing = doc.createElementNS(NS_P, "p:timing");
  const tnLst = doc.createElementNS(NS_P, "p:tnLst");
  const rootPar = doc.createElementNS(NS_P, "p:par");
  const rootCtn = doc.createElementNS(NS_P, "p:cTn");
  rootCtn.setAttribute("id", "1");
  rootCtn.setAttribute("dur", "indefinite");
  rootCtn.setAttribute("restart", "never");
  rootCtn.setAttribute("nodeType", "tmRoot");
  rootCtn.appendChild(doc.createElementNS(NS_P, "p:childTnLst"));
  rootPar.appendChild(rootCtn);
  tnLst.appendChild(rootPar);
  timing.appendChild(tnLst);
  return timing;
}

function getOrCreateTimingRoot(slideDoc: XMLDocument) {
  const timing = getDirectChildByTagName(slideDoc.documentElement, NS_P, "timing") || buildTimingRoot(slideDoc);
  if (!timing.parentNode) {
    const extLst = getDirectChildByTagName(slideDoc.documentElement, NS_P, "extLst") || null;
    slideDoc.documentElement.insertBefore(timing, extLst);
  }
  return timing;
}

function ensureSeqAttributes(seq: Element) {
  if (!seq.getAttribute("concurrent")) seq.setAttribute("concurrent", "1");
  if (!seq.getAttribute("nextAc")) seq.setAttribute("nextAc", "seek");

  const doc = seq.ownerDocument;
  if (!getDirectChildByTagName(seq, NS_P, "prevCondLst")) {
    const prevCondLst = doc.createElementNS(NS_P, "p:prevCondLst");
    const cond = doc.createElementNS(NS_P, "p:cond");
    cond.setAttribute("evt", "onPrev");
    cond.setAttribute("delay", "0");
    const tgtEl = doc.createElementNS(NS_P, "p:tgtEl");
    tgtEl.appendChild(doc.createElementNS(NS_P, "p:sldTgt"));
    cond.appendChild(tgtEl);
    prevCondLst.appendChild(cond);
    seq.appendChild(prevCondLst);
  }

  if (!getDirectChildByTagName(seq, NS_P, "nextCondLst")) {
    const nextCondLst = doc.createElementNS(NS_P, "p:nextCondLst");
    const cond = doc.createElementNS(NS_P, "p:cond");
    cond.setAttribute("evt", "onNext");
    cond.setAttribute("delay", "0");
    const tgtEl = doc.createElementNS(NS_P, "p:tgtEl");
    tgtEl.appendChild(doc.createElementNS(NS_P, "p:sldTgt"));
    cond.appendChild(tgtEl);
    nextCondLst.appendChild(cond);
    seq.appendChild(nextCondLst);
  }
}

function getOrCreateMainSequence(slideDoc: XMLDocument) {
  const timing = getOrCreateTimingRoot(slideDoc);
  const tnLst = getOrCreateChild(timing, NS_P, "p:tnLst");
  const rootPar = getOrCreateChild(tnLst, NS_P, "p:par");
  const rootCtn = getOrCreateChild(rootPar, NS_P, "p:cTn");
  if (!rootCtn.getAttribute("id")) rootCtn.setAttribute("id", "1");
  if (!rootCtn.getAttribute("dur")) rootCtn.setAttribute("dur", "indefinite");
  if (!rootCtn.getAttribute("restart")) rootCtn.setAttribute("restart", "never");
  if (!rootCtn.getAttribute("nodeType")) rootCtn.setAttribute("nodeType", "tmRoot");
  const rootChildTnLst = getOrCreateChild(rootCtn, NS_P, "p:childTnLst");
  let mainSeq = Array.from(rootChildTnLst.childNodes).find(
    (node) => node.nodeType === ELEMENT_NODE && (node as Element).namespaceURI === NS_P && (node as Element).localName === "seq" && (node as Element).getElementsByTagNameNS(NS_P, "cTn")[0]?.getAttribute("nodeType") === "mainSeq",
  ) as Element | undefined;

  if (!mainSeq) {
    mainSeq = slideDoc.createElementNS(NS_P, "p:seq");
    const mainCtn = slideDoc.createElementNS(NS_P, "p:cTn");
    mainCtn.setAttribute("id", nextTimeNodeId(slideDoc));
    mainCtn.setAttribute("dur", "indefinite");
    mainCtn.setAttribute("nodeType", "mainSeq");
    mainCtn.appendChild(slideDoc.createElementNS(NS_P, "p:childTnLst"));
    mainSeq.appendChild(mainCtn);
    rootChildTnLst.appendChild(mainSeq);
  }

  ensureSeqAttributes(mainSeq);

  return mainSeq;
}

function buildTargetElement(doc: XMLDocument, xmlShapeId: string) {
  const target = doc.createElementNS(NS_P, "p:tgtEl");
  const spTarget = doc.createElementNS(NS_P, "p:spTgt");
  spTarget.setAttribute("spid", xmlShapeId);
  target.appendChild(spTarget);
  return target;
}

function buildCommonBehavior(doc: XMLDocument, animation: SlideAnimationMutationDefinition, allocTimeNodeId: () => string) {
  const cBhvr = doc.createElementNS(NS_P, "p:cBhvr");
  const cTn = doc.createElementNS(NS_P, "p:cTn");
  cTn.setAttribute("id", allocTimeNodeId());
  cTn.setAttribute("dur", String(getAnimationDurationMs(animation)));
  cTn.setAttribute("fill", "hold");
  if (animation.repeatCount !== undefined && animation.repeatCount > 0) {
    cTn.setAttribute("repeatCount", String(animation.repeatCount));
  }
  cBhvr.appendChild(cTn);
  cBhvr.appendChild(buildTargetElement(doc, animation.targetXmlShapeId));
  return cBhvr;
}

function buildVisibilitySet(doc: XMLDocument, targetShapeId: string, allocTimeNodeId: () => string) {
  const set = doc.createElementNS(NS_P, "p:set");
  const cBhvr = doc.createElementNS(NS_P, "p:cBhvr");
  const cTn = doc.createElementNS(NS_P, "p:cTn");
  cTn.setAttribute("id", allocTimeNodeId());
  cTn.setAttribute("dur", "1");
  cTn.setAttribute("fill", "hold");
  const stCondLst = doc.createElementNS(NS_P, "p:stCondLst");
  const cond = doc.createElementNS(NS_P, "p:cond");
  cond.setAttribute("delay", "0");
  stCondLst.appendChild(cond);
  cTn.appendChild(stCondLst);
  cBhvr.appendChild(cTn);
  cBhvr.appendChild(buildTargetElement(doc, targetShapeId));
  const attrNameLst = doc.createElementNS(NS_P, "p:attrNameLst");
  const attrName = doc.createElementNS(NS_P, "p:attrName");
  attrName.textContent = "style.visibility";
  attrNameLst.appendChild(attrName);
  cBhvr.appendChild(attrNameLst);
  set.appendChild(cBhvr);
  const to = doc.createElementNS(NS_P, "p:to");
  const strVal = doc.createElementNS(NS_P, "p:strVal");
  strVal.setAttribute("val", "visible");
  to.appendChild(strVal);
  set.appendChild(to);
  return set;
}

/**
 * Build a `p:anim` property animation element with a time-value list.
 * Used by entrance animations like peekIn and growAndTurn that animate ppt_x/ppt_y
 * with explicit keyframe values rather than motion paths.
 */
function buildPropertyAnimation(
  doc: XMLDocument,
  targetShapeId: string,
  attrName: string,
  timeValues: Array<{ time: number; value: string }>,
  allocTimeNodeId: () => string,
  opts?: {
    durationMs?: number;
    fill?: string;
    decel?: string;
    accel?: string;
    delayMs?: number;
    additive?: string;
  },
): Element {
  const anim = doc.createElementNS(NS_P, "p:anim");
  anim.setAttribute("calcmode", "lin");
  anim.setAttribute("valueType", "num");

  const cBhvr = doc.createElementNS(NS_P, "p:cBhvr");
  if (opts?.additive) cBhvr.setAttribute("additive", opts.additive);
  const cTn = doc.createElementNS(NS_P, "p:cTn");
  cTn.setAttribute("id", allocTimeNodeId());
  cTn.setAttribute("dur", String(opts?.durationMs || 1000));
  if (opts?.fill !== undefined) cTn.setAttribute("fill", opts.fill);
  else cTn.setAttribute("fill", "hold");
  if (opts?.decel) cTn.setAttribute("decel", opts.decel);
  if (opts?.accel) cTn.setAttribute("accel", opts.accel);
  if (opts?.delayMs) {
    const stCondLst = doc.createElementNS(NS_P, "p:stCondLst");
    const cond = doc.createElementNS(NS_P, "p:cond");
    cond.setAttribute("delay", String(opts.delayMs));
    stCondLst.appendChild(cond);
    cTn.appendChild(stCondLst);
  }
  cBhvr.appendChild(cTn);
  cBhvr.appendChild(buildTargetElement(doc, targetShapeId));

  const attrNameLst = doc.createElementNS(NS_P, "p:attrNameLst");
  const attrNameEl = doc.createElementNS(NS_P, "p:attrName");
  attrNameEl.textContent = attrName;
  attrNameLst.appendChild(attrNameEl);
  cBhvr.appendChild(attrNameLst);
  anim.appendChild(cBhvr);

  const tavLst = doc.createElementNS(NS_P, "p:tavLst");
  for (const tv of timeValues) {
    const tav = doc.createElementNS(NS_P, "p:tav");
    tav.setAttribute("tm", String(tv.time));
    const val = doc.createElementNS(NS_P, "p:val");
    const strVal = doc.createElementNS(NS_P, "p:strVal");
    strVal.setAttribute("val", tv.value);
    val.appendChild(strVal);
    tav.appendChild(val);
    tavLst.appendChild(tav);
  }
  anim.appendChild(tavLst);

  return anim;
}

function buildEntranceAnimationNodes(doc: XMLDocument, animation: SlideAnimationMutationDefinition, allocTimeNodeId: () => string): Element[] {
  const visibilitySet = buildVisibilitySet(doc, animation.targetXmlShapeId, allocTimeNodeId);

  if (animation.type === "appear") {
    return [visibilitySet];
  }

  if (animation.type === "fade") {
    const animEffect = doc.createElementNS(NS_P, "p:animEffect");
    animEffect.setAttribute("transition", "in");
    animEffect.setAttribute("filter", "fade");
    const cBhvr = buildCommonBehavior(doc, animation, allocTimeNodeId);
    animEffect.appendChild(cBhvr);
    return [visibilitySet, animEffect];
  }

  if (animation.type === "flyIn") {
    const dir = animation.direction || "down";
    const paths: Record<string, string> = {
      left: "M -1 0 L 0 0 E",
      right: "M 1 0 L 0 0 E",
      up: "M 0 1 L 0 0 E",
      down: "M 0 -1 L 0 0 E",
    };
    const node = doc.createElementNS(NS_P, "p:animMotion");
    node.setAttribute("origin", "layout");
    node.setAttribute("path", paths[dir] || paths.down);
    node.setAttribute("pathEditMode", "relative");
    const cBhvr = buildCommonBehavior(doc, animation, allocTimeNodeId);
    const attrNameList = doc.createElementNS(NS_P, "p:attrNameLst");
    const attrX = doc.createElementNS(NS_P, "p:attrName");
    attrX.textContent = "ppt_x";
    const attrY = doc.createElementNS(NS_P, "p:attrName");
    attrY.textContent = "ppt_y";
    attrNameList.appendChild(attrX);
    attrNameList.appendChild(attrY);
    cBhvr.appendChild(attrNameList);
    node.appendChild(cBhvr);
    return [visibilitySet, node];
  }

  if (animation.type === "wipe") {
    const dir = animation.direction || "left";
    const filterMap: Record<string, string> = {
      left: "wipe(left)",
      right: "wipe(right)",
      up: "wipe(up)",
      down: "wipe(down)",
    };
    const animEffect = doc.createElementNS(NS_P, "p:animEffect");
    animEffect.setAttribute("transition", "in");
    animEffect.setAttribute("filter", filterMap[dir] || filterMap.left);
    const cBhvr = buildCommonBehavior(doc, animation, allocTimeNodeId);
    animEffect.appendChild(cBhvr);
    return [visibilitySet, animEffect];
  }

  if (animation.type === "zoomIn") {
    const node = doc.createElementNS(NS_P, "p:animScale");
    const cBhvr = buildCommonBehavior(doc, animation, allocTimeNodeId);
    node.appendChild(cBhvr);
    const from = doc.createElementNS(NS_P, "p:from");
    from.setAttribute("x", "0");
    from.setAttribute("y", "0");
    const to = doc.createElementNS(NS_P, "p:to");
    to.setAttribute("x", "100000");
    to.setAttribute("y", "100000");
    node.appendChild(from);
    node.appendChild(to);
    return [visibilitySet, node];
  }

  if (animation.type === "floatIn") {
    // Float In = upward motion + fade-in (combines animMotion up + animEffect fade)
    const motionNode = doc.createElementNS(NS_P, "p:animMotion");
    motionNode.setAttribute("origin", "layout");
    motionNode.setAttribute("path", "M 0 0.1 L 0 0 E");
    motionNode.setAttribute("pathEditMode", "relative");
    const motionCBhvr = buildCommonBehavior(doc, animation, allocTimeNodeId);
    const motionAttrNameList = doc.createElementNS(NS_P, "p:attrNameLst");
    const motionAttrX = doc.createElementNS(NS_P, "p:attrName");
    motionAttrX.textContent = "ppt_x";
    const motionAttrY = doc.createElementNS(NS_P, "p:attrName");
    motionAttrY.textContent = "ppt_y";
    motionAttrNameList.appendChild(motionAttrX);
    motionAttrNameList.appendChild(motionAttrY);
    motionCBhvr.appendChild(motionAttrNameList);
    motionNode.appendChild(motionCBhvr);
    const fadeNode = doc.createElementNS(NS_P, "p:animEffect");
    fadeNode.setAttribute("transition", "in");
    fadeNode.setAttribute("filter", "fade");
    const fadeCBhvr = buildCommonBehavior(doc, animation, allocTimeNodeId);
    fadeNode.appendChild(fadeCBhvr);
    return [visibilitySet, motionNode, fadeNode];
  }

  if (animation.type === "riseUp") {
    // Rise Up = upward fly entrance (presetID 34, similar to flyIn up but steeper motion)
    const node = doc.createElementNS(NS_P, "p:animMotion");
    node.setAttribute("origin", "layout");
    node.setAttribute("path", "M 0 1 L 0 0 E");
    node.setAttribute("pathEditMode", "relative");
    const cBhvr = buildCommonBehavior(doc, animation, allocTimeNodeId);
    const attrNameList = doc.createElementNS(NS_P, "p:attrNameLst");
    const attrX = doc.createElementNS(NS_P, "p:attrName");
    attrX.textContent = "ppt_x";
    const attrY = doc.createElementNS(NS_P, "p:attrName");
    attrY.textContent = "ppt_y";
    attrNameList.appendChild(attrX);
    attrNameList.appendChild(attrY);
    cBhvr.appendChild(attrNameList);
    node.appendChild(cBhvr);
    return [visibilitySet, node];
  }

  if (animation.type === "peekIn") {
    // Peek In = fade in + slight vertical slide from below (#ppt_y+.1 → #ppt_y)
    // Ref: presetID 42, presetSubtype 0
    const dur = getAnimationDurationMs(animation);
    const animEffect = doc.createElementNS(NS_P, "p:animEffect");
    animEffect.setAttribute("transition", "in");
    animEffect.setAttribute("filter", "fade");
    const fadeCBhvr = buildCommonBehavior(doc, animation, allocTimeNodeId);
    animEffect.appendChild(fadeCBhvr);

    const animX = buildPropertyAnimation(doc, animation.targetXmlShapeId, "ppt_x", [
      { time: 0, value: "#ppt_x" },
      { time: 100000, value: "#ppt_x" },
    ], allocTimeNodeId, { durationMs: dur });

    const animY = buildPropertyAnimation(doc, animation.targetXmlShapeId, "ppt_y", [
      { time: 0, value: "#ppt_y+.1" },
      { time: 100000, value: "#ppt_y" },
    ], allocTimeNodeId, { durationMs: dur });

    return [visibilitySet, animEffect, animX, animY];
  }

  if (animation.type === "growAndTurn") {
    // Grow & Turn = fade in + big vertical slide from below with bounce
    // Main motion: 90% of duration, decelerating, from #ppt_y+1 to #ppt_y-.03
    // Bounce: final 10% of duration, accelerating, from #ppt_y-.03 to #ppt_y
    // Ref: presetID 37, presetSubtype 0
    const dur = getAnimationDurationMs(animation);
    const mainDur = Math.round(dur * 0.9);
    const bounceDur = dur - mainDur;

    const animEffect = doc.createElementNS(NS_P, "p:animEffect");
    animEffect.setAttribute("transition", "in");
    animEffect.setAttribute("filter", "fade");
    const fadeCBhvr = buildCommonBehavior(doc, animation, allocTimeNodeId);
    animEffect.appendChild(fadeCBhvr);

    const animX = buildPropertyAnimation(doc, animation.targetXmlShapeId, "ppt_x", [
      { time: 0, value: "#ppt_x" },
      { time: 100000, value: "#ppt_x" },
    ], allocTimeNodeId, { durationMs: dur });

    // Main upward motion with deceleration
    const animYMain = buildPropertyAnimation(doc, animation.targetXmlShapeId, "ppt_y", [
      { time: 0, value: "#ppt_y+1" },
      { time: 100000, value: "#ppt_y-.03" },
    ], allocTimeNodeId, { durationMs: mainDur, decel: "100000" });

    // Bounce back (delayed by mainDur)
    const animYBounce = buildPropertyAnimation(doc, animation.targetXmlShapeId, "ppt_y", [
      { time: 0, value: "#ppt_y-.03" },
      { time: 100000, value: "#ppt_y" },
    ], allocTimeNodeId, { durationMs: bounceDur, accel: "100000", delayMs: mainDur });

    return [visibilitySet, animEffect, animX, animYMain, animYBounce];
  }

  return [visibilitySet];
}

function isEntranceAnimation(animation: SlideAnimationMutationDefinition) {
  return ENTRANCE_ANIMATION_TYPES.has(animation.type);
}

function getEntrancePresetId(animation: SlideAnimationMutationDefinition) {
  return ENTRANCE_PRESET_IDS[animation.type];
}

function getEntrancePresetSubtype(animation: SlideAnimationMutationDefinition): number | undefined {
  if (animation.type === "flyIn") return FLY_IN_DIRECTION_SUBTYPES[animation.direction || "down"];
  if (animation.type === "wipe") return WIPE_DIRECTION_SUBTYPES[animation.direction || "left"];
  if (animation.type === "floatIn") return 16; // float up
  if (animation.type === "riseUp") return 0;
  if (animation.type === "peekIn") return 0;
  if (animation.type === "growAndTurn") return 0;
  return undefined;
}

function buildAnimationNodes(doc: XMLDocument, animation: SlideAnimationMutationDefinition, allocTimeNodeId: () => string): Element[] {
  if (isEntranceAnimation(animation)) {
    return buildEntranceAnimationNodes(doc, animation, allocTimeNodeId);
  }
  if (isEmphasisColorAnimation(animation)) {
    return [buildEmphasisColorAnimationNode(doc, animation, allocTimeNodeId)];
  }
  return [buildEmphasisAnimationNode(doc, animation, allocTimeNodeId)];
}

function buildEmphasisAnimationNode(doc: XMLDocument, animation: SlideAnimationMutationDefinition, allocTimeNodeId: () => string) {
  if (animation.type === "motionPath") {
    const node = doc.createElementNS(NS_P, "p:animMotion");
    node.setAttribute("origin", animation.pathOrigin || "parent");
    node.setAttribute("path", animation.path || "M 0 0 E");
    node.setAttribute("pathEditMode", animation.pathEditMode || "relative");
    const cBhvr = buildCommonBehavior(doc, animation, allocTimeNodeId);
    const attrNameList = doc.createElementNS(NS_P, "p:attrNameLst");
    const attrX = doc.createElementNS(NS_P, "p:attrName");
    attrX.textContent = "ppt_x";
    const attrY = doc.createElementNS(NS_P, "p:attrName");
    attrY.textContent = "ppt_y";
    attrNameList.appendChild(attrX);
    attrNameList.appendChild(attrY);
    cBhvr.appendChild(attrNameList);
    node.appendChild(cBhvr);
    return node;
  }

  if (animation.type === "scale") {
    const node = doc.createElementNS(NS_P, "p:animScale");
    node.appendChild(buildCommonBehavior(doc, animation, allocTimeNodeId));
    const by = doc.createElementNS(NS_P, "p:by");
    by.setAttribute("x", String(Math.round((animation.scaleXPercent ?? 100) * 1000)));
    by.setAttribute("y", String(Math.round((animation.scaleYPercent ?? animation.scaleXPercent ?? 100) * 1000)));
    node.appendChild(by);
    return node;
  }

  const node = doc.createElementNS(NS_P, "p:animRot");
  node.setAttribute("by", String(Math.round((animation.angleDegrees ?? 360) * 60000)));
  node.appendChild(buildCommonBehavior(doc, animation, allocTimeNodeId));
  return node;
}

/**
 * Build an emphasis color animation node (p:animClr).
 *
 * Structure:
 *   <p:animClr clrSpc="hsl|rgb">
 *     <p:cBhvr>
 *       <p:cTn id="X" dur="500" fill="hold"/>
 *       <p:tgtEl><p:spTgt spid="SHAPE_ID"/></p:tgtEl>
 *       <p:attrNameLst><p:attrName>fillcolor|style.color|stroke.color</p:attrName></p:attrNameLst>
 *     </p:cBhvr>
 *     <p:to><a:srgbClr val="FF0000"/> | <a:schemeClr val="accent2"/></p:to>
 *   </p:animClr>
 */
function buildEmphasisColorAnimationNode(doc: XMLDocument, animation: SlideAnimationMutationDefinition, allocTimeNodeId: () => string) {
  const EMPHASIS_ATTR_NAMES: Record<string, string> = {
    complementaryColor: "fillcolor",
    changeFillColor: "fillcolor",
    changeLineColor: "stroke.color",
  };

  const node = doc.createElementNS(NS_P, "p:animClr");
  node.setAttribute("clrSpc", animation.colorSpace || "hsl");

  // Build cBhvr with attrNameLst for the property being animated
  const cBhvr = doc.createElementNS(NS_P, "p:cBhvr");
  const cTn = doc.createElementNS(NS_P, "p:cTn");
  cTn.setAttribute("id", allocTimeNodeId());
  cTn.setAttribute("dur", String(getAnimationDurationMs(animation)));
  cTn.setAttribute("fill", "hold");
  if (animation.repeatCount !== undefined && animation.repeatCount > 0) {
    cTn.setAttribute("repeatCount", String(animation.repeatCount));
  }
  cBhvr.appendChild(cTn);
  cBhvr.appendChild(buildTargetElement(doc, animation.targetXmlShapeId));
  const attrNameLst = doc.createElementNS(NS_P, "p:attrNameLst");
  const attrName = doc.createElementNS(NS_P, "p:attrName");
  attrName.textContent = EMPHASIS_ATTR_NAMES[animation.type] || "fillcolor";
  attrNameLst.appendChild(attrName);
  cBhvr.appendChild(attrNameLst);
  node.appendChild(cBhvr);

  // Build the <p:to> color element
  const toEl = doc.createElementNS(NS_P, "p:to");
  const colorVal = animation.toColor || "FF0000";
  if (colorVal.length === 6 && /^[0-9A-Fa-f]{6}$/.test(colorVal)) {
    // Hex color
    const srgbClr = doc.createElementNS(NS_A, "a:srgbClr");
    srgbClr.setAttribute("val", colorVal.toUpperCase());
    toEl.appendChild(srgbClr);
  } else {
    // Scheme color (e.g., "accent2", "dk1", "lt1")
    const schemeClr = doc.createElementNS(NS_A, "a:schemeClr");
    schemeClr.setAttribute("val", colorVal);
    toEl.appendChild(schemeClr);
  }
  node.appendChild(toEl);

  return node;
}

function applyPresetAttributes(cTn: Element, animation: SlideAnimationMutationDefinition) {
  if (isEntranceAnimation(animation)) {
    const presetId = getEntrancePresetId(animation);
    if (presetId !== undefined) {
      cTn.setAttribute("presetClass", "entr");
      cTn.setAttribute("presetID", String(presetId));
    }
    const subtype = getEntrancePresetSubtype(animation);
    cTn.setAttribute("presetSubtype", subtype !== undefined ? String(subtype) : "0");
    cTn.setAttribute("grpId", "0");
    return;
  }
  if (isEmphasisColorAnimation(animation)) {
    const presetId = getEmphasisPresetId(animation);
    if (presetId !== undefined) {
      cTn.setAttribute("presetClass", "emph");
      cTn.setAttribute("presetID", String(presetId));
    }
    cTn.setAttribute("presetSubtype", "0");
    cTn.setAttribute("grpId", "0");
  }
}

/**
 * Build the per-shape animation p:par node.
 * This is the innermost wrapper in the 3-level hierarchy.
 *
 * Structure:
 *   <p:par>
 *     <p:cTn presetID="..." presetClass="entr" presetSubtype="0" grpId="0"
 *            fill="hold" nodeType="clickEffect|withEffect|afterEffect"
 *            [dur="...for emphasis"]>
 *       <p:stCondLst><p:cond delay="0"/></p:stCondLst>
 *       <p:childTnLst>
 *         <!-- animation effect nodes (set, animEffect, animMotion, etc.) -->
 *       </p:childTnLst>
 *     </p:cTn>
 *   </p:par>
 */
function buildShapeAnimationPar(
  doc: XMLDocument,
  animation: SlideAnimationMutationDefinition,
  allocTimeNodeId: () => string,
  nodeType: "clickEffect" | "withEffect" | "afterEffect",
) {
  const nodes = buildAnimationNodes(doc, animation, allocTimeNodeId);
  const par = doc.createElementNS(NS_P, "p:par");
  const cTn = doc.createElementNS(NS_P, "p:cTn");
  cTn.setAttribute("id", allocTimeNodeId());
  cTn.setAttribute("fill", "hold");

  applyPresetAttributes(cTn, animation);

  // For emphasis/motion animations that don't set presetClass, set dur directly
  if (!isEntranceAnimation(animation)) {
    cTn.setAttribute("dur", String(getAnimationDurationMs(animation)));
    if (animation.repeatCount !== undefined && animation.repeatCount > 0) {
      cTn.setAttribute("repeatCount", String(animation.repeatCount));
    }
  }

  cTn.setAttribute("nodeType", nodeType);

  // Always add stCondLst with delay="0" (or the animation's delay for afterPrevious)
  const stCondLst = doc.createElementNS(NS_P, "p:stCondLst");
  const cond = doc.createElementNS(NS_P, "p:cond");
  cond.setAttribute("delay", nodeType === "afterEffect" && animation.delayMs ? String(animation.delayMs) : "0");
  stCondLst.appendChild(cond);
  cTn.appendChild(stCondLst);

  const childTnLst = doc.createElementNS(NS_P, "p:childTnLst");
  for (const node of nodes) childTnLst.appendChild(node);
  cTn.appendChild(childTnLst);
  par.appendChild(cTn);
  return par;
}

/**
 * Build a timing-group p:par (the second level in the 3-level hierarchy).
 *
 * Structure:
 *   <p:par>
 *     <p:cTn fill="hold">
 *       <p:stCondLst><p:cond delay="0"/></p:stCondLst>
 *       <p:childTnLst>
 *         <!-- per-shape p:par nodes -->
 *       </p:childTnLst>
 *     </p:cTn>
 *   </p:par>
 */
function buildTimingGroupPar(doc: XMLDocument, allocTimeNodeId: () => string, delayStr = "0") {
  const par = doc.createElementNS(NS_P, "p:par");
  const cTn = doc.createElementNS(NS_P, "p:cTn");
  cTn.setAttribute("id", allocTimeNodeId());
  cTn.setAttribute("fill", "hold");
  const stCondLst = doc.createElementNS(NS_P, "p:stCondLst");
  const cond = doc.createElementNS(NS_P, "p:cond");
  cond.setAttribute("delay", delayStr);
  stCondLst.appendChild(cond);
  cTn.appendChild(stCondLst);
  cTn.appendChild(doc.createElementNS(NS_P, "p:childTnLst"));
  par.appendChild(cTn);
  return par;
}

/**
 * Build a click-group p:par (the outermost level in the 3-level hierarchy).
 *
 * Structure:
 *   <p:par>
 *     <p:cTn fill="hold">
 *       <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
 *       <p:childTnLst>
 *         <!-- timing-group p:par nodes -->
 *       </p:childTnLst>
 *     </p:cTn>
 *   </p:par>
 */
function buildClickGroupPar(doc: XMLDocument, allocTimeNodeId: () => string) {
  const par = doc.createElementNS(NS_P, "p:par");
  const cTn = doc.createElementNS(NS_P, "p:cTn");
  cTn.setAttribute("id", allocTimeNodeId());
  cTn.setAttribute("fill", "hold");
  const stCondLst = doc.createElementNS(NS_P, "p:stCondLst");
  const cond = doc.createElementNS(NS_P, "p:cond");
  cond.setAttribute("delay", "indefinite");
  stCondLst.appendChild(cond);
  cTn.appendChild(stCondLst);
  cTn.appendChild(doc.createElementNS(NS_P, "p:childTnLst"));
  par.appendChild(cTn);
  return par;
}

/**
 * Ensure a p:bldLst exists in the timing element and add an entry for the animated shape.
 * The bldLst sits as a direct child of p:timing, after p:tnLst.
 */
function addBldLstEntry(slideDoc: XMLDocument, xmlShapeId: string) {
  const timing = getDirectChildByTagName(slideDoc.documentElement, NS_P, "timing");
  if (!timing) return;

  let bldLst = getDirectChildByTagName(timing, NS_P, "bldLst");
  if (!bldLst) {
    bldLst = slideDoc.createElementNS(NS_P, "p:bldLst");
    timing.appendChild(bldLst);
  }

  // Check if entry already exists for this shape + grpId
  const existing = Array.from(bldLst.childNodes).find(
    (node) =>
      node.nodeType === ELEMENT_NODE &&
      (node as Element).localName === "bldP" &&
      (node as Element).getAttribute("spid") === xmlShapeId &&
      (node as Element).getAttribute("grpId") === "0",
  );
  if (existing) return;

  const bldP = slideDoc.createElementNS(NS_P, "p:bldP");
  bldP.setAttribute("spid", xmlShapeId);
  bldP.setAttribute("grpId", "0");
  bldP.setAttribute("animBg", "1");
  bldLst.appendChild(bldP);
}

function addSlideAnimationInDocument(slideDoc: XMLDocument, animation: SlideAnimationMutationDefinition) {
  const allocTimeNodeId = createTimeNodeIdAllocator(slideDoc);
  const mainSeq = getOrCreateMainSequence(slideDoc);
  const mainCtn = mainSeq.getElementsByTagNameNS(NS_P, "cTn")[0];
  const mainChildren = getOrCreateChild(mainCtn, NS_P, "p:childTnLst");

  const isEntrance = isEntranceAnimation(animation);

  if (animation.start === "onClick") {
    // Create a new click group with a timing group containing the shape animation
    const clickGroup = buildClickGroupPar(slideDoc, allocTimeNodeId);
    const clickGroupChildren = clickGroup.getElementsByTagNameNS(NS_P, "childTnLst")[0];
    const timingGroup = buildTimingGroupPar(slideDoc, allocTimeNodeId);
    const timingGroupChildren = timingGroup.getElementsByTagNameNS(NS_P, "childTnLst")[0];
    const shapePar = buildShapeAnimationPar(slideDoc, animation, allocTimeNodeId, "clickEffect");
    timingGroupChildren.appendChild(shapePar);
    clickGroupChildren.appendChild(timingGroup);
    mainChildren.appendChild(clickGroup);
  } else if (animation.start === "withPrevious") {
    // Add to the last click group's last timing group
    const lastClickGroup = getLastClickGroup(mainChildren);
    if (lastClickGroup && !animation.delayMs) {
      const lastTimingGroup = getLastTimingGroup(lastClickGroup);
      if (lastTimingGroup) {
        const timingGroupChildren = getOrCreateChild(
          lastTimingGroup.getElementsByTagNameNS(NS_P, "cTn")[0],
          NS_P,
          "p:childTnLst",
        );
        const shapePar = buildShapeAnimationPar(slideDoc, animation, allocTimeNodeId, "withEffect");
        timingGroupChildren.appendChild(shapePar);
      } else {
        // No timing group yet — create one
        const clickGroupCtn = lastClickGroup.getElementsByTagNameNS(NS_P, "cTn")[0];
        const clickGroupChildren = getOrCreateChild(clickGroupCtn, NS_P, "p:childTnLst");
        const timingGroup = buildTimingGroupPar(slideDoc, allocTimeNodeId);
        const timingGroupChildren = timingGroup.getElementsByTagNameNS(NS_P, "childTnLst")[0];
        const shapePar = buildShapeAnimationPar(slideDoc, animation, allocTimeNodeId, "withEffect");
        timingGroupChildren.appendChild(shapePar);
        clickGroupChildren.appendChild(timingGroup);
      }
    } else {
      // No existing click group or has delay — create new click group
      const clickGroup = buildClickGroupPar(slideDoc, allocTimeNodeId);
      const clickGroupChildren = clickGroup.getElementsByTagNameNS(NS_P, "childTnLst")[0];
      const timingGroup = buildTimingGroupPar(slideDoc, allocTimeNodeId, animation.delayMs ? String(animation.delayMs) : "0");
      const timingGroupChildren = timingGroup.getElementsByTagNameNS(NS_P, "childTnLst")[0];
      const shapePar = buildShapeAnimationPar(slideDoc, animation, allocTimeNodeId, "withEffect");
      timingGroupChildren.appendChild(shapePar);
      clickGroupChildren.appendChild(timingGroup);
      mainChildren.appendChild(clickGroup);
    }
  } else {
    // afterPrevious — add a new timing group in the last click group (or create a new click group)
    const lastClickGroup = getLastClickGroup(mainChildren);
    if (lastClickGroup) {
      const clickGroupCtn = lastClickGroup.getElementsByTagNameNS(NS_P, "cTn")[0];
      const clickGroupChildren = getOrCreateChild(clickGroupCtn, NS_P, "p:childTnLst");
      const timingGroup = buildTimingGroupPar(slideDoc, allocTimeNodeId, "0");
      const timingGroupChildren = timingGroup.getElementsByTagNameNS(NS_P, "childTnLst")[0];
      const shapePar = buildShapeAnimationPar(slideDoc, animation, allocTimeNodeId, "afterEffect");
      timingGroupChildren.appendChild(shapePar);
      clickGroupChildren.appendChild(timingGroup);
    } else {
      // No click group — create one
      const clickGroup = buildClickGroupPar(slideDoc, allocTimeNodeId);
      const clickGroupChildren = clickGroup.getElementsByTagNameNS(NS_P, "childTnLst")[0];
      const timingGroup = buildTimingGroupPar(slideDoc, allocTimeNodeId);
      const timingGroupChildren = timingGroup.getElementsByTagNameNS(NS_P, "childTnLst")[0];
      const shapePar = buildShapeAnimationPar(slideDoc, animation, allocTimeNodeId, "afterEffect");
      timingGroupChildren.appendChild(shapePar);
      clickGroupChildren.appendChild(timingGroup);
      mainChildren.appendChild(clickGroup);
    }
  }

  // Add bldLst entry for entrance animations
  if (isEntrance) {
    addBldLstEntry(slideDoc, animation.targetXmlShapeId);
  }
}

function getLastClickGroup(mainChildren: Element): Element | undefined {
  return Array.from(mainChildren.childNodes)
    .reverse()
    .find(
      (node) => node.nodeType === ELEMENT_NODE && (node as Element).namespaceURI === NS_P && (node as Element).localName === "par",
    ) as Element | undefined;
}

function getLastTimingGroup(clickGroup: Element): Element | undefined {
  const clickGroupCtn = clickGroup.getElementsByTagNameNS(NS_P, "cTn")[0];
  if (!clickGroupCtn) return undefined;
  const childTnLst = getDirectChildByTagName(clickGroupCtn, NS_P, "childTnLst");
  if (!childTnLst) return undefined;
  return Array.from(childTnLst.childNodes)
    .reverse()
    .find(
      (node) => node.nodeType === ELEMENT_NODE && (node as Element).namespaceURI === NS_P && (node as Element).localName === "par",
    ) as Element | undefined;
}

function clearSlideAnimationsInDocument(slideDoc: XMLDocument) {
  const timing = getDirectChildByTagName(slideDoc.documentElement, NS_P, "timing");
  if (timing) {
    timing.parentNode?.removeChild(timing);
  }
}

function setSlideTransitionInDocument(slideDoc: XMLDocument, definition: SlideTransitionDefinition) {
  clearSlideTransitionNodes(slideDoc);

  if (definition.effect === "none") {
    return;
  }

  if (definition.durationMs !== undefined) {
    slideDoc.documentElement.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:mc", NS_MC);
    slideDoc.documentElement.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:p14", NS_P14);
    const ignorable = slideDoc.documentElement.getAttributeNS(NS_MC, "Ignorable") || slideDoc.documentElement.getAttribute("mc:Ignorable") || "";
    const tokens = new Set(ignorable.split(/\s+/).filter(Boolean));
    tokens.add("p14");
    slideDoc.documentElement.setAttributeNS(NS_MC, "mc:Ignorable", Array.from(tokens).join(" "));
  }

  const transitionNode = buildTransitionNode(slideDoc, definition);
  const timing = getDirectChildByTagName(slideDoc.documentElement, NS_P, "timing");
  const extLst = getDirectChildByTagName(slideDoc.documentElement, NS_P, "extLst") || null;
  slideDoc.documentElement.insertBefore(transitionNode, timing || extLst);
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
  pkg.writeText(notesSlideRelsPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIP_TYPE_NOTES_MASTER}" Target="${notesMasterRelativeTarget}"/><Relationship Id="rId2" Type="${RELATIONSHIP_TYPE_SLIDE}" Target="../slides/slide${slideNumber}.xml"/></Relationships>`);

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
  const speakerNotesShape = getSpeakerNotesShape(notesDoc);
  if (speakerNotesShape) {
    return extractTextBody(getTextBody(speakerNotesShape));
  }

  const shapes = Array.from(notesDoc.getElementsByTagNameNS(NS_P, "sp"));
  const noteBlocks = shapes
    .filter((shape) => !EXCLUDED_NOTE_PLACEHOLDER_TYPES.has(getPlaceholderType(shape) || ""))
    .map((shape) => extractTextBody(getTextBody(shape)))
    .filter(Boolean);
  return noteBlocks.join("\n\n").trim();
}

export function extractSlideTransitionFromBase64Presentation(base64: string): SlideTransitionDefinition {
  const pkg = new OpenXmlPackage(base64);
  const { slideDoc } = getFirstSlideDocument(pkg);

  const directTransition = Array.from(slideDoc.documentElement.childNodes).find(
    (node) => node.nodeType === ELEMENT_NODE && (node as Element).namespaceURI === NS_P && (node as Element).localName === "transition",
  ) as Element | undefined;
  const alternateTransition = getAlternateContentTransitionNodes(slideDoc)[0] || null;
  const transition = directTransition || alternateTransition;
  if (!transition) {
    return { effect: "none" };
  }

  const effectDetails = getTransitionEffectDetails(transition);
  const speedValue = transition.getAttribute("spd");
  const duration = transition.getAttributeNS(NS_P14, "dur") || transition.getAttribute("p14:dur") || undefined;
  return slideTransitionDefinitionSchema.parse({
    ...effectDetails,
    speed: speedValue === "med" ? "medium" : (speedValue as SlideTransitionDefinition["speed"] | null) || undefined,
    advanceOnClick: transition.hasAttribute("advClick") ? transition.getAttribute("advClick") !== "0" && transition.getAttribute("advClick") !== "false" : undefined,
    advanceAfterMs: transition.hasAttribute("advTm") ? Number(transition.getAttribute("advTm")) : undefined,
    durationMs: duration ? Number(duration) : undefined,
  });
}

export function setSlideTransitionInBase64Presentation(base64: string, definition: SlideTransitionDefinition) {
  const pkg = new OpenXmlPackage(base64);
  const { slidePath, slideDoc } = getFirstSlideDocument(pkg);
  setSlideTransitionInDocument(slideDoc, definition);
  pkg.writeText(slidePath, serializeXml(slideDoc));
  return pkg.toBase64();
}

export function addSlideAnimationInBase64Presentation(base64: string, animation: SlideAnimationDefinition, shapeIndex: number) {
  const pkg = new OpenXmlPackage(base64);
  const { slidePath, slideDoc } = getFirstSlideDocument(pkg);
  addSlideAnimationInDocument(slideDoc, {
    ...animation,
    targetXmlShapeId: resolveAnimationTargetXmlShapeId(slideDoc, shapeIndex),
  });
  const serialized = serializeXml(slideDoc);
  pkg.writeText(slidePath, serialized);
  return pkg.toBase64();
}

/**
 * Batch-add the same animation to multiple shapes in a single Open XML round-trip.
 * The first shape uses the specified `start` trigger; all subsequent shapes use `withPrevious`.
 */
export function addSlideAnimationBatchInBase64Presentation(
  base64: string,
  animation: SlideAnimationDefinition,
  shapeIndexes: number[],
) {
  const pkg = new OpenXmlPackage(base64);
  const { slidePath, slideDoc } = getFirstSlideDocument(pkg);
  for (let i = 0; i < shapeIndexes.length; i++) {
    const anim: SlideAnimationDefinition = i === 0
      ? animation
      : { ...animation, start: "withPrevious" };
    addSlideAnimationInDocument(slideDoc, {
      ...anim,
      targetXmlShapeId: resolveAnimationTargetXmlShapeId(slideDoc, shapeIndexes[i]),
    });
  }
  pkg.writeText(slidePath, serializeXml(slideDoc));
  return pkg.toBase64();
}

export function clearSlideAnimationsInBase64Presentation(base64: string) {
  const pkg = new OpenXmlPackage(base64);
  const { slidePath, slideDoc } = getFirstSlideDocument(pkg);
  clearSlideAnimationsInDocument(slideDoc);
  pkg.writeText(slidePath, serializeXml(slideDoc));
  return pkg.toBase64();
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

/** Wraps an Office.js error thrown during the slide round-trip with context about which step failed. */
function wrapRoundTripError(error: unknown, step: string, slideIndex: number): Error {
  const code = error instanceof Error ? (error as { code?: string }).code : undefined;
  const msg = error instanceof Error ? error.message : String(error);
  const codeLabel = code ? ` [${code}]` : "";
  return new Error(`Slide ${slideIndex + 1} round-trip failed while ${step} slide: ${msg}${codeLabel}`);
}

export interface RoundTripOptions {
  /** Queue extra loads on the source slide proxy before the first sync.
   *  For example, loading shapes so they're available to the mutate callback. */
  preload?: (sourceSlide: PowerPoint.Slide) => void;
}

export async function replaceSlideWithMutatedOpenXml(
  context: PowerPoint.RequestContext,
  slideIndex: number,
  mutate: (base64: string, sourceSlide: PowerPoint.Slide) => string,
  options?: RoundTripOptions,
): Promise<OpenXmlRoundTripResult> {
  if (!isPowerPointRequirementSetSupported("1.8")) {
    throw new Error("PowerPoint Open XML slide round-tripping requires PowerPointApi 1.8.");
  }

  // ── Phase 1: Load slide IDs + export source slide in one sync ──────────
  // Use getItemAt() to obtain a proxy for the source slide (and the slide
  // before it) so that exportAsBase64() can be batched with the ID loads.
  const slides = context.presentation.slides;
  slides.load("items/id");
  const sourceSlide = slides.getItemAt(slideIndex);
  sourceSlide.load("id");
  const exported = sourceSlide.exportAsBase64();
  let targetSlideProxy: PowerPoint.Slide | undefined;
  if (slideIndex > 0) {
    targetSlideProxy = slides.getItemAt(slideIndex - 1);
    targetSlideProxy.load("id");
  }
  // Allow callers to piggyback extra loads (e.g. shapes) onto this first sync.
  options?.preload?.(sourceSlide);
  try {
    await context.sync();
  } catch (e) {
    // Index-out-of-range from getItemAt or export failure both surface here.
    throw wrapRoundTripError(e, "loading and exporting", slideIndex);
  }

  if (slideIndex < 0 || slideIndex >= slides.items.length) {
    throw new Error(invalidSlideIndexMessage(slideIndex, slides.items.length));
  }

  const sourceSlideId = sourceSlide.id;
  const previousSlideIds = slides.items.map((slide) => slide.id).filter(Boolean);
  const targetSlideId = targetSlideProxy?.id;

  // ── Phase 2: Mutate the exported package ──────────────────────────────
  const mutated = mutate(exported.value, sourceSlide);

  // ── Phase 3: Insert mutated slide, delete original, load final state ──
  // Batch all three operations into a single sync to minimize IPC round-trips.
  // Office.js executes queued operations sequentially within a sync, so the
  // load will reflect the state after both insert and delete.
  context.presentation.insertSlidesFromBase64(mutated, {
    formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
    ...(targetSlideId ? { targetSlideId } : {}),
  });
  sourceSlide.delete();
  const finalSlides = context.presentation.slides;
  finalSlides.load("items/id");
  try {
    await context.sync();
  } catch (e) {
    throw wrapRoundTripError(e, "inserting mutated and deleting original", slideIndex);
  }

  // The replacement slide is the one whose ID wasn't in the original slide collection.
  const replacementSlide = finalSlides.items.find((slide) => !previousSlideIds.includes(slide.id));
  if (!replacementSlide) {
    throw new Error("Failed to locate the replacement slide after Open XML round-trip insertion.");
  }
  const finalSlideIndex = finalSlides.items.findIndex((slide) => slide.id === replacementSlide.id);

  return openXmlRoundTripResultSchema.parse({
    originalSlideId: sourceSlideId,
    replacementSlideId: replacementSlide.id,
    finalSlideIndex,
  });
}

export async function exportSlideAsBase64(context: PowerPoint.RequestContext, slideIndex: number) {
  if (!isPowerPointRequirementSetSupported("1.8")) {
    throw new Error("PowerPoint Open XML slide export requires PowerPointApi 1.8.");
  }

  const slides = context.presentation.slides;
  slides.load("items");
  const exported = slides.getItemAt(slideIndex).exportAsBase64();
  await context.sync();
  if (slideIndex < 0 || slideIndex >= slides.items.length) {
    throw new Error(invalidSlideIndexMessage(slideIndex, slides.items.length));
  }

  return exported.value;
}

/**
 * Extract the raw timing XML (`<p:timing>`) from a base64-encoded single-slide
 * PPTX, formatted for human inspection. Useful for debugging animation structures.
 * Returns the serialized XML string, or null if the slide has no timing element.
 */
export function extractTimingXmlFromBase64Presentation(base64: string): string | null {
  const pkg = new OpenXmlPackage(base64);
  const slidePath = getFirstSlidePath(pkg);
  const slideXml = pkg.readText(slidePath);

  const doc = parseXml(slideXml);
  const slideEl = doc.documentElement;
  const timing = getDirectChildByTagName(slideEl, NS_P, "timing");
  if (!timing) return null;

  return new XMLSerializer().serializeToString(timing);
}

/**
 * Extract the raw build-list XML (`<p:bldLst>`) from a base64-encoded single-slide
 * PPTX. Returns the serialized XML string, or null if there is no build list.
 */
export function extractBuildListXmlFromBase64Presentation(base64: string): string | null {
  const pkg = new OpenXmlPackage(base64);
  const slidePath = getFirstSlidePath(pkg);
  const slideXml = pkg.readText(slidePath);

  const doc = parseXml(slideXml);
  const slideEl = doc.documentElement;

  // bldLst can be at the slide level (p:sld/p:bldLst) — though typically it's
  // under the timing element. Search both locations.
  const timing = getDirectChildByTagName(slideEl, NS_P, "timing");
  if (timing) {
    const bldLst = getDirectChildByTagName(timing, NS_P, "bldLst");
    if (bldLst) return new XMLSerializer().serializeToString(bldLst);
  }

  // Fallback: directly under the slide element (rare but valid)
  const topBldLst = getDirectChildByTagName(slideEl, NS_P, "bldLst");
  if (topBldLst) return new XMLSerializer().serializeToString(topBldLst);

  return null;
}

/**
 * Extract the full slide XML from a base64-encoded single-slide PPTX.
 * Useful for comprehensive debugging of slide structure.
 */
export function extractFullSlideXmlFromBase64Presentation(base64: string): string | null {
  const pkg = new OpenXmlPackage(base64);
  const slidePath = getFirstSlidePath(pkg);
  return pkg.readText(slidePath);
}

// ── Animation summary extraction ──────────────────────────────────────────

interface AnimationSummaryEntry {
  /** Position in the sequence (1-based) */
  order: number;
  /** Start trigger: "onClick", "withPrevious", or "afterPrevious" */
  start: string;
  /** Preset class: "entr", "emph", "exit", "path", or null for custom */
  presetClass: string | null;
  /** Preset ID number */
  presetID: number | null;
  /** Preset subtype number */
  presetSubtype: number | null;
  /** Readable animation type name (e.g. "appear", "fade", "flyIn") */
  typeName: string | null;
  /** Target shape XML id (spid) */
  targetShapeId: string | null;
  /** Duration in ms */
  durationMs: number | null;
  /** Delay in ms */
  delayMs: number | null;
  /** Group ID */
  grpId: string | null;
}

interface AnimationSummary {
  /** Total number of animations */
  animationCount: number;
  /** The animation entries in sequence order */
  animations: AnimationSummaryEntry[];
  /** Shape IDs in the build list (entrance-animated shapes that start hidden) */
  buildListShapeIds: string[];
}

const PRESET_ID_TO_NAME: Record<number, Record<string, string>> = {
  // entrance
  1: { entr: "appear" },
  2: { entr: "flyIn" },
  10: { entr: "fade" },
  22: { entr: "wipe" },
  23: { entr: "zoomIn" },
  30: { entr: "floatIn" },
  34: { entr: "riseUp" },
  37: { entr: "growAndTurn" },
  42: { entr: "peekIn" },
  // emphasis
  54: { emph: "changeFillColor" },
  60: { emph: "changeLineColor" },
  70: { emph: "complementaryColor" },
};

function resolveAnimationTypeName(presetClass: string | null, presetID: number | null): string | null {
  if (presetClass === null || presetID === null) return null;
  const classMap = PRESET_ID_TO_NAME[presetID];
  if (!classMap) return null;
  return classMap[presetClass] || null;
}

const NODE_TYPE_TO_START: Record<string, string> = {
  clickEffect: "onClick",
  withEffect: "withPrevious",
  afterEffect: "afterPrevious",
};

export function extractSlideAnimationSummaryFromBase64Presentation(base64: string): AnimationSummary {
  const pkg = new OpenXmlPackage(base64);
  const slidePath = getFirstSlidePath(pkg);
  const slideXml = pkg.readText(slidePath);
  const doc = parseXml(slideXml);

  const animations: AnimationSummaryEntry[] = [];

  // Find all cTn elements with a nodeType attribute (these are per-shape animation wrappers)
  const allCTns = doc.getElementsByTagNameNS(NS_P, "cTn");
  let order = 0;
  for (let i = 0; i < allCTns.length; i++) {
    const cTn = allCTns[i];
    const nodeType = cTn.getAttribute("nodeType");
    if (!nodeType || !NODE_TYPE_TO_START[nodeType]) continue;

    order++;
    const presetClass = cTn.getAttribute("presetClass");
    const presetIDStr = cTn.getAttribute("presetID");
    const presetSubtypeStr = cTn.getAttribute("presetSubtype");
    const grpId = cTn.getAttribute("grpId");
    const presetID = presetIDStr ? Number.parseInt(presetIDStr, 10) : null;
    const presetSubtype = presetSubtypeStr ? Number.parseInt(presetSubtypeStr, 10) : null;

    // Find target shape ID — look for spTgt under this cTn's subtree
    let targetShapeId: string | null = null;
    const spTgts = cTn.getElementsByTagNameNS(NS_P, "spTgt");
    if (spTgts.length > 0) {
      targetShapeId = spTgts[0].getAttribute("spid");
    }

    // Get duration from the cTn itself (emphasis/motion) or from an inner cBhvr cTn
    let durationMs: number | null = null;
    const durStr = cTn.getAttribute("dur");
    if (durStr && durStr !== "indefinite") {
      durationMs = Number.parseInt(durStr, 10);
    }
    // If no dur on outer cTn, check first inner cBhvr's cTn (for entrance animations)
    if (durationMs === null) {
      const childTnLst = getDirectChildByTagName(cTn, NS_P, "childTnLst");
      if (childTnLst) {
        const children = Array.from(childTnLst.childNodes).filter(
          (n) => n.nodeType === ELEMENT_NODE,
        ) as Element[];
        for (const child of children) {
          const innerCBhvrs = child.getElementsByTagNameNS(NS_P, "cBhvr");
          for (let j = 0; j < innerCBhvrs.length; j++) {
            const innerCTn = getDirectChildByTagName(innerCBhvrs[j], NS_P, "cTn");
            if (innerCTn) {
              const innerDur = innerCTn.getAttribute("dur");
              if (innerDur && innerDur !== "1" && innerDur !== "indefinite") {
                durationMs = Number.parseInt(innerDur, 10);
                break;
              }
            }
          }
          if (durationMs !== null) break;
        }
      }
    }

    // Get delay from stCondLst
    let delayMs: number | null = null;
    const stCondLst = getDirectChildByTagName(cTn, NS_P, "stCondLst");
    if (stCondLst) {
      const conds = stCondLst.getElementsByTagNameNS(NS_P, "cond");
      for (let j = 0; j < conds.length; j++) {
        const delay = conds[j].getAttribute("delay");
        if (delay && delay !== "0" && delay !== "indefinite" && !conds[j].hasAttribute("evt")) {
          delayMs = Number.parseInt(delay, 10);
        }
      }
    }

    animations.push({
      order,
      start: NODE_TYPE_TO_START[nodeType],
      presetClass,
      presetID,
      presetSubtype,
      typeName: resolveAnimationTypeName(presetClass, presetID),
      targetShapeId,
      durationMs,
      delayMs,
      grpId,
    });
  }

  // Extract build list
  const buildListShapeIds: string[] = [];
  const timing = getDirectChildByTagName(doc.documentElement, NS_P, "timing");
  if (timing) {
    const bldLst = getDirectChildByTagName(timing, NS_P, "bldLst");
    if (bldLst) {
      const bldPs = bldLst.getElementsByTagNameNS(NS_P, "bldP");
      for (let i = 0; i < bldPs.length; i++) {
        const spid = bldPs[i].getAttribute("spid");
        if (spid) buildListShapeIds.push(spid);
      }
    }
  }

  return {
    animationCount: animations.length,
    animations,
    buildListShapeIds,
  };
}
