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
const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const RELATIONSHIP_TYPE_NOTES_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const RELATIONSHIP_TYPE_NOTES_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
const CONTENT_TYPE_NOTES_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";

const EXCLUDED_NOTE_PLACEHOLDER_TYPES = new Set(["sldImg", "hdr", "dt", "ftr", "sldNum"]);

export interface SlideTransitionDefinition {
  effect: "none" | "cut" | "fade" | "dissolve" | "random" | "randomBar" | "push" | "wipe" | "split" | "cover" | "pull" | "zoom";
  speed?: "slow" | "medium" | "fast";
  advanceOnClick?: boolean;
  advanceAfterMs?: number;
  durationMs?: number;
  direction?: "left" | "right" | "up" | "down" | "horizontal" | "vertical" | "in" | "out";
  orientation?: "horizontal" | "vertical";
  throughBlack?: boolean;
}

export interface SlideAnimationDefinition {
  type: "motionPath" | "scale" | "rotate";
  start: "onClick" | "withPrevious" | "afterPrevious";
  durationMs?: number;
  delayMs?: number;
  repeatCount?: number;
  shapeId: string;
  path?: string;
  pathOrigin?: "parent" | "layout";
  pathEditMode?: "relative" | "fixed";
  scaleXPercent?: number;
  scaleYPercent?: number;
  angleDegrees?: number;
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

function getTransitionEffectDetails(transition: Element) {
  const effect = Array.from(transition.childNodes).find((node) => node.nodeType === Node.ELEMENT_NODE) as Element | undefined;
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
  alternateContent.append(choice, fallback);
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

function getOrCreateChild(parent: Element, namespace: string, qualifiedName: string) {
  const localName = qualifiedName.split(":").pop() || qualifiedName;
  const existing = Array.from(parent.childNodes).find(
    (node) => node.nodeType === Node.ELEMENT_NODE && (node as Element).namespaceURI === namespace && (node as Element).localName === localName,
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
  const timing = slideDoc.documentElement.getElementsByTagNameNS(NS_P, "timing")[0] || buildTimingRoot(slideDoc);
  if (!timing.parentNode) {
    const extLst = slideDoc.documentElement.getElementsByTagNameNS(NS_P, "extLst")[0] || null;
    slideDoc.documentElement.insertBefore(timing, extLst);
  }
  return timing;
}

function getOrCreateMainSequence(slideDoc: XMLDocument) {
  const timing = getOrCreateTimingRoot(slideDoc);
  const tnLst = getOrCreateChild(timing, NS_P, "p:tnLst");
  const rootPar = getOrCreateChild(tnLst, NS_P, "p:par");
  const rootCtn = getOrCreateChild(rootPar, NS_P, "p:cTn");
  if (!rootCtn.getAttribute("id")) rootCtn.setAttribute("id", "1");
  if (!rootCtn.getAttribute("dur")) rootCtn.setAttribute("dur", "indefinite");
  if (!rootCtn.getAttribute("nodeType")) rootCtn.setAttribute("nodeType", "tmRoot");
  const rootChildTnLst = getOrCreateChild(rootCtn, NS_P, "p:childTnLst");
  let mainSeq = Array.from(rootChildTnLst.childNodes).find(
    (node) => node.nodeType === Node.ELEMENT_NODE && (node as Element).namespaceURI === NS_P && (node as Element).localName === "seq" && (node as Element).getElementsByTagNameNS(NS_P, "cTn")[0]?.getAttribute("nodeType") === "mainSeq",
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

  return mainSeq;
}

function buildStartConditions(doc: XMLDocument, start: SlideAnimationDefinition["start"], delayMs?: number) {
  const stCondLst = doc.createElementNS(NS_P, "p:stCondLst");
  const cond = doc.createElementNS(NS_P, "p:cond");
  if (start === "onClick") {
    cond.setAttribute("evt", "onClick");
  }
  if (delayMs && delayMs > 0) {
    cond.setAttribute("delay", String(delayMs));
  }
  stCondLst.appendChild(cond);
  return stCondLst;
}

function buildTargetElement(doc: XMLDocument, shapeId: string) {
  const target = doc.createElementNS(NS_P, "p:tgtEl");
  const spTarget = doc.createElementNS(NS_P, "p:spTgt");
  spTarget.setAttribute("spid", shapeId);
  target.appendChild(spTarget);
  return target;
}

function buildCommonBehavior(doc: XMLDocument, animation: SlideAnimationDefinition, allocTimeNodeId: () => string) {
  const cBhvr = doc.createElementNS(NS_P, "p:cBhvr");
  const cTn = doc.createElementNS(NS_P, "p:cTn");
  cTn.setAttribute("id", allocTimeNodeId());
  cTn.setAttribute("dur", String(animation.durationMs ?? 1000));
  cTn.setAttribute("fill", "hold");
  if (animation.repeatCount !== undefined && animation.repeatCount > 0) {
    cTn.setAttribute("repeatCount", String(animation.repeatCount));
  }
  cBhvr.append(cTn, buildTargetElement(doc, animation.shapeId));
  return cBhvr;
}

function buildAnimationNode(doc: XMLDocument, animation: SlideAnimationDefinition, allocTimeNodeId: () => string) {
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
    attrNameList.append(attrX, attrY);
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

function buildAnimationContainer(doc: XMLDocument, animation: SlideAnimationDefinition, allocTimeNodeId: () => string) {
  if (animation.start === "onClick") {
    const seq = doc.createElementNS(NS_P, "p:seq");
    const cTn = doc.createElementNS(NS_P, "p:cTn");
    cTn.setAttribute("id", allocTimeNodeId());
    cTn.setAttribute("dur", "indefinite");
    cTn.setAttribute("nodeType", "clickEffect");
    cTn.appendChild(buildStartConditions(doc, animation.start, animation.delayMs));
    const childTnLst = doc.createElementNS(NS_P, "p:childTnLst");
    const par = doc.createElementNS(NS_P, "p:par");
    const parCtn = doc.createElementNS(NS_P, "p:cTn");
    parCtn.setAttribute("id", allocTimeNodeId());
    parCtn.setAttribute("dur", String(animation.durationMs ?? 1000));
    parCtn.setAttribute("fill", "hold");
    const parChildren = doc.createElementNS(NS_P, "p:childTnLst");
    parChildren.appendChild(buildAnimationNode(doc, animation, allocTimeNodeId));
    parCtn.appendChild(parChildren);
    par.appendChild(parCtn);
    childTnLst.appendChild(par);
    cTn.appendChild(childTnLst);
    seq.appendChild(cTn);
    return seq;
  }

  const par = doc.createElementNS(NS_P, "p:par");
  const cTn = doc.createElementNS(NS_P, "p:cTn");
  cTn.setAttribute("id", allocTimeNodeId());
  cTn.setAttribute("dur", String(animation.durationMs ?? 1000));
  cTn.setAttribute("fill", "hold");
  if (animation.start === "withPrevious" && animation.delayMs) {
    cTn.appendChild(buildStartConditions(doc, "withPrevious", animation.delayMs));
  }
  if (animation.repeatCount !== undefined && animation.repeatCount > 0) {
    cTn.setAttribute("repeatCount", String(animation.repeatCount));
  }
  const childTnLst = doc.createElementNS(NS_P, "p:childTnLst");
  childTnLst.appendChild(buildAnimationNode(doc, animation, allocTimeNodeId));
  cTn.appendChild(childTnLst);
  par.appendChild(cTn);
  return par;
}

function addSlideAnimationInDocument(slideDoc: XMLDocument, animation: SlideAnimationDefinition) {
  const allocTimeNodeId = createTimeNodeIdAllocator(slideDoc);
  const mainSeq = getOrCreateMainSequence(slideDoc);
  const rootChildTnLst = mainSeq.parentElement || mainSeq.parentNode as Element;
  if (animation.start === "onClick") {
    rootChildTnLst.appendChild(buildAnimationContainer(slideDoc, animation, allocTimeNodeId));
    return;
  }

  const mainCtn = mainSeq.getElementsByTagNameNS(NS_P, "cTn")[0];
  const mainChildren = getOrCreateChild(mainCtn, NS_P, "p:childTnLst");
  if (animation.start === "withPrevious" && !animation.delayMs) {
    const lastPar = Array.from(mainChildren.childNodes)
      .reverse()
      .find((node) => node.nodeType === Node.ELEMENT_NODE && (node as Element).namespaceURI === NS_P && (node as Element).localName === "par") as Element | undefined;
    if (lastPar) {
      const lastParCtn = getOrCreateChild(lastPar, NS_P, "p:cTn");
      const lastParChildren = getOrCreateChild(lastParCtn, NS_P, "p:childTnLst");
      lastParChildren.appendChild(buildAnimationNode(slideDoc, animation, allocTimeNodeId));
      return;
    }
  }

  mainChildren.appendChild(buildAnimationContainer(slideDoc, animation, allocTimeNodeId));
}

function clearSlideAnimationsInDocument(slideDoc: XMLDocument) {
  const timing = slideDoc.documentElement.getElementsByTagNameNS(NS_P, "timing")[0];
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
  const timing = slideDoc.documentElement.getElementsByTagNameNS(NS_P, "timing")[0];
  const extLst = slideDoc.documentElement.getElementsByTagNameNS(NS_P, "extLst")[0] || null;
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

export function extractSlideTransitionFromBase64Presentation(base64: string): SlideTransitionDefinition {
  const pkg = new OpenXmlPackage(base64);
  const { slideDoc } = getFirstSlideDocument(pkg);

  const directTransition = Array.from(slideDoc.documentElement.childNodes).find(
    (node) => node.nodeType === Node.ELEMENT_NODE && (node as Element).namespaceURI === NS_P && (node as Element).localName === "transition",
  ) as Element | undefined;
  const alternateTransition = getAlternateContentTransitionNodes(slideDoc)[0] || null;
  const transition = directTransition || alternateTransition;
  if (!transition) {
    return { effect: "none" };
  }

  const effectDetails = getTransitionEffectDetails(transition);
  const speedValue = transition.getAttribute("spd");
  const duration = transition.getAttributeNS(NS_P14, "dur") || transition.getAttribute("p14:dur") || undefined;
  return {
    ...effectDetails,
    speed: speedValue === "med" ? "medium" : (speedValue as SlideTransitionDefinition["speed"] | null) || undefined,
    advanceOnClick: transition.hasAttribute("advClick") ? transition.getAttribute("advClick") !== "0" && transition.getAttribute("advClick") !== "false" : undefined,
    advanceAfterMs: transition.hasAttribute("advTm") ? Number(transition.getAttribute("advTm")) : undefined,
    durationMs: duration ? Number(duration) : undefined,
  };
}

export function setSlideTransitionInBase64Presentation(base64: string, definition: SlideTransitionDefinition) {
  const pkg = new OpenXmlPackage(base64);
  const { slidePath, slideDoc } = getFirstSlideDocument(pkg);
  setSlideTransitionInDocument(slideDoc, definition);
  pkg.writeText(slidePath, serializeXml(slideDoc));
  return pkg.toBase64();
}

export function addSlideAnimationInBase64Presentation(base64: string, animation: SlideAnimationDefinition) {
  const pkg = new OpenXmlPackage(base64);
  const { slidePath, slideDoc } = getFirstSlideDocument(pkg);
  addSlideAnimationInDocument(slideDoc, animation);
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
