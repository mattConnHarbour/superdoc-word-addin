import { DOMParser, type Document, type Element, type Node } from "@xmldom/xmldom";
import { strFromU8, unzipSync } from "fflate";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_RELS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const WE = "http://schemas.microsoft.com/office/webextensions/webextension/2010/11";
const TP = "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11";
const ADDIN = "f3925cb5-51d4-432d-b93c-b72f5eca98e3";
const CORE = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const APP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
const VOLATILE_CORE = new Set(["lastModifiedBy", "revision"]);
const VOLATILE_APP = new Set(["TotalTime", "Pages", "Words", "Characters", "CharactersWithSpaces", "Lines", "Paragraphs"]);
const RSID_ATTRIBUTES = new Set(["rsidR", "rsidRPr", "rsidDel", "rsidP", "rsidSect", "rsidRDefault"]);

function isElement(node: Node): node is Element {
  return node.nodeType === 1;
}

function relationshipTarget(part: string, relationship: Element): string {
  const owner = part.replace(/(^|\/)_rels\//u, "$1").replace(/\.rels$/u, "");
  return new URL(relationship.getAttribute("Target") ?? "", `https://package/${owner}`).pathname.slice(1);
}

function ignoredTaskPaneParts(documents: Map<string, Document>): Set<string> {
  const ignored = new Set<string>();
  for (const [name, document] of documents) {
    if (!name.startsWith("word/webextensions/")) continue;
    const references = Array.from(document.getElementsByTagNameNS(WE, "reference"));
    if (references.some((reference) => reference.getAttribute("id")?.toLowerCase() === ADDIN)) ignored.add(name);
  }
  const paneName = "word/webextensions/taskpanes.xml";
  const relationshipName = "word/webextensions/_rels/taskpanes.xml.rels";
  const pane = documents.get(paneName);
  const relationships = documents.get(relationshipName);
  if (!pane || !relationships) return ignored;
  const ignoredIds = new Set(Array.from(relationships.getElementsByTagNameNS(RELS, "Relationship"))
    .filter((relationship) => ignored.has(relationshipTarget(relationshipName, relationship)))
    .map((relationship) => relationship.getAttribute("Id")));
  if (!ignoredIds.size) return ignored;
  for (const taskpane of Array.from(pane.getElementsByTagNameNS(TP, "taskpane"))) {
    const references = Array.from(taskpane.getElementsByTagNameNS(TP, "webextensionref"));
    if (references.some((reference) => ignoredIds.has(reference.getAttributeNS(OFFICE_RELS, "id")))) taskpane.parentNode?.removeChild(taskpane);
  }
  if (!pane.getElementsByTagNameNS(TP, "taskpane").length) {
    ignored.add(paneName);
    ignored.add(relationshipName);
  }
  return ignored;
}

function volatileElement(element: Element, part: string): boolean {
  if (element.namespaceURI === W && (element.localName === "rsids" || element.localName === "rsid")) return true;
  if (part === "docProps/core.xml") {
    return (element.namespaceURI === CORE && VOLATILE_CORE.has(element.localName ?? ""))
      || (element.namespaceURI === "http://purl.org/dc/terms/" && element.localName === "modified");
  }
  return part === "docProps/app.xml" && element.namespaceURI === APP && VOLATILE_APP.has(element.localName ?? "");
}

function canonicalElement(element: Element, part: string, ignored: Set<string>): string {
  if (volatileElement(element, part)) return "";
  if (element.namespaceURI === CT && element.localName === "Override" && ignored.has((element.getAttribute("PartName") ?? "").replace(/^\//u, ""))) return "";
  if (element.namespaceURI === RELS && element.localName === "Relationship"
    && element.getAttribute("TargetMode") !== "External" && ignored.has(relationshipTarget(part, element))) return "";
  const attributes = Array.from(element.attributes)
    .filter((attribute) => attribute.namespaceURI !== "http://www.w3.org/2000/xmlns/"
      && !(attribute.namespaceURI === W && RSID_ATTRIBUTES.has(attribute.localName ?? ""))
      && !(part === "_rels/.rels" && element.namespaceURI === RELS && attribute.name === "Id"))
    .map((attribute) => [`{${attribute.namespaceURI ?? ""}}${attribute.localName ?? attribute.name}`, attribute.value])
    .sort(([left = ""], [right = ""]) => left.localeCompare(right));
  const children = Array.from(element.childNodes).map((node) => {
    if (isElement(node)) return canonicalElement(node, part, ignored);
    return node.nodeType === 3 || node.nodeType === 4 ? JSON.stringify(node.nodeValue) : "";
  }).filter(Boolean);
  if (element.namespaceURI === RELS || element.namespaceURI === CT) children.sort();
  return JSON.stringify([`{${element.namespaceURI ?? ""}}${element.localName}`, attributes, children]);
}

export async function digestBytes(bytes: Uint8Array<ArrayBuffer>): Promise<string> {
  const hash = await crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(hash), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

export async function packageRevision(bytes: Uint8Array<ArrayBuffer>): Promise<string> {
  const parts = unzipSync(bytes);
  const documents = new Map<string, Document>();
  for (const [name, content] of Object.entries(parts)) {
    if (!name.endsWith(".xml") && !name.endsWith(".rels")) continue;
    const document = new DOMParser({ onError: () => { throw new Error("Word exported an invalid XML part."); } }).parseFromString(strFromU8(content), "application/xml");
    documents.set(name, document);
  }
  const ignored = ignoredTaskPaneParts(documents);
  const hashes: string[] = [];
  for (const name of Object.keys(parts).sort()) {
    if (name.endsWith("/") || ignored.has(name)) continue;
    const part = parts[name];
    if (!part) continue;
    const root = documents.get(name)?.documentElement;
    const content = root ? new TextEncoder().encode(canonicalElement(root, name, ignored)) : new Uint8Array(part);
    hashes.push(`${name}:${await digestBytes(content)}`);
  }
  return digestBytes(new TextEncoder().encode(hashes.join("\n")));
}
