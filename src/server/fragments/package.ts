import { strFromU8 } from "fflate";
import type { Document, Element } from "@xmldom/xmldom";
import { body, CT, elements, mainDocument, parseXml, PKG, R, REL, serialize } from "../../ooxml.js";

interface PackageState {
  parts: Record<string, Uint8Array>;
  selected: Element[];
  packageXml: Document;
  contentTypes: Document;
  pending: string[];
  xml: Map<string, string>;
}

function selectedDocument(parts: Record<string, Uint8Array>, selected: Element[]): string {
  const document = mainDocument(parts);
  const documentBody = body(document);
  while (documentBody.firstChild) documentBody.removeChild(documentBody.firstChild);
  for (const block of selected) documentBody.appendChild(document.importNode(block, true));
  return serialize(document);
}

function contentType(state: PackageState, name: string): string {
  if (name.endsWith(".rels")) return "application/vnd.openxmlformats-package.relationships+xml";
  const override = Array.from(state.contentTypes.getElementsByTagNameNS(CT, "Override")).find((entry) => entry.getAttribute("PartName") === `/${name}`);
  const fallback = Array.from(state.contentTypes.getElementsByTagNameNS(CT, "Default")).find((entry) => entry.getAttribute("Extension") === name.split(".").at(-1));
  const type = override?.getAttribute("ContentType") ?? fallback?.getAttribute("ContentType");
  if (!type) throw new Error(`Missing fragment content type: ${name}`);
  return type;
}

function selectedRelationshipIds(selected: Element[]): Set<string> {
  const ids = new Set<string>();
  for (const block of selected) {
    for (const node of [block, ...Array.from(block.getElementsByTagName("*"))]) {
      for (const attr of Array.from(node.attributes)) if (attr.namespaceURI === R) ids.add(attr.value);
    }
  }
  return ids;
}

function queueRelationships(state: PackageState, name: string): void {
  const relationshipName = name.replace(/([^/]+)$/u, "_rels/$1.rels");
  const bytes = state.parts[relationshipName];
  if (!bytes) return;
  const parsed = parseXml(strFromU8(bytes));
  if (name === "word/document.xml") {
    const ids = selectedRelationshipIds(state.selected);
    for (const relationship of Array.from(parsed.getElementsByTagNameNS(REL, "Relationship"))) {
      const type = relationship.getAttribute("Type")?.split("/").at(-1);
      if (!ids.has(relationship.getAttribute("Id") ?? "") && !["styles", "numbering", "theme", "fontTable"].includes(type ?? "")) relationship.parentNode?.removeChild(relationship);
    }
  }
  state.xml.set(relationshipName, serialize(parsed));
  state.pending.push(relationshipName);
}

function queueRelationshipTargets(state: PackageState, name: string, root: Element): void {
  if (!name.endsWith(".rels")) return;
  for (const relationship of elements(root)) {
    if (relationship.getAttribute("TargetMode") === "External") continue;
    const owner = name.replace(/(^|\/)_rels\//u, "$1").replace(/\.rels$/u, "");
    state.pending.push(new URL(relationship.getAttribute("Target") ?? "", `https://package/${owner}`).pathname.slice(1));
  }
}

function appendPart(state: PackageState, name: string): void {
  const bytes = state.parts[name];
  const xml = state.xml.get(name);
  if (!bytes && !xml) throw new Error(`Missing fragment dependency: ${name}`);
  const type = contentType(state, name);
  const part = state.packageXml.createElementNS(PKG, "pkg:part");
  part.setAttributeNS(PKG, "pkg:name", `/${name}`);
  part.setAttributeNS(PKG, "pkg:contentType", type);
  const isXml = Boolean(xml) || name.endsWith(".xml") || name.endsWith(".rels") || type.endsWith("+xml");
  const data = state.packageXml.createElementNS(PKG, isXml ? "pkg:xmlData" : "pkg:binaryData");
  if (isXml) {
    const root = parseXml(xml ?? strFromU8(bytes ?? new Uint8Array())).documentElement;
    if (!root) throw new Error("Empty XML dependency.");
    data.appendChild(state.packageXml.importNode(root, true));
    queueRelationshipTargets(state, name, root);
  } else data.appendChild(state.packageXml.createTextNode(Buffer.from(bytes ?? []).toString("base64")));
  part.appendChild(data);
  state.packageXml.documentElement?.appendChild(part);
  queueRelationships(state, name);
}

export function buildFragmentPackage(parts: Record<string, Uint8Array>, selected: Element[]): string {
  const types = parts["[Content_Types].xml"];
  if (!types) throw new Error("Missing DOCX content types.");
  const state: PackageState = {
    parts, selected, packageXml: parseXml(`<pkg:package xmlns:pkg="${PKG}"/>`), contentTypes: parseXml(strFromU8(types)),
    pending: ["word/document.xml", "_rels/.rels"],
    xml: new Map([
      ["word/document.xml", selectedDocument(parts, selected)],
      ["_rels/.rels", `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ]),
  };
  const included = new Set<string>();
  while (state.pending.length) {
    const name = state.pending.pop();
    if (!name || included.has(name)) continue;
    included.add(name);
    appendPart(state, name);
  }
  const result = serialize(state.packageXml);
  if (new TextEncoder().encode(result).length > 2 * 1024 * 1024) throw new Error("A changed fragment exceeds the 2 MB patch limit. Try a smaller edit.");
  return result;
}
