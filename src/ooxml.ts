import { DOMParser, XMLSerializer, type Document, type Element, type Node } from "@xmldom/xmldom";
import { strFromU8, unzipSync } from "fflate";
import { digestBytes } from "./digest.js";

export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
export const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
export const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const VOLATILE = new Set(["rsidR", "rsidRPr", "rsidDel", "rsidP", "rsidSect", "rsidRDefault"]);
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

function isElement(node: Node): node is Element {
  return node.nodeType === 1;
}

export function elements(node: Node): Element[] {
  return Array.from(node.childNodes).filter(isElement);
}

export function parseXml(xml: string): Document {
  return new DOMParser({ onError: () => { throw new Error("The document contains unreadable XML."); } }).parseFromString(xml, "application/xml");
}

export function serialize(node: Node): string {
  return new XMLSerializer().serializeToString(node);
}

export function canonical(node: Element): string {
  if (node.namespaceURI === W && (node.localName === "proofErr" || node.localName === "lastRenderedPageBreak")) return "";
  const attrs = Array.from(node.attributes).filter((attr) => attr.namespaceURI !== "http://www.w3.org/2000/xmlns/"
    && !(attr.namespaceURI === W && VOLATILE.has(attr.localName ?? ""))
    && !(attr.namespaceURI === "http://www.w3.org/XML/1998/namespace" && attr.localName === "space" && node.namespaceURI === W && node.localName === "t")
    && !(attr.namespaceURI === W14 && (attr.localName === "paraId" || attr.localName === "textId")))
    .map((attr) => [`{${attr.namespaceURI ?? ""}}${attr.localName ?? attr.name}`, attr.value])
    .sort(([a = ""], [b = ""]) => a.localeCompare(b));
  const children: string[] = [];
  let run: { format: string; text: string } | undefined;
  const flush = () => {
    if (run) children.push(JSON.stringify(["text-run", run.format, run.text]));
    run = undefined;
  };
  for (const child of Array.from(node.childNodes)) {
    if (isElement(child)) {
      if (child.namespaceURI === W && (child.localName === "proofErr" || child.localName === "lastRenderedPageBreak")) continue;
      const content = elements(child);
      if (child.namespaceURI === W && child.localName === "r" && !Array.from(child.attributes).some((attr) => attr.namespaceURI !== "http://www.w3.org/2000/xmlns/" && !(attr.namespaceURI === W && VOLATILE.has(attr.localName ?? "")))
        && content.some((element) => element.localName === "t") && content.every((element) => element.namespaceURI === W && (element.localName === "rPr" || element.localName === "t"))) {
        const properties = content.find((element) => element.localName === "rPr");
        const format = properties && elements(properties).length ? canonical(properties) : "";
        const text = content.filter((element) => element.localName === "t").map((element) => element.textContent ?? "").join("");
        if (run?.format === format) run.text += text;
        else { flush(); run = { format, text }; }
        continue;
      }
      flush();
      children.push(canonical(child));
    } else if ((child.nodeType === 3 || child.nodeType === 4) && (child.nodeValue?.trim() || node.localName === "t" || node.localName === "instrText")) {
      flush();
      children.push(child.nodeValue ?? "");
    }
  }
  flush();
  return JSON.stringify([`{${node.namespaceURI ?? ""}}${node.localName}`, attrs, children.filter(Boolean)]);
}

export function unpack(bytes: Uint8Array): Record<string, Uint8Array> {
  let expanded = 0;
  return unzipSync(bytes, { filter: (entry) => {
    expanded += entry.originalSize;
    if (expanded > 128 * 1024 * 1024) throw new Error("The expanded DOCX exceeds 128 MB.");
    return !entry.name.endsWith("/");
  } });
}

export function mainDocument(parts: Record<string, Uint8Array>): Document {
  const xml = parts["word/document.xml"];
  if (!xml) throw new Error("The DOCX has no main document part.");
  return parseXml(strFromU8(xml));
}

export function body(document: Document): Element {
  const value = document.getElementsByTagNameNS(W, "body")[0];
  if (!value) throw new Error("The DOCX has no document body.");
  return value;
}

export function blocks(document: Document): Element[] {
  return elements(body(document)).filter((element) => element.localName !== "sectPr");
}

export function blockKind(block: Element): "p" | "tbl" {
  if (block.namespaceURI !== W || (block.localName !== "p" && block.localName !== "tbl")) throw new Error("Fragment editing currently supports body paragraphs and tables. This document has another top-level block type.");
  return block.localName;
}

export async function fingerprint(block: Element): Promise<string> {
  return digestBytes(new TextEncoder().encode(canonical(block)));
}

export function flatDocument(xml: string): Document {
  const parsed = parseXml(xml);
  const part = Array.from(parsed.getElementsByTagNameNS(PKG, "part")).find((entry) => entry.getAttributeNS(PKG, "name") === "/word/document.xml");
  const root = part?.getElementsByTagNameNS(W, "document")[0];
  if (!root) throw new Error("Word returned an unsupported OOXML fragment.");
  return parseXml(serialize(root));
}
