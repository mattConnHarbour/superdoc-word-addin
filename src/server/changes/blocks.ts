import type { SuperDocDocument } from "@superdoc/sdk";
import { parseXml, W } from "../../ooxml.js";
import type { SupportedBlock } from "./types.js";

export function paragraphIds(xml: string): string[] {
  const root = parseXml(`<root xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">${xml}</root>`);
  return Array.from(root.getElementsByTagNameNS(W, "p")).map((paragraph) => paragraph.getAttributeNS("http://schemas.microsoft.com/office/word/2010/wordml", "paraId") ?? "").filter(Boolean);
}

export function assertSupportedMarkup(xml: string): void {
  const root = parseXml(`<root xmlns:w="${W}">${xml}</root>`);
  const unsupported = ["sectPr", "commentRangeStart", "commentRangeEnd", "commentReference", "bookmarkStart", "bookmarkEnd", "footnoteReference", "endnoteReference", "fldChar", "altChunk"];
  if (unsupported.some((name) => root.getElementsByTagNameNS(W, name).length)) throw new Error("A changed fragment contains a section break, comment, bookmark, note, or complex field that this adapter cannot safely reconcile yet.");
}

export async function listBodyBlocks(document: Pick<SuperDocDocument, "blocks">): Promise<SupportedBlock[]> {
  const blocks: SupportedBlock[] = [];
  for (let offset = 0; ; offset += 100) {
    const page = await document.blocks.list({ offset, limit: 100, in: { kind: "story", storyType: "body" } });
    for (const block of page.blocks) {
      if (block.nodeType !== "paragraph" && block.nodeType !== "heading" && block.nodeType !== "listItem" && block.nodeType !== "table") throw new Error("Fragment updates support body paragraphs and tables only.");
      blocks.push({ ...block, nodeType: block.nodeType });
    }
    if (blocks.length >= page.total) return blocks;
    if (!page.blocks.length || blocks.length > 20_000) throw new Error("The SDK returned an incomplete or oversized body block list.");
  }
}

export function captureBlocks(capture: { fragment: { xml: string; namespaces: Record<string, string> } }): string {
  const root = parseXml("<root/>");
  for (const [prefix, uri] of Object.entries(capture.fragment.namespaces)) root.documentElement?.setAttribute(`xmlns${prefix ? `:${prefix}` : ""}`, uri);
  const wrapper = root.documentElement;
  if (!wrapper) throw new Error("Could not read the SDK fragment.");
  const namespaces = Array.from(wrapper.attributes).map((attribute) => `${attribute.name}="${attribute.value.replaceAll('&', '&amp;').replaceAll('"', '&quot;')}"`).join(" ");
  return `<root ${namespaces}>${capture.fragment.xml}</root>`;
}
