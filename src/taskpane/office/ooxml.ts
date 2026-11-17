import { blocks, canonical, elements, flatDocument, parseXml, PKG } from "../../ooxml.js";

export function rangeElements(xml: string, expectedCount: number) {
  const content = blocks(flatDocument(xml));
  const last = content.at(-1);
  if (content.length === expectedCount + 1 && last?.localName === "p" && last.childNodes.length === 0) content.pop();
  return content;
}

export function rangeBlocks(xml: string, expectedCount: number): string[] {
  return rangeElements(xml, expectedCount).map(canonical);
}

export function resources(xml: string): Map<string, string> {
  const parsed = parseXml(xml);
  return new Map(Array.from(parsed.getElementsByTagNameNS(PKG, "part")).flatMap((part) => {
    const name = part.getAttributeNS(PKG, "name");
    if (!name || !(/^\/word\/(styles\.xml|numbering\.xml|fontTable\.xml|theme\/|media\/)/u.test(name))) return [];
    return [[name, elements(part).map(canonical).join("")]];
  }));
}
