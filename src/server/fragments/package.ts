import { body, elements, parseXml, PKG, R, REL, serialize, W } from "../../ooxml.js";
import { captureBlocks } from "../changes/blocks.js";
import type { FragmentCapture } from "../changes/types.js";

interface PackagePart {
  contentType: string;
  xml?: string;
  base64?: string;
}
const MAIN_PART = "/word/document.xml";
const RELATIONSHIP_TYPE = "application/vnd.openxmlformats-package.relationships+xml";

function documentXml(capture: FragmentCapture): string {
  const document = parseXml(`<w:document xmlns:w="${W}"><w:body/></w:document>`);
  const root = document.documentElement;
  if (!root) throw new Error("Could not create the Word fragment document.");
  for (const [prefix, uri] of Object.entries(capture.fragment.namespaces)) root.setAttributeNS("http://www.w3.org/2000/xmlns/", prefix ? `xmlns:${prefix}` : "xmlns", uri);
  const fragment = parseXml(captureBlocks(capture)).documentElement;
  if (!fragment) throw new Error("The SDK returned an empty fragment.");
  for (const block of elements(fragment)) body(document).appendChild(document.importNode(block, true));
  return serialize(document);
}

function addDependencies(capture: FragmentCapture, parts: Map<string, PackagePart>): void {
  const relationships = new Map<string, Map<string, string>>();
  for (const dependency of capture.dependencies) {
    const source = `/${dependency.sourcePartUri.replace(/^\//u, "")}`;
    const ids = relationships.get(source) ?? new Map<string, string>();
    const relation = parseXml(`<Relationship xmlns="${REL}"/>`).documentElement;
    if (!relation) throw new Error("Could not create a fragment relationship.");
    relation.setAttribute("Id", dependency.relationshipId);
    relation.setAttribute("Type", dependency.relationshipType);
    relation.setAttribute("Target", dependency.target);
    if (dependency.targetMode === "External") relation.setAttribute("TargetMode", "External");
    const xml = serialize(relation);
    if (ids.has(dependency.relationshipId) && ids.get(dependency.relationshipId) !== xml) throw new Error("The SDK returned conflicting relationship identities.");
    ids.set(dependency.relationshipId, xml);
    relationships.set(source, ids);
    if (dependency.targetMode === "External") continue;
    const uri = new URL(dependency.target, `https://package${source}`).pathname;
    if (!dependency.contentType || dependency.dataBase64 === undefined || uri === MAIN_PART) throw new Error("The SDK fragment has an incomplete package dependency.");
    const prior = parts.get(uri);
    if (prior && (prior.base64 !== dependency.dataBase64 || prior.contentType !== dependency.contentType)) throw new Error("The SDK fragment contains conflicting package parts.");
    parts.set(uri, { contentType: dependency.contentType, base64: dependency.dataBase64 });
  }
  for (const [source, ids] of relationships) parts.set(source.replace(/([^/]+)$/u, "_rels/$1.rels"), {
    contentType: RELATIONSHIP_TYPE, xml: `<Relationships xmlns="${REL}">${Array.from(ids.values()).join("")}</Relationships>`,
  });
}

function appendPart(document: ReturnType<typeof parseXml>, uri: string, value: PackagePart): void {
  const part = document.createElementNS(PKG, "pkg:part");
  part.setAttributeNS(PKG, "pkg:name", uri);
  part.setAttributeNS(PKG, "pkg:contentType", value.contentType);
  const xml = value.xml ?? (/xml$/u.test(value.contentType) ? Buffer.from(value.base64 ?? "", "base64").toString("utf8") : undefined);
  const data = document.createElementNS(PKG, xml === undefined ? "pkg:binaryData" : "pkg:xmlData");
  if (xml === undefined) data.appendChild(document.createTextNode(value.base64 ?? ""));
  else {
    const root = parseXml(xml).documentElement;
    if (!root) throw new Error("An SDK dependency contains empty XML.");
    data.appendChild(document.importNode(root, true));
  }
  part.appendChild(data);
  document.documentElement?.appendChild(part);
}

export function buildFragmentPackage(capture: FragmentCapture): string {
  const parts = new Map<string, PackagePart>([
    [MAIN_PART, { contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", xml: documentXml(capture) }],
    ["/_rels/.rels", { contentType: RELATIONSHIP_TYPE, xml: `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>` }],
  ]);
  addDependencies(capture, parts);
  const document = parseXml(`<pkg:package xmlns:pkg="${PKG}"/>`);
  for (const [uri, part] of parts) appendPart(document, uri, part);
  const result = serialize(document);
  if (new TextEncoder().encode(result).length > 2 * 1024 * 1024) throw new Error("A changed fragment exceeds the 2 MB patch limit. Try a smaller edit.");
  return result;
}
