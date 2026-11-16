import { strFromU8, strToU8, zipSync } from "fflate";
import { blockIdentity, blockKind, blocks, body, canonical, elements, mainDocument, parseXml, serialize, unpack, W } from "../../ooxml.js";
import { packageRevision } from "../../taskpane/document-revision.js";
import type { DocumentVersions } from "../operations/types.js";
import type { ChangeSpan, PatchContext } from "./types.js";

function additivePart(before: Uint8Array | undefined, after: Uint8Array, key: string): boolean {
  if (!before) return true;
  const left = parseXml(strFromU8(before)).documentElement;
  const right = parseXml(strFromU8(after)).documentElement;
  if (!left || !right) return false;
  return elements(left).every((entry) => elements(right).some((other) => other.getAttribute(key) === entry.getAttribute(key) && canonical(entry) === canonical(other)));
}

export function openPatchContext(versions: DocumentVersions): PatchContext {
  const source = unpack(versions.source);
  const baseline = unpack(versions.baseline);
  const result = unpack(versions.updated);
  const sourceBlocks = blocks(mainDocument(source));
  const before = blocks(mainDocument(baseline));
  const after = blocks(mainDocument(result));
  [...sourceBlocks, ...before, ...after].forEach(blockKind);
  if (sourceBlocks.length !== before.length || sourceBlocks.some((block, index) => {
    const normalized = before[index];
    return !normalized || blockIdentity(block) !== blockIdentity(normalized);
  })) throw new Error("SuperDoc normalized the document's block structure. Its edits cannot be mapped safely to Word ranges.");
  return { source, baseline, result, sourceBlocks, before, after };
}

export async function assertCompatiblePackage(context: PatchContext): Promise<void> {
  const { baseline, result } = context;
  const comparison = { ...result };
  const baselineDocument = mainDocument(baseline);
  const resultDocument = mainDocument(result);
  const originalBody = body(baselineDocument);
  const resultBody = body(resultDocument);
  const beforeSection = elements(originalBody).filter((element) => element.localName === "sectPr").map(canonical);
  const afterSection = elements(resultBody).filter((element) => element.localName === "sectPr").map(canonical);
  if (JSON.stringify(beforeSection) !== JSON.stringify(afterSection)) throw new Error("Section and page setup edits are not supported by the fragment adapter yet.");
  resultBody.parentNode?.replaceChild(resultDocument.importNode(originalBody, true), resultBody);
  comparison["word/document.xml"] = strToU8(serialize(resultDocument));
  for (const [name, content] of Object.entries(result)) {
    const prior = baseline[name];
    if (name === "word/_rels/document.xml.rels" && additivePart(prior, content, "Id")) {
      if (prior) comparison[name] = prior;
      else delete comparison[name];
    }
    if ((name === "word/numbering.xml" && additivePart(prior, content, "w:numId") && additivePart(prior, content, "w:abstractNumId"))
      || (name.startsWith("word/media/") && !prior)) {
      if (prior) comparison[name] = prior;
      else delete comparison[name];
    }
  }
  if (await packageRevision(new Uint8Array(zipSync(comparison))) !== await packageRevision(new Uint8Array(zipSync(baseline)))) throw new Error("This edit changes document-wide parts (such as styles, headers, comments, or settings) outside the supported body fragments.");
}

export function assertSupportedSpans(context: PatchContext, spans: ChangeSpan[]): void {
  const { sourceBlocks, after } = context;
  for (const span of spans) {
    for (const block of [...sourceBlocks.slice(span.start, span.end), ...after.slice(span.from, span.to)]) {
      const unsupported = ["sectPr", "commentRangeStart", "commentRangeEnd", "commentReference", "bookmarkStart", "bookmarkEnd", "footnoteReference", "endnoteReference", "fldChar", "altChunk"];
      if (unsupported.some((name) => block.getElementsByTagNameNS(W, name).length)) throw new Error("A changed fragment contains a section break, comment, bookmark, note, or complex field that this adapter cannot safely reconcile yet.");
    }
  }
}
