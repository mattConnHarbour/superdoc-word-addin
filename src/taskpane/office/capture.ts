import { blockKind, blocks, canonical, fingerprint, mainDocument, unpack } from "../../ooxml.js";
import { digestBytes } from "../document-revision.js";
import type { Snapshot } from "../operations/types.js";
import { bytesToBase64, readCompressedDocument } from "./files.js";
import { rangeBlocks, resources } from "./ooxml.js";
import { setTrackedDocument, type TrackedBlock } from "./session.js";

async function loadBodyRanges(context: Word.RequestContext) {
  const paragraphs = context.document.body.paragraphs;
  const tables = context.document.body.tables;
  const bodyXml = context.document.body.getOoxml();
  paragraphs.load("items/parentTableOrNullObject/isNullObject");
  tables.load("items/nestingLevel");
  await context.sync();
  return {
    xml: bodyXml.value,
    paragraphs: paragraphs.items.filter((paragraph) => paragraph.parentTableOrNullObject.isNullObject).map((paragraph) => paragraph.getRange("Whole")),
    tables: tables.items.filter((table) => table.nestingLevel === 1).map((table) => table.getRange("Whole")),
  };
}

function matchBodyRanges(bytes: Uint8Array, captured: Awaited<ReturnType<typeof loadBodyRanges>>): TrackedBlock[] {
  const source = blocks(mainDocument(unpack(bytes)));
  const live = rangeBlocks(captured.xml, source.length);
  if (live.length !== source.length || source.some((block, index) => canonical(block) !== live[index])) throw new Error("Word changed during capture. Please retry against the latest document.");
  const paragraphs = [...captured.paragraphs];
  const tables = [...captured.tables];
  const tracked = source.map((block) => {
    const range = (blockKind(block) === "p" ? paragraphs : tables).shift();
    if (!range) throw new Error("This document's blocks cannot be mapped to Word ranges.");
    return { range, canonical: canonical(block) };
  });
  if (paragraphs.length || tables.length) throw new Error("The document's Word ranges do not match its DOCX blocks.");
  return tracked;
}

export async function captureTrackedDocument(context: Word.RequestContext): Promise<Snapshot> {
  const url = Office.context.document.url;
  const captured = await loadBodyRanges(context);
  const trackedResources = resources(captured.xml);
  const ranges = [...captured.paragraphs, ...captured.tables];
  for (const range of ranges) range.track();
  setTrackedDocument({ context, url, resources: trackedResources, blocks: ranges.map((range) => ({ range, canonical: "" })) });
  await context.sync();
  const bytes = await readCompressedDocument(Office.context.document);
  setTrackedDocument({ context, url, resources: trackedResources, blocks: matchBodyRanges(bytes, captured) });
  return {
    base64: bytesToBase64(bytes), digest: await digestBytes(bytes), url,
    blocks: await Promise.all(blocks(mainDocument(unpack(bytes))).map(async (block) => ({ kind: blockKind(block), fingerprint: await fingerprint(block) }))),
  };
}
