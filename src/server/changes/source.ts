import type { SuperDocDocument } from "@superdoc/sdk";
import { blockKind, blocks, fingerprint, mainDocument, serialize, unpack } from "../../ooxml.js";
import { assertSupportedMarkup, listBodyBlocks, paragraphIds } from "./blocks.js";
import type { ChangeSource } from "./types.js";

export async function readChangeSource(document: SuperDocDocument, source: Uint8Array): Promise<ChangeSource> {
  const journal = await document.changes.snapshot();
  const nativeBlocks = await listBodyBlocks(document);
  const wordBlocks = blocks(mainDocument(unpack(source)));
  if (wordBlocks.length !== nativeBlocks.length) throw new Error("The SDK body blocks do not match the Word snapshot.");
  const owners = new Map<string, string>();
  const protectedBlocks = new Set<string>();
  for (const [index, block] of nativeBlocks.entries()) {
    const word = wordBlocks[index];
    if (!word || (blockKind(word) === "tbl") !== (block.nodeType === "table")) throw new Error("The SDK and Word block structures do not match.");
    const xml = serialize(word);
    owners.set(block.nodeId, block.nodeId);
    for (const id of paragraphIds(xml)) { owners.set(id, block.nodeId); if (block.nodeType === "table") owners.set(`tbl:${id}`, block.nodeId); }
    try { assertSupportedMarkup(xml); } catch { protectedBlocks.add(block.nodeId); }
  }
  return {
    sessionId: document.sessionId, journal, coverage: await document.diff.capture(), blocks: nativeBlocks,
    fingerprints: await Promise.all(wordBlocks.map(fingerprint)), protectedBlocks, owners,
  };
}
