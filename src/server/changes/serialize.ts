import type { SuperDocDocument } from "@superdoc/sdk";
import type { FragmentPatch } from "../../contract.js";
import { buildFragmentPackage } from "../fragments/package.js";
import { assertSupportedMarkup, captureBlocks } from "./blocks.js";
import { extractSpan } from "./extract.js";
import type { ChangeSource, ChangeSpan, SupportedBlock } from "./types.js";

export async function serializeJournalPatches(document: SuperDocDocument, source: ChangeSource, after: SupportedBlock[], spans: ChangeSpan[]): Promise<FragmentPatch[]> {
  const patches: FragmentPatch[] = [];
  for (const { start, end, from, to } of spans) {
    if (source.blocks.slice(start, end).some((block) => source.protectedBlocks.has(block.nodeId))) throw new Error("An affected Word block contains unsupported bookmarks, comments, fields, notes, or section breaks.");
    const capture = from === to ? undefined : await extractSpan(document, after.slice(from, to));
    if (capture) assertSupportedMarkup(captureBlocks(capture));
    patches.push({ start, deleteCount: end - start, expected: source.fingerprints.slice(start, end), insertCount: to - from,
      ooxml: capture ? buildFragmentPackage(capture) : "",
    });
  }
  if (patches.reduce((size, patch) => size + new TextEncoder().encode(patch.ooxml).length, 0) > 8 * 1024 * 1024) throw new Error("The combined fragments exceed 8 MB. Try a smaller edit.");
  return patches;
}
