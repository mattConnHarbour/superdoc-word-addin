import type { FragmentPatch } from "../../contract.js";
import { fingerprint } from "../../ooxml.js";
import type { ChangeSpan, PatchContext } from "./types.js";
import { buildFragmentPackage } from "./package.js";

export async function serializePatches(context: PatchContext, spans: ChangeSpan[]): Promise<FragmentPatch[]> {
  const { sourceBlocks, after, result } = context;
  const patches = await Promise.all(spans.map(async ({ start, end, from, to }) => ({
    start, deleteCount: end - start, expected: await Promise.all(sourceBlocks.slice(start, end).map(fingerprint)),
    ooxml: from === to ? "" : buildFragmentPackage(result, after.slice(from, to)), insertCount: to - from,
  })));
  if (patches.reduce((size, patch) => size + new TextEncoder().encode(patch.ooxml).length, 0) > 8 * 1024 * 1024) throw new Error("The combined fragments exceed 8 MB. Try a smaller edit.");
  return patches;
}
