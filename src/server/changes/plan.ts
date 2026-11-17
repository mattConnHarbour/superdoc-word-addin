import type { ChangeSpan, SupportedBlock } from "./types.js";

export function planJournalPatches(before: SupportedBlock[], after: SupportedBlock[], dirty: Set<string>): ChangeSpan[] {
  const oldIds = new Set(before.map((block) => block.nodeId));
  const newIds = new Set(after.map((block) => block.nodeId));
  if (before.some((block) => !newIds.has(block.nodeId) && !dirty.has(block.nodeId))
    || after.some((block) => !oldIds.has(block.nodeId) && !dirty.has(block.nodeId))) throw new Error("The block structure changed without a corresponding SDK change record.");
  const oldAnchors = before.flatMap((block, index) => dirty.has(block.nodeId) ? [] : [{ id: block.nodeId, index }]);
  const newAnchors = after.flatMap((block, index) => dirty.has(block.nodeId) ? [] : [{ id: block.nodeId, index }]);
  if (oldAnchors.length !== newAnchors.length || oldAnchors.some((anchor, index) => anchor.id !== newAnchors[index]?.id)) throw new Error("The SDK reported an incomplete block move. Word was not updated.");
  const spans: ChangeSpan[] = [];
  let start = 0;
  let from = 0;
  for (let index = 0; index <= oldAnchors.length; index += 1) {
    const end = oldAnchors[index]?.index ?? before.length;
    const to = newAnchors[index]?.index ?? after.length;
    if (start !== end || from !== to) spans.push({ start, end, from, to });
    start = end + 1;
    from = to + 1;
  }
  return spans;
}
