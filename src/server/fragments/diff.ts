import type { Element } from "@xmldom/xmldom";
import { blockIdentity, canonical } from "../../ooxml.js";
import type { ChangeSpan } from "./types.js";

function stableMatches(before: Element[], after: Element[]): [number, number][] {
  const positions = new Map<string, number[]>();
  before.forEach((block, index) => {
    const key = blockIdentity(block);
    positions.set(key, [...positions.get(key) ?? [], index]);
  });
  const candidates: [number, number][] = [];
  after.forEach((block, index) => {
    const found = positions.get(blockIdentity(block))?.shift();
    if (found !== undefined) candidates.push([found, index]);
  });
  const tails: number[] = [];
  const previous = new Map<number, number>();
  candidates.forEach(([position], index) => {
    let low = 0;
    let high = tails.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      const candidate = candidates[tails[middle] ?? -1];
      if (candidate && candidate[0] < position) low = middle + 1;
      else high = middle;
    }
    previous.set(index, tails[low - 1] ?? -1);
    tails[low] = index;
  });
  const matches: [number, number][] = [];
  for (let index = tails.at(-1) ?? -1; index >= 0; index = previous.get(index) ?? -1) {
    const match = candidates[index];
    if (match) matches.unshift(match);
  }
  return matches;
}

export function findChangedSpans(before: Element[], after: Element[]): ChangeSpan[] {
  const matches = stableMatches(before, after);
  const spans: ChangeSpan[] = [];
  let oldPosition = 0;
  let newPosition = 0;
  for (const [oldIndex, newIndex] of [...matches, [before.length, after.length]]) {
    if (oldIndex === undefined || newIndex === undefined) continue;
    if (oldPosition !== oldIndex || newPosition !== newIndex) spans.push({ start: oldPosition, end: oldIndex, from: newPosition, to: newIndex });
    const oldBlock = before[oldIndex];
    const newBlock = after[newIndex];
    if (oldBlock && newBlock && canonical(oldBlock) !== canonical(newBlock)) spans.push({ start: oldIndex, end: oldIndex + 1, from: newIndex, to: newIndex + 1 });
    oldPosition = oldIndex + 1;
    newPosition = newIndex + 1;
  }
  const merged: typeof spans = [];
  for (const span of spans) {
    const previous = merged.at(-1);
    if (previous && previous.end === span.start && previous.to === span.from) {
      previous.end = span.end;
      previous.to = span.to;
    } else merged.push({ ...span });
  }
  return merged;
}
