import type { FragmentPatch } from "../../contract.js";
import { blocks, flatDocument } from "../../ooxml.js";
import type { FragmentApplication, Snapshot } from "../operations/types.js";
import { rangeBlocks, resources } from "./ooxml.js";
import type { TrackedBlock, TrackedDocument } from "./session.js";

interface RangeOperation {
  patch: FragmentPatch;
  range: Word.Range;
  location: "Replace" | "Before" | "After";
  check: Word.Range;
  expected: string[];
}
interface FragmentPlan {
  operations: RangeOperation[];
  checks: { xml: OfficeExtension.ClientResult<string>; expected: string[] }[];
}

function validatePatch(source: Snapshot, patch: FragmentPatch, previousEnd: number): void {
  if (patch.start < previousEnd || patch.start + patch.deleteCount > source.blocks.length || patch.expected.length !== patch.deleteCount
    || patch.expected.some((hash, index) => source.blocks[patch.start + index]?.fingerprint !== hash)
    || (patch.insertCount === 0) !== (patch.ooxml === "")
    || (patch.insertCount > 0 && blocks(flatDocument(patch.ooxml)).length !== patch.insertCount)) throw new Error("The backend returned an invalid fragment plan.");
}

function locateOperation(patch: FragmentPatch, document: TrackedDocument): RangeOperation {
  const end = patch.start + patch.deleteCount;
  const first = document.blocks[patch.start];
  const last = document.blocks[end - 1];
  if (patch.deleteCount && first && last) {
    const range = patch.deleteCount === 1 ? first.range : first.range.expandTo(last.range);
    return { patch, range, location: "Replace", check: range, expected: document.blocks.slice(patch.start, end).map((block) => block.canonical) };
  }
  const previous = document.blocks[patch.start - 1];
  const next = document.blocks[patch.start];
  const anchor = previous ?? next;
  if (!anchor) throw new Error("Word has no paragraph to anchor this insertion.");
  return {
    patch, range: anchor.range, location: previous ? "After" : "Before",
    check: previous && next ? previous.range.expandTo(next.range) : anchor.range,
    expected: [previous, next].filter((value): value is TrackedBlock => value !== undefined).map((block) => block.canonical),
  };
}

export function prepareFragmentPlan(application: FragmentApplication, document: TrackedDocument): FragmentPlan {
  let end = 0;
  const operations = application.patches.map((patch) => {
    validatePatch(application.source, patch, end);
    end = patch.start + patch.deleteCount;
    return locateOperation(patch, document);
  });
  return { operations, checks: operations.map((operation) => ({ xml: operation.check.getOoxml(), expected: operation.expected })) };
}

export function fragmentPreconditionsMatch(plan: FragmentPlan, document: TrackedDocument): boolean {
  return plan.checks.every((check) => JSON.stringify(rangeBlocks(check.xml.value, check.expected.length)) === JSON.stringify(check.expected)
    && Array.from(resources(check.xml.value)).every(([name, value]) => document.resources.get(name) === value));
}

export function queueFragmentWrites(plan: FragmentPlan): void {
  for (const { patch, range, location } of [...plan.operations].reverse()) {
    if (patch.insertCount) range.insertOoxml(patch.ooxml, location);
    else range.delete();
  }
}
