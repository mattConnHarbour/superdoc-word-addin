import assert from "node:assert/strict";
import { test } from "node:test";
import { planJournalPatches } from "../src/server/changes/plan.js";
import { readJournal, changedIds } from "../src/server/changes/journal.js";
import type { ChangePage, ChangeSnapshot, SupportedBlock } from "../src/server/changes/types.js";
import { canonical, parseXml, W } from "../src/ooxml.js";

const block = (nodeId: string): SupportedBlock => ({ nodeId, nodeType: "paragraph", ordinal: 0 });
const snapshot = (cursor: string | null): ChangeSnapshot => ({ cursor, revision: cursor ?? "0", lastChangeId: cursor, stateVector: "", at: "" });
const page = (cursor: string | null, truncated = false): ChangePage => ({ cursor, truncated, items: [], identity: "missing", collaboration: false });

test("journal-driven spans preserve unchanged anchors and combine adjacent affected blocks", () => {
  const before = ["a", "b", "c"].map(block);
  assert.deepEqual(planJournalPatches(before, ["a", "x", "b", "c"].map(block), new Set(["x", "b"])), [{ start: 1, end: 2, from: 1, to: 3 }]);
  assert.deepEqual(planJournalPatches(before, ["a", "c"].map(block), new Set(["b"])), [{ start: 1, end: 2, from: 1, to: 1 }]);
});

test("moving a recorded block creates an insertion and deletion around untouched anchors", () => {
  assert.deepEqual(planJournalPatches(["a", "b", "c"].map(block), ["c", "a", "b"].map(block), new Set(["c"])), [
    { start: 0, end: 0, from: 0, to: 1 }, { start: 2, end: 3, from: 3, to: 3 },
  ]);
});

test("unreported structural changes fail instead of invoking a fallback diff", () => {
  assert.throws(() => planJournalPatches(["a", "b"].map(block), ["b", "a"].map(block), new Set()), /incomplete block move/u);
  assert.throws(() => planJournalPatches(["a"].map(block), ["a", "x"].map(block), new Set()), /without a corresponding/u);
});

test("journal reader follows resume cursors until the captured endpoint", async () => {
  const calls: (string | undefined)[] = [];
  const events = await readJournal({ list: async (params) => {
    calls.push(params?.after);
    return page(params?.after === "one" ? "two" : "three");
  } }, snapshot("one"), snapshot("three"));
  assert.deepEqual(events, []);
  assert.deepEqual(calls, ["one", "two"]);
});

test("truncated and non-advancing journals fail before patch planning", async () => {
  await assert.rejects(readJournal({ list: async () => page("two", true) }, snapshot("one"), snapshot("two")), /evicted change records/u);
  await assert.rejects(readJournal({ list: async () => page("one") }, snapshot("one"), snapshot("two")), /incomplete change journal/u);
});

test("split/merge/table events identify all affected identities and sidecars are explicit", () => {
  const story: { kind: "story"; storyType: "body" } = { kind: "story", storyType: "body" };
  const ids = changedIds([
    { kind: "split", story, blockId: "a", tailBlockId: "b", offset: 2 },
    { kind: "merge", story, blockId: "a", mergedBlockId: "c" },
    { kind: "table", story, tableId: "table", detail: {} },
  ]);
  assert.deepEqual(Array.from(ids), ["a", "b", "c", "table"]);
  assert.throws(() => changedIds([{ kind: "sidecar", affects: "comment", ids: ["comment"] }]), /metadata changes/u);
});

test("Word proofing boundaries and equivalent split runs do not create a target conflict", () => {
  const whole = parseXml(`<w:p xmlns:w="${W}"><w:r><w:t>TechCraft builds tools.</w:t></w:r></w:p>`).documentElement;
  const split = parseXml(`<w:p xmlns:w="${W}"><w:proofErr w:type="spellStart"/><w:r><w:t>TechCraft</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r><w:t xml:space="preserve"> builds tools.</w:t></w:r></w:p>`).documentElement;
  assert.ok(whole); assert.ok(split);
  assert.equal(canonical(whole), canonical(split));
});
