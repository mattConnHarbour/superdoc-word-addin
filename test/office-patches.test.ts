import assert from "node:assert/strict";
import { test } from "node:test";
import { strToU8, zipSync } from "fflate";
import type { FragmentPatch } from "../src/contract.js";
import { blocks, flatDocument, parseXml, PKG, serialize, W } from "../src/ooxml.js";
import { applyFragmentUpdate, captureFragmentSource, releaseFragmentSource, type Snapshot } from "../src/taskpane/operations.js";

const paragraph = (text: string) => `<w:p xmlns:w="${W}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
const documentXml = (content: string) => `<w:document xmlns:w="${W}"><w:body>${content}<w:sectPr/></w:body></w:document>`;
const flat = (content: string) => `<pkg:package xmlns:pkg="${PKG}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData>${documentXml(content)}</pkg:xmlData></pkg:part></pkg:package>`;

interface Block { id: number; xml: string; }
class FakeWord {
  content: Block[] = ["First", "Second", "Third"].map((text, id) => ({ id, xml: paragraph(text) }));
  writes: string[] = [];
  tracked = new Set<FakeRange>();
  onSync: (() => void) | undefined;
  onExport: (() => void) | undefined;
  liveXml = (xml: string) => xml;
  document = {
    body: {
      getOoxml: () => ({ value: flat(this.content.map((block) => this.liveXml(block.xml)).join("")) }),
      paragraphs: { load() {}, items: this.content.map((block) => ({ parentTableOrNullObject: { isNullObject: true }, getRange: () => new FakeRange(this, block.id, block.id) })) },
      tables: { load() {}, items: [] },
    },
  };
  context = { document: this.document, sync: async () => { this.onSync?.(); } };
  exportBytes() {
    const bytes = zipSync({ "word/document.xml": strToU8(documentXml(this.content.map((block) => block.xml).join(""))) });
    this.onExport?.();
    return bytes;
  }
}
class FakeRange {
  constructor(readonly word: FakeWord, readonly first: number, readonly last: number) {}
  track() { this.word.tracked.add(this); }
  untrack() { this.word.tracked.delete(this); }
  expandTo(other: FakeRange) { return new FakeRange(this.word, this.first, other.last); }
  getOoxml() {
    const start = this.word.content.findIndex((block) => block.id === this.first);
    const end = this.word.content.findIndex((block) => block.id === this.last);
    const content = start < 0 || end < 0 ? "" : this.word.content.slice(start, end + 1).map((block) => this.word.liveXml(block.xml)).join("");
    return { value: flat(`${content}<w:p xmlns:w="${W}"/>`) };
  }
  insertOoxml(xml: string, location: string) {
    this.word.writes.push(location);
    const start = this.word.content.findIndex((block) => block.id === this.first);
    const end = this.word.content.findIndex((block) => block.id === this.last);
    const inserted = blocks(flatDocument(xml)).map((block, index) => ({ id: 100 + index, xml: serialize(block) }));
    this.word.content.splice(location === "After" ? end + 1 : start, location === "Replace" ? end - start + 1 : 0, ...inserted);
  }
  delete() {
    this.word.writes.push("Delete");
    const start = this.word.content.findIndex((block) => block.id === this.first);
    const end = this.word.content.findIndex((block) => block.id === this.last);
    this.word.content.splice(start, end - start + 1);
  }
}

async function withWord(run: (word: FakeWord) => Promise<void>) {
  const fake = new FakeWord();
  const url = "file:///test.docx";
  Object.defineProperty(globalThis, "Office", { configurable: true, value: {
    FileType: { Compressed: "compressed" }, AsyncResultStatus: { Succeeded: "succeeded" },
    context: { document: { url, getFileAsync(_type: unknown, _options: unknown, done: (result: unknown) => void) {
      const bytes = fake.exportBytes();
      done({ status: "succeeded", value: { size: bytes.length, sliceCount: 1,
        getSliceAsync(_index: number, callback: (result: unknown) => void) { callback({ status: "succeeded", value: { data: Array.from(bytes) } }); },
        closeAsync() {},
      } });
    } } },
  } });
  Object.defineProperty(globalThis, "Word", { configurable: true, value: {
    run: async (first: ((context: typeof fake.context) => Promise<unknown>) | typeof fake.document, second?: (context: typeof fake.context) => Promise<unknown>) => {
      if (typeof first === "function") return first(fake.context);
      if (second) return second(fake.context);
      throw new Error("Missing Word callback");
    },
  } });
  try { await run(fake); }
  finally { await releaseFragmentSource(); Reflect.deleteProperty(globalThis, "Office"); Reflect.deleteProperty(globalThis, "Word"); }
}

const signal = new AbortController().signal;
function applyFragmentPatches(source: Snapshot, patches: FragmentPatch[], signal: AbortSignal, beforeWrite: () => void) {
  return applyFragmentUpdate({ source, patches, signal, beforeWrite });
}
function patch(start: number, expected: string[], text: string): FragmentPatch {
  return { start, deleteCount: expected.length, expected, insertCount: 1, ooxml: flat(paragraph(text)) };
}

test("Word's live OOXML projection can differ from its DOCX export without a false capture conflict", async () => withWord(async (word) => {
  word.liveXml = (xml) => xml.replace("<w:r>", '<w:pPr><w:contextualSpacing w:val="0"/></w:pPr><w:r>');
  const source = await captureFragmentSource();
  const hash = source.blocks[0]?.fingerprint;
  assert.ok(hash);
  assert.equal(await applyFragmentPatches(source, [patch(0, [hash], "Edited first")], signal, () => {}), true);
  assert.deepEqual(word.writes, ["Replace"]);
}));

test("a real edit during DOCX capture is rejected and releases tracked ranges", async () => withWord(async (word) => {
  word.onExport = () => {
    const first = word.content[0];
    assert.ok(first);
    first.xml = paragraph("Edited while exporting");
  };
  await assert.rejects(captureFragmentSource(), /Word changed during capture/u);
  assert.equal(word.tracked.size, 0);
  assert.deepEqual(word.writes, []);
}));

test("tracked Word ranges preserve unrelated coauthor edits and apply only the target fragment", async () => withWord(async (word) => {
  const source = await captureFragmentSource();
  const hash = source.blocks[0]?.fingerprint;
  assert.ok(hash);
  const third = word.content[2];
  assert.ok(third);
  third.xml = paragraph("Concurrent change in another paragraph");
  assert.equal(await applyFragmentPatches(source, [patch(0, [hash], "Edited first")], signal, () => {}), true);
  assert.deepEqual(word.writes, ["Replace"]);
  assert.match(word.content[2]?.xml ?? "", /Concurrent change/u);
  await releaseFragmentSource();
  assert.equal(word.tracked.size, 0);
}));

test("an affected range change blocks every patch before the first write", async () => withWord(async (word) => {
  const source = await captureFragmentSource();
  const hash = source.blocks[0]?.fingerprint;
  const hash2 = source.blocks[1]?.fingerprint;
  assert.ok(hash);
  assert.ok(hash2);
  const second = word.content[1];
  assert.ok(second);
  second.xml = paragraph("Coauthor changed the target");
  assert.equal(await applyFragmentPatches(source, [patch(0, [hash], "First edited"), patch(1, [hash2], "Second edited")], signal, () => {}), false);
  assert.deepEqual(word.writes, []);
}));

test("a new block inserted inside a multi-block target is detected", async () => withWord(async (word) => {
  const source = await captureFragmentSource();
  const hashes = source.blocks.slice(0, 2).map((block) => block.fingerprint);
  word.content.splice(1, 0, { id: 44, xml: paragraph("New concurrent paragraph") });
  assert.equal(await applyFragmentPatches(source, [patch(0, hashes, "Replacement")], signal, () => {}), false);
  assert.deepEqual(word.writes, []);
}));

test("insertions use neighboring tracked ranges and detect changes at the insertion gap", async () => withWord(async (word) => {
  const source = await captureFragmentSource();
  const insertion = patch(1, [], "Inserted");
  word.content.splice(1, 0, { id: 44, xml: paragraph("New concurrent paragraph") });
  assert.equal(await applyFragmentPatches(source, [insertion], signal, () => {}), false);
  assert.deepEqual(word.writes, []);
}));

test("fragment plans are checked for invalid indices and expected content before writes", async () => withWord(async (word) => {
  const source = await captureFragmentSource();
  await assert.rejects(applyFragmentPatches(source, [patch(0, ["a".repeat(64)], "Wrong source")], signal, () => {}), /invalid fragment plan/u);
  assert.deepEqual(word.writes, []);
}));

test("paragraph insertions and deletions call range APIs without replacing the document", async () => withWord(async (word) => {
  const source = await captureFragmentSource();
  assert.equal(await applyFragmentPatches(source, [patch(1, [], "Inserted")], signal, () => {}), true);
  assert.deepEqual(word.writes, ["After"]);
  assert.deepEqual(word.content.map((block) => parseXml(block.xml).documentElement?.textContent), ["First", "Inserted", "Second", "Third"]);
}));
