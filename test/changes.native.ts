import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "node:test";
import { createSuperDocClient, type SuperDocDocument } from "@superdoc/sdk";
import { captureChangeSource, createFragmentUpdate, type ChangeSource } from "../src/server/operations.js";
import { blocks, flatDocument, parseXml, PKG, W } from "../src/ooxml.js";

async function withDocument(run: (document: SuperDocDocument, source: ChangeSource) => Promise<void>, seed?: (document: SuperDocDocument) => Promise<void>) {
  const directory = await mkdtemp(path.join(tmpdir(), "word-journal-test-"));
  const client = createSuperDocClient({ runtime: "v2", stateDir: path.join(directory, "state") });
  try {
    const document = await client.open({});
    await document.create.paragraph({ text: "First paragraph." });
    await document.create.paragraph({ text: "Second paragraph." });
    await document.create.paragraph({ text: "Third paragraph." });
    await seed?.(document);
    const file = path.join(directory, "source.docx");
    await document.save({ out: file });
    await run(document, await captureChangeSource(document, await readFile(file)));
  } finally {
    try { await client.dispose(); } finally { await rm(directory, { recursive: true, force: true }); }
  }
}
function paragraph(source: ChangeSource, text: string): { kind: "block"; nodeType: "paragraph" | "heading" | "listItem" | "table"; nodeId: string } {
  const block = source.blocks.find((entry) => entry.textPreview === text);
  assert.ok(block, `Missing test paragraph: ${text}`);
  return { kind: "block", nodeType: block.nodeType, nodeId: block.nodeId };
}

test("a read-only agent turn yields no Word patches", { timeout: 30_000 }, async () => withDocument(async (document, source) => {
  assert.deepEqual(await createFragmentUpdate(document, source), { patches: [] });
}));

test("the native journal and OOXML exporter produce only the edited paragraph", { timeout: 30_000 }, async () => withDocument(async (document, source) => {
  const target = paragraph(source, "First paragraph.");
  await document.replace({ blockId: target.nodeId, start: 0, end: 5, text: "Updated" });
  await document.format.bold({ blockId: target.nodeId, start: 0, end: 7, value: true });
  const update = await createFragmentUpdate(document, source);
  assert.equal(update.unsupportedReason, undefined);
  assert.equal(update.patches.length, 1);
  const patch = update.patches[0]; assert.ok(patch);
  assert.equal(patch.start, source.blocks.findIndex((block) => block.nodeId === target.nodeId));
  assert.equal(patch.deleteCount, 1);
  assert.match(blocks(flatDocument(patch.ooxml))[0]?.textContent ?? "", /Updated paragraph/u);
  assert.doesNotMatch(patch.ooxml, /Second paragraph|Third paragraph/u);
  const xml = parseXml(patch.ooxml);
  assert.ok(xml.getElementsByTagNameNS(W, "b").length);
  assert.ok(Array.from(xml.getElementsByTagNameNS(PKG, "part")).some((part) => part.getAttributeNS(PKG, "name") === "/word/styles.xml"));
}));

test("native insertions and deletions produce structural patches", { timeout: 30_000 }, async () => withDocument(async (document, source) => {
  await document.blocks.delete({ target: paragraph(source, "Second paragraph.") });
  await document.create.paragraph({ text: "New final paragraph." });
  const update = await createFragmentUpdate(document, source);
  assert.equal(update.unsupportedReason, undefined);
  assert.ok(update.patches.some((patch) => patch.deleteCount === 1 && patch.insertCount === 0));
  assert.ok(update.patches.some((patch) => patch.deleteCount === 0 && patch.insertCount === 1 && patch.ooxml.includes("New final paragraph.")));
}));

test("paragraph split and merge records are consolidated into final fragments", { timeout: 30_000 }, async () => withDocument(async (document, source) => {
  await document.blocks.split({ target: paragraph(source, "First paragraph."), offset: 5 });
  const update = await createFragmentUpdate(document, source);
  assert.equal(update.unsupportedReason, undefined);
  assert.ok(update.patches.some((patch) => patch.deleteCount === 1 && patch.insertCount === 2));
}));

test("created tables come from SDK extraction including shading and dependencies", { timeout: 30_000 }, async () => withDocument(async (document, source) => {
  const receipt = await document.create.table({ rows: 2, columns: 2 });
  assert.ok(receipt.table);
  await document.tables.setShading({ nodeId: receipt.table.nodeId, color: "DDEEFF" });
  const update = await createFragmentUpdate(document, source);
  assert.equal(update.unsupportedReason, undefined);
  const tablePatch = update.patches.find((patch) => patch.ooxml.includes("DDEEFF"));
  assert.ok(tablePatch);
  assert.ok(blocks(flatDocument(tablePatch.ooxml)).some((block) => block.localName === "tbl"));
}));

test("an existing table property change patches only the containing table", { timeout: 30_000 }, async () => withDocument(async (document, source) => {
  const table = source.blocks.find((block) => block.nodeType === "table"); assert.ok(table);
  await document.tables.setShading({ nodeId: table.nodeId, color: "DDEEFF" });
  const update = await createFragmentUpdate(document, source);
  assert.equal(update.unsupportedReason, undefined);
  assert.equal(update.patches.length, 1);
  const patch = update.patches[0]; assert.ok(patch);
  assert.equal(patch.deleteCount, 1);
  assert.equal(patch.insertCount, 1);
  assert.equal(blocks(flatDocument(patch.ooxml))[0]?.localName, "tbl");
}, async (document) => { await document.create.table({ rows: 2, columns: 2 }); }));

test("SDK move records become source deletion plus destination insertion", { timeout: 30_000 }, async () => withDocument(async (document, source) => {
  await document.blocks.move({ source: paragraph(source, "Third paragraph."), destination: paragraph(source, "First paragraph."), placement: "before" });
  const update = await createFragmentUpdate(document, source);
  assert.equal(update.unsupportedReason, undefined);
  assert.ok(update.patches.some((patch) => patch.deleteCount === 1 && patch.insertCount === 0));
  assert.ok(update.patches.some((patch) => patch.deleteCount === 0 && patch.insertCount === 1 && patch.ooxml.includes("Third paragraph.")));
}));

test("SDK merge records retain one replacement and remove the joined block", { timeout: 30_000 }, async () => withDocument(async (document, source) => {
  await document.blocks.merge({ first: paragraph(source, "First paragraph."), second: paragraph(source, "Second paragraph.") });
  const update = await createFragmentUpdate(document, source);
  assert.equal(update.unsupportedReason, undefined);
  assert.ok(update.patches.some((patch) => patch.deleteCount === 2 && patch.insertCount === 1));
}));

test("cell text changes resolve to their original Word table range", { timeout: 30_000 }, async () => withDocument(async (document, source) => {
  const table = source.blocks.find((block) => block.nodeType === "table"); assert.ok(table);
  const cell = Array.from(source.owners).find(([id, owner]) => owner === table.nodeId && !id.startsWith("tbl:")); assert.ok(cell);
  await document.replace({ blockId: cell[0], start: 0, end: 0, text: "Updated cell" });
  const update = await createFragmentUpdate(document, source);
  assert.equal(update.unsupportedReason, undefined);
  assert.equal(update.patches.length, 1);
  const patch = update.patches[0]; assert.ok(patch);
  assert.equal(patch.start, source.blocks.findIndex((block) => block.nodeId === table.nodeId));
  assert.match(patch.ooxml, /Updated cell/u);
}, async (document) => { await document.create.table({ rows: 2, columns: 2 }); }));

test("native coverage rejects unrepresentable document-wide parts", { timeout: 30_000 }, async () => withDocument(async (document, source) => {
  await document.customXml.parts.create({ content: "<record><value>Example</value></record>" });
  const update = await createFragmentUpdate(document, source);
  assert.deepEqual(update.patches, []);
  assert.match(update.unsupportedReason ?? "", /unsupported document part|without retained change records/u);
}));
