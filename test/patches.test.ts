import assert from "node:assert/strict";
import { test } from "node:test";
import { strToU8, zipSync } from "fflate";
import { createFragmentUpdate } from "../src/server/operations.js";
import { blocks, canonical, flatDocument, parseXml, PKG, W } from "../src/ooxml.js";

const paragraph = (id: string, text: string, properties = "") => `<w:p w14:paraId="${id}">${properties}<w:r><w:t>${text}</w:t></w:r></w:p>`;
const first = paragraph("11111111", "TechCraft builds tools.");
const second = paragraph("22222222", "Unchanged text.");
const third = paragraph("33333333", "Last paragraph.");
function doc(content: string, extra: Record<string, Uint8Array> = {}) {
  return new Uint8Array(zipSync({
    "[Content_Types].xml": strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
    "_rels/.rels": strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    "word/document.xml": strToU8(`<w:document xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${content}<w:sectPr/></w:body></w:document>`),
    ...extra,
  }));
}
const source = doc(first + second + third);

function createPatches(source: Uint8Array, baseline: Uint8Array, updated: Uint8Array) {
  return createFragmentUpdate({ source, baseline, updated });
}

test("a text replacement exports only its changed paragraph", async () => {
  const result = await createPatches(source, source, doc(first.replace("TechCraft", "MattWorks") + second + third));
  assert.equal(result.unsupportedReason, undefined);
  assert.equal(result.patches.length, 1);
  const patch = result.patches[0];
  assert.ok(patch);
  assert.equal(patch.start, 0);
  assert.equal(patch.deleteCount, 1);
  assert.equal(patch.expected.length, 1);
  assert.equal(blocks(flatDocument(patch.ooxml)).length, 1);
  assert.match(patch.ooxml, /MattWorks/u);
  assert.doesNotMatch(patch.ooxml, /Unchanged text|Last paragraph/u);
});

test("insertions, deletions, and inline formatting become separate range patches", async () => {
  const replacement = paragraph("22222222", "Unchanged text.", "<w:pPr><w:jc w:val=\"center\"/></w:pPr>");
  const inserted = paragraph("44444444", "New paragraph.");
  const result = await createPatches(source, source, doc(inserted + replacement + third));
  assert.equal(result.unsupportedReason, undefined);
  assert.deepEqual(result.patches.map((patch) => [patch.start, patch.deleteCount, patch.insertCount]), [[0, 2, 2]]);
  const removed = await createPatches(source, source, doc(first + third));
  assert.deepEqual(removed.patches.map((patch) => [patch.start, patch.deleteCount, patch.ooxml]), [[1, 1, ""]]);
  const added = await createPatches(source, source, doc(first + inserted + second + third));
  assert.deepEqual(added.patches.map((patch) => [patch.start, patch.deleteCount, patch.insertCount]), [[1, 0, 1]]);
});

test("editing a cell patches its containing table and leaves other blocks out", async () => {
  const table = `<w:tbl><w:tr><w:tc>${paragraph("44444444", "Cell text")}</w:tc></w:tr></w:tbl>`;
  const before = doc(first + table + third);
  const result = await createPatches(before, before, doc(first + table.replace("Cell text", "Edited cell") + third));
  assert.equal(result.unsupportedReason, undefined);
  assert.deepEqual(result.patches.map((patch) => [patch.start, patch.deleteCount, patch.insertCount]), [[1, 1, 1]]);
  assert.equal(blocks(flatDocument(result.patches[0]?.ooxml ?? ""))[0]?.localName, "tbl");
});

test("SDK open/save normalization is excluded by diffing the backend baseline", async () => {
  const baseline = doc(first.replace("<w:r>", "<w:r w:rsidR=\"12345678\">") + second + third);
  assert.deepEqual(await createPatches(source, baseline, baseline), { patches: [] });
});

test("Word proofing boundaries and equivalent split runs do not create a conflict", () => {
  const whole = parseXml(`<w:p xmlns:w="${W}"><w:r><w:t>TechCraft builds tools.</w:t></w:r></w:p>`).documentElement;
  const split = parseXml(`<w:p xmlns:w="${W}"><w:proofErr w:type="spellStart"/><w:r><w:t>TechCraft</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r><w:t xml:space="preserve"> builds tools.</w:t></w:r></w:p>`).documentElement;
  assert.ok(whole);
  assert.ok(split);
  assert.equal(canonical(whole), canonical(split));
  const bold = parseXml(`<w:p xmlns:w="${W}"><w:r><w:rPr><w:b/></w:rPr><w:t>TechCraft builds tools.</w:t></w:r></w:p>`).documentElement;
  assert.ok(bold);
  assert.notEqual(canonical(whole), canonical(bold));
});

test("document-wide edits and range-spanning annotations never trigger a full replacement", async () => {
  const before = doc(first + second, { "word/styles.xml": strToU8('<styles name="before"/>') });
  const after = doc(first + second, { "word/styles.xml": strToU8('<styles name="after"/>') });
  const result = await createPatches(before, before, after);
  assert.deepEqual(result.patches, []);
  assert.match(result.unsupportedReason ?? "", /document-wide/u);
  const marked = first.replace("<w:r>", '<w:bookmarkStart w:id="1" w:name="bookmark"/><w:r>');
  const markedSource = doc(marked + second);
  const protectedResult = await createPatches(markedSource, markedSource, doc(marked.replace("TechCraft", "MattWorks") + second));
  assert.deepEqual(protectedResult.patches, []);
  assert.match(protectedResult.unsupportedReason ?? "", /bookmark/u);
});

test("fragments contain Flat OPC relationships and required style dependencies", async () => {
  const extras = {
    "word/_rels/document.xml.rels": strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'),
    "word/styles.xml": strToU8(`<w:styles xmlns:w="${W}"/>`),
  };
  const before = doc(first + second, extras);
  const after = doc(first.replace("TechCraft", "MattWorks") + second, extras);
  const result = await createPatches(before, before, after);
  assert.equal(result.unsupportedReason, undefined);
  const xml = parseXml(result.patches[0]?.ooxml ?? "");
  const names = Array.from(xml.getElementsByTagNameNS(PKG, "part")).map((part) => part.getAttributeNS(PKG, "name"));
  assert.deepEqual(names.sort(), ["/_rels/.rels", "/word/_rels/document.xml.rels", "/word/document.xml", "/word/styles.xml"]);
});
