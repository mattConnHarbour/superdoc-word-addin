import assert from "node:assert/strict";
import { test } from "node:test";
import { strToU8, zipSync } from "fflate";
import { packageRevision } from "../src/taskpane/document-revision.js";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const WE = "http://schemas.microsoft.com/office/webextensions/webextension/2010/11";
const TP = "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11";
const OFFICE_RELS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const ADDIN = "f3925cb5-51d4-432d-b93c-b72f5eca98e3";

function exportedParts(withTaskPane: boolean, rsid: string) {
  const parts: Record<string, Uint8Array> = {
    "word/document.xml": strToU8(`<w:document xmlns:w="${W}"><w:body><w:p w:rsidRDefault="${rsid}"><w:r><w:t>Original text</w:t></w:r></w:p></w:body></w:document>`),
    "word/settings.xml": strToU8(`<w:settings xmlns:w="${W}"><w:rsids><w:rsid w:val="${rsid}"/></w:rsids><w:defaultTabStop w:val="720"/></w:settings>`),
    "[Content_Types].xml": strToU8(`<Types xmlns="${CT}"><Override PartName="/word/document.xml" ContentType="document"/>${withTaskPane ? '<Override PartName="/word/webextensions/taskpanes.xml" ContentType="taskpanes"/><Override PartName="/word/webextensions/webextension1.xml" ContentType="webextension"/>' : ""}</Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${OFFICE_RELS}/officeDocument" Target="word/document.xml"/>${withTaskPane ? '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="word/webextensions/taskpanes.xml"/>' : ""}<Relationship Id="${withTaskPane ? "rId3" : "rId2"}" Type="properties" Target="docProps/core.xml"/></Relationships>`),
  };
  if (withTaskPane) {
    parts["word/webextensions/webextension1.xml"] = strToU8(`<we:webextension xmlns:we="${WE}" id="session"><we:reference id="${ADDIN}"/><we:properties/></we:webextension>`);
    parts["word/webextensions/taskpanes.xml"] = strToU8(`<tp:taskpanes xmlns:tp="${TP}" xmlns:r="${OFFICE_RELS}"><tp:taskpane visibility="0" width="350"><tp:webextensionref r:id="rId1"/></tp:taskpane></tp:taskpanes>`);
    parts["word/webextensions/_rels/taskpanes.xml.rels"] = strToU8(`<Relationships xmlns="${RELS}"><Relationship Id="rId1" Target="webextension1.xml" Type="http://schemas.microsoft.com/office/2011/relationships/webextension"/></Relationships>`);
  }
  return parts;
}

async function revision(parts: Record<string, Uint8Array>) {
  return packageRevision(new Uint8Array(zipSync(parts)));
}

test("untouched Word exports stay equal when session IDs and this task pane's persistence change", async () => {
  assert.equal(await revision(exportedParts(true, "00AABBCC")), await revision(exportedParts(false, "00112233")));
});

test("real formatting, text, and document setting edits still change the revision", async () => {
  const baseline = await revision(exportedParts(true, "00AABBCC"));
  for (const [part, from, to] of [
    ["word/document.xml", "Original text", "New text"],
    ["word/document.xml", "<w:r>", "<w:r><w:rPr><w:b/></w:rPr>"],
    ["word/settings.xml", 'w:val="720"', 'w:val="1440"'],
  ]) {
    assert.ok(part && from && to);
    const parts = exportedParts(false, "00112233");
    const bytes = parts[part];
    assert.ok(bytes);
    parts[part] = strToU8(new TextDecoder().decode(bytes).replace(from, to));
    assert.notEqual(await revision(parts), baseline);
  }
});

test("persistence belonging to another add-in remains part of the revision", async () => {
  const parts = exportedParts(true, "00AABBCC");
  const extension = parts["word/webextensions/webextension1.xml"];
  assert.ok(extension);
  parts["word/webextensions/webextension1.xml"] = strToU8(new TextDecoder().decode(extension).replace(ADDIN, "another-addin"));
  assert.notEqual(await revision(parts), await revision(exportedParts(false, "00112233")));
});

test("XML prefix and attribute serialization changes are harmless while text whitespace remains meaningful", async () => {
  const before = { "word/document.xml": strToU8(`<w:document xmlns:w="${W}"><w:body><w:p w:rsidR="AAA" w:rsidP="BBB"><w:r><w:t xml:space="preserve"> word </w:t></w:r></w:p></w:body></w:document>`) };
  const after = { "word/document.xml": strToU8(`<x:document xmlns:x="${W}"><x:body><x:p x:rsidP="CCC" x:rsidR="DDD"><x:r><x:t xml:space="preserve"> word </x:t></x:r></x:p></x:body></x:document>`) };
  assert.equal(await revision(before), await revision(after));
  after["word/document.xml"] = strToU8(new TextDecoder().decode(after["word/document.xml"]).replace("> word <", ">word<"));
  assert.notEqual(await revision(before), await revision(after));
});
