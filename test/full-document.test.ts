import assert from "node:assert/strict";
import { test } from "node:test";
import { MAX_FULL_REPLACEMENT_BYTES } from "../src/contract.js";
import { createFullDocumentUpdate } from "../src/server/operations.js";
import { applyFullDocumentUpdate } from "../src/taskpane/operations.js";

test("explicit full-document operations produce and insert the same DOCX with import options", async () => {
  const update = createFullDocumentUpdate(new Uint8Array([80, 75, 3, 4]));
  const inserted: unknown[] = [];
  let synced = false;
  const context = { document: { insertFileFromBase64: (...args: unknown[]) => inserted.push(args) }, sync: async () => { synced = true; } };
  Object.defineProperty(globalThis, "Word", { configurable: true, value: {
    InsertLocation: { replace: "Replace" }, run: async (callback: (batch: typeof context) => Promise<void>) => callback(context),
  } });
  try {
    await applyFullDocumentUpdate(update);
    assert.equal(inserted.length, 1);
    assert.equal(synced, true);
    assert.deepEqual(inserted[0], [update.documentBase64, "Replace", {
      importTheme: true, importStyles: true, importParagraphSpacing: true, importPageColor: true,
      importChangeTrackingMode: true, importCustomProperties: true, importCustomXmlParts: true, importDifferentOddEvenPages: true,
    }]);
  } finally { Reflect.deleteProperty(globalThis, "Word"); }
});

test("full-document apply rejects oversized or invalid files before calling Office", async () => {
  const tooLarge = new Uint8Array(MAX_FULL_REPLACEMENT_BYTES + 1);
  tooLarge.set([80, 75, 3, 4]);
  await assert.rejects(applyFullDocumentUpdate(createFullDocumentUpdate(tooLarge)), /no larger than 4 MB/u);
  await assert.rejects(applyFullDocumentUpdate(createFullDocumentUpdate(new Uint8Array([1, 2, 3, 4]))), /requires a DOCX file/u);
});
