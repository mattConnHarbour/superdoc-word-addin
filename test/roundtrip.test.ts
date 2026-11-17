import assert from "node:assert/strict";
import { test } from "node:test";
import type { RunRequest, RunResponse } from "../src/contract.js";
import { bytesToBase64, readCompressedDocument } from "../src/taskpane/office/files.js";
import type { Snapshot } from "../src/taskpane/operations.js";
import { updateDocument, WordPatchError, type DocumentBridge } from "../src/taskpane/roundtrip.js";

const snapshot: Snapshot = { base64: "UEsDBA==", digest: "a".repeat(64), blocks: [], url: "file:///document.docx" };
const signal = new AbortController().signal;

function response(request: RunRequest): RunResponse {
  return { requestId: request.requestId, sourceDigest: snapshot.digest, message: "Updated.", documentBase64: "UEsDBA==", toolsUsed: [], patches: [{ start: 0, deleteCount: 0, expected: [], ooxml: "fragment", insertCount: 1 }] };
}

function documentBridge(conflict = false) {
  const writes: string[] = [];
  let captures = 0;
  const bridge: DocumentBridge = {
    capture: async () => { captures += 1; return snapshot; },
    apply: async ({ beforeWrite }) => { if (conflict) return false; beforeWrite(); writes.push("fragment"); return true; },
    release: async () => {},
  };
  return { bridge, writes };
}

test("applies fragment patches after validating the request identity", async () => {
  const { bridge, writes } = documentBridge();
  const statuses: string[] = [];
  const outcome = await updateDocument("Shorten it", [], signal, (status) => statuses.push(status), bridge, async (request) => response(request));
  assert.equal(outcome.status, "applied");
  assert.deepEqual(writes, ["fragment"]);
  assert.equal(statuses.at(-1), "Updating Word…");
});

test("retains the agent version without writing when an affected Word range changes", async () => {
  const { bridge, writes } = documentBridge(true);
  const outcome = await updateDocument("Edit", [], signal, () => {}, bridge, async (request) => response(request));
  assert.equal(outcome.status, "conflict");
  assert.equal(outcome.response.documentBase64, "UEsDBA==");
  assert.deepEqual(writes, []);
});

test("rejects a response bound to another snapshot", async () => {
  const { bridge, writes } = documentBridge();
  await assert.rejects(updateDocument("Edit", [], signal, () => {}, bridge, async (request) => ({ ...response(request), sourceDigest: "b".repeat(64) })), /does not match/u);
  assert.deepEqual(writes, []);
});

test("cancellation after the backend completes does not write to Word", async () => {
  const controller = new AbortController();
  const { bridge, writes } = documentBridge();
  await assert.rejects(updateDocument("Edit", [], controller.signal, () => {}, bridge, async (request) => {
    controller.abort();
    return response(request);
  }), { name: "AbortError" });
  assert.deepEqual(writes, []);
});

test("a backend failure does not write to Word", async () => {
  const { bridge, writes } = documentBridge();
  await assert.rejects(updateDocument("Edit", [], signal, () => {}, bridge, async () => { throw new Error("offline"); }), /offline/u);
  assert.deepEqual(writes, []);
});

test("retains the exported version and warns about partial writes if Office rejects a patch", async () => {
  const { bridge } = documentBridge();
  bridge.apply = async ({ beforeWrite }) => { beforeWrite(); throw new Error("InvalidArgument"); };
  await assert.rejects(updateDocument("Edit", [], signal, () => {}, bridge, async (request) => response(request)), (error: unknown) => {
    assert.ok(error instanceof WordPatchError);
    assert.equal(error.response.documentBase64, "UEsDBA==");
    return true;
  });
});

test("Word file handles are closed when a later export slice fails", async () => {
  const closed: boolean[] = [];
  Object.defineProperty(globalThis, "Office", { configurable: true, value: {
    FileType: { Compressed: "compressed" }, AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
  } });
  const file: Office.File = {
    size: 4, sliceCount: 2,
    getSliceAsync(index, callback) {
      callback?.({
        status: index === 0 ? Office.AsyncResultStatus.Succeeded : Office.AsyncResultStatus.Failed,
        value: { data: [80, 75], index, size: 2 },
        error: { name: "TestError", code: 1, message: "Slice failed" },
        asyncContext: undefined,
        diagnostics: undefined,
      });
    },
    closeAsync() { closed.push(true); },
  };
  const host: Pick<Office.Document, "getFileAsync"> = {
    getFileAsync(_type: Office.FileType, options?: Office.GetFileOptions | ((result: Office.AsyncResult<Office.File>) => void), callback?: (result: Office.AsyncResult<Office.File>) => void) {
      const done = typeof options === "function" ? options : callback;
      done?.({ status: Office.AsyncResultStatus.Succeeded, value: file, error: { name: "", code: 0, message: "" }, asyncContext: undefined, diagnostics: undefined });
    },
  };
  try {
    await assert.rejects(readCompressedDocument(host), /Slice failed/u);
    assert.deepEqual(closed, [true]);
  } finally {
    Reflect.deleteProperty(globalThis, "Office");
  }
});

test("base64 conversion handles documents larger than the JavaScript argument limit", () => {
  const bytes = new Uint8Array(250_000).fill(243);
  assert.equal(bytesToBase64(bytes), Buffer.from(bytes).toString("base64"));
});
