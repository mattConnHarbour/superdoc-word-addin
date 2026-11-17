import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { strToU8, zipSync } from "fflate";
import { blocks, body, fingerprint, flatDocument, mainDocument, serialize, unpack } from "../src/ooxml.js";
import { test } from "node:test";
import { createSuperDocClient, dispatchSuperDocTool } from "@superdoc/sdk";
import type OpenAI from "openai";
import { runAgent, type Complete } from "../src/server/agent.js";
import { createApp } from "../src/server/app.js";
import { ADDIN_ORIGIN, runResponseSchema } from "../src/contract.js";
import { digestBytes } from "../src/digest.js";
import type { Snapshot } from "../src/taskpane/operations.js";
import { updateDocument } from "../src/taskpane/roundtrip.js";

function completion(message: OpenAI.ChatCompletionMessage): OpenAI.ChatCompletion {
  return { id: "test", object: "chat.completion", created: 0, model: "deterministic-test", choices: [{ index: 0, finish_reason: message.tool_calls ? "tool_calls" : "stop", logprobs: null, message }] };
}

test("edits through real SDK tools and applies only the returned OOXML fragments to a reopenable DOCX", { timeout: 90_000 }, async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "word-agent-proof-"));
  const client = createSuperDocClient({ runtime: "v2", stateDir: path.join(directory, "state") });
  try {
    const input = path.join(directory, "blank.docx");
    const document = await client.open({});
    await document.save({ out: input });
    const source = new Uint8Array(await readFile(input));
    const sourceSnapshot: Snapshot = {
      base64: Buffer.from(source).toString("base64"), digest: await digestBytes(source), blocks: await Promise.all(blocks(mainDocument(unpack(source))).map(async (block) => ({ kind: "p", fingerprint: await fingerprint(block) }))), url: "file:///blank.docx",
    };
    const actions = [
      { name: "superdoc_inspect", args: { blockLimit: 10 } },
      { name: "superdoc_perform_action", args: { action: "insert_paragraphs", texts: ["Edited by the SuperDoc backend agent."] } },
      { name: "superdoc_inspect", args: { blockLimit: 10 } },
    ];
    let turn = 0;
    const complete: Complete = async (messages, tools) => {
      assert.equal(tools.length, 4);
      const action = actions[turn++];
      if (!action) {
        assert.match(JSON.stringify(messages.at(-1)), /Edited by the SuperDoc backend agent/u);
        return completion({ role: "assistant", content: "Added the requested paragraph.", refusal: null });
      }
      return completion({ role: "assistant", content: null, refusal: null, tool_calls: [{ id: `call-${turn}`, type: "function", function: { name: action.name, arguments: JSON.stringify(action.args) } }] });
    };
    const app = createApp((request, signal) => runAgent(request, complete, signal));
    const output = path.join(directory, "word-patched.docx");
    let patchesApplied = 0;
    const result = await updateDocument("Add a paragraph", [], new AbortController().signal, () => {}, {
      capture: async () => sourceSnapshot,
      apply: async ({ patches, beforeWrite }) => {
        beforeWrite();
        const parts = unpack(source);
        const xml = mainDocument(parts);
        const documentBody = body(xml);
        const originalBlocks = blocks(xml);
        for (const patch of [...patches].reverse()) {
          patchesApplied += 1;
          const anchor = originalBlocks[patch.start] ?? documentBody.lastChild;
          if (patch.ooxml) for (const block of blocks(flatDocument(patch.ooxml))) documentBody.insertBefore(xml.importNode(block, true), anchor);
          for (const block of originalBlocks.slice(patch.start, patch.start + patch.deleteCount)) documentBody.removeChild(block);
        }
        parts["word/document.xml"] = strToU8(serialize(xml));
        await writeFile(output, zipSync(parts));
        return true;
      },
      release: async () => {},
    }, async (request) => {
      const response = await app.fetch(new Request(`${ADDIN_ORIGIN}/api/run`, {
        method: "POST", headers: { host: "localhost:3000", origin: ADDIN_ORIGIN, "Content-Type": "application/json" }, body: JSON.stringify(request),
      }));
      assert.equal(response.status, 200, await response.clone().text());
      return runResponseSchema.parse(await response.json());
    });
    assert.equal(result.status, "applied", result.response.unsupportedReason);
    assert.equal(patchesApplied, 1);
    assert.equal(result.response.unsupportedReason, undefined);
    assert.equal(result.response.patches[0]?.deleteCount, 0);
    const exported = await client.open({ doc: output, runtime: "v2" });
    const inspection = await dispatchSuperDocTool(exported, "superdoc_inspect", { blockLimit: 10 }, { preset: "full" });
    assert.match(JSON.stringify(inspection), /Edited by the SuperDoc backend agent/u);
  } finally {
    try { await client.dispose(); }
    finally { await rm(directory, { recursive: true, force: true }); }
  }
});
