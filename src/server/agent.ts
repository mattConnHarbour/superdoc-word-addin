import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { createSuperDocClient } from "@superdoc/sdk";
import OpenAI from "openai";
import { captureChangeSource, createFragmentUpdate, createFullDocumentUpdate } from "./operations.js";
import { MAX_DOCX_BYTES, type RunRequest, type RunResponse } from "../contract.js";
import { runAgentTools, type Complete } from "./agent-loop.js";
export type { Complete } from "./agent-loop.js";

export function createCompletion(apiKey: string): Complete {
  const openai = new OpenAI({ apiKey, maxRetries: 0 });
  return (messages, tools, signal) => openai.chat.completions.create({
    model: "gpt-5.4",
    messages,
    tools,
    parallel_tool_calls: false,
    store: false,
  }, { signal });
}

export async function runAgent(request: RunRequest, complete: Complete, signal: AbortSignal): Promise<RunResponse> {
  signal.throwIfAborted();
  const source = Buffer.from(request.documentBase64, "base64");
  if (source.length < 4 || source.length > MAX_DOCX_BYTES || source.toString("base64") !== request.documentBase64 || source.readUInt32LE(0) !== 0x04034b50) {
    throw new Error("Send a DOCX no larger than 32 MB.");
  }
  const directory = await mkdtemp(path.join(tmpdir(), "superdoc-word-agent-"));
  const client = createSuperDocClient({ runtime: "v2", stateDir: path.join(directory, "state"), requestTimeoutMs: 30_000 });
  try {
    const input = path.join(directory, "input.docx");
    const output = path.join(directory, "output.docx");
    await writeFile(input, source, { mode: 0o600 });
    const document = await client.open({ doc: input, runtime: "v2" });
    const changeSource = await captureChangeSource(document, source);
    const result = await runAgentTools(document, request, complete, signal);
    signal.throwIfAborted();
    await document.save({ out: output });
    const updated = await readFile(output);
    return {
      requestId: request.requestId,
      sourceDigest: createHash("sha256").update(source).digest("hex"),
      message: result.message,
      toolsUsed: result.toolsUsed,
      ...createFullDocumentUpdate(updated),
      ...await createFragmentUpdate(document, changeSource),
    };
  } finally {
    try { await client.dispose(); }
    finally { await rm(directory, { recursive: true, force: true }); }
  }
}
