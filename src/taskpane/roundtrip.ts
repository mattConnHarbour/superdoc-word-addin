import { RUN_TIMEOUT_MS, runResponseSchema, type ChatMessage, type RunRequest, type RunResponse } from "../contract.js";
import { captureFragmentSource, applyFragmentUpdate, releaseFragmentSource, type Snapshot, type FragmentApplication } from "./operations.js";

export interface DocumentBridge {
  capture(): Promise<Snapshot>;
  apply(application: FragmentApplication): Promise<boolean>;
  release(): Promise<void>;
}
export type AgentTransport = (request: RunRequest, signal: AbortSignal) => Promise<RunResponse>;
export type Outcome = { status: "applied" | "conflict" | "unsupported" | "unchanged"; response: RunResponse };

export class WordPatchError extends Error {
  constructor(readonly response: RunResponse, cause: unknown) {
    super("Word could not finish applying the fragments. Some edits may have been applied. Inspect Word before retrying, or download the agent version.", { cause });
    this.name = "WordPatchError";
  }
}

export async function requestAgent(request: RunRequest, signal: AbortSignal): Promise<RunResponse> {
  const response = await fetch("/api/run", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request),
    signal: AbortSignal.any([signal, AbortSignal.timeout(RUN_TIMEOUT_MS + 5_000)]),
  });
  if (!response.ok) {
    const body: unknown = await response.json().catch(() => null);
    if (body && typeof body === "object" && "error" in body && typeof body.error === "string") throw new Error(body.error);
    throw new Error(`The backend returned HTTP ${response.status}.`);
  }
  return runResponseSchema.parse(await response.json());
}

export async function updateDocument(
  prompt: string,
  history: ChatMessage[],
  signal: AbortSignal,
  setStatus: (text: string) => void,
  bridge: DocumentBridge = { capture: captureFragmentSource, apply: applyFragmentUpdate, release: releaseFragmentSource },
  transport: AgentTransport = requestAgent,
): Promise<Outcome> {
  setStatus("Reading the Word document…");
  const source = await bridge.capture();
  try {
    signal.throwIfAborted();
    const requestId = crypto.randomUUID();
    setStatus("Agent is working on your document…");
    const response = await transport({ requestId, prompt, history: history.slice(-12), documentBase64: source.base64 }, signal);
    signal.throwIfAborted();
    if (response.requestId !== requestId || response.sourceDigest !== source.digest) throw new Error("The agent response does not match this document request.");
    if (response.unsupportedReason) return { status: "unsupported", response };
    if (!response.patches.length) return { status: "unchanged", response };
    setStatus("Checking affected ranges in Word…");
    let writing = false;
    try {
      const applied = await bridge.apply({ source, patches: response.patches, signal, beforeWrite: () => {
        signal.throwIfAborted();
        writing = true;
        setStatus("Updating Word…");
      } });
      return { status: applied ? "applied" : "conflict", response };
    } catch (error) {
      if (writing) throw new WordPatchError(response, error);
      throw error;
    }
  } finally {
    await bridge.release().catch(() => {});
  }
}
