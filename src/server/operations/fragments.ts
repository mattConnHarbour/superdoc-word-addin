import type { SuperDocDocument } from "@superdoc/sdk";
import { errorMessage, type FragmentUpdate } from "../../contract.js";
import { listBodyBlocks } from "../changes/blocks.js";
import { assertChangeCoverage } from "../changes/coverage.js";
import { resolveChangedBlocks } from "../changes/extract.js";
import { changedIds, readJournal } from "../changes/journal.js";
import { planJournalPatches } from "../changes/plan.js";
import { serializeJournalPatches } from "../changes/serialize.js";
import { readChangeSource } from "../changes/source.js";
import type { ChangeSource } from "../changes/types.js";

export async function captureChangeSource(document: SuperDocDocument, source: Uint8Array): Promise<ChangeSource> {
  return readChangeSource(document, source);
}

export async function createFragmentUpdate(document: SuperDocDocument, source: ChangeSource): Promise<FragmentUpdate> {
  try {
    if (document.sessionId !== source.sessionId) throw new Error("The change source belongs to another SDK document session.");
    const end = await document.changes.snapshot();
    const events = await readJournal(document.changes, source.journal, end);
    if (!await assertChangeCoverage(document, source, events.length)) return { patches: [] };
    const after = await listBodyBlocks(document);
    const dirty = await resolveChangedBlocks(document, source, after, changedIds(events));
    const spans = planJournalPatches(source.blocks, after, dirty);
    if (!spans.length) throw new Error("The SDK changes could not be mapped to supported Word fragments.");
    const patches = await serializeJournalPatches(document, source, after, spans);
    if ((await document.changes.snapshot()).revision !== end.revision) throw new Error("The SDK document changed while its fragments were being prepared.");
    return { patches };
  } catch (error) {
    return { patches: [], unsupportedReason: errorMessage(error) };
  }
}
