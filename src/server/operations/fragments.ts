import { errorMessage } from "../../contract.js";
import { assertCompatiblePackage, assertSupportedSpans, openPatchContext } from "../fragments/compatibility.js";
import { findChangedSpans } from "../fragments/diff.js";
import { serializePatches } from "../fragments/serialize.js";
import type { DocumentVersions, FragmentUpdate } from "./types.js";

export async function createFragmentUpdate(versions: DocumentVersions): Promise<FragmentUpdate> {
  try {
    const context = openPatchContext(versions);
    await assertCompatiblePackage(context);
    const spans = findChangedSpans(context.before, context.after);
    assertSupportedSpans(context, spans);
    return { patches: await serializePatches(context, spans) };
  } catch (error) {
    return { patches: [], unsupportedReason: errorMessage(error) };
  }
}
