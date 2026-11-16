import { captureTrackedDocument } from "../office/capture.js";
import { fragmentPreconditionsMatch, prepareFragmentPlan, queueFragmentWrites } from "../office/fragment-plan.js";
import { getTrackedDocument, releaseTrackedDocument } from "../office/session.js";
import type { FragmentApplication, Snapshot } from "./types.js";

export async function captureFragmentSource(): Promise<Snapshot> {
  await releaseTrackedDocument();
  try {
    return await Word.run(captureTrackedDocument);
  } catch (error) {
    await releaseTrackedDocument().catch(() => {});
    throw error;
  }
}

export async function applyFragmentUpdate(application: FragmentApplication): Promise<boolean> {
  const document = getTrackedDocument();
  if (!document || document.url !== application.source.url || Office.context.document.url !== application.source.url) return false;
  return Word.run(document.context.document, async (context) => {
    const plan = prepareFragmentPlan(application, document);
    await context.sync();
    application.signal.throwIfAborted();
    if (!fragmentPreconditionsMatch(plan, document)) return false;
    application.beforeWrite();
    queueFragmentWrites(plan);
    await context.sync();
    return true;
  });
}

export { releaseTrackedDocument as releaseFragmentSource } from "../office/session.js";
