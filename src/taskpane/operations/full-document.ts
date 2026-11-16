import { queueFullDocumentReplacement, validateFullDocumentUpdate } from "../office/full-document.js";
import type { FullDocumentUpdate } from "./types.js";

export async function applyFullDocumentUpdate(update: FullDocumentUpdate): Promise<void> {
  validateFullDocumentUpdate(update);
  await Word.run(async (context) => {
    queueFullDocumentReplacement(context, update);
    await context.sync();
  });
}
