import { MAX_DOCX_BYTES } from "../../contract.js";
import type { FullDocumentUpdate } from "./types.js";

export function createFullDocumentUpdate(document: Uint8Array): FullDocumentUpdate {
  if (!document.length || document.length > MAX_DOCX_BYTES) throw new Error("The updated DOCX exceeds the 32 MB export limit.");
  return { documentBase64: Buffer.from(document).toString("base64") };
}
