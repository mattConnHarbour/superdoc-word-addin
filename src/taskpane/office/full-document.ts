import { MAX_FULL_REPLACEMENT_BYTES, type FullDocumentUpdate } from "../../contract.js";

export function validateFullDocumentUpdate(update: FullDocumentUpdate): void {
  const bytes = Uint8Array.from(atob(update.documentBase64), (character) => character.charCodeAt(0));
  if (bytes.length < 4 || bytes.length > MAX_FULL_REPLACEMENT_BYTES) throw new Error("Whole-document replacement requires a DOCX no larger than 4 MB.");
  if (bytes[0] !== 80 || bytes[1] !== 75 || bytes[2] !== 3 || bytes[3] !== 4) throw new Error("Whole-document replacement requires a DOCX file.");
}

export function queueFullDocumentReplacement(context: Word.RequestContext, update: FullDocumentUpdate): void {
  context.document.insertFileFromBase64(update.documentBase64, Word.InsertLocation.replace, {
    importTheme: true,
    importStyles: true,
    importParagraphSpacing: true,
    importPageColor: true,
    importChangeTrackingMode: true,
    importCustomProperties: true,
    importCustomXmlParts: true,
    importDifferentOddEvenPages: true,
  });
}
