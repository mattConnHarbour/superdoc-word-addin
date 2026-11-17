import type { SuperDocDocument } from "@superdoc/sdk";
import type { ChangeSource } from "./types.js";

export async function assertChangeCoverage(document: SuperDocDocument, source: ChangeSource, eventCount: number): Promise<boolean> {
  const comparison = await document.diff.compare({ targetSnapshot: source.coverage });
  if (comparison.summary.changedComponents.some((component) => component !== "body" && component !== "parts")) throw new Error("This edit changes document-wide styles, numbering, comments, or headers outside the supported body fragments.");
  for (const part of comparison.partChanges ?? []) {
    const uri = `/${part.partUri.replace(/^\//u, "")}`;
    if (["/word/document.xml", "/word/_rels/document.xml.rels", "/[Content_Types].xml"].includes(uri)) continue;
    if (uri.startsWith("/word/media/") && part.kind === "added") continue;
    throw new Error(`The edit changes an unsupported document part: ${uri}`);
  }
  if (comparison.summary.hasChanges && !eventCount) throw new Error("SuperDoc changed the document without retained change records. Word was not updated.");
  return comparison.summary.hasChanges;
}
