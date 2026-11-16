export type { FragmentUpdate, FullDocumentUpdate } from "../../contract.js";

export interface DocumentVersions {
  source: Uint8Array;
  baseline: Uint8Array;
  updated: Uint8Array;
}
