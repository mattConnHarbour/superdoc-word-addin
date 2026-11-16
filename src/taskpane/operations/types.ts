import type { FragmentPatch } from "../../contract.js";

export type { FullDocumentUpdate } from "../../contract.js";
export interface Snapshot {
  base64: string;
  digest: string;
  blocks: { kind: "p" | "tbl"; fingerprint: string }[];
  url: string;
}
export interface FragmentApplication {
  source: Snapshot;
  patches: FragmentPatch[];
  signal: AbortSignal;
  beforeWrite(): void;
}
