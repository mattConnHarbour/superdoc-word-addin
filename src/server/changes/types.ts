import type { SuperDocDocument } from "@superdoc/sdk";

export type ChangeSnapshot = Awaited<ReturnType<SuperDocDocument["changes"]["snapshot"]>>;
export type ChangePage = Awaited<ReturnType<SuperDocDocument["changes"]["list"]>>;
export type ChangeEvent = ChangePage["items"][number]["changes"][number];
export type NativeBlock = Awaited<ReturnType<SuperDocDocument["blocks"]["list"]>>["blocks"][number];
export type SupportedBlock = NativeBlock & { nodeType: "paragraph" | "heading" | "listItem" | "table" };
export type OoxmlCapture = Awaited<ReturnType<SuperDocDocument["selection"]["extractOoxml"]>>;
export type FragmentCapture = Pick<OoxmlCapture, "fragment" | "dependencies">;
export type DiffSnapshot = Awaited<ReturnType<SuperDocDocument["diff"]["capture"]>>;
export interface ChangeSource {
  sessionId: string;
  journal: ChangeSnapshot;
  coverage: DiffSnapshot;
  blocks: SupportedBlock[];
  fingerprints: string[];
  protectedBlocks: Set<string>;
  owners: Map<string, string>;
}
export interface ChangeSpan {
  start: number;
  end: number;
  from: number;
  to: number;
}
