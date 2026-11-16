import type { Element } from "@xmldom/xmldom";

export interface ChangeSpan {
  start: number;
  end: number;
  from: number;
  to: number;
}
export interface PatchContext {
  source: Record<string, Uint8Array>;
  baseline: Record<string, Uint8Array>;
  result: Record<string, Uint8Array>;
  sourceBlocks: Element[];
  before: Element[];
  after: Element[];
}
