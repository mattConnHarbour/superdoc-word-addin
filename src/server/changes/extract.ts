import type { SuperDocDocument } from "@superdoc/sdk";
import { captureBlocks, paragraphIds } from "./blocks.js";
import type { ChangeSource, FragmentCapture, SupportedBlock } from "./types.js";

async function extractBlock(document: Pick<SuperDocDocument, "selection">, block: SupportedBlock): Promise<FragmentCapture> {
  const node = { kind: "block", nodeType: block.nodeType, nodeId: block.nodeId };
  const capture = await document.selection.extractOoxml({ at: {
    kind: "selection",
    start: { kind: "nodeEdge", node: { ...node, kind: "block" }, edge: "before" },
    end: { kind: "nodeEdge", node: { ...node, kind: "block" }, edge: "after" },
  } });
  return capture;
}

export async function extractSpan(document: Pick<SuperDocDocument, "selection">, selected: SupportedBlock[]): Promise<FragmentCapture> {
  if (!selected.length) throw new Error("An inserted fragment has no SDK block address.");
  const captures: FragmentCapture[] = [];
  const namespaces: Record<string, string> = {};
  for (const block of selected) {
    const capture = await extractBlock(document, block);
    for (const [prefix, uri] of Object.entries(capture.fragment.namespaces)) {
      if (namespaces[prefix] && namespaces[prefix] !== uri) throw new Error("SDK fragments use incompatible XML namespace bindings.");
      namespaces[prefix] = uri;
    }
    captures.push(capture);
  }
  return { fragment: { placement: "blocks", xml: captures.map((capture) => capture.fragment.xml).join(""), namespaces },
    dependencies: captures.flatMap((capture) => capture.dependencies),
  };
}

export async function resolveChangedBlocks(document: SuperDocDocument, source: ChangeSource, after: SupportedBlock[], ids: Set<string>): Promise<Set<string>> {
  const owners = new Map(source.owners);
  for (const block of after) owners.set(block.nodeId, block.nodeId);
  const liveIds = new Set(after.map((block) => block.nodeId));
  if (Array.from(ids).some((id) => !owners.has(id)) || source.blocks.some((block) => block.nodeType === "table" && !liveIds.has(block.nodeId))) {
    for (const table of after.filter((block) => block.nodeType === "table")) {
      const capture = await extractSpan(document, [table]);
      for (const id of paragraphIds(captureBlocks(capture))) { owners.set(id, table.nodeId); owners.set(`tbl:${id}`, table.nodeId); }
    }
  }
  const dirty = new Set<string>();
  for (const id of ids) {
    const originalOwner = source.owners.get(id);
    if (originalOwner) dirty.add(originalOwner);
    const owner = owners.get(id);
    if (owner) dirty.add(owner);
  }
  return dirty;
}
