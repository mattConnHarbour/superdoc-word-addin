import type { SuperDocDocument } from "@superdoc/sdk";
import type { ChangeEvent, ChangeSnapshot } from "./types.js";

export async function readJournal(changes: Pick<SuperDocDocument["changes"], "list">, before: ChangeSnapshot, after: ChangeSnapshot): Promise<ChangeEvent[]> {
  const events: ChangeEvent[] = [];
  let cursor = before.cursor;
  for (let pageNumber = 0; cursor !== after.cursor; pageNumber += 1) {
    if (pageNumber >= 1_000) throw new Error("The SDK change journal exceeded the request limit.");
    const page = await changes.list({ ...(cursor ? { after: cursor } : {}), limit: 100 });
    if (page.truncated) throw new Error("The SDK evicted change records. Retry against the latest Word document.");
    events.push(...page.items.flatMap((item) => item.changes));
    if (!page.cursor || page.cursor === cursor) throw new Error("The SDK returned an incomplete change journal.");
    cursor = page.cursor;
  }
  return events;
}

export function changedIds(events: ChangeEvent[]): Set<string> {
  const ids = new Set<string>();
  for (const event of events) {
    if (event.kind === "sidecar") throw new Error("Comment, bookmark, and revision metadata changes are not supported by this Word adapter.");
    if (event.story.storyType !== "body") throw new Error("The fragment adapter only applies document body changes.");
    if (event.kind === "table") ids.add(event.tableId);
    else {
      ids.add(event.blockId);
      if (event.kind === "split") ids.add(event.tailBlockId);
      if (event.kind === "merge") ids.add(event.mergedBlockId);
    }
  }
  return ids;
}
