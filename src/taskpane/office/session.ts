export interface TrackedBlock {
  range: Word.Range;
  canonical: string;
}
export interface TrackedDocument {
  context: Word.RequestContext;
  blocks: TrackedBlock[];
  url: string;
  resources: Map<string, string>;
}
let session: TrackedDocument | undefined;

export function getTrackedDocument(): TrackedDocument | undefined {
  return session;
}

export function setTrackedDocument(value: TrackedDocument): void {
  session = value;
}

export async function releaseTrackedDocument(): Promise<void> {
  const current = session;
  session = undefined;
  if (!current) return;
  await Word.run(current.context.document, async (context) => {
    for (const block of current.blocks) block.range.untrack();
    await context.sync();
  });
}
