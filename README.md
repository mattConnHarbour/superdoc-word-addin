# SuperDoc Word Add-in

A desktop Microsoft Word task pane with agent chat. Each message exports the
current DOCX, sends it to a local backend agent, edits it through the SuperDoc
SDK, and applies changed body fragments to the open Word document with Office.js.

Based on SuperDoc's [Embedded Agent Word Add-in](https://github.com/superdoc/docx-embedded-agents/tree/5bb32d96b500e5a01b9c570adac95a00b7f56025/agent-word-addin).
The manifest and compressed-file export started from that example. The chat
bridge uses tracked Word ranges and OOXML fragment patches. Whole-document
replacement is exposed separately for explicit callers and is never a fallback. Chat
and the SDK backend replace the example's embedded-registrar discovery and
automatic document polling.
The correctly sized icons are reused from the
[Agent Revision Markup add-in](https://github.com/superdoc/agent-revision-markup/tree/0f7488334180a8c46dbc3eaf97d713acf2d46893/apps/word-addin)
and its demo favicon.

## Installation and reuse

See [AGENTS.md](AGENTS.md) for local setup, certificates, macOS/Windows add-in
installation, and troubleshooting. It also explains how to extract the separated
backend and Office.js operations into your own application: which files and
dependencies to copy, how to connect your agent and API, and how to choose
fragment updates or explicit full-document replacement.

Clone this repository, then follow [AGENTS.md](AGENTS.md#install-and-run-the-add-in)
to install dependencies, configure the backend key, and sideload the manifest:

```sh
git clone https://github.com/mattConnHarbour/superdoc-word-addin.git
cd superdoc-word-addin
```

## Document flow and change tracking

See the [fragment patching flow diagram](docs/diagrams/word-fragment-patching-v3.png)
for a visual overview of capture, backend edits, position tracking, and application.

1. Before the request, the pane reads the body OOXML, tracks the top-level Word
   paragraph/table ranges, and exports the source DOCX with
   `getFileAsync(Compressed)`. Source hashes bind the response and target ranges
   to that capture. Live body OOXML is checked again after export to reject edits
   during capture. Word preconditions use the live OOXML baseline; DOCX hashes
   use the file export. These two export formats are not compared for equality.
2. The backend opens the DOCX in a fresh SuperDoc SDK session.
   `captureChangeSource(document, source)` captures a native Changes API cursor,
   lists the SDK body block addresses, and maps them to the original Word block
   hashes. The checkpoint stays in backend memory in that same document session.
3. The agent inspects and edits through SDK tools. After it finishes,
   `createFragmentUpdate(document, checkpoint)` reads `changes.list` from the
   saved cursor to a fixed endpoint, following all resume cursors. Truncated or
   incomplete journals fail without producing patches.
4. Native journal records identify affected blocks, including text, formatting,
   creation, removal, split, merge, move, and table events. The planner maps
   nested cell changes to their containing table and uses unchanged block IDs
   as Word range anchors. It combines adjacent affected blocks; it does not
   compare XML or run a sequence-diff algorithm to discover changes.
5. `selection.extractOoxml` exports each selected SDK block and its dependencies.
   Complete blocks are extracted individually, then combined into the planned
   span. A separate serializer wraps SDK-provided XML and dependency data into
   Flat OPC packages for Office.js. It does not discover dependencies by walking
   the exported DOCX.
6. SuperDoc's native `diff.capture` checkpoint and `diff.compare` summary provide
   a separate completeness guard for net-zero edits and unsupported global
   components/parts. These opaque native payloads stay engine-owned. They do not
   select fragments, and `diff.apply` is not called. No baseline DOCX export or
   custom XML diff remains.
7. The backend saves the final DOCX once for recovery download, returns the
   patches and request/source identity, then disposes the SDK and temporary
   files. The normal chat path never applies this full DOCX to Word.
8. Office.js checks every affected tracked range and insertion gap before
   queuing any writes, then applies the patches in reverse order in one batch.
   It releases tracked ranges in `finally`. Unchanged blocks elsewhere survive.

Only the original Word DOCX capture and final recovery DOCX export remain.
The journal cursor, SDK block-address checkpoint, and native coverage snapshot
are in-memory API state, not intermediate DOCX exports. A target conflict,
missing journal history, unsupported component, or SDK-session mismatch never
falls back to whole-document replacement.

## How the SDK pairs with Office.js

The adapter translates resulting document structure, not every SDK operation.
The SDK owns editing semantics; OOXML is the interchange format; Office.js owns
locating, checking, and updating the live Word ranges.

| SuperDoc result | Office.js application |
| --- | --- |
| Text/run formatting or paragraph property change | `Range.insertOoxml(fragment, "Replace")` on the containing paragraph span |
| Paragraphs/table inserted | `Range.insertOoxml(fragment, "Before" / "After")` at a tracked neighbor |
| Paragraphs/table removed | `Range.delete()` on the affected span |
| Cell text, shading, borders, or table structure changed | Replace the containing top-level table fragment |
| No effective body edit | No Word write |
| Header, section, global style, or other unsupported edit | No Word write; explain and offer the result DOCX download |

A patch contains `start`, `deleteCount`, `expected` source hashes, `insertCount`,
and `ooxml`. Indices address the original snapshot's ordered body blocks, then
resolve to tracked Office ranges. They are never reapplied as current paragraph
indices after the backend returns. The pane validates ordering, bounds, expected
hashes, and the inserted block count before queuing writes.

XML namespace prefixes, attribute order, session IDs, proofing marks, and
identically formatted runs split only by spellcheck do not create content
conflicts. Actual text whitespace, inline/paragraph/table formatting, and
structural changes remain meaningful. Word's range export can add an empty
terminal paragraph; the bridge removes that one export artifact when comparing
the expected block count.

## Scope and limits

- Desktop Word with WordApi 1.5 and CompressedFile 1.1. Word on the web cannot
  provide this DOCX snapshot through the same compressed-file API.
- Default write mode is always fragment patches. There is no automatic whole
  document or whole-body replacement path. A user request that changes every
  body block can still produce a large changed span; normal local edits only
  touch their affected spans.
- Fragment granularity is top-level paragraphs and tables, with adjacent edits
  combined. A cell edit replaces its containing table, not just that cell.
- Supported edits include body text, direct formatting, paragraph properties,
  table insertion/removal, and table content/formatting when required package
  parts remain compatible. Documents with other top-level body block types
  cannot yet be mapped by this bridge.
- Document-wide style, header/footer, comment, and settings edits are not
  reconciled. Changed fragments containing bookmarks, comments, notes, section
  breaks, complex fields, or imported chunks are also rejected. Relationship or
  resource rewrites outside the supported native coverage policy may be rejected.
  These cases return an explanation and recovery download without applying any
  patches. Untouched annotations elsewhere remain in Word.
- Preconditions are optimistic, not an atomic coauthor lock. Another edit can
  arrive between checking and applying. Office batches are not transactions:
  an Office error may leave some patches applied. The pane warns about partial
  writes and does not automatically retry or roll back over coauthor edits.
- The local prototype accepts DOCX files up to 32 MiB, expanded packages up to
  128 MiB, each fragment up to 2 MiB of UTF-8 OOXML, and combined fragments up to
  8 MiB. These are application caps, not claims about Office platform limits.
  Large documents still incur full upload, SDK open/save, recovery download
  payload, and initial block tracking costs. Large image-heavy tables can hit a
  fragment cap even when a text edit is small. The journal must remain retained
  in the same SDK session; the adapter fails explicitly when it is truncated. This is not a large-document
  performance qualification.
- Each prompt sends the DOCX to the local backend; inspected document content
  goes to OpenAI. The provider key stays on the backend, chat history stays in
  pane memory, and document content is not logged. Stop prevents applying a
  pending result and is disabled once Word writes begin. Running SDK operations
  finish before cleanup.
- The server binds only to loopback and checks exact Origin/Host. A hosted
  deployment needs your application's authentication, authorization, request
  ownership, and lifecycle persistence.

## Source map for an agent modifying this demo

| File | Responsibility |
| --- | --- |
| `manifest.xml` | Word entry point, host requirements, permissions, icons |
| `src/taskpane/taskpane.ts` | Chat state, pending/cancel UI, conflict/unsupported/recovery messages |
| `src/taskpane/operations.ts` | Public Office.js operations entry point for callers |
| `src/taskpane/operations/fragments.ts` | Short capture/apply/release fragment operations |
| `src/taskpane/operations/full-document.ts` | Explicit whole-document apply operation |
| `src/taskpane/office/` | Word file access, capture, tracking, preconditions, write plans, full-document insertion helpers |
| `src/taskpane/roundtrip.ts` | Request binding, cancellation, outcomes, cleanup, partial-write errors |
| `src/contract.ts` | Shared Zod request/response/patch schemas and limits |
| `src/ooxml.ts` and `src/digest.ts` | Shared XML parsing and Word target fingerprints; no document diff |
| `src/server/agent.ts` | SDK session lifecycle, Changes API checkpoint, agent call, final recovery export, cleanup |
| `src/server/operations.ts` | Public backend operations entry point for callers |
| `src/server/operations/fragments.ts` | Short fragment-update preparation operation |
| `src/server/operations/full-document.ts` | Full DOCX update preparation operation |
| `src/server/changes/` | Native journal capture/reading, block ownership, supported-change checks, range planning, SDK extraction |
| `src/server/fragments/package.ts` | SDK capture/dependency manifest to Office.js Flat OPC serialization |
| `src/server/agent-loop.ts` | SDK tool discovery and agent turn loop |
| `src/server/app.ts` | Local API validation, limits, concurrency, sanitized failures, static UI |
| `src/server/index.ts` | HTTPS startup, trusted development certificate, server-only provider key |
| `scripts/build.ts` | Browser bundle and static assets |

For integration instructions and the supported caller entry points, follow
[the extraction guide in AGENTS.md](AGENTS.md#extract-the-adapters-into-a-users-application).
Keep installation and extraction procedures there so this README and the agent
instructions stay consistent.

## Verification

```sh
pnpm typecheck
pnpm build
pnpm test
pnpm test:native
pnpm validate
```

Unit tests cover journal pagination/truncation, recorded structural plans, Word
range application, unrelated coauthor
edits, changed target spans/gaps, request identity, cancellation, file-handle
cleanup, unsupported features, relationship packaging, origin/host checks, and
sanitized failures. Native tests exercise the real SDK journal and extraction for text, formatting,
insertions, deletions, splits, merges, moves, tables, and table cells. A deterministic
agent test applies returned OOXML to a source package and verifies it after reopening. It needs no OpenAI key; it does not exercise real Office.
The GitHub Actions workflow runs typecheck, build, unit tests, and the native
SDK test on Linux.

Desktop Word smoke tests: replace one paragraph's text; insert a table between
two existing paragraphs; update a table cell and shade its header; insert and
remove a paragraph. Verify untouched blocks remain unchanged. While an agent
request is pending, edit an unrelated paragraph and then an affected paragraph
on separate runs. The unrelated edit should survive; the affected edit should
prevent application. Save/reopen the test DOCX to inspect persistence. These
checks establish observed behavior, not full OOXML conformance or exhaustive
feature compatibility.
