# Word Agent installation and integration guide

Use this guide when installing the add-in, modifying its adapters, or extracting
its code into another application. [README.md](README.md) explains the document
flow, snapshot timing, supported features, and limits. Paths below are relative
to this repository root.

## Install and run the add-in

Use Node 22+, the pnpm version pinned in `package.json`, and desktop Word supporting WordApi 1.5 and CompressedFile
1.1. This export adapter does not work in Word on the web. From the repository root:

```sh
pnpm install --frozen-lockfile
[ -f .env ] || cp .env.example .env
```

Set `OPENAI_API_KEY` in the local `.env`; preserve an existing key. Never put it
in the manifest, browser bundle, source control, or diagnostic output. Then run:

```sh
pnpm certs
pnpm dev
```

The certificate installer may require OS approval. The HTTPS server uses the
trusted localhost certificate and serves both pane and API on port 3000. Check
`https://localhost:3000/api/health`: `ready: true` means the provider key is
configured, not that a model call has been tested. Run commands from this package
so environment loading and static asset paths resolve correctly.

### Sideload in Word on macOS

Copy this package's manifest into Word's development add-in directory:

```sh
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
cp manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml
```

If this add-in is already installed under another filename, update that copy
instead; do not leave duplicate manifests with the same add-in ID. If macOS
blocks terminal access to Word's container, use Finder → Go → Go to Folder to
open the directory and copy the manifest there. See [Microsoft's Mac sideloading
instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac).
Restart Word, open a disposable DOCX, then choose **Home → SuperDoc**; if needed,
find **SuperDoc Agent** under **Home → Add-ins**.

### Sideload in Word on Windows

Run the server on the Windows machine where Word is running; localhost in the
manifest refers to that machine. For a development shared-folder catalog, copy
`manifest.xml` into the shared folder, register its network path in Word under
**File → Options → Trust Center → Trust Center Settings → Trusted Add-in
Catalogs**, enable **Show in Menu**, and restart Word. Select the add-in from
**Home → Add-ins → Advanced → Shared Folder**. Follow [Microsoft's Windows
catalog instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
for the full setup. A shared-folder catalog is a testing mechanism, not a
production distribution method.

After code changes, rebuild and restart with `pnpm dev`, then close and reopen
the Word pane. Manifest/ribbon changes may require reinstalling the manifest and
restarting Word. Do not clear all Office caches as the first troubleshooting
step. If the pane cannot connect, check the running server, certificate trust,
manifest URLs, and health response before changing the document adapter.

## Extract the adapters into a user's application

Keep browser and server code in separate builds. Callers import operations from
`src/taskpane/operations.ts` or `src/server/operations.ts`; the short mode-specific
implementations stay under their respective `operations/` directories. Keep
longer logic in imported helpers rather than growing these public functions.
Do not bundle the backend, SDK runtime, filesystem access, or provider key into
the task pane. The two operations entry points are local modules, not an SDK
package that can be installed independently from this example.

### What to copy

| Integration slice | Copy together |
| --- | --- |
| Shared adapter foundation | `src/contract.ts`, `src/ooxml.ts`, and `src/digest.ts` |
| Word fragment adapter | `src/taskpane/operations.ts`, `src/taskpane/operations/`, and `src/taskpane/office/` |
| Backend journal adapter | `src/server/operations.ts`, `src/server/operations/`, `src/server/changes/`, and `src/server/fragments/` |
| Request lifecycle and conflict UI behavior | `src/taskpane/roundtrip.ts`; adapt its transport and status callbacks to the host app |
| Existing OpenAI + SDK agent, if wanted | `src/server/agent.ts` and `src/server/agent-loop.ts`; otherwise keep the host app's agent |
| Standalone Word add-in shell, if needed | `manifest.xml`, `assets/`, `src/taskpane/taskpane.html`, and the host's task-pane UI/build wiring |

Copy the shared foundation with either adapter. Copying an operations barrel
alone omits its transitive helpers. Preserve relative paths initially, then move
modules with their imports intact. The Changes API adapter also needs
`@superdoc/sdk`; it is independent of OpenAI, the agent loop, HTTP, and the UI.

For **full replacement only**, copy `src/server/operations/full-document.ts`
or `src/taskpane/operations/full-document.ts` with that side's operations types
and `src/contract.ts`. The Word side also needs
`src/taskpane/office/full-document.ts`. Make a host-app
operations entry point exporting only this mode instead of copying a barrel
that still re-exports omitted fragment modules. A full-only integration must
supply its own Word DOCX capture, transport, request identity, and freshness
checks, or retain this example's capture/orchestration dependencies as well.

### Dependencies and host wiring

- Take compatible versions from [package.json](package.json); the SDK version
  is pinned. Shared fragment code needs `@xmldom/xmldom`, `fflate`, and `zod`.
  Copying `src/contract.ts` requires `zod` in either mode. Browser development
  needs `@types/office-js` and the Office.js runtime
  script used by `src/taskpane/taskpane.html`.
- The existing backend agent additionally needs `@superdoc/sdk` and `openai`
  plus Node 22's Web Crypto and AbortSignal support. The journal adapter does
  not require an LLM; it observes mutations in the host agent's open SDK document.
- Reuse the host's HTTP framework and HTTPS deployment. `src/server/app.ts` and
  `src/server/index.ts` are optional local-server scaffolding, not required
  adapter dependencies. Hono and `@hono/node-server` are needed only if copied.
- Reuse the host's chat UI and styling. The example's styles are local to
  `src/taskpane/taskpane.css`; they are not operation dependencies. Keep the host's
  own build tooling, or adapt `scripts/build.ts` for a new standalone add-in.
- Give a separately branded add-in its own manifest ID and replace its display
  metadata, icon, source, and task-pane resource URLs. Keep the required Word
  permissions/API sets and load Office.js before use.
- Adapt the transport in `src/taskpane/roundtrip.ts` to the app's authenticated
  API. The demo uses a same-origin `/api/run`; do not copy localhost Origin/Host
  checks unchanged into a hosted service or simply remove access control.
  Keep keys and document processing server-side, and retain validation, request
  limits, cancellation, cleanup, and sanitized errors.

### Wire the fragment flow

Use `src/server/agent.ts` and `src/taskpane/roundtrip.ts` as executable examples:

1. Capture with `captureFragmentSource`; retain its snapshot and tracked ranges
   in the same Word pane session. Send the DOCX, request UUID, prompt, and history
   to the backend. Never serialize the live Word ranges to the server.
   Preserve the live-body checks bracketing DOCX export. Keep live Word
   preconditions separate from DOCX hashes; the two XML projections can differ.
2. Open the source in SuperDoc and call `captureChangeSource(document, source)`
   **before agent mutations**. Run the agent on this same live document, then
   call `createFragmentUpdate(document, checkpoint)`. Keep both calls inside the
   SDK client lifetime; saving/reopening or switching sessions loses the journal
   pairing. Do not substitute before/after DOCX comparisons.
3. Return the patch plan with the request ID and SHA-256 of the exact source
   bytes. Keep `runResponseSchema` or an equivalent validated contract in sync
   on both sides. The current response also requires a full recovery DOCX from
   `createFullDocumentUpdate`; it is not automatically applied.
4. Verify identity and cancellation, handle `unsupportedReason` without writing,
   then call `applyFragmentUpdate({ source: snapshot, patches, signal,
   beforeWrite })`. A `false` result means a precondition failed; show a conflict
   instead of reporting success. `beforeWrite` marks when cancellation/retry UI
   must stop promising that no writes occurred.
5. Release with `releaseFragmentSource` in `finally`, including on transport
   failure, cancellation, unsupported edits, or conflicts. Permit only one
   capture/apply lifecycle at a time per pane; tracked state is session-local.

The backend operation reads all Changes API pages through a captured endpoint,
rejects evicted or incomplete history, and verifies the session/revision. Journal
records supply affected identities; SDK block lists supply order; SDK
`selection.extractOoxml` supplies complete block XML and dependencies. Nested
cell events resolve to the containing top-level table. Keep extraction and Flat
OPC packaging separate from journal reading and Word range planning.
A native `diff.capture`/`diff.compare` checkpoint is used only to reject unsupported
global changes and detect net-zero edits. Do not inspect its opaque payload or
use it to discover fragment targets. Source fingerprints are Word preconditions,
not another document diff.

The SDK-to-Office mapping is structural: resulting paragraphs/tables become
OOXML passed to tracked Word ranges, rather than a duplicate Office.js
implementation of each SDK command. Preserve target/gap and resource checks;
unchanged blocks elsewhere in Word must survive. Tables are the smallest table
patch unit. Unsupported features must not trigger whole-document fallback.

### Wire explicit full-document replacement

Pair backend `createFullDocumentUpdate(updated)` with browser
`applyFullDocumentUpdate(fullDocument)`. This calls Word's document-level
`insertFileFromBase64` and imports the configured document options. The backend
can prepare a recovery export up to 32 MiB, while this apply operation rejects
files over 4 MiB; preparation alone does not guarantee Word can insert it.
The caller must verify source freshness, coordinate with pending fragment work,
and explicitly choose this mode. It is not the chat default or conflict retry.

## Verify extraction or changes

From this package run `pnpm typecheck`, `pnpm build`, `pnpm test`,
`pnpm test:native`, and `pnpm validate` as appropriate. Carry the adapter tests
into the host's test runner; `test/agent.native.ts` demonstrates a real SDK run
with deterministic model replies. No test provider key is required.
In desktop Word, test text edits, table insertion/formatting, deletion, and
save/reopen. On separate requests, edit an unrelated paragraph and a targeted
paragraph while the agent runs: preserve the former and reject the latter.
Test unsupported edits and Office failures without automatic retry; a failed
Office batch can leave partial writes. See the README for remaining limits.
