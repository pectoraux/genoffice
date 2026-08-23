# Phase 2 Increment 15A — Pivot boundary hardening + rename concurrency tests

Source commit: `a37c3db` (rename serialization + pivot boundary + cleanup)
Top of tree: `6a501b4` (only added skill/download files; no Sheets code changes)

## Audit findings (a37c3db)

Forensic audit of `a37c3db` against the user's 7 architectural requirements
revealed the following remaining gaps:

1. `ElectronXlsxSidecarEngine.readArchiveEntry()` still uses unchecked
   `as Record<string, unknown>` and `as Array<Record<string, unknown>>`
   casts on the sidecar response — Issue 1.
2. `SpreadsheetService.readPivotDefinition()` returns `Promise<unknown>` —
   Issue 2. No real runtime-independent `WorkbookPivotDefinition` contract.
3. `SpreadsheetEngine.readArchiveEntry(handle, entryName)` exposes a generic
   ZIP entry API — Issue 3. Should be a Sheets-specific
   `readPivotDefinition(handle, pivotTablePath, cacheDefinitionPath)`.
4. `SpreadsheetServiceImpl.readPivotDefinition()` performs a dynamic
   `await import('@genoffice/xlsx-gateway/...')` — this is a violation of
   the existing `services-sheets` architecture test (which forbids
   xlsx-gateway imports); the test pattern only catches `from`/`require`
   and misses dynamic `import()`.
5. Rename race tests (`sheets-rename-race.test.ts`) all require the real
   Rust sidecar — Issue 4. Must be mocked.
6. Coordinator's `renameWorkbook()` does NOT update the legacy
   `SessionInfo.path` mirror — Issue 4 verification gap.
7. Existing `sheets-pivot-rename.test.ts` still passes `sidecarClient` to
   the coordinator constructor (TS2353 errors at HEAD).

## Plan

1. Define a real runtime-independent `WorkbookPivotDefinition` contract in
   `runtime-contracts` (structurally matching xlsx-gateway's
   `PivotDefinition`, no dependency on xlsx-gateway).
2. Replace `SpreadsheetEngine.readArchiveEntry(handle, entryName)` with
   `readPivotDefinition(handle, pivotTablePath, cacheDefinitionPath):
   Promise<WorkbookPivotDefinition>` — Sheets-specific, no generic ZIP API.
3. Change `SpreadsheetService.readPivotDefinition()` return type to
   `Promise<WorkbookPivotDefinition>`. Remove the dynamic
   `await import('@genoffice/xlsx-gateway/...')` from the service —
   the engine does the parsing (it already has xlsx-gateway as a dep).
4. Rewrite `ElectronXlsxSidecarEngine.readArchiveEntry()` as
   `readPivotDefinition()`:
   - generate workDir once via `mkdtempSync` (deterministic, single path)
   - call sidecar `read_entries` for both XML parts in parallel
   - runtime-validate the sidecar response via type guards (ZERO `as`
     casts)
   - parse via `parsePivotDefinition` from xlsx-gateway
   - clean up workDir in `finally`
5. Coordinator: type the return; add `onWorkbookRenamed` callback to deps
   (invoked after successful rename — no push, coordinator already
   pushed). This closes the legacy mirror gap.
6. `sheets-runtime.ts`: accept + plumb `onWorkbookRenamed`.
7. `sheets-main.ts`: extract `updateLegacySessionPath()` from
   `sheetsFileRenamed()`; wire it as the coordinator's
   `onWorkbookRenamed` callback.
8. Handler: keep thin — validate → coordinator → return. The return type
   is now `WorkbookPivotDefinition`; the existing `schema.parse()` runs
   as the IPC-shape sanity check.
9. Rewrite rename race tests with MOCKED service/engine/filesystem —
   no Rust sidecar. Verify: rename vs save, vs close, vs teardown,
   concurrent rename, no stale originalPath, no stale legacy mirror,
   no duplicate workbook:renamed event.
10. Keep exactly one real-sidecar integration test
    (`sheets-pivot-rename.test.ts`); fix its stale `sidecarClient` coordinator
    dep errors; update it for the typed `WorkbookPivotDefinition` return.
11. Re-run the real CDP smoke test.
12. Git diff verification: renderer/shared/preload unchanged, no generated
    artifacts, architecture tests intact.

---
Task ID: 0
Agent: Principal Architect (main)
Task: Forensic audit of a37c3db and design plan for the 7 architectural corrections.

Work Log:
- Reviewed commit a37c3db diff stat: 10 files, 481 insertions, 130 deletions.
- Confirmed top-of-tree commit 6a501b4 only added skill/download files; no Sheets code changes.
- Read the 7 issue specifications from the user message.
- Read SpreadsheetEngine contract — found generic `readArchiveEntry(handle, entryName)` at lines 521-539.
- Read SpreadsheetService contract — found `readPivotDefinition` returns `Promise<unknown>` at line 423.
- Read SpreadsheetServiceImpl — found dynamic `await import('@genoffice/xlsx-gateway/...')` at line 251 (services-sheets arch test violation).
- Read ElectronXlsxSidecarEngine — found `readArchiveEntry()` impl with `as Record` + `as Array` casts at lines 538, 542.
- Read coordinator — confirmed `renameWorkbook` already uses `withSessionLock` and updates `ShellWorkbookSession.originalPath`, but does NOT update legacy `SessionInfo.path` mirror.
- Read `sheets-main.ts` — confirmed `sheetsFileRenamed()` updates legacy mirror + pushes 'workbook:renamed' event (would duplicate if coordinator also pushed).
- Read `sheets-rename-race.test.ts` — confirmed all 6 tests use the real Rust sidecar.
- Read `sheets-pivot-rename.test.ts` — confirmed 5 stale `sidecarClient: sidecarClient` coordinator constructor args (TS2353 errors at HEAD).
- Read `architecture.test.ts` — confirmed existing pivot/rename guards; identified gaps to add.
- Read xlsx-gateway `PivotDefinition` source — confirmed runtime-independent shape (no Electron/node:*).
- Ran `tsc --noEmit` — confirmed pre-existing TS errors including the 5 `sidecarClient` errors in sheets-pivot-rename.test.ts.
- Ran existing tests — they pass (vitest does not enforce type-check at runtime).

Stage Summary:
- Architectural gaps identified: 7 issues confirmed against a37c3db.
- Implementation plan finalized (12 steps).
- No code changes yet — design phase only.

---
Task ID: 1
Agent: Principal Architect (main)
Task: Implement the 7 architectural corrections.

Work Log:
- Created packages/runtime-contracts/src/services/pivot-definition.ts with
  WorkbookPivotDefinition + 9 supporting types (WorkbookPivotSharedItem,
  WorkbookPivotFieldItem, WorkbookPivotDateUnit, WorkbookPivotFieldGrouping,
  WorkbookPivotCacheField, WorkbookPivotLayoutLine, WorkbookPivotShowDataAs,
  WorkbookPivotDataField, WorkbookPivotLabelFilter, WorkbookPivotValueFilter,
  WorkbookPivotFilterDef). Structurally mirrors xlsx-gateway's PivotDefinition
  (no dependency on xlsx-gateway — runtime-contracts is Layer 1, xlsx-gateway
  is Layer 4a).
- Re-exported pivot-definition from runtime-contracts/src/index.ts.
- Updated SpreadsheetEngine contract: removed generic readArchiveEntry
  (ZIP-entry escape-hatch); added Sheets-specific readPivotDefinition
  returning Promise<WorkbookPivotDefinition>.
- Updated SpreadsheetService contract: readPivotDefinition now returns
  Promise<WorkbookPivotDefinition> (was Promise<unknown>).
- Updated SpreadsheetServiceImpl: removed the dynamic
  `await import('@genoffice/xlsx-gateway/...')` (services-sheets arch test
  violation); delegates directly to engine.readPivotDefinition() with
  zero parsing logic in the service.
- Added validateReadEntriesResponse to sidecar-validators.ts: type-guard
  based runtime validator (ZERO `as` casts on the response). Returns
  ReadArchiveEntriesResult { entries: readonly ReadArchiveEntry[] }.
- Updated ElectronXlsxSidecarEngine: removed readArchiveEntry; added
  readPivotDefinition — generates workDir once via mkdtempSync (single
  deterministic path), single sidecar read_entries call for BOTH XML
  parts, runtime-validate via validateReadEntriesResponse, parse via
  xlsx-gateway's parsePivotDefinition, cleanup workDir in finally.
  PivotParseError mapped to typed InvalidInputError.
- Added module-level helper findEntryPath (pure, no I/O, no `as`).
- Updated coordinator: readPivotDefinition returns
  Promise<WorkbookPivotDefinition>; added onWorkbookRenamed callback to
  SheetsShellCoordinatorDeps; renameWorkbook invokes the callback AFTER
  successful rename + after pushing the event exactly once (no duplicate).
- Updated sheets-runtime.ts: SheetsCoordinatorConfig carries
  onWorkbookRenamed; initSheetsRuntime accepts + plumbs it to the
  coordinator.
- Updated sheets-main.ts: extracted updateLegacySessionPath() from
  sheetsFileRenamed() (no push — only the caller decides to push);
  wired updateLegacySessionPath as the coordinator's onWorkbookRenamed
  callback via initSheetsRuntime({ ... }, { onWorkbookRenamed }).
- Updated migrated handler comment to reflect the new architecture
  (engine is the single translation point).
- Rewrote sheets-rename-race.test.ts: 8 sidecar-free tests using a
  MockSpreadsheetService (no real sidecar, no real engine). Verifies:
  rename vs save, rename vs close, rename vs teardown, concurrent
  rename, no stale originalPath, no stale legacy mirror, no duplicate
  workbook:renamed event, failed-rename no-callback.
- Fixed sheets-pivot-rename.test.ts: removed 5 stale `sidecarClient`
  coordinator dep args (TS2353); updated pivot read to use the typed
  WorkbookPivotDefinition return (no `as { outputRef, fields }` cast).
- This is the SINGLE real-sidecar integration test (per user directive).
- Added architecture guards:
  * runtime-contracts/dependency-direction.test.ts: pivot-definition.ts
    has ZERO Electron/node:/xlsx-gateway imports; exports
    WorkbookPivotDefinition; engine contract declares readPivotDefinition
    (NOT readArchiveEntry).
  * services-sheets/architecture.test.ts: ZERO static AND dynamic
    xlsx-gateway imports (the prior test missed dynamic imports);
    readPivotDefinition returns Promise<WorkbookPivotDefinition>;
    ZERO raw sidecar command construction; ZERO filesystem refs;
    ZERO Electron refs.
  * platform-electron/architecture.test.ts: engine declares
    readPivotDefinition (NOT readArchiveEntry); returns
    Promise<WorkbookPivotDefinition>; ZERO unchecked `as` casts on
    sidecar response; workDir via mkdtempSync + rmSync(workDir) in
    finally; delegates parsing to parsePivotDefinition; validator
    exists in sidecar-validators.ts and has ZERO `as` casts.
  * apps/sheets/architecture.test.ts: coordinator readPivotDefinition
    returns Promise<WorkbookPivotDefinition>; coordinator has
    onWorkbookRenamed dep; renameWorkbook invokes the callback;
    coordinator has ZERO raw sidecar command construction; coordinator
    has ZERO global caller state; sheets-runtime plumbs
    SheetsCoordinatorConfig; sheets-main wires
    updateLegacySessionPath as the callback; updateLegacySessionPath
    is extracted from sheetsFileRenamed and does NOT push.

Stage Summary:
- All 7 architectural gaps closed.
- runtime-contracts: 6 dependency-direction tests pass (including 3 new
  pivot-related guards).
- services-sheets: 23 architecture tests pass (including 4 new pivot
  purity guards + 1 enhanced xlsx-gateway dynamic-import guard).
- platform-electron: 11 architecture tests pass (including 7 new
  engine-contract hardening guards).
- apps/sheets: 68 architecture tests pass (including 8 new coordinator
  + sheets-main + sheets-runtime guards).
- apps/sheets: 8 sidecar-free rename race tests pass (MockSpreadsheetService).
- apps/sheets: 9 real-sidecar pivot/rename integration tests pass (the
  SINGLE real-sidecar integration test, per user directive).
- Pre-existing 8 workbook-skill-tools failures unrelated to this work
  (verified by stash + retest).

---
Task ID: 2
Agent: Principal Architect (main)
Task: Real CDP smoke test + final architecture assertions + git diff verification.

Work Log:
- Re-ran the real Electron CDP smoke test (scripts/sheets-cdp-smoke.mjs):
  ALL CHECKS PASSED. The full real-production path was exercised:
    open → read → save → read (session continuity) → save response fidelity →
    save content fidelity (re-open saved file via sidecar) → PDF export →
    screen capture → files (attachment read) → pivot read (real sidecar
    `read_entries` → parsePivotDefinition → typed WorkbookPivotDefinition) →
    auto-rename (refused — fixture not untitled, as expected) → stale session
    (InvalidSessionError reaches the renderer).
- Sidecar process identity: exactly ONE sidecar process (PID verified).
- PDF output: 14651 bytes, header "%PDF-" — real printToPDF path.
- Pivot read: outputRef=A3:B7, fields=3 — real parsePivotDefinition output.
- Verified git status: ZERO changes to apps/sheets/src/renderer,
  apps/sheets/src/shared, or apps/sheets/src/preload. The user's "renderer /
  shared / preload unchanged" invariant is preserved.
- Verified ZERO generated artifacts in the working tree: apps/sheets/out is
  gitignored (matches `out/` pattern in .gitignore); no out/ or dist/ files
  are tracked.
- Verified architecture tests strengthened (not weakened):
    apps/sheets/tests/architecture.test.ts: 497 → 601 lines (+21%)
    packages/runtime-contracts/tests/dependency-direction.test.ts: 102 → 149 (+46%)
    packages/services-sheets/tests/architecture.test.ts: 220 → 293 (+33%)
    packages/platform-electron/tests/architecture.test.ts: 48 → 152 (+217%)
  All four files grew — no guards were removed or weakened. New guards added
  for: WorkbookPivotDefinition contract purity, ZERO xlsx-gateway dynamic
  imports in services-sheets, ZERO unchecked `as` casts on sidecar responses,
  coordinator onWorkbookRenamed callback wiring, sheets-main
  updateLegacySessionPath extraction, etc.
- Created scripts/final-arch-assertions.mjs — a single script that checks
  every invariant from the user's directive:
    runtime-contracts: ZERO Electron, ZERO node:*, ZERO apps/sheets
    services-sheets: ZERO Electron, ZERO filesystem, ZERO raw sidecar
                     protocol, ZERO Promise<unknown> for pivot
    coordinator: ZERO raw sidecar command construction, ZERO global
                 caller state
    migrated handlers: ZERO pivot parser, ZERO sidecar client,
                       ZERO filesystem implementation, ZERO type assertions
  Plus extra engine-contract and engine-impl invariants.
- Ran the script: 28/28 assertions PASSED, 0 FAILED.

Stage Summary:
- All 7 architectural gaps from the user directive are CLOSED.
- Real CDP smoke test PASSES (the actual production path — real Electron,
  real Rust sidecar, real BrowserWindow + printToPDF, real desktopCapturer).
- 28 final architecture assertions PASS.
- Test sweep across affected packages:
    runtime-contracts: 64 tests pass (5 test files)
    services-sheets: 64 tests pass (2 test files)
    platform-electron: 72 tests pass (3 test files: architecture + engine + translator)
    apps/sheets: 122 tests pass (5 test files: architecture + rename races +
                 pivot/rename E2E + legacy adoption + real sidecar adoption)
- Pre-existing 8 workbook-skill-tools.test.ts failures are unrelated to this
  work (verified by git stash + retest at HEAD before my changes).
- Renderer/shared/preload: ZERO changes (confirmed via git status).
- Generated artifacts: ZERO tracked (out/ is gitignored).
- Architecture tests: STRENGTHENED (all 4 arch test files grew; no guards
  removed or weakened).
- The increment is COMPLETE per the user's directive. The single real-sidecar
  integration test (sheets-pivot-rename.test.ts) verifies the production
  path; the sidecar-free rename race tests (sheets-rename-race.test.ts) verify
  concurrency invariants deterministically without environment dependencies.
