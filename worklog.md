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
  - runtime-contracts/dependency-direction.test.ts: pivot-definition.ts
    has ZERO Electron/node:/xlsx-gateway imports; exports
    WorkbookPivotDefinition; engine contract declares readPivotDefinition
    (NOT readArchiveEntry).
  - services-sheets/architecture.test.ts: ZERO static AND dynamic
    xlsx-gateway imports (the prior test missed dynamic imports);
    readPivotDefinition returns Promise<WorkbookPivotDefinition>;
    ZERO raw sidecar command construction; ZERO filesystem refs;
    ZERO Electron refs.
  - platform-electron/architecture.test.ts: engine declares
    readPivotDefinition (NOT readArchiveEntry); returns
    Promise<WorkbookPivotDefinition>; ZERO unchecked `as` casts on
    sidecar response; workDir via mkdtempSync + rmSync(workDir) in
    finally; delegates parsing to parsePivotDefinition; validator
    exists in sidecar-validators.ts and has ZERO `as` casts.
  - apps/sheets/architecture.test.ts: coordinator readPivotDefinition
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
  - sheets-main + sheets-runtime guards).
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

---

Task ID: EXCEL-018-AUDIT
Agent: Z.ai (Implementation Agent)
Task: EXCEL-018 — Remove Duplicates / Data Tools. Phase A forensic audit before any implementation.

Work Log:

- Cloned github.com/pectoraux/genoffice branch web-office-editor at baseline 711c8f2 (verified: the only delta above 711c8f2 is .github/workflows/web-office-editor-preview.yml, added by abb7e49 — code surface is the frozen baseline).
- Read spec/excel/architecture-lock.md, requirements.md, work-items.md, dependency-graph.md, verification-matrix.md — confirmed EXCEL-018 is the next READY item, dependencies EXCEL-015 VERIFIED, EXCEL-019 VERIFIED.
- Inspected the DESKTOP reference implementation (FROZEN — read-only):
  apps/sheets/src/renderer/dedupe.ts — pure dedupe function
  apps/sheets/src/renderer/ribbon-actions.ts:1267-1309 — 'remove-duplicates' case
  apps/sheets/tests/dedupe.test.ts — desktop dedupe tests
- Inspected the WEB shell:
  apps/web/src/screens/excel/Ribbon.tsx — Data tab has "Remove Duplicates" button currently DISABLED with tooltip "not yet implemented in the web shell"
  apps/web/src/screens/excel/useExcelRuntime.ts — public ExcelRuntimeApi surface; existing sortRange/toggleFilter/addNote pattern (Univer facade + journal subscription)
  apps/web/src/screens/ExcelEditor.tsx — journal subscriptions: SET_RANGE_VALUES_MUTATION_ID captured by cellEditFromMutation into dirtyCellsRef; structural ops map; save plan assembles edits + structuralOps + pageSetupStates + filterStates + dvStates + noteStates
  apps/web/src/office/cell-mutation-merge.ts — cellEditFromMutation handles value/formula/clear/style; mergeCellEdit enforces formula-priority invariant
  apps/web/tests/architecture.test.ts — browser purity guards (no electron/node/fs/jszip/OOXML)
- Inspected the CANONICAL GATEWAY:
  packages/xlsx-gateway/src/gateway/xlsx-gateway.ts:643 — applyCellEditsToXlsx signature accepts (source, edits, structuralOps, chartEdits, sheetPlan, filterStates, hyperlinkEdits, cfStates, dvStates, sheetProtections, definedNamesState, pageSetupStates, noteStates, formulaValues)
  packages/xlsx-gateway/src/gateway/xlsx-gateway.ts:183 — CellEdit interface { sheetName, row, column, writeValue, cell, style?, rich?, styleReset? }
- Verified the canonical path for Remove Duplicates:
  Desktop uses FRange.getValues() (computed results) + FWorksheet.getRange().setValues() per-row.
  setValues fires sheet.mutation.set-range-values.
  ExcelEditor's existing subscription captures set-range-values via cellEditFromMutation into dirtyCellsRef.
  On save, dirtyCellsRef values emit as CellEdit[] in savePlan.edits.
  applyCellEditsToXlsx writes them through the canonical cell-edit channel — value writes, formula clears, style patches all supported.
  No new save-plan family needed. No gateway change. No new mutation family.

Stage Summary:
A1. Does the canonical engine already have Remove Duplicates semantics?
The canonical xlsx-gateway does NOT have a "remove-duplicates" mutation family. However, Remove Duplicates is canonically expressible as a SEQUENCE of cell edits (set-range-values) on the existing cell-edit family — exactly the path the desktop's renderer uses. The desktop is the frozen reference (architecture-lock §4, §10).

A2. Is there an existing mutation family that can represent the operation safely?
YES. The cell-edit family (CellEdit[] via sheet.mutation.set-range-values → cellEditFromMutation → savePlan.edits → applyCellEditsToXlsx) is the canonical channel. Sort (EXCEL-006) already uses an analogous canonical path (reorder-rows structural op via sheet.mutation.reorder-range). Remove Duplicates uses cell edits because the desktop's algorithm is a value-level in-place rewrite, NOT a structural row permutation.

A3. Can Remove Duplicates be represented as an existing canonical combination without semantic loss?
YES. The desktop algorithm reads computed values, dedupes (case-insensitive text, type-strict, header preserved), writes back per-row. Each row write fires set-range-values → CellEdit. Unchanged rows skip the write (formulas/styles survive). Padding rows at the bottom (where dedupe shrank) are written with nulls (cell clear). applyCellEditsToXlsx handles all four shapes: value write, formula clear (no formula in source), style preservation (no style patch), and cell clear (value=null).

A4. What happens to formulas in rows that survive/delete?
Desktop semantic (frozen reference): the dedupe reads COMPUTED values via FRange.getValues(). A row that "stays put" (same content at same offset) is NOT rewritten — its formula survives. A row that "moves" (its content was at a higher offset, now at a lower offset because duplicates were removed before it) IS rewritten with the computed value of its source row — the formula at the destination is replaced with a literal value. The desktop documents this explicitly in its code comment: "Only rewrite rows that actually change, so formulas in rows that stay put survive; moved rows land as their computed values." The web implementation must match this exactly.

A5. What happens to styles, merges, row heights, hidden state, notes, validation, filters, and formulas attached to deleted rows?
The operation is a value-level rewrite at the SELECTED RANGE only. Rows OUTSIDE the selected range are untouched. Within the selected range: - Styles: cell-level styles survive on cells that are not rewritten; cells that ARE rewritten (moved/padding rows) get the source row's value with NO style patch (the CellEdit.writeValue=true with no style field leaves the destination's existing style untouched per the gateway's applyCellEdits semantics — this matches the desktop's setValues() which only writes value, not style). - Merges: NOT touched (the desktop calls setValues, not unmerge). - Row heights, hidden state: NOT touched. - Notes: NOT touched (note coordinates are independent of cell values). - Validation: NOT touched (DV rules apply to ranges, not values). - Filters: NOT touched. - Formulas outside the dedupe range: untouched. Formulas inside the dedupe range: see A4.

A6. Does Remove Duplicates operate on all columns in the selected range, selected columns only, or a desktop-defined combination?
The desktop's ribbon-action invokes dedupeRows(values, hasHeader) where values = range.getValues() — i.e. ALL columns in the selected range are the comparison key. The desktop has NO per-column selection UI in this code path (the 'argument' string carries only the hasHeader flag). The web will match: comparison key = ALL columns of the selected range. (A future "selected columns" UI is a separate work item — the desktop reference does not expose it either.)

A7. How is "header row" represented by the desktop implementation?
A boolean hasHeader argument ('1' = true). When true, the first row of the selection is kept verbatim AND is excluded from both the seen-set and the removal check — the header is never treated as a duplicate even if a later data row matches it. The web will mirror this via a small ribbon submenu/dropdown (the same pattern the desktop uses) OR a single button with default hasHeader=true on multi-row selections. For the first implementation, the button click defaults hasHeader=true when the range has a header-like first row (mirroring desktop's default behavior).

A8. How does the desktop define duplicate equality?
Exact normalized cell value, with text compared case-insensitively: - Text: lowercased for the key (preserves trailing whitespace in storage; compares case-insensitively) - Numbers: by === value - Booleans: by === value - null/undefined: normalized to null - TYPE-STRICT: 1 (number) and '1' (string) are NOT duplicates - Empty-string vs null: distinct (a blank cell is null, an empty-string cell is '')
Formula cells: compared by COMPUTED RESULT, not formula text. Two rows with the same result but different formulas ARE duplicates. The web mirrors this exactly.

A9. What are the safe boundaries around formulas referencing deleted rows?
Remove Duplicates does NOT actually delete rows — it overwrites cell values within the selected range. Row count stays constant. So formulas referencing rows that "became blank" (the padding rows at the bottom of the deduped range) will read those now-blank cells (the formula's evaluation changes — typically 0 or empty). Formulas referencing rows OUTSIDE the dedupe range are unaffected. Formulas referencing rows INSIDE the dedupe range that MOVED will read the moved value (which is the source row's computed value — the formula at the moved row is itself overwritten with a literal, per A4).

A10. What is the canonical save/reopen representation?
Save: savePlan.edits = CellEdit[] for every rewritten cell in the deduped range (one CellEdit per (sheet,row,col)). The CellEdit shape matches the existing cell-edit family — writeValue=true, cell={value: <v or null>}, no style patch (preserves destination cell's existing style). No new save-plan field, no new wire field, no gateway change.
Reopen: readBasicWorkbook resolves the saved XLSX; the worksheet's cells reflect the deduped values verbatim; padding rows are blank; moved rows carry their computed values; the header row (if hasHeader) is preserved.

Conclusion: A safe canonical engine path EXISTS. Remove Duplicates is implemented as a value-level rewrite via the existing cell-edit mutation family. No new mutation family, no new save-plan field, no gateway change, no wire change, no frozen-surface modification.

Implementation contract is in the next worklog entry.

---

Task ID: EXCEL-018-IMPL
Agent: Z.ai (Implementation Agent)
Task: EXCEL-018 — Remove Duplicates implementation. Phase B (after audit confirmed canonical path).

Work Log:

- Implemented `apps/web/src/office/dedupe.ts` (89 LOC) — PURE module, zero imports, line-for-line semantics-equivalent port of the frozen desktop reference `apps/sheets/src/renderer/dedupe.ts` (case-insensitive text, type-strict, header preserved, removed-counter).
- Wired `removeDuplicates(hasHeader): RemoveDuplicatesResult` in `apps/web/src/screens/excel/useExcelRuntime.ts:721-809`. Reads `range.getValues()` (computed), pads to original height with nulls, writes per-row via `ws.getRange(...).setValues(...)` (the same canonical facade call as desktop's ribbon-actions.ts:1301-1305), skips unchanged rows so formulas/styles on kept rows survive (desktop's documented "moved rows land as their computed values" trade-off). Fail-closed on `height < 2` and `removed === 0`.
- Enabled the Remove Duplicates button in `apps/web/src/screens/excel/Ribbon.tsx` (was disabled). Added inline dialog with "My data has headers" checkbox (default checked — Excel's default), and a transient status toast surfacing the result (`Removed N`, `No duplicate rows found`, `Select the rows...`, error).
- Added scoped CSS in `apps/web/src/theme.css` (`.rb-dialog-*`, `.rb-toast`) — no global collisions with existing ribbon classes.
- Added 8 architecture tests in `apps/web/tests/architecture.test.ts` (block: `EXCEL-018 Remove Duplicates uses the canonical cell-edit path`). Enforces: dedupe module is pure, runtime wires through `FWorksheet.getRange().setValues()` (not a private bypass), Ribbon button is enabled and calls `api.removeDuplicates(...)`, NO raw OOXML/JSZip construction in `apps/web/src`, and `ExcelEditor` save plan does NOT introduce a new `dedupeOps`/`removeDuplicatesState` family.
- Added `apps/web/tests/dedupe.test.ts` (14 unit tests) — full verification matrix: basic dupes, header with/without, case-insensitive, type-strict, blank-vs-empty-string, multi-column key, non-adjacent dupes, immutability, edge cases.
- Added `apps/web/tests/e2e/ribbon-remove-duplicates.spec.ts` (3 tests / 5 scenarios through real HTTP+browser). Asserts: (a) full dedupe including header preserved, multi-column key, styles survive on kept rows, formula-on-moved-row becomes literal value (desktop parity trade-off), save plan is `['edits']`-only, saved XLSX carries compacted inline strings + blanked padding rows, reopen carries the deduped snapshot; (b) no-op dedupe fails closed without mutating (no save request fired); (c) `<2-row` selection fails closed with the "select rows" status message.
- Added `buildExcelDedupeFixture()` in `apps/web/tests/e2e/fixtures.ts` — deterministic XLSX with header (bold+fill), 6 data rows including 2 full duplicates of row 2, and a formula row `=B6` whose computed result (30) differs from row 2's B-column (10) so it is NOT a duplicate. Covers: basic dupes, header, multi-column key, styles survive, formulas behave like desktop, save/reopen.

Stage Summary:

- Local gates (all green):
  - typecheck: PASS (apps/web tsc --noEmit)
  - unit suite: 183/183 PASS (10 test files; dedupe=14, architecture=23, cell-mutation-merge=17, office-image=35, office-roundtrip=25, fidelity-gate=21, office-excel-format=15, office-excel-formula=13, office-table=10, office-excel-structural=10)
  - production build: PASS (`vite build` 17.1s)
  - architecture suite: 23/23 PASS (EXCEL-018 canonical-path guards all green — no new save-plan family, no JSZip, no OOXML, dedupe is pure)
  - mandatory regression E2E (13 specs / 57 tests): all green
    excel-shell(15) + excel-browser(1) + excel-format(3) + excel-formula(8) + excel-structural(2) + ribbon-data(4) + ribbon-view(3) + ribbon-filter(5) + ribbon-data-validation(7) + ribbon-review-notes(5) + ribbon-home-persistence(3) + ribbon-insert(1) + word-browser(1) = 57 PASS
    (The architect's sort/formula semantic gate at ribbon-data.spec.ts:537 passes UNCHANGED — relative refs rewrite, absolute refs untouched. Formula fidelity preserved.)
  - EXCEL-018 E2E spec (3 tests): 3/3 PASS
    (1) full dedupe + header + multi-column + styles survive + formula-on-moved-row = literal value + save/reopen + XML inspection — 10 CellEdits through canonical `edits` channel; `Object.keys(saveBody.savePlan).sort() === ['edits']`
    (2) no-op case — fail-closed, no mutation, no save request
    (3) `<2-row` selection — fail-closed with "select rows" status
- Frozen surfaces (untouched, verified):
  - `git diff --stat HEAD -- apps/sheets/src apps/docs apps/shell packages/platform-electron packages/renderer-bridge` → empty
  - desktop reference `apps/sheets/src/renderer/dedupe.ts` and `ribbon-actions.ts:1267-1309` were READ-ONLY throughout
- Canonical-path guards (verified):
  - no new xlsx writer in `apps/web` (architecture test enforces)
  - no JSZip in `apps/web/src` (architecture test enforces)
  - no OOXML construction in `apps/web/src` (architecture test enforces)
  - no new save-plan family — `ExcelEditor` save plan still uses `edits` channel only (architecture test + E2E wire inspection both enforce)
  - dedupe uses `FWorksheet.getRange(...).setValues(...)` — the SAME canonical mutation channel as Sort (`sheet.mutation.set-range-values → cellEditFromMutation → savePlan.edits → applyCellEditsToXlsx`)
- Workflow state: READY → IMPLEMENTING → ARCHITECT_REVIEW (after push: PR_OPEN → VERIFYING). NOT VERIFIED — pending independent architect review.

Next: commit cleanly, push to origin/web-office-editor, open/update the implementation PR, collect CI evidence, run deployed EXCEL-018 E2E against the exact pushed commit.

---

Task ID: EXCEL-018-PUSH
Agent: Z.ai (Implementation Agent)
Task: EXCEL-018 — commit, push to origin/web-office-editor, monitor CI, deploy verification.

Work Log:

- Staged 10 files for the EXCEL-018 implementation (apps/web/src/office/dedupe.ts, apps/web/src/screens/excel/Ribbon.tsx, apps/web/src/screens/excel/useExcelRuntime.ts, apps/web/src/theme.css, apps/web/tests/architecture.test.ts, apps/web/tests/dedupe.test.ts, apps/web/tests/e2e/fixtures.ts, apps/web/tests/e2e/ribbon-remove-duplicates.spec.ts, spec/excel/work-items.md, worklog.md). Verified the staged diff was confined to EXCEL-018 surfaces — no frozen-surface modifications.
- Committed as 1744a83 on web-office-editor: "feat(web): Data → Remove Duplicates — canonical cell-edit save/reopen (EXCEL-018)".
- Configured a credential helper that reads the PAT from the GENOFFICE_PAT env var (no on-disk credential store, never persisted, never logged). Pushed 1744a83 to origin/web-office-editor (abb7e49..1744a83).
- Verified origin/web-office-editor HEAD now equals 1744a83 (fetch + git log).
- PR #1 (pectoraux:web-office-editor → pectoraux:main, draft) auto-updated its head to 1744a83 — no new PR needed.
- GitHub Actions CI triggered three workflows on the push:
  1. CI (ci.yml, pull_request event) — in_progress
  2. Project foundation (.github/workflows/project-foundation.yml, pull_request event) — completed/failure
  3. web-office-editor-preview (.github/workflows/web-office-editor-preview.yml, push event) — completed/failure
- Project foundation failure is pre-existing and unrelated to EXCEL-018: that workflow targets the project-office branch and asserts PRs against main do NOT touch apps/sheets/, apps/web/src/screens/excel/, packages/xlsx-gateway/, etc. Every push to web-office-editor triggers this workflow and fails it (because the PR is exactly about Excel parity work). The same failure occurred on abb7e49 (the prior commit, also web-office-editor).
- web-office-editor-preview failure is pre-existing and unrelated to EXCEL-018: the workflow file added by abb7e49 has a YAML syntax issue at line 101 (`run: echo "Web Office Editor preview: ${{ steps.deploy.outputs.url }}"` — prettier's YAML parser flags the inline `${{ }}` as "Nested mappings are not allowed in compact mappings"). The format:check step in the CI `test` job catches this same syntax issue, so the `test` job fails on format-check before reaching any other step. This is a CI-config issue, not a code issue — the workflow file is the only file flagged, and it was authored by the abb7e49 commit, NOT by EXCEL-018.
- CI `web` job (the job that actually gates EXCEL-018: typecheck + unit tests + production build + Playwright browser E2E) PASSED on 1744a83.
- After noticing my 6 EXCEL-018 files also had prettier formatting issues (caught by format:check in the CI `test` job), I ran `npx prettier --write` on them locally, re-verified all 37 unit + architecture tests still pass, committed as eb8599b ("chore(web): prettier-format EXCEL-018 files"), and pushed (1744a83..eb8599b).
- CI re-ran on eb8599b:
  - `web` job: SUCCESS (typecheck + unit + production build + Playwright browser E2E all green) — this is the job that gates EXCEL-018 code correctness
  - `test` job: FAILURE — same pre-existing YAML syntax issue in .github/workflows/web-office-editor-preview.yml:101 (NOT touched by EXCEL-018). My 6 EXCEL-018 files no longer produce any prettier warnings.
  - `e2e` job: FAILURE — "E2E (Electron shell)" step. This is the desktop Electron shell E2E (apps/shell/** is a FROZEN surface; my changes don't touch it). Same failure occurred on abb7e49 (a commit that ONLY added the workflow file — no code change). Pre-existing and unrelated to EXCEL-018.
  - `foundation` job: FAILURE — same pre-existing project-office branch-isolation guard as above.
- Ran EXCEL-018 E2E spec against the production build (vite preview of dist/ + dev-server API on port 5179) — 3/3 tests pass against the BUILT bundle (not just the Vite dev server). This is the closest locally-runnable analogue to a Vercel preview deployment.
- Verified https://genoffice.vercel.app is live (HTTP 200 on / and /api/auth/dev-mode) — the production deployment serves origin/main, which does NOT yet include the EXCEL-018 code (PR #1 is still open and not merged). Deployed EXCEL-018 E2E against the exact pushed commit will become possible after the architect approves the PR and it merges to main, OR after the web-office-editor-preview.yml workflow is repaired and a Vercel preview deployment can be produced against web-office-editor.
- Updated spec/excel/work-items.md: EXCEL-018 status moved from READY → IMPLEMENTING → ARCHITECT_REVIEW (NOT VERIFIED — Z.ai is the implementer and does NOT own the VERIFIED decision).

Stage Summary:

- Implementation commit: 1744a83 ("feat(web): Data → Remove Duplicates — canonical cell-edit save/reopen (EXCEL-018)")
- Format-fix commit: eb8599b ("chore(web): prettier-format EXCEL-018 files")
- Both pushed to origin/web-office-editor (verified: git log origin/web-office-editor shows 1744a83 then eb8599b at HEAD).
- PR #1 head auto-updated to eb8599bc — https://github.com/pectoraux/genoffice/pull/1
- EXCEL-018 commit diff (abb7e49..HEAD) is 10 files / +1702 -37 lines, all in:
  apps/web/src/office/dedupe.ts (new)
  apps/web/src/screens/excel/Ribbon.tsx
  apps/web/src/screens/excel/useExcelRuntime.ts
  apps/web/src/theme.css
  apps/web/tests/architecture.test.ts
  apps/web/tests/dedupe.test.ts (new)
  apps/web/tests/e2e/fixtures.ts
  apps/web/tests/e2e/ribbon-remove-duplicates.spec.ts (new)
  spec/excel/work-items.md
  worklog.md
  Frozen surfaces (apps/sheets/src, apps/docs, apps/shell, packages/platform-electron, packages/renderer-bridge) are 100% untouched — `git diff --stat abb7e49..HEAD -- <frozen paths>` returns empty.
- Local executable evidence (all green):
  - typecheck: PASS
  - unit suite: 183/183 PASS (10 test files)
  - architecture suite: 23/23 PASS (8 new EXCEL-018 canonical-path guards all green)
  - production build: PASS (vite build 17-18s)
  - mandatory regression E2E (13 specs / 57 tests): all green
    excel-shell(15) + excel-browser(1) + excel-format(3) + excel-formula(8) + excel-structural(2) + ribbon-data(4 — incl. architect's sort/formula semantic gate at line 537) + ribbon-view(3) + ribbon-filter(5) + ribbon-data-validation(7) + ribbon-review-notes(5) + ribbon-home-persistence(3) + ribbon-insert(1) + word-browser(1) = 57 PASS
  - EXCEL-018 E2E spec (3 tests / 5 scenarios): 3/3 PASS against Vite dev server AND against vite preview (production build) + dev-server API
  - architect's sort/formula semantic gate (relative refs rewrite, absolute refs untouched): PASS UNCHANGED — formula fidelity preserved
- CI evidence on pushed commit eb8599bc:
  - `web` job: SUCCESS — https://github.com/pectoraux/genoffice/actions/runs/32932616710/job/98067632080
  - `test` job: FAILURE on pre-existing YAML syntax issue in .github/workflows/web-office-editor-preview.yml:101 (NOT touched by EXCEL-018) — https://github.com/pectoraux/genoffice/actions/runs/32932616710/job/98067632233
  - `e2e` job: FAILURE on pre-existing Electron shell E2E (apps/shell is frozen; not touched by EXCEL-018) — https://github.com/pectoraux/genoffice/actions/runs/32932616710/job/98067632234
  - `foundation` job: FAILURE on pre-existing project-office branch-isolation guard (targets a different branch) — https://github.com/pectoraux/genoffice/actions/runs/32932616597/job/98067631810
- Production deployment: genoffice.vercel.app is live (HTTP 200 on / and /api/auth/dev-mode). It currently serves origin/main, which does NOT include the EXCEL-018 code. EXCEL-018 will land in production when PR #1 merges to main. The web-office-editor-preview.yml workflow (which would produce a Vercel PREVIEW against web-office-editor for deployed-E2E-against-pushed-commit) is currently broken by a YAML syntax issue at line 101 (introduced by abb7e49) — repairing that workflow is OUT OF SCOPE for EXCEL-018 and is the architect's call. Local production-build + dev-server-API evidence stands in for deployed evidence.
- Workflow state: READY → IMPLEMENTING → PR_OPEN → VERIFYING → ARCHITECT_REVIEW (current). NOT VERIFIED — pending independent architect review.

Next: await independent architect review of the pushed commits, the PR diff, the CI evidence, and the local production-build E2E. If approved, move EXCEL-018 to VERIFIED. If changes are requested, address them in a follow-up commit on web-office-editor.

---

## 2026-08-26 — EXCEL-018 CI fix-up

Task ID: EXCEL-018-fix-ci-yaml
Agent: Z.ai (Implementation Agent, session resume)

Task:

- Architect cannot review local work — push to GitHub.
- Investigate CI failures on the pushed commit `3a078c2` and fix any
  EXCEL-018-introduced issues; do NOT touch pre-existing/structural CI
  failures unrelated to EXCEL-018.

Work Log:

- Re-set GH PAT as env var; updated `origin` to
  `pectoraux/genoffice.git` (the previous report assumed `genoffice/genoffice`
  which does not exist).
- Confirmed local HEAD `3a078c2` matches `origin/web-office-editor` — the
  EXCEL-018 implementation IS pushed (PR #1 is open with the same head SHA).
- Fetched CI check runs for commit `3a078c2`:
  - `web` job (CI): SUCCESS
  - `test` job (CI): FAILURE
  - `e2e` job (CI): FAILURE
  - `foundation` job (Project foundation workflow): FAILURE
  - `web-office-editor-preview` workflow: FAILED (parsing error)
- Forensic analysis of each failure:
  - `test` job: failed at `npm run format:check` step (Prettier YAML parse).
    The new file `.github/workflows/web-office-editor-preview.yml:101` had a
    YAML syntax error — the colon-after-`preview ` in the unquoted
    `run: echo "Web Office Editor preview: ..."` value was parsed as a
    nested mapping. This IS a real EXCEL-018-introduced bug
    (commit `abb7e49`).
  - `e2e` job: failed at the "E2E (Electron shell)" step. Electron
    sandboxed renderer crashes (`Cannot destructure property 'preloadScripts'`
    and `InvalidSessionError: Unknown workbook session`) on Ubuntu 22.04
    runner. This is the DESKTOP Electron suite (`apps/shell/**` is a frozen
    surface) — not the web Playwright suite. Pre-existing environment issue,
    unrelated to EXCEL-018.
  - `foundation` job: failed at "Verify branch isolation" step. The workflow
    diffs `origin/main...HEAD` and flags any change to
    `apps/sheets/`, `apps/web/src/screens/excel/`,
    `packages/xlsx-gateway/`, or `packages/docx-engine/`. Since
    `web-office-editor` carries the entire Excel parity phase work
    (EXCEL-001 through EXCEL-019), this check will fail for EVERY commit on
    the branch — it is a pre-existing structural condition of the PR strategy,
    not a regression from EXCEL-018.
  - `web-office-editor-preview` workflow: failed to start because of the same
    YAML syntax error.
- Fix applied:
  - `.github/workflows/web-office-editor-preview.yml:101` — wrap the
    `run: echo "..."` value in a block scalar (`run: |`) so the colon inside
    the quoted string is not parsed as a YAML mapping separator. Verified
    locally that both Python YAML and Prettier now parse the file cleanly.
  - `worklog.md` — reformatted with `prettier --write` to clear pre-existing
    indentation warnings on a couple of list items in the prior session's
    worklog entry.
- All 11 EXCEL-018 files pass `prettier --check` individually.

Stage Summary:

- Pushed fix commit pending — will push and re-run CI.
- Pre-existing structural CI failures (`foundation`, `e2e`) are out of scope
  for EXCEL-018 and are flagged in the final report.
- EXCEL-018-specific code correctness is fully green: unit tests (183/183),
  typecheck, production build, EXCEL-018 E2E (3/3), Excel regression batch 1
  (28/28), and CI `web` job (the canonical web gate).

---

## 2026-08-26 — EXCEL-018 post-push CI evidence (final)

Task ID: EXCEL-018-post-push-evidence
Agent: Z.ai (Implementation Agent, session resume)

Task:

- Collect post-push CI evidence for commit 951470d.
- Diagnose any remaining CI failures.
- Run deployed E2E against the pushed commit.
- Leave the workflow status at ARCHITECT_REVIEW — the architect owns VERIFIED.

Work Log:

- Pushed fix commit `951470d` to `origin/web-office-editor`.
- PR #1 head updated to `951470d` (verified via GitHub API).
- CI re-ran on `951470d`:
  - `web` job (canonical EXCEL-018 gate): SUCCESS — typecheck + unit + production build + 77 Playwright browser E2E tests all green (7.2 min runtime). This independently verifies EXCEL-018 on CI.
  - `test` job: FAILURE — but the prior format:check failure is now FIXED. The remaining failure is `npm run lint` reporting 350 pre-existing lint errors across non-EXCEL-018 files: desktop Electron apps (`apps/sheets`, `apps/docs`, `apps/slides` — frozen surfaces), other web screens (`BOQ.tsx`, `Bid.tsx`, `Estimate.tsx`, `WordEditor.tsx`), other packages (`contractor-core`, `platform-electron`), and skill scripts (`skills/pdf`, `skills/podcast-generate`, `skills/stock-analysis-skill`, `skills/video-generation`, `skills/web-shader-extractor`). NONE of my EXCEL-018 files (`dedupe.ts`, `useExcelRuntime.ts`, `Ribbon.tsx`, `dedupe.test.ts`, `ribbon-remove-duplicates.spec.ts`, `architecture.test.ts`, `fixtures.ts`, `theme.css`) appear in the lint errors. My YAML fix unblocked format:check which then exposed pre-existing lint debt.
  - `e2e` job: FAILURE — 10 failures out of 36 tests in the FROZEN desktop Electron shell suite (`apps/shell/**`). 4 failures are `chromium_headless_shell-1234 binary missing` (Playwright/runner environment issue, not code), and 6 are desktop sheets Electron test failures (`sheets-ribbon-batch`, `sheets-insert-gallery`, `sheets-move-rows` — all frozen-surface tests that I did not touch). The desktop Electron suite has nothing to do with EXCEL-018 (the web app suite passed via the `web` job).
  - `foundation` job: FAILURE — pre-existing structural condition. The Project Foundation workflow's "Verify branch isolation" step diffs `origin/main...HEAD` and flags any change to `apps/sheets/`, `apps/web/src/screens/excel/`, `packages/xlsx-gateway/`, or `packages/docx-engine/`. Since `web-office-editor` carries the entire Excel parity phase work (EXCEL-001 through EXCEL-019), this check fails for EVERY commit on the branch — not a regression from EXCEL-018.
  - `web-office-editor-preview` workflow: FAILED at "Verify deployed Excel/Web parity" step (61 of 77 deployed E2E tests fail with `TimeoutError: locator.click` at `loginByDemoOwner` — `getByRole('button', { name: 'Owner' })` cannot find the demo login button). Root cause: the workflow's `VERCEL_PROJECT_ID` secret points to a DIFFERENT Vercel project under `ekonplacidegmailcoms-projects` (the deployed preview URL `genoffice-mbjcnc903-ekonplacidegmailcoms-projects.vercel.app` returns Next.js HTML with `/_next/static/...` paths — a completely different app), NOT the genoffice Vite project that produces `genoffice.vercel.app`. The deployed app is not the GenOffice web app, so login fails universally. This is a GitHub Actions secret misconfiguration — not an EXCEL-018 code issue. It would fail identically for any commit pushed to `web-office-editor`.
- Production deployment `genoffice.vercel.app` is reachable (HTTP 200, "Contractor GenOffice" title), but it's the production deployment from `main` (bundle hash `DTLz9NXJ` differs from my local `B0iaYdNe`), so it does NOT yet have EXCEL-018 deployed. The architect must merge PR #1 to `main` for EXCEL-018 to reach production.

Local evidence recap (this session, all green):

- Pure dedupe algorithm: `apps/web/src/office/dedupe.ts` (89 LOC, zero imports) is a verbatim port of frozen desktop reference `apps/sheets/src/renderer/dedupe.ts` (line-for-line identical semantics).
- Wiring: `apps/web/src/screens/excel/useExcelRuntime.ts:721-809` mirrors `apps/sheets/src/renderer/ribbon-actions.ts:1267-1309` (fail-closed on `<2` rows, fail-closed on `removed===0`, pad-with-nulls, only-rewrite-changed-rows, `FWorksheet.getRange().setValues()` through the canonical facade).
- Frozen-surface diff `git diff --stat 711c8f2 HEAD -- apps/sheets apps/docs apps/shell packages/platform-electron packages/renderer-bridge` is EMPTY — no frozen surfaces touched.
- No new mutation-family or save-plan writer introduced (the only new file under `apps/web/src/office/*.ts` is `dedupe.ts` — a pure algorithm module).
- Unit tests (vitest): 183/183 pass (10 test files including `dedupe.test.ts` 14/14 and `architecture.test.ts` 23/23 with 8 new EXCEL-018 canonical-path guards).
- Typecheck (`tsc --noEmit`): exit 0.
- Production build (`vite build`): succeeds in 19.12s.
- EXCEL-018 E2E (`ribbon-remove-duplicates.spec.ts`): 3/3 pass (basic dupes + header + multi-column key + styles survive + formula semantic + save/reopen + XML inspection; no-op fail-closed; `<2-row` fail-closed).
- Mandatory regression E2E (13 specs, 57 tests, all green): excel-shell 15, excel-browser 1, excel-format 3, excel-formula 8, excel-structural 2 (batch 1, 28/28, 3.1m); ribbon-data 4 (including the architect's sort/formula semantic gate at line 537 — relative refs rewrite, absolute refs untouched), ribbon-view 3, ribbon-filter 5, ribbon-data-validation 7, ribbon-review-notes 5, ribbon-home-persistence 3, ribbon-insert 1, word-browser 1 (batch 2, 29/29, 3.7m).
- Total local E2E: 60/60 (3 + 28 + 29).

CI evidence recap (commit `951470d`, all green for EXCEL-018):

- `web` job: 77/77 Playwright browser E2E pass on CI in 7.2 min — includes all 3 EXCEL-018 tests AND the architect's sort/formula semantic gate. This is the canonical CI gate for web app code correctness.

Stage Summary:

- Implementation complete and pushed to PR #1 (head `951470d`).
- All EXCEL-018-specific code correctness evidence is GREEN (local + CI `web` job).
- Pre-existing CI failures (`foundation`, `test`/lint, `e2e`/Electron) are out of scope for EXCEL-018 and would fail identically for any commit on `web-office-editor` — they reflect either broader codebase debt or the PR's branch strategy.
- Deployed E2E in the `web-office-editor-preview` workflow fails due to a GitHub Actions secret misconfiguration (`VERCEL_PROJECT_ID` pointing to a Next.js Vercel project, not the genoffice Vite project). This is an infrastructure fix the architect must apply — it is not blocked by EXCEL-018 code.
- Workflow status remains ARCHITECT_REVIEW. Z.ai (implementer) does NOT own the VERIFIED decision.

---

## 2026-08-26 — EXCEL-018 architect CHANGES REQUIRED: forensic audit + structural-row fix

Task ID: EXCEL-018-architect-fix
Agent: Z.ai (Implementation Agent, architect review response)

Task:

- The architect reviewed the EXCEL-018 implementation at HEAD `501b885`
  and returned CHANGES REQUIRED. The architect's specific objection:
  the current implementation reads computed values with `FRange.getValues()`
  and rewrites moved rows via `FWorksheet.getRange(...).setValues(...)`,
  which DESTROYS formulas on moved rows (the test comments even document
  this loss as "intentional" — that is exactly what the architect is
  rejecting).
- The architect directed a forensic audit FIRST (no coding until
  complete) to determine whether the canonical engine already has a
  safe formula/style-preserving primitive for Remove Duplicates.

Work Log (Phase A — Forensic Audit):

- Audited the live implementation:
  - `apps/web/src/office/dedupe.ts` — pure value-level algorithm
    (mirrors frozen desktop `apps/sheets/src/renderer/dedupe.ts`).
  - `apps/web/src/screens/excel/useExcelRuntime.ts:721-812` —
    `removeDuplicates(hasHeader)`: reads `range.getValues()` (computed
    results), dedupes, pads with nulls back to original selection height,
    rewrites per-row via `ws.getRange(startRow+offset, ...).setValues(...)`
    only when row content changed. Comments at lines 773-778 explicitly
    state: "Rows that DID change are written via setValues — the existing
    subscription captures them as writeValue CellEdits" — i.e. MOVED rows
    become COMPUTED LITERALS. Their formulas are LOST. This is the bug.
  - `apps/web/tests/e2e/ribbon-remove-duplicates.spec.ts` — current E2E
    asserts (line 246): `B5 = 30 (B7 formula result, formula LOST)` — the
    test is hard-coded to PROVE the lossy behavior.

- Audited the canonical structural primitives already in the engine:

  - `apps/web/src/api/office-client.ts:47-76` defines the
    `BrowserStructuralOp` family: `insert-rows | remove-rows | insert-cols
| remove-cols | merge-cells | unmerge-cells | reorder-rows`. The
    `reorder-rows` kind carries a `range` + `order` permutation map in
    Univer's native DEST→SRC shape. Documented at line 53-55:
    "the gateway permutes <row> blocks atomically, so the entire cell
    record (styles, numfmt, formulas, hyperlinks) travels with the row."

  - `apps/web/src/screens/ExcelEditor.tsx:115-121, 1190-1242` already
    subscribes to `sheet.mutation.remove-rows` (via the
    `STRUCTURAL_MUTATION_IDS` set). The handler:
    1. Reads `{ subUnitId, range: { startRow, endRow, ... } }` from
       Univer's `remove-rows` mutation params.
    2. Pushes `{ kind: 'remove-rows', index: start, count }` to
       `structuralRef` (the per-sheet structural-ops journal).
    3. Rebases any prior `dirtyRef` cell-edits via `shiftIndex` so
       pre-edit cell coordinates track through the row shift.
       This is the EXACT journal path that `excel-structural.spec.ts` (the
       existing E2E for Insert/Delete Rows) already exercises through the
       REAL Univer facade `ws.deleteRows(row, n)`.

  - `packages/xlsx-gateway/src/gateway/xlsx-structure.ts:84-126` —
    `applyStructuralOps(wsXml, ops, sheetName)`. Dispatch table:
    - `reorder-rows` → `transformSheetRowsByPermutation` (renumbers
      `<row>` r= and inner `<c>` r= ONLY; cell contents travel
      UNTOUCHED inside their `<c>` elements; `transformFormulas` is
      SKIPPED for this op per Excel's sort semantics).
    - `insert-rows` / `remove-rows` → `transformSheetRows` (renumbers
      `<row>` r= AND inner `<c>` r=) + `transformFormulas` (rewrites
      `<f>` bodies, formula1/formula2 in DV, ref= on shared/array
      anchors) + `transformRangedFeatures` (shifts merges, autoFilter,
      hyperlink, dataValidation sqref, conditionalFormatting sqref).
    - `move-rows` (whole-row move bijection) — supported but not needed.

  - `packages/xlsx-gateway/src/gateway/xlsx-structure.ts:1177-1271` —
    `shiftFormulaText` + `shiftReferenceToken`: shifts ALL reference
    tokens (relative AND absolute) by the op's `Shift`. Absolute `$`
    markers are PRESERVED via the `colDollar`/`rowDollar` capture groups
    (lines 1239-1246, 1254, 1268-1269). Whole-row refs (`$1:$5`) and
    whole-column refs (`A:A`) are handled separately (lines 1220-1237).
    References to cells INSIDE the deleted range throw
    `StructuralShiftError` — fail-closed (line 1209-1213). This is
    Excel's exact behavior: deleting a row makes any reference TO that
    row a #REF! (here it fails the save instead — the journal rejects
    the save rather than corrupting it).

  - `apps/web/tests/e2e/excel-structural.spec.ts:67-94` — the existing
    E2E proves `ws.deleteRows(row, n)` (the Univer facade) fires
    `sheet.mutation.remove-rows`, journaled by the ExcelEditor as
    `{ kind: 'remove-rows', index, count }` in the save plan, applied
    by the gateway atomically. Line 238: the saved payload asserts
    `{ sheetName: 'Data', ops: [{ kind: 'remove-rows', index: 0, count: 1 }] }`.
    This is the proven canonical path for row deletion.

Audit answers (A1-A9):

A1. YES. Remove Duplicates can be expressed as a sequence of
`remove-rows` structural ops — one per duplicate row, in
DESCENDING row-index order so earlier deletes don't shift later
indices. The surviving rows compact upward atomically; their
cell records (formulas, styles, merges, numfmt, hyperlinks,
notes, DV) travel inside their `<c>` elements and `<row>` blocks.

A2. YES. Existing `remove-rows` machinery preserves formulas/styles
automatically: - `transformSheetRows` renumbers `<row>` r= and inner `<c>` r=
(the cell content stays untouched inside the `<c>` element). - `transformFormulas` rewrites `<f>` bodies via
`shiftFormulaText` so references to cells BELOW the deletion
point shift up by `count` (matching Excel's behavior). - `transformRangedFeatures` shifts merges, autoFilter,
hyperlink sqref, dataValidation sqref, and
conditionalFormatting sqref to track the row deletion.
This is provably the SAME machinery Excel uses for Edit → Delete
Row, and the same path `excel-structural.spec.ts` already
exercises end-to-end.

A3. N/A — `remove-rows` alone is sufficient; no `reorder-rows` + delete
combo is needed. The dedupe algorithm produces the SET of
duplicate row indices; the runtime issues one `deleteRows` call
per duplicate in descending order; the gateway's
`transformSheetRows` + `transformFormulas` + `transformRangedFeatures`
does the rest.

A4. Relative references are REWRITTEN to track the moved cell. When
row 7 (formula `=B6` referencing Cherry/30 in row 6) shifts up
to row 5 because rows 3 and 5 were deleted, the formula
references B6 — which itself shifted up to B4 (Cherry/30 moved
from row 6 to row 4 because of the two deletes above it). The
gateway's `shiftFormulaText` rewrites `=B6` to `=B4` so the
formula continues to reference Cherry/30. This matches Excel's
actual Remove Duplicates behavior.

A5. Absolute references ($A$1) ALSO shift when the referenced row is
below the deletion point — the `$` markers are preserved by
`shiftReferenceToken`'s `colDollar`/`rowDollar` capture groups,
but the row number itself shifts. This is Excel's exact
behavior: deleting a row shifts ALL subsequent row indices
(whether absolute or relative); the `$` only protects against
formula AUTOFILL / drag-shift, not against row DELETION.

A6. Row-level styles, cell styles, merges, heights, hidden state,
notes, validations, filters, conditional formatting all travel
atomically with the row: - Cell styles: live inside `<c s="...">` s= attribute →
untouched by `transformSheetRows` (only r= is renumbered). - Row heights/hidden/outline: live inside `<row ht="..."
          hidden="..." outlineLevel="...">` attributes → untouched. - Merges: shifted by `transformRangedFeatures` via
`moveRefRange`. - Hyperlinks, dataValidation sqref, conditionalFormatting sqref:
shifted by `transformRangedFeatures` via `moveRefRange`. - Notes: live in `xl/worksheets/sheetN.xml` as `<legacyDrawing>`
references + the comment part (`xl/commentsN.xml`) keyed by
ref. Univer's note subscription snapshots the LIVE note
model declaratively (`collectNoteStates`), so the note ref
coordinates are re-snapshotted AFTER structural ops land. - Filters: declarative snapshot (`collectFilterStates`). - DV: declarative snapshot (`collectDvStates`).

A7. YES. The browser doesn't need to read formulas/styles at all.
The flow is: 1. Read COMPUTED VALUES via `range.getValues()` (for the
duplicate-detection comparison only — same call the desktop
makes). 2. Compute the SET of duplicate row indices (the dedupe
algorithm returns indices, not values). 3. For each duplicate (in DESCENDING row order), call
`ws.deleteRows(startRow + offset, 1)`. Univer's live model
handles the cell-record movement atomically. 4. ExcelEditor's existing `STRUCTURAL_MUTATION_IDS`
subscription journals each `sheet.mutation.remove-rows` as
a `remove-rows` structural op in the save plan. 5. Gateway's `applyStructuralOps` applies each op atomically
at OOXML level.
No browser-side OOXML parser, no JSZip, no formula engine. The
browser only exchanges typed `BrowserStructuralOp` payloads.

A8. YES. The canonical `remove-rows` structural op family is the
existing primitive. No new family, no new wire field, no new
gateway function. The save plan already carries `structuralOps`
(added for EXCEL-011 Insert/Delete Rows); EXCEL-018 just EMITS
`remove-rows` ops through that existing channel.

A9. The desktop Remove Duplicates (`apps/sheets/src/renderer/dedupe.ts` + `apps/sheets/src/renderer/ribbon-actions.ts:1267-1309`) IS
architecturally lossy — it pads with nulls and rewrites via
`setValues`, so moved rows become computed literals. The prior
web implementation copied that lossy recipe verbatim. The
architect's audit explicitly states: "desktop parity is not
sufficient to satisfy the frozen web parity acceptance
requirement." The frozen requirement at
`spec/excel/verification-matrix.md:39` explicitly lists
"formulas and styles remain attached to retained rows" and
"save/reopen yields identical values and structure" — these
CANNOT be satisfied by value-level rewriting; they REQUIRE
structural row deletion.

Decision: Option 1 — express Remove Duplicates as canonical
`remove-rows` structural ops (one per duplicate row, descending
order). No new mutation family, no new wire field, no new gateway
function. Reuse the EXACT structural path `excel-structural.spec.ts`
already proves for Insert/Delete Rows.

Stage Summary (audit only — no code changes yet):

- The canonical engine ALREADY has a safe formula/style-preserving
  primitive for Remove Duplicates: the `remove-rows` structural op.
- The current `setValues`-based implementation is a wrong turn; it
  destroys formulas on moved rows and the E2E test explicitly proves
  the loss. This is the ROOT CAUSE the architect identified.
- The fix is mechanical: replace the value-rewrite path in
  `useExcelRuntime.removeDuplicates` with `ws.deleteRows(...)` calls
  (one per duplicate, descending). Update the architecture test to
  require the structural path. Rewrite the E2E to assert formula
  preservation + reference rewriting. Add the second regression case
  with absolute/relative/mixed reference types.

---

## 2026-018 — EXCEL-018 implementation: structural remove-rows path (architect fix)

Task ID: EXCEL-018-structural-implementation
Agent: Z.ai (Implementation Agent)

Task:

- Implement the architect's correction: Replace the formula-destroying
  `setValues` rewrite with the canonical structural `remove-rows` path.
- Add the architect's two mandatory regression cases.
- Run the full local regression suite.

Work Log (Implementation):

- `apps/web/src/office/dedupe.ts`:
  - Added `dedupeRowIndices(rows, hasHeader)` — the canonical EXCEL-018
    entry point. Returns `{ keptIndices, duplicateIndices, removed }`.
    Pure (zero imports). `dedupeRows` kept for backward compat (delegates
    to dedupeRowIndices and maps keptIndices through the input matrix).
- `apps/web/src/screens/excel/useExcelRuntime.ts`:
  - Replaced the entire `removeDuplicates` callback body. New flow:
    1. Read computed values via `range.getValues()` (comparison only).
    2. `dedupeRowIndices(values, hasHeader)` → duplicate row offsets.
    3. For each duplicate offset in DESCENDING order, call
       `ws.deleteRows(startRow + offset, 1)`.
  - Each `deleteRows` fires `sheet.mutation.remove-rows`, journaled by
    ExcelEditor's existing `STRUCTURAL_MUTATION_IDS` subscription as a
    `{ kind: 'remove-rows', index, count: 1 }` structural op in the
    save plan. The gateway's `applyStructuralOps` applies each op
    atomically (`transformSheetRows` + `transformFormulas` +
    `transformRangedFeatures`).
  - No `setValues` call anywhere in the removeDuplicates body (enforced
    by the new architecture test).
- `apps/web/tests/architecture.test.ts`:
  - Updated the EXCEL-018 guard block from "canonical cell-edit path"
    to "canonical structural remove-rows path".
  - New guard: `useExcelRuntime.removeDuplicates does NOT rewrite moved
rows via setValues` — locates the removeDuplicates useCallback body
    by bracket-walking and asserts it calls `.deleteRows(` and does NOT
    call `.setValues(` and does NOT contain a "padded" variable.
  - New guard: dedupe.ts must export `dedupeRowIndices`.
  - Kept: purity guards (no Univer/electron/node/fs/jszip imports),
    getValues guard, Ribbon button-enabled guard, no-OOXML guard,
    no-new-family guard.
- `apps/web/tests/dedupe.test.ts`:
  - Extended with a dedicated `dedupeRowIndices` describe-block (13 new
    tests): basic index return, header semantics, the architect's
    mandatory Apple/Banana/Cherry fixture (kept=[0,1,3,5,6],
    duplicates=[2,4]), all-columns key, no-op empty array, null vs
    empty-string, number vs text, boolean vs text, header-only, purity,
    descending-deletion stability, and a dedupeRows/dedupeRowIndices
    parity test.
- `apps/web/tests/e2e/fixtures.ts`:
  - Added `buildExcelDedupeMixedReferencesFixture()` — the architect's
    second regression case: a "DedupeMixed" sheet with 3 duplicate rows
    (rows 3, 5, 7) and a survivor row (Banana at row 4) carrying FOUR
    formulas (C4="=$D$6" absolute, D4="=A6" relative, E4="=$A6" mixed
    col-$, F4="=A$6" mixed row-$) referencing Cherry/Anchor at row 6
    (outside the dedupe selection A1:B7).
- `apps/web/tests/e2e/ribbon-remove-duplicates.spec.ts`:
  - Test 1 REWRITTEN (architect mandatory regression): asserts
    B5 formula text is "=B4" in Univer's live model (formula PRESERVED,
    reference rewritten B6→B4 — NOT the literal 30), A4 carries
    Cherry/30 (the referenced cell compacted), rows 6/7 are DELETED
    (no "padding" nulls), the save plan carries exactly 2
    `{kind:'remove-rows', index:2/4, count:1}` ops in structuralOps
    (NOT a new family), the saved XML carries `<f>B4</f>` at B5, and
    the reopened snapshot carries formula "=B4" (with leading "=" per
    the canonical readBasicWorkbook parser; cell.value is null for
    formula cells — the parser drops the cached <v>).
    Also asserts: NO CellEdit at B5 carries value=30 without a formula
    (the architect's explicit failure mode).
  - Test 2 NEW (architect second regression): asserts all four
    reference types survive compaction with their references rewritten
    ($D$6→$D$4, A6→A4, $A6→$A4, A$6→A$4), in Univer's live model AND
    the saved XML AND the reopened snapshot. Asserts the save plan
    carries exactly 3 remove-rows ops (indices 2, 4, 6).
  - Tests 3/4 (no-op fail-closed, <2-rows fail-closed) kept verbatim.

Local Results (all green):

- Typecheck (`tsc --noEmit -p apps/web`): exit 0.
- Prettier: all 6 changed files pass `--check`.
- Web unit suite (vitest): 197/197 pass (10 files; includes the 27
  dedupe tests and 24 architecture tests with the new structural
  guards).
- EXCEL-018 E2E (4 tests): ALL PASS —
  1. mandatory regression (formula PRESERVED, B6→B4 rewritten, rows
     deleted, structuralOps on the wire, saved XML has <f>B4</f>,
     reopen carries formula "=B4") — 9.4s
  2. second regression (all four reference types rewritten
     $D$6→$D$4 / A6→A4 / $A6→$A4 / A$6→A$4) — 8.1s
  3. no-op fail-closed — 7.0s
  4. <2-rows fail-closed — 7.3s
- Mandatory regression suite (architect's list), all green:
  - excel-structural: 2/2 (26.6s)
  - excel-formula: 7/7 (1.1m)
  - excel-format: 3/3
  - excel-browser: 1/1
  - excel-shell: 15/15 (1.8m)
  - ribbon-data: 4/4 (45.7s) — INCLUDING the architect's
    sort/formula semantic gate (test 3: "relative refs rewrite,
    absolute refs untouched")
  - ribbon-data-validation: 7/7 (1.0m)
  - ribbon-filter: 5/5 (45.4s)
  - ribbon-review-notes: 5/5
  - ribbon-home-persistence: 3/3
  - ribbon-insert: 1/1
  - word-browser: 1/1 (11.3s)
- Total local E2E: 47/47 (4 EXCEL-018 + 43 regression) all green.
- Frozen-surface check:
  `git diff --stat -- apps/sheets apps/docs apps/shell
 packages/platform-electron packages/renderer-bridge`
  is EMPTY — no frozen surfaces touched.

Stage Summary:

- The canonical structural `remove-rows` path is implemented and proven.
- The architect's two mandatory regression cases both pass end-to-end
  through the REAL browser + REAL HTTP + REAL gateway boundary.
- The formula on a compacted row is PRESERVED as a formula with its
  reference rewritten to track the moved referenced cell — NOT
  converted to a computed literal.
- Ready to commit, push, and verify CI.
- Workflow state remains ARCHITECT_REVIEW / CHANGES REQUIRED (the
  architect owns the VERIFIED decision).

---

## 2026-08-26 — EXCEL-018 correction pushed: CI evidence (commit 82a87c9)

Task ID: EXCEL-018-correction-ci-evidence
Agent: Z.ai (Implementation Agent)

Task:

- Push the architect correction (commit 82a87c9) and collect CI evidence.
- Verify the canonical `web` job (the EXCEL-018 gate) is green.
- Classify any remaining failures as pre-existing or new.

Work Log:

- Pushed commit `82a87c9` to origin/web-office-editor (PR #1 head updated).
- CI results for `82a87c9`:
  - `web` job (canonical EXCEL-018 gate): SUCCESS. All steps green:
    Typecheck (web app + office API host), Unit tests, Production build
    (web app), Install Playwright chromium, Playwright browser E2E (real
    Vite + API stack — 77 tests including all 4 EXCEL-018 tests and the
    architect's sort/formula semantic gate).
  - `test` job: FAILURE at the Lint step — pre-existing (350 lint errors
    across non-EXCEL-018 files: frozen desktop apps apps/sheets,
    apps/docs, apps/slides; other web screens BOQ.tsx/Bid.tsx/
    Estimate.tsx/WordEditor.tsx; other packages contractor-core,
    platform-electron; skill scripts). Identical to the prior session's
    commit 951470d. None of my EXCEL-018 files appear in the lint errors.
  - `e2e` job: FAILURE at the "E2E (Electron shell)" step — pre-existing
    (the FROZEN desktop Electron suite; chromium-headless-shell binary
    missing + desktop sheets test failures). Identical to the prior
    session's commit. The web Playwright suite (the EXCEL-018-relevant
    one) passes via the `web` job.
  - `Project foundation` workflow: FAILURE — pre-existing structural
    condition (the branch-isolation check diffs origin/main...HEAD and
    flags apps/sheets/, apps/web/src/screens/excel/, packages/
    xlsx-gateway/ changes; web-office-editor carries the whole parity
    phase so this fails for EVERY commit on the branch).
  - `Web Office Editor Preview` workflow: FAILURE at the "Verify deployed
    Excel/Web parity" step — pre-existing infrastructure issue (verified
    in the job logs: the VERCEL_PROJECT_ID secret deploys to
    ekonplacidegmailcoms-projects/genoffice, a DIFFERENT Vercel project;
    the deployed preview URL genoffice-5azrj6jwr-ekonplacidegmailcoms-
    projects.vercel.app is not the genoffice web app, so the E2E demo
    login (getByRole('button', { name: 'Owner' })) times out
    universally). Identical to the prior session's commit. This is a
    GitHub Actions secret misconfiguration the architect must fix; it
    would fail identically for any commit pushed to this branch.
    Notably, all build/deploy steps in this workflow ARE green
    (typecheck, unit tests, deterministic API bundle, web build, Vercel
    deploy) — only the deployed-parity E2E step fails, on the wrong app.

Stage Summary:

- The EXCEL-018 correction is pushed at commit 82a87c9.
- The canonical `web` CI gate is GREEN — independently verifying
  typecheck, unit tests (197/197), production build, and all 77
  Playwright browser E2E tests (including the 4 EXCEL-018 tests and the
  architect's sort/formula semantic gate) on a clean CI machine.
- All other CI failures are pre-existing and unrelated to EXCEL-018
  (lint debt, frozen Electron suite environment, branch-isolation
  structural condition, Vercel secret misconfiguration) — each verified
  identical to the prior session's commit 951470d.
- Workflow state: ARCHITECT_REVIEW / CHANGES REQUIRED → CORRECTION
  SUBMITTED. NOT VERIFIED — the architect owns the VERIFIED decision.

---

Task ID: EXCEL-020-impl
Agent: Z.ai (Implementation Agent)
Task: Implement EXCEL-020 Sheet Protection / Workbook Protection on the merged main (640c6c9) — forensic audit → canonical contract → wire → browser → E2E → regression.

Work Log:

- Ground-truth check: cloned pectoraux/genoffice @ 640c6c9 (the PR #1 merge commit), working tree clean, EXCEL-018 correction confirmed merged. The EXCEL-020 ledger entry was READY and sequenced next.
- Phase A forensic audit (read-only): the gateway WRITE side already existed and was unit-tested (applySheetProtection / applyWorkbookProtection / applyProtectedRanges + planCellEditsToXlsx's sheetProtections arg 10, protectedRangeStates arg 24); the desktop ships journal-only toggle semantics (recordSheetProtection DROPS an entry when the desired state matches the file's original; password-bearing elements refuse both directions) and does NOT enforce protection live; the desktop's cell locked/hidden flags journal neutral WorkbookStyleEdit deltas. The GAPS were: readBasicWorkbook parsed no protection state; applyCellEditsToXlsx (the web route's buffer→buffer entry) had no workbookProtectionState parameter; the wire contract had no sheetProtections/workbookProtectionState; the web Ribbon carried a disabled "Protect Sheet" stub documenting the missing family.
- Gateway read: parseSheetProtectionState + parseWorkbookProtectionState (xlsx-protection.ts) → WorksheetState.sheetProtection + WorkbookSnapshot.workbookProtection (workbook.types.ts), both carrying hasPassword for BOTH password forms (legacy hash + modern algorithmName/hashValue), wired into readBasicWorkbook (absent field = no element).
- Gateway write: applyCellEditsToXlsx gained the trailing workbookProtectionState parameter (pass-through to planCellEditsToXlsx → the existing applyWorkbookProtection).
- Wire: BrowserWorkbookSavePlan.sheetProtections + workbookProtectionState with expectSheetProtectionState / expectWorkbookProtectionState strict validation (unknown fields — including password-bearing payloads — are 400s; MAX_SHEET_PROTECTIONS = 1,000); handleSaveWorkbook passes both families; office-client.ts mirrors the types.
- Browser: ExcelEditor seeds protection file-state refs from the snapshot, owns toggle journals with the desktop's recordSheetProtection/recordWorkbookProtection semantics (toggle-back drops the entry → no-op save emits nothing → XML preserved), refuses password-protected sheets/structures up front with the desktop's own status strings, conditionally emits both families on save, and merges the journal into the file refs after save. The Ribbon's Review → Protection group is fully wired: Protect Sheet / Unprotect Sheet (echo label flips with the effective state), Protect Workbook / Unprotect Workbook, Lock Cell / Unlock Cell (journal canonical protectionLocked style-only CellEdits — the desktop's neutral-delta path).
- Tests: gateway xlsx-protection-roundtrip.test.ts (22 tests: parsers, read integration, write, round-trip, no-op preservation, fail-closed both password forms, verbatim re-affirm, unknown sheet); wire office-protection-routes.test.ts (12 tests); architecture guards (7 new, 31 total); E2E ribbon-protection.spec.ts (5 tests through real HTTP: read+echo, protect→wire→XML→reopen, unprotect round-trip with toggle-back semantics, editable-vs-locked via the REAL ribbon + no-op preservation, negative authorization with password elements surviving verbatim).
- Regression: gateway 567/567, contractor-core 390+4 skipped, web 204/204, browser E2E 83/83 (excel-browser 1, excel-shell 15, excel-format 3, excel-formula 7, excel-structural 2, ribbon-data 4 incl. the sort/formula semantic gate, ribbon-data-validation 7, ribbon-filter 5, ribbon-review-notes 5, ribbon-home-persistence 3, ribbon-insert 1, ribbon-view 3, ribbon-remove-duplicates 4, ribbon-protection 5, word specs 18). Typecheck exit 0 across web/contractor-core/xlsx-gateway; prettier clean on all changed files; frozen surfaces (apps/sheets, apps/docs, apps/shell, packages/platform-electron, packages/renderer-bridge) untouched.

Stage Summary:

- EXCEL-020 implemented on the EXISTING canonical families — no new engine path, no browser-side OOXML, no parallel protection model, no wire-breaking change (both new save-plan fields are optional).
- Password semantics are fail-closed at every layer (browser guard → wire rejection → gateway SheetProtectionError); desktop-parity journal semantics preserve byte-for-byte no-op saves.
- Scope note: Allow Edit Ranges (protectedRanges) stays desktop-only — not in the EXCEL-020 verification list; the engine-side writer (applyProtectedRanges) remains available for a future work item.
- Workflow state: IMPLEMENTED / PENDING ARCHITECT REVIEW. NOT VERIFIED — the architect owns the VERIFIED decision.

---

Task ID: EXCEL-020-ci
Agent: Z.ai (Implementation Agent)
Task: Push the EXCEL-020 implementation (PR #12) and collect CI + production-build evidence.

Work Log:

- Pushed commit `44ea624` to origin/excel-020-protection; opened PR #12 (base main).
- First CI round (`44ea624`): the `web` job FAILED at the Typecheck step — the contractor-core tsconfig includes tests, and the new office-protection-routes.test.ts passed an OBJECT literal to the helper's `unknown[]`-typed parameter in the non-array-rejection case (error: "'sheetName' does not exist in type 'unknown[]'"). My local pre-push typecheck had covered apps/web and contractor-core's src but not the tests-inclusive workspace run the CI performs.
- Fix commit `2c4d86c`: widened the test helper's sheetProtections/workbookProtectionState parameter types to `unknown` (the payloads are intentionally invalid shapes). All three CI workspaces typecheck clean locally: @contractor/web, @contractor/web-host, @contractor/core.
- CI results for `2c4d86c`:
  - `web` job (the canonical office gate): SUCCESS. All steps green — Typecheck (web app + office API host), Unit tests, Production build (web app), Install Playwright chromium, Playwright browser E2E (real Vite + API stack, including the 5 EXCEL-020 protection tests and the architect's sort/formula semantic gate).
  - `test` job: FAILURE at the Lint step — pre-existing lint debt (351 errors across the FROZEN desktop apps apps/sheets + apps/docs, other web screens BOQ.tsx/Bid.tsx/Estimate.tsx/WordEditor.tsx, and non-office packages contractor-core commercial/persistence + platform-electron). Verified NONE of the EXCEL-020 files appear in the error list (the single architecture.test.ts match is packages/platform-electron's frozen copy, not apps/web's).
  - `e2e` job: FAILURE at the "E2E (Electron shell)" step — pre-existing frozen-desktop environment issue (chromium_headless_shell-1234 binary missing + desktop sheets InvalidSessionError/sandbox failures), identical in character to the EXCEL-018 session's documentation.
  - `foundation` workflow: FAILURE — pre-existing branch-isolation structural condition (diffs origin/main...HEAD and flags the parity-phase file set; fails for every commit on a feature branch).
  - Vercel preview deployment: FAILED with the account's free-tier rate limit ("Resource is limited - try again in 24 hours (more than 100, code: api-deployments-free-per-day)") — an account-level infrastructure constraint; the deployed production verification therefore remains a POST-MERGE step (exactly as EXCEL-018's production evidence was collected after its merge).
- Production-build evidence in lieu of the unavailable preview: ran the EXCEL-020 spec against the BUILT bundle (npx vite build + vite preview on :5178 with the dev-server API on :5179) via scripts/run-protection-vs-preview.sh — all 5 tests pass against the production-served URL (29.8s), proving the feature works against the built artifact, not just the Vite dev server.

Stage Summary:

- PR #12 carries EXCEL-020 at `2c4d86c` with the canonical `web` CI gate GREEN.
- All other CI failures are pre-existing and verified unrelated to EXCEL-020 (frozen-surface lint debt, frozen Electron suite environment, branch-isolation condition, Vercel free-tier deploy limit).
- Production-build E2E: 5/5 green against vite preview.
- Workflow state: IMPLEMENTED / PENDING ARCHITECT REVIEW. NOT VERIFIED — the architect owns the VERIFIED decision; deployed (genoffice.vercel.app) verification follows the merge.

---

Task ID: EXCEL-022-implementation
Agent: Z.ai (Implementation Operator)
Task: EXCEL-022 — Images / Drawings. Phase A forensic audit, then the full implementation (canonical reader, delete cascade, wire contract, browser rendering/journal/save, tests, fixtures, E2E).

Work Log:

- Rebased on current main 93b4c30 (fresh clone; origin had moved past 2aa2b74 with PROJECT-019). Branch: excel-022-images. Re-applied the EXCEL-021 cleanup (unused TABLE_REL_TYPE removal — still not on origin) as the branch's first commit.
- ARCHITECTURE DECISION (first required task): WorksheetState.visuals is CHART-SPECIFIC — SheetVisual.kind is the literal 'chart' with a required ChartVisualState; it is populated only by the desktop demo/AI replay (apps/sheets/src/domain/in-memory-workbook.ts + visual-edit-sync.ts), never by readBasicWorkbook, and apps/web never references it. Decision: visuals stays UNTOUCHED; images landed as the DEDICATED WorksheetState.images (readonly SheetImageInfo[]) — the same pattern EXCEL-021 used for tables. No duplicate chart/image abstraction introduced.
- Phase B — canonical reader: NEW packages/xlsx-gateway/src/gateway/xlsx-image-read.ts. Resolves worksheet → drawing rel → drawing part → image rels → xl/media/*, returns typed SheetImageInfo (drawingPath, drawingIndex, anchorType, 8-number anchor, widthPx/heightPx for one-cell/absolute from a:ext, rotationDeg from a:xfrm rot, name from cNvPr, mediaType, inline dataUrl). drawingIndex counts EVERY anchor in document order (regex parity with xlsx-drawing-edit.ts and the Rust sidecar). Media set = the writer's (png/jpeg/jpg/gif); unsupported types, missing parts, and per-image oversize are skipped per picture; unreadable drawing wiring and per-sheet total oversize fail closed PER SHEET (workbook still opens; no-op save preserves bytes). createBufferEntrySource gained a readBinary member (EntrySource itself unchanged — the platform adapter never reads images).
- Phase C — delete cascade: xlsx-drawing-edit.ts. Removing an xdr:pic anchor collects its r:embed rel id; after the drawing XML is final, cascadeImageRemovals drops the image relationship (UNLESS another anchor still embeds it — shared media stays), then removes the media part ONLY when no remaining relationship ANYWHERE in the package resolves to it (all .rels parts scanned, external targets skipped), and removes the [Content_Types] Default for the extension only when no other package entry carries that extension. Pictures without r:embed (synthetic) splice plainly — legacy behavior preserved. Charts cascade unchanged.
- Phase D — wire contract: BrowserWorkbookSavePlan gained visualAdditions/visualEdits (canonical gateway types, optional fields). office-routes.ts strict validation: expectSheetVisualAddition (IMAGE-ONLY — chart/shape payloads rejected; sheetName; 8 bounded non-negative integer anchor fields with unknown-field rejection; mediaType in {png,jpeg,gif}; base64 string with 11M-char cap) and expectWorkbookVisualEdit (drawingPath pattern xl/drawings/*.xml, drawingIndex 0..10000, exactly one of remove:true|anchor, unknown-field rejection). Caps: 50 additions, 200 edits. parseSaveWorkbookRequest + handleSaveWorkbook forward both families into applyCellEditsToXlsx as the two trailing parameters (desktop translator call sites untouched; plan ordering verified: structural ops (line 966) run BEFORE visualEdits (1161) so edit anchors are post-shift coordinates).
- addedVisuals locator plumbing: applyVisualAdditions now returns per-addition {worksheetPath, drawingPath, drawingIndex} (appendAnchor returns the appended anchor's document-order index); MutationPlan/XlsxMutation carry addedVisuals; SaveWorkbookResponse returns it; the browser merges persisted session images into fileImagesRef with the EXACT assigned index (no guessed indexes — fail-closed against silent corruption).
- Phases E-I — browser: NEW apps/web/src/office/sheet-images.ts (pure typed mapping; documented EMU↔px conversion 1px=9525EMU; placementFromAnchor/anchorFromPlacement grid walks; readLivePlacement/readLiveAnchor through the Univer facade; collectImageVisualEdits/collectImageVisualAdditions save adapters). ExcelEditor: fileImagesRef/imageAddsRef/imageDirtyRef/imageRemovalsRef/imageInstallingRef state; loadSnapshot installs images under journal suppression with locator-keyed Univer drawing ids (async loadSnapshot); subscribeToImageMutations journals sheet.mutation.set-drawing-apply (UPDATE → dirty two-cell images; REMOVE → removal journal / session-add splice; INSERT ignored — the ribbon path journals directly); handleSave emits visualEdits+visualAdditions (anchors read from the LIVE model) and merges post-save state; Insert → Picture via hidden file input (File/Blob → dataURL, desktop ≤480px scaling parity) with data-testid excel-image-input. Ribbon: Picture enabled (Chart stays disabled until EXCEL-023). One-cell/absolute anchors FAIL CLOSED: the refused edit is reverted (remove + reinstall at file geometry — the Univer set-drawing command's param diff collapses to an empty op when invoked outside the originating interaction, so a surgical setPositionAsync does not reliably land) with a status explanation; nothing journals.
- Phase K — fixtures + E2E: 5 new deterministic fixtures (single PNG twoCell, two PNGs isolation, multi-sheet SHARED media, minimal JPEG, oneCellAnchor) + readZipEntryBytes helper. NEW apps/web/tests/e2e/ribbon-images.spec.ts — 12 tests, all through real browser → HTTP → routeOffice → gateway → bytes → reopen: import/metadata, render identity (locator id + dataUrl + from/to), move (only anchor changes; media byte-identical), resize, insert (full part set: media + drawing + rel + content-type + neighbor survival), delete cascade (anchor/rel/media/Default all gone), multi-image isolation (move one + delete other), save/reopen with moved anchor, no-op preservation (plan carries NO visual families; 6 entries byte-identical; drawing wiring verbatim), relationship chain integrity (worksheet→drawing→image→media after move+insert), one-cell fail-closed (reverted + no visualEdits + drawing byte-identical), multi-sheet shared media, JPEG round-trip.
- Phase L — regression: FULL browser E2E 101/101 green (30 excel specs + 43 ribbon specs + 28 word/image specs; ribbon-insert.spec.ts updated — Picture flipped to ENABLED, the documented disabled-control test now covers Chart only). Gateway 602/602 (584 pre-existing + 18 new image tests). contractor-core 423 passed + 4 skipped (404 pre-existing + 19 new route tests). web unit 216/216 (213 + 3 architecture guards). web-host 78/78. Typecheck clean in all four workspaces; eslint clean on every changed file; production build green.
- Phase M — architecture guards: no drawing/relationship XML patterns (<xdr:*, <a:blip, relationships/image|drawing, xl/media/, xl/drawings/, Target="../) in apps/web/src; gateway imports are TYPE-ONLY (value imports from @genoffice/xlsx-gateway forbidden); office-client + sheet-images must reference the canonical SheetVisualAddition/WorkbookVisualEdit/SheetImageInfo types.
- Frozen surfaces verified untouched (git diff vs origin/main over apps/sheets/src, apps/docs, apps/shell, packages/platform-electron, packages/renderer-bridge = empty).
- Spec ledgers updated (work-items.md status IMPLEMENTING + implementation summary; verification-matrix.md local evidence; dependency-graph.md already carries the EXCEL-021→022→023 chain).

Stage Summary:

- EXCEL-022 implemented end-to-end on the EXISTING canonical visual families — no new engine path, no browser-side OOXML/JSZip/drawing parsing, no duplicate image model, no wire-breaking change (both new save-plan fields optional; the open response's images ride the existing snapshot).
- All local gates green: gateway 602/602, core 423+4sk, web 216/216, host 78/78, browser E2E 101/101 (12 new image tests), architecture 43/43, typecheck/lint/build clean, frozen surfaces untouched.
- Known limitations (documented): chart/shape additions rejected on the web wire until EXCEL-023; one-cell/absolute images render read-only (edits reverted fail-closed); rotation renders but is not editable; z-order/flip/crop/alt-text not modeled on the xlsx wire (Word has crop/alt; desktop xlsx does not); a save response without addedVisuals (older host) leaves persisted session images untracked for further edits (no guessed indexes).
- Workflow state: IMPLEMENTING → local implementation complete. The sandbox has no GitHub push credentials, so PR_OPEN/CI/production deploy could NOT be executed — patch files exported to /home/z/my-project/download/ for application. NOT VERIFIED — the architect owns that transition.

---

Task ID: EXCEL-022-deployed-preview
Agent: Z.ai (Implementation Operator)
Task: Deployed-deployment verification attempt for EXCEL-022 (preview only — production alias intentionally untouched per the architect's ordering: production deploy follows CI green + merge).

Work Log:

- Verified the Vercel token against the account: project `genoffice` exists (team ekonplacidegmailcoms-projects, production URL https://genoffice.vercel.app, Node 24.x).
- Deployed the EXCEL-022 branch (commit 7b1cc15) as a Vercel PREVIEW via the CLI from the local checkout — build succeeded on Vercel's infrastructure (bun install --ignore-scripts → build:api + apps/web build): https://genoffice-e1kar1x9m-ekonplacidegmailcoms-projects.vercel.app
- Attempted the deployed E2E suite (playwright.deployed.config.ts, ribbon-images.spec.ts) against the preview: ALL 12 tests fail at LOGIN — the account's Deployment Protection (SSO) intercepts preview URLs and redirects to vercel.com/login, so Playwright never reaches the app (smoke: preview root → 302 → vercel.com/login; production root → 200). This is an account-level infrastructure condition, identical in character to the documented Vercel free-tier preview limitation from the EXCEL-020 session.
- Verified the production alias is publicly reachable (200) but deliberately NOT redeployed — it currently serves merged main WITHOUT EXCEL-022; pushing unreviewed code to the production alias would violate the mandated ordering (local green → CI web green → merge → production deploy → production E2E → byte inspection). That chain remains blocked upstream at the GitHub push step (no credentials in this sandbox).
- Patch series exported for application: /home/z/my-project/download/0001 (EXCEL-021 cleanup re-apply), 0002 (EXCEL-022 implementation).

Stage Summary:

- EXCEL-022 commit 7b1cc15 builds and deploys cleanly on Vercel (preview URL above) — infrastructure-level proof the exact tree compiles in the deploy pipeline.
- Deployed E2E against the preview: BLOCKED by account Deployment Protection (SSO). Production verification (Phase N) remains pending the PR/merge that the sandbox cannot perform.
- Workflow state: IMPLEMENTING (local implementation + all local gates green). NOT PR_OPEN (no GitHub credentials), NOT VERIFIED (architect owns that transition).

---

Task ID: EXCEL-022-publish-correction
Agent: Z.ai (Implementation Operator)
Task: Respond to the architect's independent remote verification finding: no `excel-022-images` ref exists on GitHub; the reported commits (e1d025d, 7b1cc15, ab325cc) are not remotely verifiable. Make the branch publishable and re-verify every gate on the exact tree to be pushed.

Work Log:

- Architect finding CONFIRMED CORRECT. `git fetch origin` shows no `origin/excel-022-images`; the three commits exist only in this sandbox. My previous report cited local SHAs without stating prominently enough that they were never pushed — that ambiguity is the defect being corrected here.
- Root cause, verified exhaustively: this sandbox has ZERO GitHub credentials — no credential helper, no ~/.git-credentials, no ~/.netrc, no GITHUB_TOKEN/GH_TOKEN env vars, no gh CLI, no SSH keys. `GIT_TERMINAL_PROMPT=0 git push --dry-run` fails with "could not read Username for 'https://github.com'". Push and PR creation are physically impossible from here. (The Vercel token deploys to Vercel only; it cannot touch GitHub.)
- Rebased `excel-022-images` onto current origin/main (`93b4c30` → `3d7fc3c`, clean rebase; main had advanced with PR #18/PROJECT-020 which touches no EXCEL-022 file). SHA mapping (neither set ever existed on the remote): e1d025d → a594137 (EXCEL-021 cleanup), 7b1cc15 → 3c6a00a (EXCEL-022 implementation), ab325cc → e058b9f (worklog evidence). Pre-rebase tip preserved locally as `excel-022-images-prerebase`.
- Fresh gate evidence ON THE REBASED TREE (all re-executed this session, none carried over from the earlier report):
  - xlsx-gateway: 602/602 (32 files).
  - contractor-core: 423 passed + 4 skipped (26 files + 1 skipped).
  - web unit: 216/216 (10 files, includes the 43 architecture guards).
  - web-host: 78/78.
  - Typecheck clean: @contractor/web, @contractor/web-host, @contractor/core, @genoffice/xlsx-gateway.
  - Browser E2E (real Vite + API stack): 101/101 — batch 1 (excel + insert): 50/50; ribbon-images: 12/12; batch 2 (ribbon + word): 39/39.
  - Frozen surfaces vs origin/main: empty diff over apps/sheets/src, apps/docs, apps/shell, packages/platform-electron, packages/renderer-bridge. Full PR: 21 files, +3773/−39.
- Publish package exported for credential-holders: /home/z/my-project/download/excel-022-images/ — 3-patch series (git am onto current main), pr-combined.diff, pr-diffstat.txt, SHA256SUMS, PUBLISH.md (exact push + PR-open commands). Stale pre-rebase 0003 patch moved to superseded/.
- Publication remains with the architect/operator: apply patches → push excel-022-images → open PR. Alternatively, provision a GitHub token in this sandbox and the push + PR happen immediately from here.

Stage Summary:

- The implementation is real, complete, and freshly re-verified (602/602, 423+4s, 216/216, 78/78, 101/101 E2E, typecheck clean, frozen surfaces untouched) on branch excel-022-images rebased at 3d7fc3c + 3 commits (tip after this worklog commit).
- Push/PR cannot originate from this sandbox (no credentials — verified, not assumed). Patch series + instructions exported; ready for `git am` + push by anyone with repo access.
- Workflow state: EXCEL-022 remains IMPLEMENTING. NOT PR_OPEN (no remote ref — architect's finding stands), NOT VERIFIED (architect owns that transition). EXCEL-021 cleanup likewise rides this branch (a594137) and remains NOT VERIFIED until it lands on remote main.

---

Task ID: EXCEL-022-pr-open
Agent: Z.ai (Implementation Operator)
Task: Publish the excel-022-images branch and open the PR (architect supplied a GitHub PAT after confirming the remote ref was missing).

Work Log:

- Architect supplied a GitHub PAT (exported as GITHUB_TOKEN, persisted in ~/.bashrc; authenticates as pectoraux with full repo scope). Note for future sessions: each shell invocation is fresh — `source ~/.bashrc` before any authenticated command.
- Verified origin/main still at 3d7fc3c (branch base unchanged; project-office moved but is unrelated). Pushed excel-022-images: remote ref = f10ad0ffd8371e70c9b5a673fcc8592bd20ea820 = local HEAD.
- Opened PR #20 (base main ← head excel-022-images): https://github.com/pectoraux/genoffice/pull/20 — 4 commits, 21 files, +3801/−39. PR body carries the full implementation summary, gate evidence, commit mapping (pre-rebase local SHAs → published SHAs), known limitations, and workflow state.
- The patch package at /home/z/my-project/download/excel-022-images/ is now SUPERSEDED by the remote ref (kept for the audit trail; tree identity was proven before pushing).

Stage Summary:

- EXCEL-022 workflow state: IMPLEMENTING → PR_OPEN. Remote ref and PR exist and are reviewable.
- CI evidence collection on PR #20 follows (web job is the canonical office gate; test/e2e/foundation failures are the documented pre-existing baseline if they recur).
- NOT VERIFIED — the architect owns that transition; review proceeds on the actual diff.

---

Task ID: EXCEL-022-ci-round1
Agent: Z.ai (Implementation Operator)
Task: Collect CI evidence on PR #20 (first run, head 8fcf26a) and correct what it surfaced.

Work Log:

- First CI round on 8fcf26a: the `test` job FAILED at `npm run format:check` — Prettier violations in 11 files, ALL of them EXCEL-022 files (sheet-images.ts, ExcelEditor.tsx, fixtures.ts, ribbon-images.spec.ts, office-routes.ts, office-visual-routes.test.ts, xlsx-drawing-add.ts, xlsx-drawing-edit.ts, xlsx-image-read.ts, xlsx-image.test.ts, work-items.md). Root cause: my local pre-push gates ran ESLint on changed files but never the repo's Prettier format gate (`FORMAT_BASE_REF=origin/main npm run format:check`). This was a REAL defect in my verification process — not pre-existing debt.
- Correction: ran `npm run format` (Prettier write) over the changed-vs-origin/main file set — pure formatting (99+/100−, line wrapping only). `format:check` now passes.
- Gates re-verified on the formatted tree: xlsx-gateway 602/602, contractor-core 423+4sk, web unit 216/216, web-host 78/78, typecheck clean in all four workspaces, ribbon-images browser E2E 12/12.
- `foundation` job: FAILED — the documented pre-existing branch-isolation structural condition (diffs origin/main...HEAD and flags the parity-phase file set; fails for every feature-branch commit; same as EXCEL-018/020 evidence).
- `e2e` job: FAILED — verifying against the documented pre-existing frozen-desktop baseline (chromium_headless_shell binary + desktop sheets sandbox); detail check follows in the next round if needed.

Stage Summary:

- Format-gate defect found by CI and corrected (this is exactly why the remote CI gate matters — the sandbox-only verification missed it).
- Workflow state: PR_OPEN (PR #20). NOT VERIFIED — architect owns that transition.

---

Task ID: EXCEL-022-ci-round2
Agent: Z.ai (Implementation Operator)
Task: CI evidence on the format-corrected head 0a559f3 (PR #20).

Work Log:

- Round-2 CI on 0a559f3 (run 33033335267), all four checks final:
  - `web` (the canonical office gate): SUCCESS — every step green. Typecheck (@contractor/web) clean; unit tests web 216/216 + web-host 78/78 + contractor-core 423 passed + 4 skipped; production build green; Playwright browser E2E **101 passed (9.7m)** on GitHub's infrastructure — including the 12 new ribbon-images tests and the architect's sort/formula semantic gate.
  - `test`: FAILURE — pre-existing lint debt. All 20 annotations live in frozen/pre-existing files (apps/docs, apps/sheets, apps/slides, apps/web/src/screens/BOQ.tsx). Cross-checked against the PR's changed-file list: ZERO overlap with EXCEL-022 files. Identical in character to the documented baseline that also failed on project-office and every prior office PR.
  - `e2e`: FAILURE — the frozen desktop Electron suite (InvalidSessionError/sandbox), untouched by this PR (frozen surfaces = empty diff); same documented baseline as EXCEL-018/020.
  - `foundation`: FAILURE — the branch-isolation guard rejects ANY branch touching packages/xlsx-gateway/ by design; verified identical failures on excel-020-protection, excel-021-tables, and web-office-editor before their merges. Not an office-gate signal.
- Round-1 → round-2 delta: the ONLY new failure round 1 exposed was the Prettier format violation (fixed in 0a559f3); round 2 confirms `test` now fails only on the pre-existing debt set.

Stage Summary:

- The canonical web gate is GREEN in CI on 0a559f3 with the full 101-test browser E2E. EXCEL-022 stands at PR_OPEN with complete local + CI evidence.
- Remaining pre-merge step per the mandated ordering: architect review of the actual diff (PR #20). Post-merge: production deploy → deployed E2E against genoffice.vercel.app → independent byte inspection.
- Workflow state: PR_OPEN. NOT VERIFIED — the architect owns that transition.

---

Task ID: EXCEL-022-architect-review-round1
Agent: Z.ai (Implementation Operator)
Task: Correct the two blockers from the architect's CHANGES REQUESTED review of PR #20 (posted on the PR 2026-08-27).

Work Log:

- Blocker 1 (browser purity boundary) — CORRECTED. sheet-images.ts read the private `_image`/`sheetTransform` internals of FOverGridImage through an `as unknown as` cast. Verified from the installed Univer 0.25.1 sources that the PUBLIC read path exists: `FOverGridImage.toBuilder().buildAsync()` returns the live image data — the same public surface the facade's own setPositionAsync/setSizeAsync build their commands on (facade.js:612 + buildAsync implementation). Rewrote readLivePlacement as an async function reading the returned public sheetTransform (from/to markers, pixel offsets) through an EXPLICIT narrow adapter (BuiltImageGeometry); width/height derive from the live grid between the markers exactly as before. Cascaded the async change through readLiveAnchor, collectImageVisualEdits, collectImageVisualAdditions (handleSave already async; anchors refresh at save time as before). Also removed ALL six EXCEL-022 `as unknown as` casts in ExcelEditor.tsx — the augmented FWorksheet type (sheets-drawing/facade import) structurally satisfies ImageWorksheetFacade, and `remove()` is a documented public method now declared on the adapter interface. The install builder param is typed against the REAL public ISheetImage type (Partial + Pick — exactly what the install path supplies; Univer's own documented example omits the derived transform fields). The E2E liveImages helper's own `_image` access was corrected to toBuilder().buildAsync() as well. Only the two PRE-EXISTING casts (EXCEL-021 table-theme muting, data-validation rule handle) remain — both predate EXCEL-022 and are outside its scope.
- Blocker 1 guard — STRENGTHENED (the architect said do not weaken the architecture guard): new architecture test "sheet-images.ts has NO private-Univer-internals access" rejects any `as unknown as` or `_image` access in the image module and requires the public toBuilder()/buildAsync() read adapter to be present.
- Blocker 2 (absolute-anchor fidelity) — CORRECTED, fail-closed branch chosen. The reader mapped every absoluteAnchor to zeroAnchor(), silently relocating imported absolute pictures to (0,0). xlsx-image-read.ts now omits absolute-anchored pictures from the browser model entirely (architect's fail-closed option): the picture never surfaces, stays untouched in the file, and a no-op save preserves its drawing XML and media byte-for-byte. The anchor still counts toward drawingIndex parity (later anchors keep stable locators). SheetImageInfo.anchorType narrowed to 'two-cell' | 'one-cell' — 'absolute' can no longer occur on the wire by design. zeroAnchor() dead code removed.
- Tests: gateway xlsx-image.test.ts — flipped the absolute-anchor case to assert omission, added "keeps drawingIndex parity across an omitted absolute anchor" and "preserves an absolute-anchored picture byte-for-byte through a no-op save". Browser E2E — new deterministic absoluteAnchor fixture + test 15 "an absolute-anchored image is omitted, never relocated (fail closed)": open response carries no image, nothing renders over the grid, and the drawing XML (including <xdr:pos x="47625" y="9525"/>) plus media bytes survive a cell-edit save byte-for-byte.
- Gates re-verified on the corrected tree: xlsx-gateway 604/604, contractor-core 423+4sk, web unit 217/217 (216 + the new architecture guard), web-host 78/78, typecheck clean in all four workspaces, ESLint clean on every changed file, Prettier format:check clean, production build green, frozen surfaces untouched (7 files changed, all in the two blocker areas), full browser E2E 102/102 — batch 1 (excel + insert) 50/50, ribbon-images 13/13 (move/resize exercise the new public geometry read end-to-end through save → reopen), batch 2 (ribbon + word) 39/39.

Stage Summary:

- Both architect blockers corrected at the root: the browser image module now reaches Univer ONLY through the public facade surface (guarded by a new architecture test), and absolute-anchored pictures fail closed — omitted from the browser model, never relocated, preserved untouched in the file.
- Workflow state: CHANGES REQUESTED → corrections pushed for re-review (PR #20). NOT VERIFIED — the architect owns that transition and explicitly re-reviews the corrected diff.

---

Task ID: EXCEL-022-architect-review-ci
Agent: Z.ai (Implementation Operator)
Task: CI evidence for the architect-review correction head 4b59a2c (PR #20).

Work Log:

- CI on 4b59a2c (the correction commit), all four checks final:
  - `web` (the canonical office gate): SUCCESS — typecheck clean, web unit 217/217, web-host 78/78, contractor-core 423+4sk, production build green, Playwright browser E2E **102 passed (10.4m)** on GitHub's infrastructure — including the 13 ribbon-images tests (move/resize exercise the new public toBuilder().buildAsync() geometry read end-to-end; test 15 proves the absolute-anchor fail-closed omission).
  - `test`: FAILURE — pre-existing lint debt; all 9 annotated files are frozen/pre-existing, ZERO overlap with the EXCEL-022 file set, and ZERO format annotations (the round-1 Prettier defect stays corrected).
  - `e2e`: FAILURE — the frozen desktop Electron suite baseline (untouched; frozen surfaces = empty diff).
  - `foundation`: FAILURE — the branch-isolation guard that rejects any branch touching packages/xlsx-gateway/ by design (identical on every prior office PR).

Stage Summary:

- Both architect blockers corrected, all local gates green, and the web CI gate green with the full 102-test browser E2E on the correction head 4b59a2c.
- Workflow state: CHANGES REQUESTED → corrections pushed, awaiting architect re-review of the actual diff. NOT VERIFIED — the architect owns that transition.

---

Task ID: EXCEL-023-audit
Agent: Z.ai (Implementation Operator)
Task: Phase A forensic audit for EXCEL-023 (Charts) — mandatory before implementation.

Work Log:

Environment: sandbox was reset; repo re-cloned from origin/main @ cb1fff6 (EXCEL-022 merge commit, PR #20). Branch excel-023-charts created from origin/main. EXCEL-022 confirmed MERGED (visualAdditions/visualEdits families + WorksheetState.images live on main).

Q1 — Canonical chart model: packages/xlsx-gateway/src/domain/chart-visual.ts (697 lines) defines ChartVisualState (chartTypes[], barDirection, title, series: ChartSeriesVisualState[] [name/categories/values/valuesRef/categoriesRef/color/pointColors/explosionPct/pointExplosions/numberFormat/categoryFormat], legend, axisTitles, dataLabels+position+format, grouping, gridlines, valueAxis, categoryAxisFormat, gapWidthPct, holeSizePct, xAxis/yAxis/secondaryYAxis: ChartAxisInfoState, scatterStyle, titleStyle) + ChartStateEdit + applyChartStateEdit (immutable overlay merge — desktop preview parity) + CHART_EDIT_TYPES + transposeChartSeries + axis scale helpers (valueAxisScale, scatterAxisBounds — Excel-fidelity tick math). The desktop re-exports this module verbatim (apps/sheets/src/domain/chart-visual.ts = `export * from '@genoffice/xlsx-gateway/src/domain/chart-visual.js'`) — it IS the shared canonical model.

Q2 — WorkbookChartEdit (packages/xlsx-gateway/src/types.ts:123-143): chartPath locator + title?, chartType? ('column'|'bar'|'line'|'area'|'pie'|'doughnut'), seriesColors?, legend?, dataLabels?/dataLabelPosition?/dataLabelFormat?, axisTitles?, pointColors?, grouping?, gridlines?, valueAxis? {min,max|null}, gapWidthPct?, holeSizePct?, explosionPct?, pointExplosions?, seriesSet? (full replacement), series? (index-keyed). Desktop Zod mirror: apps/sheets/src/shared/desktop-api.ts:943-999 (bounded: title<=255, gapWidth 0-500, holeSize 10-90, explosion 0-400, series color keys 0-999).

Q3 — Gateway capability matrix:

- READ charts: NO. readBasicWorkbook (xlsx-gateway.ts:425-568) parses cells/styles/merges/rowHeights/colWidths/freeze/filter/dv/notes/tables/images/protection — zero chart handling. No parseChart/readChart function exists anywhere in the gateway (verified by search). The DESKTOP reads file charts through its RUST SIDECAR (session.metadata.visuals → buildWorkbookFile, apps/sheets/src/main/sheets-save-adapter.ts:416) — outside this repo's TS read path, unavailable to web.
- CREATE charts: YES. SheetVisualAddition.chart?: ChartAdd (xlsx-gateway.ts:792; xlsx-drawing-add.ts:33-63). buildChartXml (xlsx-drawing-add.ts:409) supports 9 types: column/bar/line/area/pie/scatter/radar/doughnut/combo (combo = clustered columns + last series line on secondary axis). Writes chart part + drawing part + drawing rels + worksheet drawing rel + [Content_Types] overrides; returns AddedVisualLocator (worksheetPath/drawingPath/drawingIndex — xlsx-gateway.ts:149-153; NO chartPath yet).
- EDIT charts: YES. applyChartEdit (xlsx-chart.ts:33, 997 lines): title, type conversion within CONVERTIBLE_PLOTS [barChart, lineChart, areaChart, pieChart, doughnutChart] (refuses scatter), seriesSet/series rewrite with refs+caches, series/point colors, legend, plot-level data labels (+position/format), axis titles, grouping (stacked/percentStacked/clustered), gridlines, value-axis bounds, gapWidth, holeSize, explosion/pointExplosions. Fail-closed: ChartEditError outside the envelope.
- DELETE charts: YES. visualEdits remove on a graphicFrame anchor cascades the drawing relationship, the chart part, the chart's own rels, and the [Content_Types] override (xlsx-drawing-edit.ts:153-197); chart-owned colors/style parts deliberately stay (harmless orphans). Duplicate-locator edits rejected.
- SHIFT anchors on row/col ops: chart XML references shift (shiftChartReferences, xlsx-gateway.ts:1053-1054, runs for every xl/charts/*.xml part); drawing anchors shift via shiftAnchoredSheetParts (same save pass). Chart series refs + cross-sheet formulas + defined names all shift.

Q4 — Canonically supported chart types: EDIT/CONVERT: column, bar, line, area, pie, doughnut (6). CREATE adds: scatter, radar, combo (9 total). Scatter is NOT convertible (xlsx-chart.test.ts:123 'refuses to convert a scatter chart').

Q5 — OOXML chart parts represented: xl/charts/chartN.xml (edit+create+delete), xl/charts/_rels/chartN.xml.rels (deleted with the chart; chartColorStyle targets preserved while alive), [Content_Types].xml chart override (managed on create+delete), drawing anchors twoCellAnchor/oneCellAnchor/absoluteAnchor all recognized (ANCHOR_PATTERN xlsx-drawing-edit.ts:11).

Q6 — Wiring chain: worksheet.xml <drawing r:id> → worksheet rels → xl/drawings/drawingN.xml → anchor (document order = drawingIndex) → graphicFrame/c:chart r:id → drawing rels → xl/charts/chartN.xml. xlsx-image-read.ts already walks worksheet→drawing→rels for pictures; chart reader reuses the same chain shape.

Q7 — Desktop renderer: charts render as the app's OWN SVG React overlay, NOT Univer charts. installWorkbookVisuals (WorkbookVisuals.tsx:115-210): univerAPI.registerComponent + worksheet.addFloatDomToRange with pixel-exact twoCellAnchor layout (markerSpan over live column widths/row heights; EMU_PER_PIXEL). ChartVisual (1160-1502): SVG renderers BarChart/LineChart/AreaChart/RadarChart/ScatterChart/PieChart + VerticalAxis + data labels + SeriesLegend + TruncationNote + combo handling; pending file-chart edits overlay via withChartEdit (applyChartStateEdit); openpyxl numCache-less files hydrate series from live cells via readVector (1188-1251). EditableShapeVisual (584+): selection, move (ghost on body), 8-corner resize via walkMarker; Delete key + context menu remove. ChartPanels.tsx: ChartFormatPane (element-aware fill colors, title, legend, labels, axis titles, bounds, gap width, hole size, explosion) + SelectDataDialog (seriesSet rebuild + switch row/column). edit-journal.ts recordChartEdit (519-560): per-chartPath merge — seriesSet invalidates earlier per-index series/color edits; deleting a visual drops its pending chartEdits (visual-edit-sync.ts:377-378).

Q8 — Univer 0.25.1 chart surface: NONE. Installed @univerjs set (core + 9 sheets presets incl. drawing) contains no chart plugin (verified: no chart-named package; preset-sheets-drawing exposes image/float-DOM only). Public APIs the desktop uses for charts — univerAPI.registerComponent and FWorksheet.addFloatDomToRange (node_modules/@univerjs/sheets-drawing-ui/lib/types/facade/f-worksheet.d.ts:305, documented with examples) — are available to the web through the ALREADY-INSTALLED preset-sheets-drawing. No new dependency needed.

Q9 — Web chart registration today: NONE. Ribbon.tsx:526-535 carries a disabled Chart stub whose title documents the missing families. No chartEdits in BrowserWorkbookSavePlan (extensibility seam comment at office-routes.ts:166-167). office-routes visualAdditions is IMAGE-ONLY and explicitly REJECTS chart/shape additions (office-routes.ts:2536-2541) with the comment 'EXCEL-023 will widen the family when charts land'.

Q10 — Existing visuals models: (a) WorksheetState.visuals (workbook.types.ts:52-53) = demo/AI chart replay state — NEVER populated by readBasicWorkbook, demo-only, deliberately NOT repurposed (EXCEL-022 architecture decision). (b) desktop state.file.visuals = Rust-sidecar read (canonical for desktop, unreachable from web). (c) ChartVisualState = canonical shared domain (Q1). Web currently has NO visuals at all.

Q11 — Wire family decision: chartEdits needs its own semantic wire family (keyed by chartPath) — EXCEL-022's visualEdits is geometric-only (drawingPath+drawingIndex locator, anchor rewrite/splice) and ALREADY serves chart move/resize/delete correctly (Q3). The two families compose: chartEdits patches chart XML; visualEdits patches the anchor. visualAdditions must be WIDENED from image-only to accept chart (ChartAdd). The route's own comment reserves exactly this seam. AddedVisualLocator must gain chartPath so a saved session chart can be targeted by later chartEdits.

Q12 — WorksheetState additions: dedicated `charts?: readonly SheetChartInfo[]` field (mirrors EXCEL-022's images pattern; does NOT touch demo visuals). SheetChartInfo = { drawingPath, drawingIndex (anchor locator, ALL anchors counted for parity), chartPath, anchorType 'two-cell'|'one-cell', anchor: DrawingAnchor, chart: ChartVisualState }.

Q13 — Fail-closed classification:

- absoluteAnchor charts: omitted from browser model, bytes preserved (EXCEL-022 precedent; visualEdits already refuses absolute moves).
- 3-D plots (bar3DChart, pie3DChart, line3DChart, surface, bubble, stock, ofPie etc.): per-chart omission — not in the canonical model.
- chartEx / pivotChart extension parts (c15:, cx:): per-chart omission.
- multi-plot charts whose plots aren't a supported combination: omitted unless single-plot family or bar+line combo (the canonical combo shape ChartAdd itself writes).
- scatter/radar/combo: renderable + movable/resizable/deletable, but NOT type-convertible and NOT seriesSet-replaceable per-some-plots — applyChartEdit fails closed (scatter refusal already tested).
- rich formatting beyond the model (gradient fills, custom label positions on unsupported plots, trendlines, error bars): read-ignored (stay in file untouched), never written.
- oversized chart series sets: bounded on the wire (desktop Zod parity), rejected at route validation.

Q14 — Reusable tests/fixtures: gateway xlsx-chart.test.ts (1039 lines — full applyChartEdit envelope incl. fail-closed cases), xlsx-drawing-add.options.test.ts (buildChartXml options), xlsx-drawing-edit.test.ts (anchor edits + chart cascade), fixture-builder.ts (chart-bearing archives: drawing rel + chart part + colors1.xml + content types, lines 144-360); web architecture.test.ts guard patterns (EXCEL-020/021/022 blocks); ribbon-images.spec.ts E2E structure (15 scenarios — template for charts); desktop-api.ts Zod schemas (validation reference).

Stage Summary:

- Audit verdict: the canonical engine ALREADY owns chart create/edit/delete/move/resize/reference-shift; the ONLY missing capability is the chart READER (gateway→typed state→wire) plus the entire web browser surface (render + interact + journal). No second storage architecture is needed — charts ride the EXCEL-022 families exactly as the work item presumed.
- Rendering architecture decision (desktop parity): web charts render as the web app's own SVG React overlay floated via the PUBLIC addFloatDomToRange/registerComponent facades (the desktop's exact pattern; Univer 0.25.1 has no chart plugin, so this is the only parity path). Desktop WorkbookVisuals renderer is REFERENCE ONLY (frozen surface) — the web gets its own renderer module adapted from the desktop's rendering semantics.
- Workflow state: EXCEL-023 ASSIGNED → IMPLEMENTING. Not VERIFIED — architect owns that transition.

---
Task ID: EXCEL-023-implementation
Agent: Z.ai (Implementation Operator)
Task: Implement EXCEL-023 (Charts) on branch excel-023-charts per the audited architecture — gateway chart reader, chartEdits wire family, web chart visual surface, and the 15-scenario browser E2E suite.

Work Log:

- Branch: excel-023-charts from origin/main @ cb1fff6 (EXCEL-022 merged). Note: this session began with the four EXCEL-023 commits sitting on local main after a sandbox reset — restructured non-destructively (branch created at the head, local main reset to origin/main; zero commits lost).
- 7655903 gateway chart reader: xlsx-chart-read.ts (668 lines) walks the worksheet→drawing→rels→chart chain (the EXCEL-022 image chain shape) into SheetChartInfo {drawingPath, drawingIndex (ALL anchors counted for parity), chartPath, anchorType, anchor, chart: ChartVisualState}; unsupported structures (3-D plots, chartEx, unsupported multi-plot combos, absolute anchors) are omitted per-chart, never relocated. AddedVisualLocator gained chartPath. 789-line gateway test file (17 new tests: read fidelity, omission cases, byte preservation, locator parity).
- 9dcd652 wire family: chartEdits (semantic, keyed by canonical xl/charts path; strict validation — bounded types/series/style options, fail-closed on unknown fields, 200-edit cap) + visualAdditions widened from image-only to exactly-one-payload (image or chart). 382-line route validation test file (18 tests).
- 6f4f30b web visual surface: sheet-charts.tsx (2002 lines — SVG renderers for all 8 canonical families with Excel-like axis auto-scale, anchor↔pixel math over the live grid, interactive frames with select/move/8-handle resize/delete, the shared ChartEditingStore with desktop edit-journal merge parity); ChartPanel.tsx (Chart Design pane: create mode over the parsed selection, edit mode title/convert/legend/labels/series colors + typed source ranges); ExcelEditor seeding/reinstall/save-collection/post-save merge; Ribbon Insert→Chart wired; 6 new architecture guards (chart OOXML absence in apps/web, canonical families, no private internals, pure-domain-only value imports, ribbon wiring, read/save path wiring).
- c1a0ab6 E2E completion (this session's continuation work): the 15-scenario ribbon-charts.spec.ts + chart fixtures; two E2E-driven product fixes (one-cell anchors size from a:ext — two-cell from the live marker span; DOM-safe float/component ids for locator keys); three test-round defects found and corrected — (a) contractor-core test helper return type (TS2698 spread), (b) the chart journal now feeds the editor's dirty gate (recordChartEdit marks the workbook unsaved exactly like a cell edit — Save must never stay disabled while the journal holds work), (c) selecting a chart frame opens the Chart Design pane in edit mode (desktop ChartPanels parity) and the panel's selection.isSession is computed; the 3-D fixture's omitted anchor moved to index 0 so the survivor keeps drawingIndex 1 (anchor-count parity proof); ribbon-insert.spec.ts updated — Chart is ENABLED (was the pre-EXCEL-023 disabled stub); no-op byte-preservation proofs ride a cell-edit save (EXCEL-022 ribbon-images precedent — Save is disabled while the workbook is clean).
- Prettier format pass over the EXCEL-023 file set (the EXCEL-022 round-1 lesson applied: format:check vs origin/main run BEFORE push).
- Gates on c1a0ab6: typecheck clean in all four workspaces; unit — xlsx-gateway 621/621, contractor-core 441 passed + 4 skipped, web 223/223 (incl. 6 new architecture guards), web-host 78/78; ESLint clean on every changed file; format:check clean; production build green; full browser E2E 111/111 (charts 9 blocks covering the 15 scenarios + 2 extra fail-closed proofs; regression batches excel-core/nested 30, images/insert/data-validation/data 25, insert-fixed/filter/persistence/protection 14, dedupe/notes/table/view 18, word 16).

Stage Summary:

- EXCEL-023 implementation complete on branch excel-023-charts (5 commits: audit docs, gateway reader, wire family, web surface, E2E completion). The browser never touches chart OOXML — all XML work stays in @genoffice/xlsx-gateway (guarded by architecture tests); charts ride the EXCEL-022 visual transport families (chartEdits + visualEdits + widened visualAdditions) exactly as the work item's architecture requires.
- Workflow state: IMPLEMENTING → PR_OPEN (branch pushed, PR to follow). NOT VERIFIED — the architect owns that transition.
