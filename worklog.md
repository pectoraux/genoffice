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
