# GenOffice Excel Parity — Work Items

## Workflow rule

Each work item is independently tracked and follows:

```text
DRAFT → READY → ASSIGNED → IMPLEMENTING → PR_OPEN → VERIFYING → ARCHITECT_REVIEW → APPROVED → MERGED → VERIFIED
```

A work item has one active implementation PR at a time. A failed verification returns the item to IMPLEMENTING. Architecture changes require an Architecture Change Request.

## Completed foundation/parity work

### EXCEL-INC-001 — Real browser workbook pipeline

Objective: Connect the web Excel editor to canonical XLSX open/save through the office API.
Dependencies: EXCEL-002.
Evidence: real HTTP route tests, XLSX round-trip tests, architecture purity checks.
Status: VERIFIED.

### EXCEL-INC-002 — Browser editor fidelity foundation

Objective: Make Excel change-driven and Word/block serialization safe; establish runtime validation and document handles.
Status: VERIFIED.

### EXCEL-INC-003 — Browser editor fidelity

Objective: Preserve Excel workbook structure and establish Word structural fidelity.
Status: VERIFIED.

### EXCEL-INC-004 — Word inline/run fidelity

Objective: Preserve run-level formatting and stable DOCX identity.
Status: VERIFIED.

### EXCEL-INC-005 — Dirty-state/API hardening

Objective: Fingerprint unchanged blocks, require response guards, strengthen source identity.
Status: VERIFIED.

## Phase 3 Excel work

### EXCEL-010 — Excel formatting

Objective: Render and persist common cell formatting through canonical style deltas.
Dependencies: EXCEL-001, EXCEL-002.
Status: VERIFIED.
Acceptance highlights: styles, fills, font variants, alignments, wrapping, number-format read/write subset, merges, dimensions.

### EXCEL-011 — Row/column structural operations

Objective: Insert/delete rows and columns, shift dirty coordinates, preserve styles/formulas.
Dependencies: EXCEL-002, EXCEL-005.
Status: VERIFIED.

### EXCEL-012 — Formula-bar fidelity

Objective: Protect formula-vs-value semantics against Univer recalculation echoes and explicit clears.
Dependencies: EXCEL-003, EXCEL-005.
Status: VERIFIED.

### EXCEL-013 — Word image fidelity

Objective: Preserve healthy image metadata, alt text, transformations, relationships, and media bytes.
Dependencies: unrelated to Excel parity but part of the office pipeline.
Status: VERIFIED.

## Phase 4 workspace and command parity

### EXCEL-014 — Sheets workspace shell

Objective: Recreate the desktop Excel workspace shell in web: title bar, ribbon, name box, formula bar, sheet tabs, status bar, zoom, theme.
Dependencies: EXCEL-001.
Status: VERIFIED.

### EXCEL-015 — Core ribbon mutation persistence

Objective: Make Home/View/Data/Formulas command subsets real rather than decorative, with canonical save/reopen paths.
Dependencies: EXCEL-014.
Status: VERIFIED.

### EXCEL-016 — Data Filter

Objective: Import, render, edit, clear, persist, and reopen AutoFilter state.
Dependencies: EXCEL-015, EXCEL-009.
Status: VERIFIED.

### EXCEL-017 — Data Validation

Objective: Import, render, edit, clear, persist, and reopen supported validation rules.
Dependencies: EXCEL-016, EXCEL-009.
Status: VERIFIED.

### EXCEL-019 — Comments / Notes

Objective: Import existing cell notes/comments, display them, create/edit/delete supported notes, and persist through `noteStates`.
Dependencies: EXCEL-009, EXCEL-015.
Status: VERIFIED.
Implementation commit: 68cbb9d1a36c54b1731b2f43460da547fde1e437 (bundle: 8c05c18; pushed to origin/web-office-editor through 26f5f54).
Evidence:

- Gateway read/write: `parseCommentsPart` + `resolveCommentsPath` in `packages/xlsx-gateway/src/gateway/xlsx-notes.ts`; `WorksheetState.notes` in `packages/xlsx-gateway/src/domain/workbook.types.ts`; per-sheet fail-closed (`NoteReadError`) on unreadable refs, out-of-sheet refs, missing text, oversized sets.
- Canonical save: `noteStates` at `applyCellEditsToXlsx` argument 13 → `applySheetNotes` (comments part + VML note shapes + rels + content types + `<legacyDrawing>`); empty notes list removes the part.
- Wire: `BrowserWorkbookSavePlan.noteStates` with strict `expectSheetNoteState`/`expectSheetNote` validation in `packages/contractor-core/src/api/office-routes.ts`.
- Browser: journal-suppressed import via `createOrUpdateNote` (no Undo pollution — proven), `sheet.mutation.update-note`/`remove-note` dirty marks, live-model `collectNoteStates` snapshot at save, Review → New Comment via the real `sheet.operation.add-note-popup`.
- Tests: gateway `packages/xlsx-gateway/tests/xlsx-notes.test.ts` (19 tests: all required verification points — import, author/text/multi-line/special-XML preservation, multiple notes, write→reopen, edit, delete-one, delete-all, no-op byte preservation, malformed rejection, oversized rejection, out-of-sheet legacy form, readBasicWorkbook integration, per-sheet fail-closed); wire `packages/contractor-core/tests/unit/office-notes-routes.test.ts` (13 tests); browser E2E `apps/web/tests/e2e/ribbon-review-notes.spec.ts` (5 tests / 10 scenarios through real HTTP: existing-notes render, no-undo-on-load, create, edit, delete-one-of-two, save/reopen with XML + typed wire inspection, untouched-note survival, no-op byte preservation).
- Production: deployed to genoffice.vercel.app; 25-assertion live pipeline verification (read/save/reopen/clear/no-op/validation) all green; CI `web` job green on the final commit.

## Remaining implementation roadmap

### EXCEL-018 — Remove Duplicates / Data Tools

Objective: Implement a real Remove Duplicates command using canonical workbook mutations, or explicitly prove that a safe canonical engine path does not exist and record the feature as deferred.
Dependencies: EXCEL-015.
Required verification: duplicate-row fixtures, header handling, multiple selected columns, formulas/styles preservation, save/reopen.
Status: ARCHITECT_REVIEW / CHANGES REQUIRED → CORRECTION SUBMITTED (2026-08-26). The architect rejected the first implementation (value-level `setValues` rewrite destroys formulas on compacted rows — "moved rows become computed values" is explicitly NOT accepted as formula preservation). The correction replaces the entire write path with the canonical STRUCTURAL `remove-rows` family: `dedupeRowIndices()` returns the duplicate row offsets, the runtime issues `ws.deleteRows(startRow+offset, 1)` per duplicate in DESCENDING order, each firing `sheet.mutation.remove-rows` journaled by the existing `STRUCTURAL_MUTATION_IDS` subscription as a `{kind:'remove-rows', index, count:1}` structural op in the save plan. The gateway's `applyStructuralOps` applies each op atomically: `transformSheetRows` renumbers `<row>` r= and inner `<c>` r= (cell contents travel UNTOUCHED inside their `<c>` elements), `transformFormulas` rewrites `<f>` bodies via `shiftFormulaText` (relative + absolute + mixed references all track the moved cells — `$` markers preserved by `shiftReferenceToken`'s colDollar/rowDollar capture groups), and `transformRangedFeatures` shifts merges/autoFilter/hyperlink sqref/dataValidation sqref/conditionalFormatting sqref. This is the EXACT canonical path `excel-structural.spec.ts` already proves for Insert/Delete Rows — no new mutation family, no wire change, no gateway change, no frozen-surface modification. Desktop parity is documented as INSUFFICIENT (the desktop's own dedupe is architecturally lossy); the frozen web requirement at `spec/excel/verification-matrix.md` explicitly requires formulas/styles preservation + save/reopen, which is only satisfiable through structural row deletion. Independent architect verification PENDING (Z.ai is the implementer and does NOT own the VERIFIED decision).
Evidence (correction, 2026-08-26):

- Pure dedupe algorithm: `apps/web/src/office/dedupe.ts` — TWO exports: `dedupeRows` (legacy value-level, kept for compat) and `dedupeRowIndices` (the canonical entry point returning `{keptIndices, duplicateIndices, removed}`). ZERO imports — verified by the architecture test `dedupe.ts is a PURE module`. Equality semantics unchanged (JSON.stringify key with text lowercased, type-strict, header preserved).
- Structural wiring: `apps/web/src/screens/excel/useExcelRuntime.ts` — `removeDuplicates(hasHeader)` reads `range.getValues()` (computed values, comparison only), calls `dedupeRowIndices`, then issues `ws.deleteRows(startRow+offset, 1)` per duplicate in DESCENDING offset order. NO `setValues` call anywhere in the body (enforced by the architecture test). Fail-closed on `height < 2` and on `removed === 0` (unchanged).
- Canonical-path guards: `apps/web/tests/architecture.test.ts` — the EXCEL-018 block updated to `canonical structural remove-rows path` (24 tests total in the file). New guard `removeDuplicates does NOT rewrite moved rows via setValues` locates the useCallback body by bracket-walking and asserts it calls `.deleteRows(` and does NOT call `.setValues(` and does NOT contain a "padded" variable.
- Architect's mandatory regression case (test 1): B7 `=B6` (computed 30) on the survivor row. After dedupe (2 remove-rows ops), B7's content compacts to B5, the formula is PRESERVED with its reference REWRITTEN B6→B4 (Cherry/30 moved from row 6 to row 4). Asserted in Univer's live model (`b5.f === '=B4'`, `b5.v === 30`), on the wire (save plan carries exactly 2 `{kind:'remove-rows', index:2/4, count:1}` structural ops), in the saved XML (`<c r="B5"...><f>B4</f>`), and in the reopened snapshot (`cells.B5.formula === '=B4'`). NO CellEdit at B5 carries value=30 without a formula (the architect's explicit failure mode).
- Architect's second regression case (test 2, NEW): a "DedupeMixed" fixture with a survivor row (Banana) carrying FOUR formulas — C4=`=$D$6` (absolute), D4=`=A6` (relative), E4=`=$A6` (mixed col-$), F4=`=A$6`(mixed row-$) — referencing Cherry/Anchor at row 6 (outside the dedupe selection A1:B7). After dedupe (3 remove-rows ops at rows 3/5/7), ALL FOUR references are rewritten to track the moved target (row 6 → row 4):`$D$6→$D$4`, `A6→A4`, `$A6→$A4`, `A$6→A$4`. `$` markers preserved (absolute stays absolute). Asserted in Univer's live model, the saved XML, AND the reopened snapshot — distinguishing "formula text survived and was rewritten correctly" from "computed result happens to be the same".
- Unit tests: `apps/web/tests/dedupe.test.ts` — 27 tests (14 legacy `dedupeRows` + 13 new `dedupeRowIndices` covering the architect's mandatory fixture, descending-deletion stability, and the dedupeRows/dedupeRowIndices parity).
- EXCEL-018 E2E (`apps/web/tests/e2e/ribbon-remove-duplicates.spec.ts`, 4 tests through real HTTP): (1) mandatory regression — formula preserved + reference rewritten + rows deleted (no padding nulls) + structuralOps on the wire + saved XML `<f>B4</f>` + reopen formula `=B4`; (2) second regression — all four reference types rewritten; (3) no-op fail-closed; (4) `<2-row` fail-closed. All 4 pass.
- Mandatory regression suite (all green): excel-structural (2), excel-formula (7), excel-format (3), excel-browser (1), excel-shell (15), ribbon-data (4 — INCLUDING the architect's sort/formula semantic gate), ribbon-data-validation (7), ribbon-filter (5), ribbon-review-notes (5), ribbon-home-persistence (3), ribbon-insert (1), word-browser (1). Total 47/47 local E2E green.
- Web unit suite: 197/197 pass. Typecheck: exit 0. Prettier: clean.
- Frozen surfaces untouched: `git diff --stat -- apps/sheets apps/docs apps/shell packages/platform-electron packages/renderer-bridge` is EMPTY.
- Workflow state: READY → IMPLEMENTING → PR_OPEN → VERIFYING → ARCHITECT_REVIEW → CHANGES REQUIRED → CORRECTION SUBMITTED (current). NOT VERIFIED — pending independent architect verification of the structural-path correction.

### EXCEL-020 — Sheet Protection / Workbook Protection

Objective: Add Review protection controls using canonical `sheetProtections` and workbook-protection families.
Dependencies: EXCEL-019 (VERIFIED — see the completed Phase 4 section above).
Required verification: protect/unprotect, password semantics if supported, protected-cell behavior, save/reopen.
Status: READY, but sequenced AFTER EXCEL-018 per architect direction (2026-08-25 review): `EXCEL-018 Remove Duplicates → EXCEL-020 Protection`.

### EXCEL-021 — Tables

Objective: Import existing Excel tables, render structured-table semantics, create/delete tables, and persist table additions/deletions and supported table options.
Dependencies: EXCEL-009, EXCEL-015.
Required verification: table metadata, filters, styles, row/column edits, no-op preservation, create/save/reopen/delete.

### EXCEL-022 — Images / Drawings

Objective: Import existing worksheet images and supported drawing metadata, render them, create/remove supported images, and persist via `visualAdditions`.
Dependencies: EXCEL-009, EXCEL-021.
Required verification: media bytes, relationships, anchor position, dimensions, deletion, no-op preservation.

### EXCEL-023 — Charts

Objective: Import supported charts, render them through the shared visual layer, create/edit/delete supported chart types, and persist `chartEdits`.
Dependencies: EXCEL-022.
Required verification: chart type, source ranges, dimensions, series, style, reopen.

### EXCEL-024 — Conditional Formatting

Objective: Import supported CF rules, render them, edit/create/delete rules, and persist `cfStates`.
Dependencies: EXCEL-021.
Required verification: expression rules, cell-value rules, color scales/icon sets/data bars where canonical support exists, fail-closed for unsupported extensions.

### EXCEL-025 — Named Ranges / Name Manager

Objective: Import workbook and sheet-scoped defined names, provide Name Manager UI, support create/edit/delete of supported names, and persist `definedNamesState`.
Dependencies: EXCEL-015.
Required verification: workbook scope, sheet scope, print names if supported, formulas consuming names, reopen.

### EXCEL-026 — Freeze panes and View persistence expansion

Objective: Complete page setup and view-state persistence beyond frozen rows/columns, including supported gridline/formula-view semantics.
Dependencies: EXCEL-015.
Status: freeze foundation already implemented; remaining view/page-setup work pending.

### EXCEL-027 — Advanced cell formatting

Objective: Complete borders, number formats, indentation, text rotation, named cell styles, paste special, and format painter where canonical engine support exists.
Dependencies: EXCEL-010.
Required verification: style XML and rendered comparison for representative variants.

### EXCEL-028 — Autosave and crash recovery

Objective: Provide browser-safe autosave and recovery semantics equivalent to the desktop's user-visible behavior.
Dependencies: EXCEL-011, EXCEL-015.
Required verification: timer/blur save, crash simulation, recovery prompt, recovery discard, concurrency/identity safety.

### EXCEL-029 — Undo/Redo and journal parity

Objective: Move from partial browser dirty maps toward shared semantic journaling and correct undo/redo grouping/suppression.
Dependencies: EXCEL-019 (VERIFIED), EXCEL-021, EXCEL-024.
Required verification: mutation-family coverage, load suppression, batch grouping, redo, save/reopen after undo.

### EXCEL-030 — Theme and locale parity

Objective: Match desktop light/dark/system theme and supported localized Excel UI.
Dependencies: EXCEL-014.
Required verification: live theme switching, persistence, locale packs, key ribbon/dialog labels.

### EXCEL-031 — Ribbon parity completion

Objective: Replace disabled placeholders across Insert, Page Layout, Formulas, Data, Review, View with real commands whenever canonical paths are available.
Dependencies: EXCEL-018 through EXCEL-027.
Required verification: command-by-command browser and desktop interaction gates.

### EXCEL-032 — Dialog parity

Objective: Port the desktop Project/Excel dialog surfaces needed for Excel parity: Format Cells, Go To, Insert Function, Name Manager, Sort/Filter dialogs, Conditional Formatting, Data Validation, Protection, Page Setup, etc.
Dependencies: EXCEL-027, EXCEL-031.
Required verification: dialog input/output parity and save/reopen effects.

### EXCEL-033 — Shared sheets editor core extraction

Objective: Extract transport-neutral Univer bootstrap, fidelity patches, journal, ribbon primitives, dialogs, and import/collect transforms into a shared package only after behavior is stable.
Dependencies: EXCEL-029, EXCEL-031, EXCEL-032.
Required verification: desktop and web behavior remains unchanged before/after extraction.

### EXCEL-034 — AI panel parity

Objective: Provide the desktop Sheets AI panel in web through a transport-neutral chat/agent interface without making browser UI authoritative for workbook semantics.
Dependencies: EXCEL-033.
Required verification: streaming, tool invocation, safe workbook mutations, CAS/drift guard, undo.

### EXCEL-035 — Large-workbook and performance parity

Objective: Introduce session/range streaming and memory-safe loading when workbook size exceeds web snapshot limits.
Dependencies: EXCEL-010, EXCEL-021, EXCEL-024, EXCEL-033.
Required verification: large fixtures, scrolling latency, memory ceilings, correctness under incremental loading.

### EXCEL-036 — Final parity acceptance

Objective: Demonstrate desktop/web renderer parity against the frozen requirements and verification matrix.
Dependencies: EXCEL-018 through EXCEL-035.
Required verification: full local suite, desktop E2E, web E2E, production E2E, golden-file suite, architecture checks, no-op byte preservation.

## Rule for advancement

No work item may be skipped because a later feature is visually attractive. A dependency must be VERIFIED before a dependent item becomes READY.
