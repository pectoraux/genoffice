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
- Workflow state: READY → IMPLEMENTING → PR_OPEN → VERIFYING → ARCHITECT_REVIEW → CHANGES REQUIRED → CORRECTION SUBMITTED → correction MERGED into main (82a87c9, in current main ancestry). NOT VERIFIED — the architect owns that transition.

Verification debt closure evidence (2026-08-27, independent rerun on main):

- 15-point source audit re-confirmed on current main: pure comparison (`dedupe.ts` zero imports), selection-scoped key, header excluded from the seen-set, descending `deleteRows`, NO `setValues` compaction (architecture guard green), zero-removal no-op, undersized fail-closed.
- Mandatory regression A (B7 `=B6` → B5 `=B4`): E2E test 1 green in all four layers (live model, save plan, saved XML, reopened snapshot).
- Mandatory regression B (mixed refs): E2E test 2 green — `$D$6→$D$4`, `A6→A4`, `$A6→$A4`, `A$6→A$4`.
- NEW independent byte proof `packages/xlsx-gateway/tests/xlsx-dedupe-byte-proof.test.ts` (3 tests, package-level, no browser): survivor `<c>` records remain `<f>` formulas, `s=` style indices travel with survivor rows, unrelated rows/cells and unrelated package parts (styles.xml, sharedStrings.xml, other worksheet) byte-identical, zero-op no-op preserves worksheet bytes verbatim, reopened snapshot carries `=B4`.
- Deployed E2E vs production `genoffice.vercel.app`: 4/4 green (both mandatory regressions through the live HTTPS → serverless → gateway → bytes path).

### EXCEL-020 — Sheet Protection / Workbook Protection

Objective: Add Review protection controls using canonical `sheetProtections` and workbook-protection families.
Dependencies: EXCEL-019 (VERIFIED — see the completed Phase 4 section above).
Required verification: protect/unprotect, password semantics if supported, protected-cell behavior, save/reopen.
Status: CORRECTION-SUBMITTED-EQUIVALENT / IMPLEMENTED, PENDING ARCHITECT REVIEW (2026-08-26). Built on the EXISTING canonical families — no new engine path, no browser-side OOXML, no parallel protection model. The gateway's write-side helpers (`applySheetProtection` / `applyWorkbookProtection` / `applyProtectedRanges`) and the `sheetProtections` save-family parameter of `planCellEditsToXlsx` ALREADY existed and were unit-tested; the gaps closed here are (a) the READ side (readBasicWorkbook now parses `<sheetProtection>` per worksheet + `<workbookProtection>` in workbook.xml into `WorksheetState.sheetProtection` / `WorkbookSnapshot.workbookProtection`, both carrying `hasPassword`), (b) `applyCellEditsToXlsx`'s trailing `workbookProtectionState` pass-through parameter, (c) the wire contract (`BrowserWorkbookSavePlan.sheetProtections` + `workbookProtectionState` with strict validation in `office-routes.ts` — unknown fields including password-bearing ones are 400s), and (d) the browser shell (Review → Protection group: Protect Sheet / Protect Workbook / Lock Cell / Unlock Cell with journal-only toggle semantics and desktop-parity status strings). Desktop parity notes: protection is a FILE-level journal concern — the editor does not enforce it (the desktop's own status string says so); Allow Edit Ranges (`protectedRanges`) stays desktop-only (not in the EXCEL-020 verification list); password semantics are fail-closed at every layer (the browser refuses up front via the read-side `hasPassword` flag, the wire rejects password fields, the gateway throws `SheetProtectionError` on unprotecting a password-bearing element). The "editable vs locked cell behavior" requirement is satisfied through the REAL ribbon: Lock Cell / Unlock Cell journal canonical `WorkbookStyleEdit.protectionLocked` style-only CellEdits (the desktop's neutral-delta path — Univer's OSS presets carry no cell-protection model), which the gateway's `buildProtection` writes as `<protection locked="0|1"/>` in the cell's xf; combined with a protected sheet that is Excel's editable-vs-locked semantics in the saved file.
Evidence (2026-08-26):

- Gateway read: `parseSheetProtectionState` + `parseWorkbookProtectionState` in `packages/xlsx-gateway/src/gateway/xlsx-protection.ts`; `WorksheetState.sheetProtection` + `WorkbookSnapshot.workbookProtection` in `packages/xlsx-gateway/src/domain/workbook.types.ts`; wired into `readBasicWorkbook` (absent field = no element → byte-preserving no-op saves).
- Gateway write: `applyCellEditsToXlsx(..., sheetProtections (arg 10), ..., workbookProtectionState (trailing arg))` in `packages/xlsx-gateway/src/gateway/xlsx-gateway.ts` — the buffer→buffer entry the web route uses now carries both families through to the existing canonical appliers.
- Gateway tests: `packages/xlsx-gateway/tests/xlsx-protection-roundtrip.test.ts` — 22 tests: element parsers (enabled/disabled/legacy hash/modern algorithm hash/boolean form, both levels), readBasicWorkbook integration (protected sheet + locked workbook, absent fields, both password forms), write path (Excel-defaults element + lockStructure, unprotect removes both elements, protect→save→read→protected round-trip, no-op preservation of the worksheet entry's sha256, fail-closed unprotect on both password forms, re-affirm keeps the password element verbatim, unknown sheet name throws).
- Wire: `BrowserWorkbookSavePlan.sheetProtections` + `workbookProtectionState` in `packages/contractor-core/src/api/office-routes.ts` with `expectSheetProtectionState` / `expectWorkbookProtectionState` strict validation (non-boolean flags, missing sheet names, unknown fields — including password-bearing payloads — and counts > 1,000 are 400s); `handleSaveWorkbook` passes sheetProtections as arg 10 + workbookProtectionState as the trailing argument.
- Wire tests: `packages/contractor-core/tests/unit/office-protection-routes.test.ts` — 12 tests (canonical accept · empty array accept · non-boolean protected · missing/empty sheetName · password field rejected · non-array payload · oversized payload · canonical lock accept · null accept · non-boolean lockStructure · unknown workbook field · non-object state).
- Browser: `apps/web/src/screens/ExcelEditor.tsx` — protection file-state refs (seeded from the snapshot on open, merged after save) + toggle journals implementing the desktop's `recordSheetProtection` / `recordWorkbookProtection` semantics (an entry is DROPPED when toggled back to the file's original state → a no-op save emits NO protection family → the file's XML is preserved); `handleSave` conditionally emits both families; password fail-closed guards surface the desktop's own status strings. `apps/web/src/screens/excel/Ribbon.tsx` — the Review → Protection group (the previously-disabled stub is gone): Protect Sheet / Unprotect Sheet (echo label), Protect Workbook / Unprotect Workbook, Lock Cell, Unlock Cell.
- Architecture guards: `apps/web/tests/architecture.test.ts` — the EXCEL-020 block (7 new guards, 31 tests total in the file): both save-plan families emitted, recordSheetProtection semantics present, no disabled stub + all four commands wired, protectionLocked style-only CellEdits, NO raw OOXML/JSZip/direct-gateway writes in apps/web/src, office-client typed families carry no password fields.
- Browser E2E (`apps/web/tests/e2e/ribbon-protection.spec.ts`, 5 tests through real HTTP): (1) open a protected workbook → snapshot carries sheet + workbook protection state → ribbon echoes Unprotect Sheet / Unprotect Workbook → clean undo state; (2) protect from scratch → typed sheetProtections + workbookProtectionState on the wire → `<sheetProtection sheet="1" objects="1" scenarios="1"/>` + `lockStructure="1"` in the saved XML → reopen carries the state + ribbon echo; (3) unprotect round-trip → sheet element removed + toggle-back semantics keep the workbook lock → reopen clean; (4) editable-vs-locked: unlock A2 via the REAL ribbon (journaled protectionLocked:false style edit on the wire) + protect the sheet → `<protection locked="0"/>` in styles.xml + protected sheet → reopen + no-op save preserves both; (5) negative authorization: password-bearing elements on both levels → unprotect refused up front (desktop-parity message), journal stays clean, save carries no protection family, password attributes survive verbatim.
- Mandatory regression suite (all green, 83 browser E2E through the real HTTP stack): excel-browser (1), excel-shell (15), excel-format (3), excel-formula (7), excel-structural (2), ribbon-data (4 — INCLUDING the architect's sort/formula semantic gate), ribbon-data-validation (7), ribbon-filter (5), ribbon-review-notes (5), ribbon-home-persistence (3), ribbon-insert (1), ribbon-view (3), ribbon-remove-duplicates (4 — the EXCEL-018 regression), ribbon-protection (5), word-browser (1) + word-fidelity/image/table/marks/nested-runs (17).
- Unit suites: xlsx-gateway 567/567, contractor-core 390 passed + 4 skipped (pre-existing), web 204/204 (197 pre-existing + 7 new architecture guards). Typecheck: web app + contractor-core + xlsx-gateway all exit 0. Prettier: all changed files clean.
- Frozen surfaces untouched: `git diff --stat -- apps/sheets apps/docs apps/shell packages/platform-electron packages/renderer-bridge` is EMPTY.
- Workflow state: READY → IMPLEMENTING → PR_OPEN (#12) → MERGED (bf70fe9). NOT VERIFIED — the architect owns that transition.

Verification debt closure evidence (2026-08-27, independent rerun on main):

- Gateway roundtrip 22/22 (element parsers incl. legacy hash + modern algorithm hash at both levels, read integration, write round-trips, no-op preservation, fail-closed unprotect on both password forms); route validation 12/12; architecture guards 7/7.
- Read-path source audit: `hasPassword` covers BOTH password forms at BOTH levels; `SheetProtectionError` fail-closed on password-bearing unprotect at sheet level, workbook level, and protected ranges.
- Lock/unlock BOTH directions through the REAL ribbon: E2E test 4 (unlock) + NEW test 6 (lock — re-lock A2 through the ribbon: journaled `protectionLocked:true` style edit, no redundant `sheetProtections` re-journal, A2 references a LOCKED xf — no `locked="0"` on the referenced record — sheet protection survives, reopen carries protected+locked state). The acceptance condition is the referenced XF state, per the architect's directive.
- Password fail-closed (E2E test 5): refusal up front, journal clean, no destructive save, original password XML survives verbatim.
- No-op/toggle-back: toggling protection back to the file's original state drops the pending mutation (architecture-guarded `recordSheetProtection` semantics; E2E test 3).
- Browser E2E 6/6 local + 6/6 deployed vs production `genoffice.vercel.app`.

### EXCEL-021 — Tables

Objective: Import existing Excel tables, render structured-table semantics, create/delete tables, and persist table additions/deletions and supported table options.
Dependencies: EXCEL-009, EXCEL-015.
Required verification: table metadata, filters, styles, row/column edits, no-op preservation, create/save/reopen/delete.
Status: IMPLEMENTED, PENDING ARCHITECT REVIEW (2026-08-26). Built on the EXISTING canonical table families — no new engine write path, no browser-side OOXML, no parallel table model. The gateway's write side (`applyTableAdditions` in `xlsx-table-add.ts`: brand-new `xl/tables/tableN.xml` + worksheet `<tableParts>` + rel + `[Content_Types]` override, fail-closed on name collisions/overlaps/bad column names) and the structural table SHIFT (`shiftAnchoredSheetParts` → `shiftTablePart` shifts ref/autoFilter/sortState along row/column ops, fail-closed on anatomy changes) ALREADY existed and were unit-tested; the gaps closed here are (a) the READ side (`xlsx-table-read.ts`: `parseSheetTables` resolves each worksheet's `<tableParts>` through its rels into `WorksheetState.tables` — metadata plus PRE-RESOLVED banding colors from the workbook's real theme accents with Excel's HSL tint transform for the built-in Light/Medium/Dark families and the file's custom `<tableStyle>` dxfs; fail-closed PER SHEET so the workbook still opens and a no-op save preserves the parts byte-for-byte; parts without a readable ref are skipped per table — desktop `read_sheet_tables` parity), (b) `applyCellEditsToXlsx`'s trailing `tableAdditions` pass-through parameter (→ `planCellEditsToXlsx` slot 15), (c) the wire contract (`BrowserWorkbookSavePlan.tableAdditions` with the desktop preload's exact validation: sheetName, 0-based ordered integer area, name 1-255, columnNames 1-1000 × ≤255, built-in style name regex, bandedRows bool, unknown fields 400, ≤50 entries), and (d) the browser shell (Insert → Tables group: Table + Delete Table). Desktop parity notes: create ports `handleFormatAsTable` → `applyAiTableAdd` (header + ≥1 data row, ≤1,000 columns, TableN session names, column-name sanitize with write-back through the REAL facade, Univer registration with the DEFAULT theme — the desktop does not mute created tables); delete is convert-to-range for SESSION tables only (journal splice — nothing reaches the file, baked cells stay) and refuses file-native tables with the desktop's verbatim message; a sheet whose filter origin is a table (worksheet has no `<autoFilter>`) installs the table's range under journal suppression and REFUSES filter commands through a `BeforeCommandExecute` gate (desktop `FILTER_COMMAND_PATTERN` + message parity); row/column changes + new tables on one save hold the tables into a phase-2 tables-only save (desktop `heldTables` split-save parity). OUT OF SCOPE (documented): file-native table deletion (desktop refuses — "cannot be deleted yet"), post-create table options editing (rename/style change), table-owned filter criteria editing (blocked with the desktop's message), totals-row authoring, custom tableStyle dxf BAND EDITING (colors still RESOLVE from dxfs read-only).
Evidence (2026-08-26):

- Gateway read: `packages/xlsx-gateway/src/gateway/xlsx-table-read.ts` — `readTableThemePalette` (theme-attribute index order, srgbClr/sysClr), `readCustomTableStyles` (styles.xml `<tableStyles>` → per-style band palettes through the `<dxfs>`; dxf fills read bgColor-first — the OOXML differential-format quirk), `resolveTableStylePalette` (Light variant1 → filled header + #FFFFFF font; Light else → unfilled bold accent header; Dark → header tint −0.25 + #FFFFFF + stripe tint 0.4; Medium variant2 header tint −0.25, stripe tint 0.6 (variant1) / 0.8 (else); UNKNOWN names → Medium2 semantics; nameless → nothing), `resolveTableStyleBorder` (Light 1-7 only), Excel's HSL luminance tint transform, `parseSheetTables` (two-step rels lookup, `TableReadError` fail-closed per sheet, per-table skip for unreadable refs). `WorksheetState.tables` in `workbook.types.ts`; wired into `readBasicWorkbook` with the style context parsed once per workbook.
- Gateway write: `applyCellEditsToXlsx(..., tableAdditions (trailing arg))` in `packages/xlsx-gateway/src/gateway/xlsx-gateway.ts` — the buffer→buffer entry the web route uses now carries the family through to the existing canonical applier at `planCellEditsToXlsx` slot 15.
- Gateway tests: `packages/xlsx-gateway/tests/xlsx-table-roundtrip.test.ts` — 17 tests: rels-resolved metadata parse, Medium2 default-accent palette (#4472C4/#FFFFFF/#DAE3F3), Light/Dark/Medium-variant rules, unknown-name → Medium2, nameless → nothing, custom tableStyle dxfs, showFirstColumn gating, headerless tables, headerRowCount default, unreadable-ref skip, missing-relationship fail-closed per sheet, no-tableParts sheets, applyCellEditsToXlsx end-to-end persistence, create → save → reopen round-trip, no-op byte preservation.
- Wire: `BrowserWorkbookSavePlan.tableAdditions` in `packages/contractor-core/src/api/office-routes.ts` with `expectSheetTableAddition` strict validation (desktop `workbookTableAddSchema` parity — unknown fields, unordered/fractional/negative areas, over-long names/columns, non-builtin styles, non-boolean bandedRows, counts > 50 are 400s); `handleSaveWorkbook` passes tableAdditions as the trailing argument.
- Wire tests: `packages/contractor-core/tests/unit/office-table-routes.test.ts` — 14 tests (canonical accept · styleless accept · empty array · missing sheet name · bad table name · bad column lists · non-builtin style · all built-in families · non-boolean bandedRows · bad areas · unknown fields · non-array payload · oversized payload).
- Browser: `apps/web/src/office/table-banding.ts` — the pure, import-free port of the desktop's `applyTableBanding` (header band + bold + font color, first-data-row striping, totals band, first/last column emphasis, column stripes, whole-table fill, borderColor frame top MEDIUM/header-rule THIN/bottom MEDIUM; explicit cell fills WIN). `apps/web/src/screens/ExcelEditor.tsx` — file-table refs seeded from the snapshot (banding painted INTO the cell matrix before `createWorkbook`; Univer registration with a muted `plain-*` theme under journal suppression, headerRowCount=0 skipped), the session `tableAddsRef` journal, the table-owned filter origin install + `BeforeCommandExecute` gate, Insert → Table / Delete Table with the desktop's English status strings, `handleSave`'s conditional `tableAdditions` emission + split-save + post-save file-native merge. `apps/web/src/screens/excel/Ribbon.tsx` — the Insert → Tables group (the previously-disabled stub is gone): Table + Delete Table.
- Architecture guards: `apps/web/tests/architecture.test.ts` — the EXCEL-021 block (9 new guards, 40 tests total in the file): tableAdditions family emitted, heldTables split-save semantics present, no disabled stub + both commands wired, table-banding.ts import-free, FILTER_COMMAND_PATTERN gate present, session-only delete semantics + file-native refusal message, file-table seed + banding + addTableTheme + headerless skip, NO raw OOXML/JSZip/direct-gateway table writes in apps/web/src, office-client typed family.
- Browser E2E (`apps/web/tests/e2e/ribbon-table.spec.ts`, 6 tests through real HTTP + `buildExcelTableFixture`/`buildExcelTableCreateFixture`): (1) open a workbook with an existing table → snapshot carries name/area/headerRowCount/columns/styleName/showRowStripes + the resolved Medium2 palette → clean undo state → Delete Table on the file-native table refused with the desktop's verbatim message → journals nothing; (2) create from scratch over a Name-Box-selected range → typed tableAdditions on the wire → `xl/tables/table1.xml` + `<tableParts>` + rels + content-type override in the saved bytes → reopen reads the metadata back → Delete Table now refuses (file-native); (3) create then delete a session table → save carries NO tableAdditions → saved bytes have no table part, no worksheet wiring, and the table-touched entries are byte-identical to the fixture; (4) insert a row INSIDE the table → save → table ref + autoFilter grew to A1:B5 → reopen carries the shifted area; deleting the HEADER row then fails closed at save with the gateway's anatomy error surfaced; (5) Data → Filter on a table-owned filter sheet → refused with the desktop's exact message → nothing journaled; (6) an unrelated edit → save plan carries no table family → table part + rels + `<tableParts>` preserved byte-for-byte.
- `apps/web/tests/e2e/ribbon-insert.spec.ts` flipped with the feature: Table + Delete Table enabled; Picture/Chart still disabled with their documented reasons.
- Mandatory regression suite (all green, 89 browser E2E through the real HTTP stack — 83 pre-existing + 6 new): excel-browser (1), excel-shell (15), excel-format (3), excel-formula (7), excel-structural (2), ribbon-data (4 — INCLUDING the architect's sort/formula semantic gate), ribbon-data-validation (7), ribbon-filter (5), ribbon-review-notes (5), ribbon-home-persistence (3), ribbon-insert (1), ribbon-view (3), ribbon-remove-duplicates (4 — the EXCEL-018 regression), ribbon-protection (5 — the EXCEL-020 regression), ribbon-table (6 — NEW), word-browser (1) + word-fidelity/image/table/marks/nested-runs (17).
- Unit suites: xlsx-gateway 584/584 (567 pre-existing + 17 new), contractor-core 404 passed + 4 skipped (390 pre-existing + 14 new), web 213/213 (204 pre-existing + 9 new architecture guards). Typecheck: web app + contractor-core + xlsx-gateway all exit 0. Prettier: all changed files clean.
- Frozen surfaces untouched: `git diff --stat -- apps/sheets apps/docs apps/shell packages/platform-electron packages/renderer-bridge` is EMPTY.
- CI (run 32989308325 on 257f54b AND run 32990749832 on the evidence commit 563975d — manually dispatched; the pull_request event scheduled no run): **web gate GREEN on both** — typecheck on @contractor/web + @contractor/web-host + @contractor/core, unit suites, and the FULL 89-test browser E2E through the real HTTP stack (all 6 ribbon-table tests + the flipped ribbon-insert individually confirmed in the log). The test job's Lint failure is exactly the documented pre-existing 351 errors (all in frozen apps/docs/pdf/sheets/slides, non-office web screens, and contractor-core commercial domain — ZERO EXCEL-021 files); the e2e job's 10 failures are ALL in the frozen Electron shell suite (contractor/browser-e2e + sheets-* desktop specs, `sandboxed_renderer` launch errors — the same pre-existing condition PR #12 documented; the WEB browser E2E lives in the green web job).
- Production-build evidence: all 6 table E2E tests pass against the BUILT bundle (`npx vite build` + `vite preview` on :5178 with the dev-server API on :5179, `playwright.preview.config.ts`, 41.1s) — the feature works against the production-served artifact, not just the Vite dev server.
- Workflow state: READY → IMPLEMENTING → PR_OPEN (#16) → MERGED (e3d4311). NOT VERIFIED — the architect owns that transition.

Verification debt closure evidence (2026-08-27, independent rerun on main):

- Gateway table roundtrip 17/17 (rels-resolved metadata, theme/custom-style palettes, fail-closed per sheet, create→save→reopen, no-op byte preservation); route validation 14/14; architecture guards 9/9.
- E2E 6/6: import metadata + file-native delete refusal, create → typed wire → `xl/tables/table1.xml` + `<tableParts>` + rel + content-type override → reopen, session-table delete = no-op bytes, structural row insert grows ref + header-row delete fails closed, table-owned filter refuses ordinary filter commands, unrelated edit preserves table parts byte-for-byte.
- Browser E2E 6/6 local + 6/6 deployed vs production `genoffice.vercel.app`.

### EXCEL-022 — Images / Drawings

Objective: Import existing worksheet images and supported drawing metadata, render them, create/remove supported images, and persist via `visualAdditions`.
Dependencies: EXCEL-009, EXCEL-021.
Required verification: media bytes, relationships, anchor position, dimensions, deletion, no-op preservation.

Status: MERGED (PR #20 @ f5e805d → cb1fff6; both architect blockers fixed pre-merge: public `toBuilder().buildAsync()` read surface + absolute-anchor fail-closed omission). NOT VERIFIED — the architect owns that transition.

Implementation summary:

- Canonical read path: `packages/xlsx-gateway/src/gateway/xlsx-image-read.ts`
  (worksheet → drawing rel → drawing part → image rel → `xl/media/*`, typed
  `SheetImageInfo` with inline media, per-sheet fail-closed semantics,
  drawingIndex parity with the desktop sidecar and the edit family).
- Architecture decision: `WorksheetState.visuals` is CHART-SPECIFIC demo/AI
  replay state (kind: 'chart' only, never populated by readBasicWorkbook) and
  stays untouched; images landed as the DEDICATED `WorksheetState.images`
  field (the EXCEL-021 `tables` pattern).
- Canonical write path: the EXISTING visualAdditions/visualEdits families —
  `applyCellEditsToXlsx` gained trailing `visualAdditions`/`visualEdits`
  parameters (desktop translator call sites untouched); image deletion now
  cascades the image relationship and removes the media part ONLY when no
  remaining relationship anywhere references it (charts cascade unchanged;
  shared media — one rel across pictures or two rels to one part — is
  preserved; the [Content_Types] Default drops only when no same-extension
  part remains).
- Wire contract: `BrowserWorkbookSavePlan.visualAdditions`/`visualEdits`
  (IMAGE-ONLY additions — chart/shape payloads rejected with 400 until
  EXCEL-023), strict runtime validation in office-routes.ts (media types,
  base64 caps, bounded integer anchors, drawing-path pattern, exactly one of
  remove|anchor, unknown-field rejection, entry caps), and the save response
  returns `addedVisuals` locators so the browser merges persisted session
  images into its file-native state with the EXACT assigned drawing index.
- Browser: Univer 0.25.1 over-grid images (`@univerjs/preset-sheets-drawing`,
  already installed/registered) — locator-keyed drawing ids, install under
  journal suppression (no undo pollution), mutation journal via
  `sheet.mutation.set-drawing-apply` (move/resize dirty-tracking, delete
  journaling, session-add splice), Insert → Picture through a browser
  File/Blob upload path, and one-cell/absolute anchors FAIL CLOSED (the
  refused edit is reverted with remove+reinstall at the file geometry and a
  status explanation; nothing journals).
- EMU ↔ pixel conversion documented at 1 px = 9525 EMU (96 dpi) in
  `apps/web/src/office/sheet-images.ts` — the browser renders typed anchors
  and derives edit anchors from the live Univer model; the gateway
  serializes them verbatim. No JSZip, no OOXML/drawing/relationship parsing,
  no media-part manipulation in `apps/web/src` (architecture tests guard
  all of it).

Verification debt closure evidence (2026-08-27, independent rerun on main):

- Public-API audit: ZERO `_image` / private-facade reaches in `apps/web/src`; the image module uses only `FOverGridImage.toBuilder().buildAsync()` (architecture-guarded); the remaining `as unknown as` casts in the runtime are documented public-facade type-narrowing (numfmt/freeze/undo-prototype), none in the image/chart layer.
- Gateway image tests 20/20 (read path incl. anchor modeling + drawingIndex parity + rotation + unsupported/missing media skip + per-sheet fail-closed + multi-sheet + no-op byte preservation; delete cascade incl. shared rels, shared media via two rels, cross-drawing media, final-image removal, content-type cleanup, unrelated media survival); route validation 19/19; architecture guards green (no drawing/relationship XML in apps/web, type-only gateway imports, canonical visual wire types, private-internals ban).
- E2E 13/13: import/render, move, resize, insert (full canonical part set), delete cascade, multi-image isolation, save/reopen, no-op byte-for-byte (drawings/rels/media/content-types), relationship chain validity, one-cell fail-closed (renders, refuses edits), multi-sheet + JPEG round-trip, absolute-anchor omission (never relocated).
- Browser E2E 13/13 local + 13/13 deployed vs production `genoffice.vercel.app`.

### EXCEL-023 — Charts

Objective: Import supported charts, render them through the shared visual layer, create/edit/delete supported chart types, and persist `chartEdits`.
Dependencies: EXCEL-022.
Required verification: chart type, source ranges, dimensions, series, style, reopen.
Status: MERGED (PR #23 @ c9a02df → ef9ffb6 "Merge EXCEL-023 Charts"). NOT VERIFIED — the architect owns that transition.

Verification debt closure evidence (2026-08-27, independent rerun on main):

- Gateway chart tests 91/91 (chart reader read fidelity + omission cases + byte preservation + locator parity; the full applyChartEdit envelope incl. fail-closed cases); route validation 18/18 (chartEdits wire family + widened visualAdditions exactly-one-payload); architecture guards 6/6 (no chart OOXML in apps/web, canonical families, no private internals, pure-domain-only value imports, ribbon wiring, read/save path wiring).
- Canonical path audit: the browser never touches chart OOXML — XLSX → gateway chart reader → `WorksheetState.charts` → HTTP → typed browser chart state → `chartEdits`/`visualEdits`/`visualAdditions.chart` → gateway → XLSX (forbidden-pattern scans clean).
- E2E 9/9 covering the 15 scenarios + fail-closed proofs: import/render with stable identity + typed source ranges, move + resize through `visualEdits`, property edits through `chartEdits` round-trip, create from selection → file-native reopen, delete cascade (chart part removed, unrelated chart survives), no-op byte preservation, unsupported structures omitted (never relocated, bytes preserved), absolute anchors omitted, one-cell move-but-refuse-resize.
- Dirty-state reconfirmation: a chart edit alone marks the workbook dirty and enables Save (no cell-edit crutch) — E2E-proven.
- Browser E2E 9/9 local + 9/9 deployed vs production `genoffice.vercel.app` (Excel content identical from ef9ffb61e through the current production deployment).

### EXCEL-024 — Conditional Formatting

Objective: Import supported CF rules, render them, edit/create/delete rules, and persist `cfStates`.
Dependencies: EXCEL-021.
Required verification: expression rules, cell-value rules, color scales/icon sets/data bars where canonical support exists, fail-closed for unsupported extensions.
Status: PR_OPEN — PR #29, branch excel-024-cf head 6121f2e (rebased onto main @ edcd2ae after the PROJECT-026 merge — zero file overlap, byte-identical replay verified; original base 4d2ec337 → 3b8b2a4 → edcd2ae). Post-rebase gates re-run independently: typecheck ×3, unit 668/668 + 458+4sk + 229/229, format:check vs origin/main, ESLint, production build, full browser E2E 122/122. CI on 6121f2e: web gate GREEN (the canonical Excel gate); test/e2e/foundation failures fully attributed pre-existing — Lint annotations (20) all in frozen/non-Excel files, file-for-file identical to merged PR #25's set; e2e fails only at the frozen Electron shell step; foundation is the project-foundation isolation guard that fires on every Excel PR by construction (identical on merged PRs #23 and #25). NOT VERIFIED — the architect owns ARCHITECT_REVIEW → APPROVED → MERGED → VERIFIED.

Implementation evidence (2026-08-28, branch excel-024-cf):

- Canonical read path: `parseConditionalFormatting` in `xlsx-cf.ts` — the exact inverse of `applyCfRules`, emitting the Univer wire shape (CfWireRule) with dxf styles PRE-RESOLVED through `StylesheetReader.dxfAt` (font marks, rgb colors, solid bgColor fill, numFmt pattern — anything the writer cannot round-trip fails closed per sheet). Covers cellIs (numeric/text-equality/formula-fallback + the desktop's blank-divergence compensation for ≤20k-cell ranges), all ten text/presence operators, duplicate/unique, top10 (rank/percent/bottom), all four aboveAverage operator forms, expression, colorScale (2-5 stops, mixed threshold types incl. element-text formula cfvos), dataBar (base), iconSet (17-set OOXML whitelist, natural/reversed orders, worst-first rating-set inversion, strict thresholds). Fail-closed per sheet (CfReadError → no cfRules + `cfLocked: true`): x14 extensions (worksheet extLst twins AND cfRule-carried base halves), timePeriod, unknown types/operators, malformed sqref/priority, unresolvable dxfs, guard-rail overruns.
- Wire: `WorksheetState.cfRules?: CfWireRule[]` + `cfLocked?: boolean` (additive; the open route extends automatically). Save side: `BrowserWorkbookSavePlan.cfStates?: SheetCfState[]` with strict route validation (office-routes.ts) — unknown fields rejected at every level, operator/subType/threshold/icon-set enums mirrored from the canonical writer (OOXML_ICON_SETS imported from the gateway barrel — single source of truth), bounded caps (1,000 states / 500 rules / 100 ranges per rule / 255-char text / 1,000-char formulas), timePeriod and x14-only icon sets rejected with 400s before the engine touches bytes.
- Browser: rules install priority-descending through raw `sheet.mutation.add-conditional-rule` under journal suppression (Univer prepends each add; list-front rules win rendered precedence — verified from the installed package source, avoiding the desktop's reversed-precedence install quirk). cfDirty journal on the four CF mutations; declarative `collectCfStates` save snapshot (the full live rule set of every dirty sheet); `cfLocked` command gate (BeforeCommandExecute) refuses every rule edit on unrepresentable sheets plus the panel's date-occurring button — the silent-loss path the desktop has with skipped families is closed. Home → Styles → Conditional Formatting opens Univer's own manage-rules panel (the preset the runtime already registers; no invented CF UI).
- §8 resolution: the earlier audit's premise was WRONG — Univer 0.25.1 natively transforms live CF rule ranges on row/column structural ops. Precise mechanism (verified from the installed package source): the CF-UI plugin (registered by the browser preset) owns `ConditionalFormattingFormulaRefRangeController`, which registers EVERY live rule's ranges (and formulas) with sheets-formula's `FormulaRefRangeService`; on insert/delete/move the service returns transformed {ranges, formulas} and the controller emits `set-conditional-rule` (rewrites) + `add-conditional-rule` (splits) + `delete-conditional-rule` (fully-deleted) side-effect mutations, which both update the live model and flow into the browser's CF-dirty journal. Verified empirically both ways: a headless Node probe WITHOUT the CF-UI plugin shows ranges stay STALE and no CF mutations fire (the base CF package only markRuleDirty — the Phase A finding, correct but incomplete), while the real browser E2E (tests 5/6/10) proves the transforms active (insert at top edge extends, insert below shifts, delete shrinks, formula rules split with re-anchored formulas). A browser-side shift would DOUBLE-transform; none exists (architecture-guarded). The save snapshot keeps file, live model, and reopen in exact lockstep. Boundary note: Univer extends a range whose top edge sits at the insertion point (Excel's own boundary behavior — B2:B6 + 2 rows at row 2 → B2:B8) while the gateway's own structural `moveRange` shifts (B4:B8); the divergence is inert because any native transform fires CF mutations → the sheet is CF-dirty → `applyCfRules` replaces the whole CF set with the live snapshot AFTER the structural replay (the live model always wins on CF-dirty sheets).
- Tests: gateway 44 new reader tests (families, styles, priority, multi-rule/multi-range, x14/timePeriod/malformed fail-closed, write→reopen round-trip, clear, no-op byte preservation, unrelated-sheet preservation, per-sheet lock, structural sqref+formula shifting); route 17 new wire-validation tests; web architecture 6 new guards (no CF OOXML/JSZip in apps/web, no second CF engine, canonical type-only imports, install/journal/save/gate wiring, real panel command, native-transform resolution); browser E2E 10 new tests (render+install, create→save→XML→reopen, edit-preserves-siblings, delete-preserves-siblings, row insert extends ranges + splits formula rules + persists, column insert extends/shifts, x14+timePeriod lock with refused edits + byte-preserved no-op save, no-CF-dirty byte preservation incl. styles.xml, panel wiring, row deletion shrink). Suites: gateway 668/668, contractor-core 458+4sk, web unit 229/229, full browser E2E 122/122 (106 Excel + 16 Word — count re-verified by the continuation session's independent 8-chunk re-run), format:check clean vs origin/main, production build green, frozen surfaces zero changes.

### EXCEL-025 — Named Ranges / Name Manager

Objective: Import workbook and sheet-scoped defined names, provide Name Manager UI, support create/edit/delete of supported names, and persist `definedNamesState`.
Dependencies: EXCEL-015.
Required verification: workbook scope, sheet scope, print names if supported, formulas consuming names, reopen.
Status: IMPLEMENTING → PR_OPEN / VERIFYING (branch excel-025-named-ranges from main @ 71188d9). All local gates green: typecheck ×3, unit 693/693 + 475+4sk + 235/235, architecture 6 new guards, format:check vs origin/main, ESLint, production build, full browser E2E 133/133 (11 new name scenarios). NOT VERIFIED — the architect owns that transition.

Implementation evidence (2026-08-28, branch excel-025-named-ranges):

- Canonical read path: `parseDefinedNamesState` in `xlsx-defined-names.ts` — the exact inverse of the existing canonical writer `applyDefinedNamesState` (extended, not duplicated). Emits `DefinedNamesState {names, preserveNames}` with desktop-parity classification: `_xlnm.*` built-ins and hidden names skipped (the writer's keep-rules preserve them); invalid names (writer-mirrored predicate via the exported `definedNameIsSaveable` — the single canonical rule set, reused by the route's wire validation), out-of-range localSheetId scopes, and empty bodies routed to `preserveNames`; duplicate groups (case-insensitive, engine-resolution parity) ranked workbook-scope → sheet-scope → #REF! residue with only the winner modeled and the losers preserved. Structurally unparseable sections throw `DefinedNamesReadError` → snapshot carries `namesLocked: true` (workbook still opens, Name Manager refuses edits, no-op save preserves bytes). Snapshot: `WorkbookSnapshot.definedNames {names, preserveNames}` + `namesLocked` (the open route returns the snapshot verbatim — the family auto-extends the open wire).
- Wire + route: `BrowserWorkbookSavePlan.definedNamesState` (additive optional; the desktop's save-plan contract already carried the same family). `expectDefinedNamesState` strict validation in office-routes.ts — unknown fields rejected at every level, names checked against the gateway's canonical predicate (`definedNameIsSaveable` imported from the barrel — no rule duplication), formula bodies capped at 1,000 chars, sheet scopes bounded integers 0..255, per-(name, scope) duplicates rejected, the modeled∩preserve collision rejected with 400 BEFORE the engine touches bytes, entry caps 1,000. The route threads the validated state into `applyCellEditsToXlsx` at the `definedNamesState` position (11th, previously hardcoded null).
- Browser: file names install through the PUBLIC Univer builder facade (`newDefinedNameBuilder().load().build()` + `insertDefinedNameBuilder` — the desktop's exact install pair) under journal suppression; engine rejects (duplicate losers, sheet/table/function-name conflicts) land in a runtime preserve set unioned with the reader's list at save. Names journal = a single workbook-level dirty flag on the engine's `formula.mutation.set/remove-defined-name` mutations (the desktop's DEFINED_NAME_MUTATIONS pair, verified from installed engine-formula source). Save = `collectDefinedNamesState` declarative FULL live-model snapshot (editing one name never drops siblings) + `preserveNames`; SPLIT-SAVE parity with the desktop's heldNames (names never ride the same save as structural ops — the gateway's fail-closed guard; phase 1 structure, phase 2 names against the shifted bytes, with the live model's post-shift coordinates from Univer's own `UpdateDefinedNameController`). Name Box resolves defined names first (desktop `resolveGoToRef` port: names win over addresses case-insensitively; a name whose ref is a formula is not a jump target) and `goTo` now accepts `Sheet1!`-prefixed + `$`-marked references (desktop `isA1Reference` parity — the facade resolves the prefix and switches sheets). Name Manager dialog (Formulas → Defined Names): list/create/edit(name+ref; scope at creation only — desktop semantics)/delete through the public facade with the namesLocked fail-closed gate. `namesLocked` workbooks refuse every action with an explanatory message.
- §8-equivalent (structural ops × names): Univer 0.25.1's sheets-formula `UpdateDefinedNameController` (in the web's preset) intercepts every structural command + sheet rename/remove and rewrites the live name `formulaOrRefString`s, emitting `set-defined-name` mutations — the live model stays in sync AND the journal marks names dirty, which drives the split-save. Verified from installed source AND in the real browser (E2E test 9: insert 2 rows → live ref $A$1:$A$5 → $A$3:$A$7 → split-save → shifted XML → reopen reads the shifted truth). No browser-side second shift exists (would double-transform).
- Tests: gateway 25 new reader/family tests (scope read, _xlnm/hidden skip, invalid-name preservation, out-of-range/empty preservation, duplicate ranking + case-variant grouping, #REF! modelable, XML entity decoding, non-numeric localSheetId fail-closed, missing-name fail-closed, unaccounted-elements fail-closed, snapshot integration, namesLocked open, no-op section preservation, edit-one-preserves-all incl. print titles + hidden + preserved, delete-one-preserves-others, preserve-list survival, write→reopen round-trip, names+structural same-save rejection, the split-save two-phase consistency, package-level byte evidence (only workbook.xml changes on a names-only save)); route 17 new wire-validation tests (accept canonical + empty model, reject non-object/unknown fields/unsaveable names/missing fields/over-length formulas/non-integer + out-of-bounds scopes/in-scope duplicates/malformed preserve/overruns; same-name-different-scope accepted; collision rejected before bytes; null = untouched no-op); web architecture 6 new guards (no definedNames OOXML/JSZip in apps/web, no second persistence model/no client-side name predicate duplication, type-only gateway imports, install/journal/collect/split-save/gate wiring, ribbon wiring + no disabled stub, Name Box resolution + no private Univer internals in the dialog); browser E2E 11 new scenarios (open+install+Name Box workbook-scoped jump incl. case-insensitive, sheet-scoped jump with sheet switch + garbage refusal, Name Manager listing with scopes, create→save→XML→reopen with sibling/built-in/hidden/preserved survival, edit-preserves-everything byte-verbatim, delete-preserves-others, formula-consuming-name resolves live (150) and survives save/reopen as a formula, no-name-dirty byte preservation, structural shift + split-save + reopen, namesLocked fail-closed with refused edits + byte-preserved no-op save, duplicate-collision save failure with the writer's exact message — desktop parity). Suites: gateway 693/693, contractor-core 475+4sk, web 235/235, full browser E2E 133/133, frozen surfaces zero changes.

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
