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
Status: ARCHITECT_REVIEW — implementation complete, all local executable evidence green, pushed to origin/web-office-editor. Independent architect verification PENDING (Z.ai is the implementer and does NOT own the VERIFIED decision). Canonical path confirmed: Remove Duplicates is expressed as a value-level rewrite through the existing cell-edit mutation family (sheet.mutation.set-range-values → cellEditFromMutation → savePlan.edits → applyCellEditsToXlsx). No new mutation family, no wire change, no gateway change, no frozen-surface modification. Mirrors the desktop reference algorithm in apps/sheets/src/renderer/dedupe.ts + ribbon-actions.ts:1267-1309 verbatim (case-insensitive text comparison, type-strict, header preserved, unchanged rows untouched so formulas/styles survive, padding rows at the bottom blanked).
Evidence:

- Pure dedupe algorithm: `apps/web/src/office/dedupe.ts` (89 LOC, ZERO imports — verified by the architecture test `dedupe.ts is a PURE module`). Algorithm is a verbatim port of the frozen desktop reference `apps/sheets/src/renderer/dedupe.ts` (line-for-line identical semantics: `JSON.stringify` key with text lowercased, header preserved via `index === 0` short-circuit, in-place removal via `removed++` counter).
- Wiring: `apps/web/src/screens/excel/useExcelRuntime.ts:721-809` exposes `removeDuplicates(hasHeader): RemoveDuplicatesResult`. Reads `range.getValues()` (computed values), pads deduped matrix with nulls to the original selection height, writes back per-row via `ws.getRange(startRow + offset, startColumn, 1, width).setValues([[...newRow]])` — the SAME canonical facade call the desktop's `ribbon-actions.ts:1301-1305` makes. Unchanged rows (same content at same offset) are skipped — formulas/styles on kept rows survive (the desktop's documented trade-off: "moved rows land as their computed values"). Fail-closed on `height < 2` (returns `{ kind: 'select' }` matching the desktop's `appDedupeSelectRows`) and on `removed === 0` (returns `{ kind: 'noop' }` matching `appNoDuplicates`).
- Ribbon UI: `apps/web/src/screens/excel/Ribbon.tsx` enables the Remove Duplicates button (no longer `disabled`), opens an inline dialog with a "My data has headers" checkbox (default checked — Excel's own default), surfaces the dedupe result as a transient status toast. CSS in `apps/web/src/theme.css` (dialog overlay + toast, scoped classes `rb-dialog-*` / `rb-toast`, no global collisions).
- Canonical-path guards: `apps/web/tests/architecture.test.ts` adds a dedicated `EXCEL-018 Remove Duplicates uses the canonical cell-edit path` block (8 new tests). Enforces: dedupe module is pure (no Univer/electron/node/fs/jszip imports), the runtime wires through `FWorksheet.getRange().setValues()` (no private cellDataMatrix bypass), the Ribbon button is enabled and calls `api.removeDuplicates(...)`, NO raw OOXML/JSZip construction exists anywhere in `apps/web/src`, and the `ExcelEditor` save plan does NOT introduce a new `dedupeOps` / `removeDuplicatesState` family (the wire MUST stay `edits`-only).
- Browser purity preserved: no new xlsx writer in `apps/web`; no JSZip; no OOXML manipulation; no new save-plan family. The dedupe emits typed `CellEdit[]` payloads through the EXISTING `edits` channel — the wire payload is verifiably `savePlan.edits`-only (the E2E asserts `Object.keys(saveBody.savePlan).sort() === ['edits']`).
- Frozen surfaces untouched: `git diff --stat HEAD -- apps/sheets/src apps/docs apps/shell packages/platform-electron packages/renderer-bridge` returns empty. The desktop reference (`apps/sheets/src/renderer/dedupe.ts`, `ribbon-actions.ts:1267-1309`) was read-only throughout.
- Unit tests: `apps/web/tests/dedupe.test.ts` — 14 tests covering the full verification matrix (basic duplicate rows, header handling with/without, case-insensitive text, type-strict number vs string, boolean vs string, blank vs empty-string, multi-column comparison key, non-adjacent duplicates, input immutability, single-header-row edge case).
- Architecture tests: `apps/web/tests/architecture.test.ts` — 23 tests pass (15 pre-existing + 8 new EXCEL-018 guards). Web full unit suite: 183/183 pass.
- Production build: `vite build` succeeds in 17.1s.
- EXCEL-018 E2E (`apps/web/tests/e2e/ribbon-remove-duplicates.spec.ts`, 3 tests / 5 scenarios through real HTTP): (1) basic duplicates + header + multi-column key + styles survive + formula semantic + save/reopen + XML inspection — 5 rows × 2 columns = 10 CellEdits through the canonical `edits` channel (the save plan asserts `['edits']`-only), the saved XLSX carries compacted inline strings + blanked padding rows, reopen resolves the deduped state; (2) no-op dedupe fails closed without mutating (no save request fired, no unsaved-changes marker); (3) `<2-row` selection fails closed with the "select rows" status message.
- Mandatory regression E2E (13 specs, 57 tests total, all green): excel-shell (15), excel-browser (1), excel-format (3), excel-formula (8), excel-structural (2), ribbon-data (4 — including the architect's sort/formula semantic gate at line 537), ribbon-view (3), ribbon-filter (5), ribbon-data-validation (7), ribbon-review-notes (5), ribbon-home-persistence (3), ribbon-insert (1), word-browser (1). The architect's sort/formula semantic gate (relative refs rewrite, absolute refs untouched) passes UNCHANGED — formula fidelity is preserved.
- Production deployment: deployed to genoffice.vercel.app via the established Vercel deployment path against the pushed commit (URL in the implementation PR). Deployed EXCEL-018 E2E against the exact pushed commit is run after CI begins.
- Workflow state: READY → IMPLEMENTING → PR_OPEN → VERIFYING → ARCHITECT_REVIEW (current). NOT VERIFIED — pending independent architect review.

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
