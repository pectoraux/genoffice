# GenOffice Excel Parity — Verification Matrix

## Acceptance rule

A work item is VERIFIED only when every required evidence column for that work item is green. Local tests alone do not satisfy production or desktop evidence requirements. An implementation-agent report without independent evidence does not satisfy a criterion.

## Evidence classes

- `UNIT`: deterministic package/unit tests.
- `ARCH`: static architecture/purity checks.
- `GOLDEN`: deterministic workbook fixture with exact semantic/XML expectations.
- `BROWSER`: real Playwright browser through HTTP/API stack.
- `DESKTOP`: real Electron renderer/host test.
- `DEPLOYED`: browser E2E against production `genoffice.vercel.app`.
- `BYTE`: byte/package preservation assertions for no-op or untouched features.
- `CI`: authoritative GitHub Actions result for the exact commit.

## Completed work

| Work item                  | UNIT | ARCH | GOLDEN | BROWSER | DESKTOP                     | DEPLOYED | BYTE | CI       | Status   |
| -------------------------- | ---- | ---- | ------ | ------- | --------------------------- | -------- | ---- | -------- | -------- |
| EXCEL-010 Formatting       | ✓    | ✓    | ✓      | ✓       | inherited desktop reference | ✓        | ✓    | web gate | VERIFIED |
| EXCEL-011 Structural Ops   | ✓    | ✓    | ✓      | ✓       | existing desktop behavior   | ✓        | ✓    | web gate | VERIFIED |
| EXCEL-012 Formula Fidelity | ✓    | ✓    | ✓      | ✓       | desktop formula path        | ✓        | ✓    | web gate | VERIFIED |
| EXCEL-014 Workspace Shell  | ✓    | ✓    | N/A    | ✓       | read-only reference         | ✓        | N/A  | web gate | VERIFIED |
| EXCEL-015 Core Ribbon      | ✓    | ✓    | ✓      | ✓       | reference                   | ✓        | ✓    | web gate | VERIFIED |
| EXCEL-016 Filter           | ✓    | ✓    | ✓      | ✓       | reference                   | ✓        | ✓    | web gate | VERIFIED |
| EXCEL-017 Data Validation  | ✓    | ✓    | ✓      | ✓       | reference                   | ✓        | ✓    | web gate | VERIFIED |

## Required future verification

### EXCEL-018 Remove Duplicates

Required:

- duplicate rows removed only within the selected range;
- header row semantics correct;
- multiple selected columns act as the comparison key;
- formulas and styles remain attached to retained rows;
- no unintended row changes outside the selected range;
- save/reopen yields identical values and structure;
- unsupported cases fail closed;
- browser + production E2E.

### EXCEL-019 Notes / Comments

Required:

- import existing notes;
- render author/text metadata;
- create/edit/delete;
- multiple notes;
- no-op byte preservation;
- save/reopen;
- browser + production E2E.

Status: VERIFIED (commit 68cbb9d1a36c54b1731b2f43460da547fde1e437; see `spec/excel/work-items.md` for the full evidence record).

Verification evidence map:

- import existing notes — `xlsx-notes.test.ts` "readBasicWorkbook integration" + E2E test 1 (snapshot + live model);
- render author/text metadata — E2E test 1 (author split convention proven in the live model);
- create/edit/delete — E2E tests 3+4 (create, edit, delete through the real facade);
- multiple notes — E2E tests 1, 5+6 (two notes; delete-one isolation);
- no-op byte preservation — gateway test 10 + E2E test 10 + live production assertion;
- save/reopen — gateway test 6 (write→reopen) + E2E tests 7+8+9 (typed wire + XML + reopened live model);
- browser + production E2E — `ribbon-review-notes.spec.ts` (5 tests, real HTTP) + the 25-assertion live pipeline verification against genoffice.vercel.app.

### EXCEL-020 Protection

Required:

- protect/unprotect sheet;
- editable vs locked cell behavior;
- workbook protection where supported;
- password semantics where canonical engine supports them;
- save/reopen;
- negative authorization tests;
- browser + production E2E.

### EXCEL-021 Tables

Required:

- import existing tables;
- correct range/header/total-row state;
- table style rendering;
- create/delete;
- row/column changes within a table;
- table-owned filter behavior;
- save/reopen;
- no-op byte preservation;
- browser + production E2E.

### EXCEL-022 Images / Drawings

Required:

- import supported image types;
- media relationship preservation;
- anchor position and dimensions;
- create/delete;
- no-op media byte identity;
- save/reopen;
- browser + production E2E.

Status: MERGED (PR #20). Local verification complete AND the post-merge
steps are now complete (2026-08-27): browser E2E 13/13 locally + 13/13
against production `genoffice.vercel.app`, ZIP-entry byte inspection in
suite, independent rerun of gateway (20/20) + route (19/19) + architecture
guards on current main. Awaiting the architect's VERIFIED decision.

Evidence (local):

- Gateway unit: `tests/xlsx-image.test.ts` — 18 tests covering the read
  path (two-cell/one-cell/absolute modeling, drawingIndex parity over
  non-picture anchors, rotation, unsupported-media skip, missing-media
  skip, per-sheet fail-closed wiring, multi-sheet, no-op byte
  preservation) and the delete cascade (single image, one-of-two,
  shared rel, shared media via two rels, cross-drawing media, final-image
  removal, move-only edits, composed delete+move).
- Route unit: `office-visual-routes.test.ts` — 19 validation tests
  (image-only wire, media types, base64 shape/caps, anchor bounds,
  unknown fields, both/neither edit kinds, drawing-path pattern, count
  caps).
- Architecture: 3 new guards (no drawing/relationship XML work in
  apps/web/src, type-only gateway imports, canonical visual wire types).
- Browser E2E: `ribbon-images.spec.ts` — 12 tests through browser → HTTP
  → routeOffice → gateway → XLSX bytes → reopen (import, render,
  identity, move, resize, insert, delete cascade, multi-image isolation,
  save/reopen, no-op preservation + plan inspection, relationship chain,
  one-cell fail-closed, multi-sheet shared media, JPEG round-trip).

### EXCEL-023 Charts

Required:

- import supported chart types;
- chart source ranges;
- series and category references;
- dimensions/position;
- create/edit/delete supported chart families;
- save/reopen;
- chart part/relation preservation;
- browser + production E2E.

### EXCEL-024 Conditional Formatting

Required:

- import supported rule categories;
- render correct visual outcome;
- create/edit/delete;
- rule priority and applies-to range;
- unsupported extensions fail closed;
- save/reopen;
- no-op byte preservation;
- browser + production E2E.

Evidence (2026-08-28, branch excel-024-cf @ head 6121f2e — rebased onto main @ edcd2ae after the PROJECT-026 merge, zero file overlap, byte-identical replay; every gate re-executed from scratch on the rebased tree, none trusted from prior reports):

- UNIT: xlsx-gateway 668/668 (44 new reader tests: every saveable family, dxf subset, priority/multi-range/multi-rule, x14/timePeriod/malformed fail-closed, write→reopen round-trip, clear, no-op byte preservation, structural sqref+formula shift, unrelated-part byte identity); contractor-core 458 passed + 4 skipped (17 new cfStates wire-validation tests); web unit 229/229 (6 new architecture guards).
- ARCH: 6 new guards — no CF OOXML/JSZip in apps/web/src; no second CF engine; canonical type-only imports; install/journal/snapshot/gate wiring; real panel command; native-transform §8 resolution (no browser-side double shift).
- GOLDEN/BYTE: package-level no-op proof (CF sections + styles.xml dxfs + unrelated sheet byte-identical through a cell-edit save); supported-edit proof (only expected CF parts change; workbook.xml gains only the documented fullCalcOnLoad marker; rels/content-types/unrelated sheet byte-identical); browser-level byte proofs in E2E tests 7/8.
- BROWSER: 10/10 ribbon-conditional-formatting.spec.ts (open+render, create→XML→reopen, edit/delete preserving siblings, row-insert/column-insert/row-delete live-range transforms persisted + reopened, x14+timePeriod lock with refused edits, no-CF-dirty byte preservation, ribbon→real panel). Full-suite regression 122/122 (106 Excel + 16 Word) re-run in 9 chunks on the rebased head 6121f2e.
- CI (PR #29, head 6121f2e): the canonical `web` gate GREEN. `test` fails ONLY at the Lint step — its 20 annotations are file-for-file identical to merged PR #25's set, ALL in frozen/non-Excel surfaces (apps/docs, apps/sheets, apps/slides, .github, pre-existing BOQ.tsx), zero in EXCEL-024 files (verified via the check-run annotations API). `e2e` fails ONLY at the frozen Electron shell step. `foundation` fails at the project-foundation isolation guard, which fires on every Excel PR by construction (diffs vs main and rejects Excel-surface changes — the guard's stated purpose is keeping project branches off Excel surfaces; identical failure on merged+verified PRs #23 and #25).
- §8 resolution (decisive): Univer 0.25.1's CF-UI `ConditionalFormattingFormulaRefRangeController` + `FormulaRefRangeService` natively transform live CF rule ranges on structural ops (extend-at-top-edge on insert, shrink on delete, formula-rule splits) and fire set/add/delete-conditional-rule mutations that journal the sheet CF-dirty — verified from installed package source AND empirically (headless Node probe WITHOUT the CF-UI plugin shows ranges stay STALE and no mutations fire; the real browser E2E tests 5/6/10 prove the transforms active). A browser-side shift would double-transform; none exists (architecture-guarded).
- DEPLOYED: NOT YET — branch pushed (excel-024-cf @ 6121f2e) and PR #29 OPEN; production E2E against genoffice.vercel.app must run after merge with the deployed SHA exactly matching merged main (the architect owns the merge + VERIFIED transitions).

### EXCEL-025 Named Ranges

Required:

- workbook scope;
- worksheet scope;
- create/edit/delete;
- formulas consuming names;
- print-area/print-title semantics where supported;
- save/reopen;
- browser + production E2E.

Evidence (2026-08-28, branch excel-025-named-ranges from main @ 71188d9 — every gate re-executed on the final tree):

- UNIT: xlsx-gateway 693/693 (25 new: reader classification — scope read, _xlnm/hidden skip, invalid-name/out-of-range/empty-body preservation, duplicate ranking workbook→sheet→#REF! with case-insensitive grouping, #REF! modelable, XML entity decoding, non-numeric localSheetId + missing-name + unaccounted-elements fail-closed, namesLocked snapshot, write→reopen round-trip, names+structural same-save rejection, the split-save two-phase consistency, package-level byte evidence — only workbook.xml changes on a names-only save); contractor-core 475 passed + 4 skipped (17 new wire-validation tests: canonical accepts, unknown fields rejected at both levels, unsaveable names rejected via the gateway's own predicate, length caps, scope bounds, in-scope duplicates, collision 400-before-bytes, null no-op); web 235/235 (6 new architecture guards).
- ARCH: 6 new guards — no definedNames OOXML/JSZip in apps/web/src; no second persistence model (no client-side applyDefinedNamesState/parseDefinedNamesState/definedNameIsSaveable); type-only gateway imports; install/journal/collect/split-save/namesLocked wiring; Ribbon Formulas → Name Manager wired (no disabled stub); Name Box resolution present + no private Univer internals in the dialog.
- GOLDEN/BYTE: package-level proofs inside the gateway suite — a no-op save preserves the <definedNames> section byte-for-byte (workbook.xml delta = the documented fullCalcOnLoad marker only); editing ONE name preserves every sibling + the _xlnm.Print_Titles built-in + hidden names + preserve-listed entries byte-verbatim; a names-only save touches ONLY workbook.xml (worksheets/styles/rels/content-types byte-identical); the split-save two-phase flow (structural phase shifts the section; names phase rewrites from the live post-shift snapshot) produces consistent bytes that reopen correctly.
- BROWSER: 11/11 ribbon-named-ranges.spec.ts through the real stack (browser → Vite → HTTP → vercel-handler → routeOffice → xlsx-gateway): open+install with Name Box workbook-scoped jump (incl. case-insensitive resolution), sheet-scoped jump switching to the Other sheet, Name Manager listing with scope labels, create→save→XML→reopen with sibling/built-in/hidden/preserved survival, edit-preserves-everything, delete-preserves-others, formula-consuming-name resolving live (SUM(GlobalTotal)=150) and surviving save/reopen as a FORMULA, no-name-dirty byte preservation, structural insert shifting the live ref ($A$1:$A$5→$A$3:$A$7) + split-save + shifted XML + reopen, namesLocked fail-closed (refused edits + byte-preserved no-op save), duplicate-collision save failure with the canonical writer's exact message (desktop parity). Full-suite regression 133/133.
- ARCHITECT BLOCKER CORRECTION (2026-08-28, PR #31 CHANGES REQUESTED → corrected): the reader collapsed same-name cross-scope definitions; corrected to (case-insensitive name, scope) grouping — the writer's exact uniqueness key — at reader, route, and writer; browser holds BOTH same-name entries (engine service is id-keyed; siblings ride the public insertDefinedNameBuilder facade), Name Box resolves with Excel scope precedence (active-sheet-scoped shadows workbook), Name Manager create/rename enforce the scope rule. Gates re-executed on the corrected tree: gateway 701/701 (+8 reader/save pair tests), contractor-core 477+4sk (+2 case-rule wire tests), web 235/235, BROWSER 17/17 names scenarios (test 11 redesigned around a GENUINE same-scope duplicate; 6 new pair scenarios: open/install/list both, Name Box scope precedence, edit-one-preserves-other both directions, delete-one-preserves-other, structural shift of both, create-at-new-scope + same-scope refusal) and full-suite regression 139/139. Known limitation documented: Univer 0.25.1 formula-engine name resolution is name-keyed (scope-blind) for AMBIGUOUS same-names in cell formulas; the file preserves both scopes correctly.
- CI (PR #31, correction head 753f13a): the canonical `web` gate GREEN. `test` fails ONLY at the Lint step — its 20 annotations are file-for-file identical to the established pre-existing set (frozen apps/docs/apps/sheets/apps/slides/.github surfaces + the pre-existing BOQ.tsx, which is NOT in the EXCEL-025 diff), zero in EXCEL-025 files (verified via the check-run annotations API on 753f13a). `e2e` fails ONLY at the frozen Electron shell step. `foundation` fails at the project-foundation isolation guard, which structurally fires on every Excel PR (identical on merged+verified PRs #23/#25/#29); `desktop-e2e` is green and outside the Excel scope.
- SECOND ARCHITECT BLOCKER CORRECTION (2026-08-28, PR #31 CHANGES STILL REQUIRED → corrected at head 73158ce): the modeled∩preserve collision checks at the writer and route compared `preserveNames` case-sensitively — a same-scope case-variant pair (`Foo` + `foo`) could have BOTH variants serialized instead of failing closed. Corrected to a case-insensitive collision predicate at all four layers: writer `applyDefinedNamesState` (lowercased preserve set for both the collision check and the keep-rule — every case-variant element of a preserved name kept verbatim, no silent deletion), route `expectDefinedNamesState` (same check as a 400 before bytes, every case combination), browser `namePreserved` (create/rename refused up front in any case), reader poisoned-name sweep (a case-variant sibling of an unmodelable element stays file-only rather than poisoning every names-dirty save). Gates re-executed on the corrected tree: gateway 706/706 (+3 writer/+2 reader regressions), contractor-core 479+4sk (+2 route regressions), web 235/235, BROWSER 19/19 names scenarios (18: Foo+foo same-scope duplicate fails the save closed with the writer's exact message + creating FOO at the Data scope refused up front; 19: the case-variant cross-scope pair Total+total installs both and round-trips with the twin byte-verbatim) and full-suite regression 141/141; independent byte evidence outside the test files 7/7 green.
- CI (PR #31, second correction head 73158ce): the canonical `web` gate GREEN. `test` fails ONLY at the Lint step — 20 annotations file-for-file identical to the established pre-existing set (BOQ.tsx NOT in the EXCEL-025 diff — verified), zero in EXCEL-025 files. `e2e` fails ONLY at the frozen Electron shell step. `foundation` fails only at the project-foundation isolation guard (identical on merged+verified PRs #23/#25/#29); `desktop-e2e` green.
- DEPLOYED: NOT YET — branch pushed (excel-025-named-ranges @ 73158ce, second blocker correction) and PR #31 OPEN pending architect re-review; production E2E against genoffice.vercel.app must run after merge with the deployed SHA exactly matching merged main (the architect owns the merge + VERIFIED transitions).

### EXCEL-026 View / Page Setup

Required:

- freeze rows/columns;
- gridline visibility where persisted by canonical model;
- formula-view state where persisted;
- print orientation/margins/page setup supported by the canonical engine;
- save/reopen;
- browser + production E2E.

Evidence (2026-08-28, branch `excel-026-view-persistence`, PR pending):

- UNIT (gateway): `xlsx-view-state.test.ts` 25/25 through the REAL entry points (readBasicWorkbook / applyCellEditsToXlsx slot) — reader exposes ONLY non-default `<sheetView>` flags (existing, absent, explicit defaults, xsd "true"/"false" literals, malformed values ignored-for-modeling, malformed sibling of a valid flag, missing sheetViews); parseFrozenPane unit gap filled (frozen, split rejected, frozenSplit, zero splits); write→reopen, edit→save→reopen, clear/reset for both flags AND pane removal (0/0), malformed→canonical replacement, freeze set→clear round-trip; the print family through the slot (orientation/margins/paperSize/scale/fitToPage incl. an untouched `blackAndWhite` attribute staying verbatim); no-op save leaves both worksheet parts byte-identical (touchedEntries + sha256 proof); per-sheet isolation (editing Data leaves Other's non-default view + bytes intact); unknown sheet name fails closed. Full gateway suite 731/731.
- UNIT (route): `office-page-setup-routes.test.ts` 20/20 — freeze regression (0/0 clear accepted; non-integer/negative/beyond-OOXML rejected), boolean view flags (wrong types rejected per field), print family (orientation enum + non-string, margins enum, paperSize 1..118, scale 10..400, fit axes 0..1000, fitToPage boolean), the documented unknown-key seam (accepted, but a WIRED invalid field still 400s), cap (100 accepted / 101 rejected), non-array rejected. Full contractor-core suite 499 passed + 4 skipped.
- ARCH: 6 new guards — no view/page OOXML, JSZip, `<sheetView>`/`<pageSetup>`/`<pane>` construction in apps/web/src; no second persistence model (no applyPageSetupState/setSheetViewAttr/MARGIN_PRESETS/setFrozenPane client-side); type-only wire (pageSetupStates + showGridlines/showFormulas fields, no gateway value imports); the single generalized journal family wired (PageSetupJournalEntry + pageSetupRef + SET_FROZEN 0/0 clear + 'sheet.mutation.toggle-gridlines' + toggleShowFormulas/applyShowFormulasView/installFormulaViewInterceptor + the journal-emission line + import seeding from sheet.view + applyPageLayout); ribbon wiring (Show Formulas + the five live Page Layout selects, no disabled stubs); the formula-view module's shape (RENDER_RAW_FORMULA_KEY + INTERCEPTOR_POINT.CELL_CONTENT, no gateway imports). Full web suite 241/241.
- BROWSER: `ribbon-view-persistence.spec.ts` 6 tests / 11 scenarios through the real stack: (1) the file's `<pane>` imports (snapshot + live model); (2) freeze clear + re-freeze + save/reopen with exact XML assertions; (3) the freeze-CLEAR defect fix end-to-end — the save plan carries {0,0}, the `<pane>` element is REMOVED from the saved XML, reopen stays unfrozen; (4) showGridLines="0" imports (snapshot + live model + ribbon echo) with the sibling sheet default; (5) gridline restore persists as the dropped attribute + reopen default; (6) hide→save→show→save returns the exact original XML state; (7) showFormulas="1" imports AND renders (pixel-level proof: toggle-off changes the grid render, toggle-back returns the byte-exact imported render); (8) formula-view save/reopen round-trip (save plan + showFormulas="1" in XML + reopened render identical); (9) an unrelated cell edit preserves all three view truths (showGridLines="0" + showFormulas="1" + frozen pane) with `<sheetViews>` byte-verbatim and the untouched sibling sheet byte-identical, and a default-view file gains NO view attributes through a cell-edit save; (11) the Page Layout print family (orientation landscape, margins narrow, size A4, fit width 1 page) round-trips with exact attribute assertions (fitToHeight="0", pageSetUpPr fitToPage, margins 0.25) and a clean reopen. `ribbon-view.spec.ts` regression 3/3. Full-suite regression 147/147 in 7 batches.
- Frozen surfaces: ZERO changes (diff confined to apps/web + packages/xlsx-gateway + packages/contractor-core + spec).
- CI (PR #35, head 67500a7): the canonical `web` gate GREEN — typecheck (web/web-host/core) clean; unit web 241/241, web-host 78/78, contractor-core 499+4sk; production build green; Playwright browser E2E **147 passed (18.3m)**, exactly matching local. `test` fails ONLY at Lint — 20 annotations file-for-file identical to the established pre-existing set (all frozen/non-Excel files incl. BOQ.tsx which is NOT in the EXCEL-026 diff — verified via the check-run annotations API and git), zero in EXCEL-026 files; Check formatting SUCCESS. `e2e` fails ONLY at the frozen Electron shell step (job-steps API). `foundation` fails ONLY at Verify branch isolation (the structural cross-domain guard, identical on merged+verified PRs #23/#25/#29/#31); its desktop-e2e and web-e2e jobs both SUCCESS. CI evidence comment: issuecomment 5452797708.
- DEPLOYED: NOT YET — awaiting ARCHITECT_REVIEW → merge; production E2E against genoffice.vercel.app runs after merge with the deployed SHA exactly matching merged main (the architect owns the merge + VERIFIED transitions).

### EXCEL-027 Advanced Formatting

Required:

- borders and line styles;
- text rotation;
- indentation;
- built-in and custom number formats;
- named cell styles where supported;
- paste special;
- format painter;
- save/reopen;
- XML style assertions;
- browser + production E2E.

Evidence (LOCAL, branch `excel-027-advanced-formatting`):

- Implemented: per-edge borders (all 13 OOXML ST_BorderStyle line styles, hex colors, null clears, side isolation, diagonal + diagonalUp/Down verbatim preservation), text rotation (1..90 ccw / 91..180 cw / 255 stacked / 0 clear), indentation (read + journal + wire 0..250; UI deferred to EXCEL-032 with the desktop's Format Cells dialog), and number-format import render seeding — all through the ONE canonical `WorkbookStyleEdit` family on the existing cell-edit save pipeline (no second persistence model; the browser never learns border/rotation XML). Named cell styles / paste special / format painter = documented fail-closed/deferred (no canonical engine path). Route defect fixed: border null-clears previously dropped by validation (silently keeping the file's border); indent bound aligned 15→250 with integer enforcement.
- UNIT: gateway `xlsx-advanced-formatting.test.ts` 23/23; gateway suite 754/754. Route `office-style-routes.test.ts` 11/11; contractor-core 510 passed + 4 skipped. Web 257/257 (unit +9 journal-mapping, architecture +7 EXCEL-027 guards). Typecheck ×3 clean; `FORMAT_BASE_REF=origin/main format:check` clean; ESLint clean on all 16 changed files; production build green (18.6s).
- BROWSER E2E: `ribbon-advanced-formatting.spec.ts` 7/7 covering all 14 mandated scenarios (live model → save request → saved XLSX/XML → reopened model at every step; pixel-level render proofs). FULL regression 154/154 in 9 batches (147 prior + 7 new). Independent byte evidence captured (read → side-isolated border write + rotation write → reopen with exact XML + untouched sibling sheet byte-identical).
- Frozen surfaces: ZERO changes (`apps/sheets`, `apps/docs`, `apps/slides`, `apps/shell`, `packages/platform-electron`, `packages/renderer-bridge`).
- CI: pending push (the canonical `web` gate must go green on the PR head).
- DEPLOYED: NOT YET — awaiting ARCHITECT_REVIEW → merge; production E2E against genoffice.vercel.app runs after merge with the deployed SHA exactly matching merged main (the architect owns the merge + VERIFIED transitions).

### EXCEL-028 Autosave / Recovery

Required:

- configurable autosave interval;
- blur/visibility-triggered save behavior where required;
- deterministic recovery copy identity;
- crash/reload simulation;
- restore/discard flows;
- recovery copy cleanup;
- stale-session safety;
- browser + production E2E.

### EXCEL-029 Undo / Redo

Required:

- every accepted mutation family journals correctly;
- load-time imports do not pollute undo;
- one logical action groups correctly;
- undo restores prior workbook semantics;
- redo restores subsequent semantics;
- save after undo/redo produces correct bytes;
- browser and desktop evidence.

### EXCEL-030 Theme / Locale

Required:

- light/dark/system theme;
- theme propagation to Univer and shell;
- persistence/reload;
- all supported locale packs load without runtime errors;
- ribbon/dialog/status text changes consistently;
- browser + desktop evidence.

### EXCEL-031 Ribbon Completion

Required:

- all seven primary tabs rendered;
- each enabled command has a real canonical path;
- disabled commands have documented architectural reason;
- no fake success actions;
- keyboard shortcuts where required;
- browser + desktop E2E.

### EXCEL-032 Dialog Parity

Required:

- each required dialog has equivalent inputs and outputs;
- dialog actions produce semantic commands;
- cancel produces no mutation;
- validation errors are deterministic;
- save/reopen effects verified;
- browser + desktop E2E.

### EXCEL-033 Shared Sheets Editor Core

Required:

- pure extraction with no behavior change;
- desktop before/after E2E identical;
- web before/after E2E identical;
- no Electron/Node imports in browser-safe package;
- architecture guards updated;
- bundle/build checks;
- CI green.

### EXCEL-034 AI Panel

Required:

- transport-neutral Chat/Agent interface;
- streaming response;
- approved workbook tools;
- safe semantic mutation application;
- CAS/drift guard;
- inline undo;
- authorization/security boundaries;
- browser + desktop + deployed E2E.

### EXCEL-035 Large Workbook / Streaming

Required:

- large workbook fixture exceeds ordinary snapshot threshold;
- range/chunk loading;
- scroll/sheet-switch correctness;
- low memory growth;
- edits outside viewport persist;
- no-op preservation;
- desktop and web performance evidence.

### EXCEL-036 Final Parity Acceptance

Final acceptance requires:

1. all prior work items VERIFIED;
2. desktop and web E2E suites green;
3. deployed web E2E green;
4. deterministic golden workbook suite green;
5. no-op byte preservation green across supported families;
6. architecture/purity checks green;
7. formula-semantics regression gate green;
8. structural-operation regression gate green;
9. no forbidden browser imports;
10. CI green for the web implementation gate;
11. any repository-wide failures independently attributed and proven unrelated;
12. Architect Review records no open blocking findings.

## Failure rule

A failed criterion returns the current work item to IMPLEMENTING. Do not weaken or delete the criterion merely to achieve green status. Unsupported capability should be represented as an explicit deferred requirement, not as a silently degraded implementation.
