## EXCEL-025 — Named Ranges / Name Manager (browser Excel defined-name read/resolve/edit/create/delete + persistence)

Branch `excel-025-named-ranges` from `main` @ `71188d9` (EXCEL-024 merged). 1 commit, audit-first per the work item's Phase A mandate (Phase A recorded in worklog as task `EXCEL-025-phase-a` before any implementation code).

### Architecture / data flow (canonical defined-name transport — no browser OOXML, no second persistence model)

```
XLSX → xlsx-gateway CF-names reader (parseDefinedNamesState — the exact
         inverse of the pre-existing canonical writer applyDefinedNamesState;
         model/preserve split with desktop-parity classification:
         _xlnm.*/hidden auto-kept, invalid names + out-of-range scopes +
         empty bodies + duplicate losers → preserveNames verbatim,
         ranked workbook→sheet→#REF! per name, case-insensitive)
      → WorkbookSnapshot.definedNames {names, preserveNames}
         (+ namesLocked fail-closed marker on unparseable sections)
      → HTTP → browser installs via the PUBLIC Univer builder facade
         (newDefinedNameBuilder().load().build() + insertDefinedNameBuilder —
         the desktop's exact install pair) under journal suppression;
         engine rejects → runtime preserve set
      → live Univer defined-name model (Name Manager dialog, Name Box
         resolution, formulas consuming names)
      → names-dirty journal on formula.mutation.set/remove-defined-name
         (the desktop's DEFINED_NAME_MUTATIONS pair)
      → typed save plan: collectDefinedNamesState — the FULL live model +
         preserveNames (reader list ∪ runtime rejects)
      → routeOffice strict validation (400 before bytes) → gateway
         applyDefinedNamesState (rewrites <definedNames>, keeps every
         preserved entry byte-verbatim) → XLSX
```

### Supported vs fail-closed classification

**Supported (canonical round-trip proven):**
- workbook-scoped names; sheet-scoped names (localSheetId ↔ tab order, remapped by the gateway on sheet reorder/delete — pre-existing)
- create / edit (name + refers-to; scope fixed at creation — desktop semantics) / delete through the Name Manager
- Name Box navigation: names win over addresses (case-insensitive), jump selects the range and switches sheets for qualified refs; formulas as refs are not jump targets
- formulas consuming names (the engine's DefinedNamesService resolves them; `=SUM(GlobalTotal)` computes 150 live and survives save/reopen as a formula)
- `#REF!` residue names (round-trip verbatim); external-reference text bodies (verbatim formula text)
- structural row/column ops: Univer 0.25.1's `UpdateDefinedNameController` (sheets-formula, in the web preset) rewrites live name refs AND fires set-defined-name mutations → names dirty → **split-save** (phase 1 structure, phase 2 names against the shifted bytes — desktop `heldNames` parity; the gateway's own guard rejects names+structural in one pass)

**Fail-closed (never silent loss):**
- `_xlnm.*` built-ins (Print_Area / Print_Titles — owned by the page-setup family) and hidden names: preserved byte-verbatim by the writer's keep-rules, never modeled, never editable
- invalid names (writer-predicate failures: cell-ref lookalikes, reserved prefixes, bad starts, >255), out-of-range scopes, empty bodies: reader-classified → preserveNames, never modeled — editing another name never drops them
- duplicate-name losers (ranked; case-variant twins): preserved verbatim; a workbook whose winner collides with a preserved twin fails the SAVE closed with the writer's exact message (E2E test 11 — desktop parity)
- unparseable `<definedNames>` sections: whole family locks (`namesLocked`) — the workbook opens, the Name Manager refuses every action, a no-op save preserves the bytes

### The critical preservation invariant

**Editing one supported defined name never silently deletes or rewrites unrelated names.** Proven three ways: gateway byte tests (edit one name → siblings + print titles + hidden + preserved entries byte-verbatim; names-only save touches ONLY workbook.xml), browser E2E XML assertions (create/edit/delete each preserve every unrelated name), and the writer's own collision/duplicate guards mirrored as 400s at the route before bytes.

### Test counts

| Suite | Result |
|---|---|
| xlsx-gateway unit | **693/693** (25 new: reader classification, fail-closed paths, round-trip, split-save consistency, byte proofs) |
| contractor-core unit | **475 passed + 4 skipped** (17 new wire-validation tests) |
| web unit | **235/235** (6 new architecture guards) |
| browser E2E (full) | **133/133** (11 new name scenarios + 122 regression) |

### Byte-preservation evidence

- No-op save: the whole `<definedNames>` section byte-identical (workbook.xml delta = the documented `fullCalcOnLoad` marker only) — gateway test + E2E test 8
- Names-only save: worksheets / styles / rels / content-types byte-identical; only workbook.xml changes — gateway package-level test
- Edit one name: every sibling + `_xlnm.Print_Titles` + hidden + preserve-listed entries byte-verbatim — gateway test + E2E tests 4/5/6
- namesLocked workbook + no-op save: section byte-for-byte — E2E test 10

### Browser E2E evidence

11 scenarios through the real stack (browser → Vite → HTTP → vercel-handler → routeOffice → xlsx-gateway): open+install, Name Box workbook-scoped jump (case-insensitive), sheet-scoped jump with sheet switch, Name Manager listing with scopes, create→save→XML→reopen, edit-preserves-everything, delete-preserves-others, formula-consuming-name resolves live + survives reopen as a formula, no-op byte preservation, structural insert → live ref shift ($A$1:$A$5 → $A$3:$A$7) → split-save → shifted XML → reopen, namesLocked fail-closed, duplicate-collision save failure (desktop parity).

### CI evidence

Local gates mirror CI exactly (typecheck ×3 clean; format:check vs origin/main clean; ESLint clean on all changed files; production build green 19.7s). The canonical `web` gate will be attributed on this PR's head after CI runs — per the established pattern, `test` (frozen-file lint annotations), `e2e` (frozen Electron shell), and `foundation` (the project-isolation guard that structurally fires on every Excel PR) failures will be attributed against the merged-PR baseline.

### Known limitations

- Scope is chosen at creation and cannot change afterwards (desktop Name Manager parity)
- No name comments / hidden-flag editing (the wire model carries name/formula/sheetIndex only — desktop parity)
- A name that wins a duplicate group while its twin is preserved makes ANY name edit's save fail closed with an explicit error (the canonical writer's collision guard — desktop-identical behavior, surfaced as E2E test 11)
- `namesLocked` workbooks (structurally unparseable sections) are read-preserving but not editable
- Sheet rename/delete in-session: the web shell performs no sheet management; the gateway's rename/delete XML-level interplay for names (renameSheetReferencesInDefinedNames, localSheetId remap) is pre-existing and desktop-covered
- DESKTOP divergence avoided: the desktop's Rust reader silently skips empty-body/invalid names WITHOUT a preserve list — the web reader routes them to preserveNames so a rewrite can never drop them

### Frozen-surface result

Zero changes to `apps/sheets`, `apps/docs`, `apps/slides`, `apps/shell`, `packages/platform-electron`, `packages/renderer-bridge` (verified by diff vs origin/main).

Not VERIFIED — the architect owns `ARCHITECT_REVIEW → APPROVED → MERGED → VERIFIED`.
