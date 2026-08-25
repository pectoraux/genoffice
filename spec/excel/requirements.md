# GenOffice Excel Parity — Requirements

## Scope

These requirements govern the Excel parity sprint for both Electron and Web hosts. Completed increments remain documented here so future agents can reason from repository state rather than conversation history.

## Product requirements

### EXCEL-001 — Workspace shell parity
The web editor shall present a GenOffice-native Excel workspace equivalent to the desktop experience: application title bar, ribbon, name box, formula bar, worksheet surface, sheet tabs, status bar, zoom, save state, theme controls, and host navigation.

Acceptance: one coherent shell, no duplicate generic Office chrome, core controls have deterministic interactions.

### EXCEL-002 — Workbook open/save fidelity
Opening and saving a workbook shall preserve supported workbook semantics through the canonical gateway. No-op saves shall preserve unsupported content rather than silently rewriting it.

Acceptance: golden fixtures and byte-level package assertions for accepted feature families.

### EXCEL-003 — Formula fidelity
Formula edits, formula clears, recalculation echoes, structural changes, and save/reopen shall preserve formula semantics.

Acceptance: relative-reference, absolute-reference, cross-sheet, clear-to-literal, clear-to-blank, structural-interaction, and recalculation regression gates.

### EXCEL-004 — Cell formatting
The web editor shall support and persist the desktop-supported common formatting surface, including font, size, bold, italic, underline, strike, font/fill color, alignment, wrap, merge, and number format. Advanced formatting shall not silently disappear when edited.

Acceptance: format fixtures, save-plan assertions, XML preservation, reopen assertions.

### EXCEL-005 — Row/column structural operations
Insert/delete rows and columns, merge/unmerge, and accepted future structural families shall be represented as semantic operations and replayed by the canonical engine.

Acceptance: exact coordinate transformations and save/reopen tests.

### EXCEL-006 — Sort and formula synchronization
Sort operations shall preserve row-associated relative references and absolute references according to Excel semantics.

Acceptance: deterministic browser + gateway regression gate covering in-session values, save plan, saved XML, and reopen/recalc.

### EXCEL-007 — AutoFilter
Existing filters shall import into the browser, render correctly, be editable, save through the canonical filter family, and reopen identically.

Acceptance: read/import, create, modify, clear, unsupported-criteria fail-closed, no-op preservation, and production E2E.

### EXCEL-008 — Data validation
Existing validation rules shall import and render, user-created rules shall persist, and unsupported validation semantics shall fail closed.

Acceptance: whole/decimal/list/date/time/text/custom, prompts/errors, multiple rules, clear, unsupported/x14 behavior, production E2E.

### EXCEL-009 — Feature-family persistence
The wire contract shall expose canonical mutation families as the corresponding ribbon commands become enabled. Each enabled command must have a canonical save/reopen path.

Initial families include:

- filterStates
- dvStates
- noteStates
- sheetProtections
- tableAdditions
- visualAdditions
- chartEdits
- cfStates
- definedNamesState
- pageSetupStates

### EXCEL-010 — Import parity
If the canonical engine writes a feature family, the workbook reader should expose enough state for the web renderer to display it before that feature is declared fully supported.

No feature may be considered parity-complete based solely on write support.

### EXCEL-011 — Undo/redo and journal parity
The two hosts shall converge on semantic command/journal behavior, including suppression of load-time mutations and correct undo/redo grouping.

### EXCEL-012 — Autosave and recovery
The web experience shall eventually provide desktop-equivalent autosave and crash-recovery semantics using browser-safe storage/transport while retaining canonical workbook semantics.

### EXCEL-013 — Theme and locale parity
The web and desktop shall converge on light/dark/system theme behavior and supported localized Excel UI.

### EXCEL-014 — Ribbon parity
The web ribbon shall converge on the desktop's seven primary tabs and contextual capabilities. Controls may be disabled only while their canonical persistence/semantic path is genuinely unavailable, and disabled controls must identify the architectural reason.

### EXCEL-015 — Advanced workbook surfaces
The final parity target includes tables, charts, images, conditional formatting, filters, data validation, notes/comments, freeze panes, named ranges, protection, page setup, pivots, sparklines, hyperlinks, and other accepted workbook features present in the desktop target.

### EXCEL-016 — AI surface parity
The eventual target includes the desktop Sheets AI panel and its approved tools, with a transport-neutral interface for web and Electron.

### EXCEL-017 — Performance and scale
The web shall remain usable for normal office workbooks within the documented payload ceiling. If workbook size demands it, the architecture shall permit session/range streaming without changing canonical semantics.

## Current completion ledger

### Completed

- EXCEL-001: initial workspace shell parity — substantially implemented.
- EXCEL-002: core XLSX open/save pipeline — implemented.
- EXCEL-003: formula fidelity — implemented with dedicated regression gates.
- EXCEL-004: common cell formatting + number formats — implemented for accepted subset.
- EXCEL-005: initial row/column operations and merge/unmerge — implemented.
- EXCEL-006: sort/formula synchronization — implemented and independently gated.
- EXCEL-007: AutoFilter — implemented.
- EXCEL-008: Data Validation — implemented.

### Remaining

The remaining requirements are staged into the work-item backlog and may not be considered complete until the corresponding verification matrix entries pass.
