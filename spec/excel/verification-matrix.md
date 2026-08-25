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

| Work item | UNIT | ARCH | GOLDEN | BROWSER | DESKTOP | DEPLOYED | BYTE | CI | Status |
|---|---|---|---|---|---|---|---|---|---|
| EXCEL-010 Formatting | ✓ | ✓ | ✓ | ✓ | inherited desktop reference | ✓ | ✓ | web gate | VERIFIED |
| EXCEL-011 Structural Ops | ✓ | ✓ | ✓ | ✓ | existing desktop behavior | ✓ | ✓ | web gate | VERIFIED |
| EXCEL-012 Formula Fidelity | ✓ | ✓ | ✓ | ✓ | desktop formula path | ✓ | ✓ | web gate | VERIFIED |
| EXCEL-014 Workspace Shell | ✓ | ✓ | N/A | ✓ | read-only reference | ✓ | N/A | web gate | VERIFIED |
| EXCEL-015 Core Ribbon | ✓ | ✓ | ✓ | ✓ | reference | ✓ | ✓ | web gate | VERIFIED |
| EXCEL-016 Filter | ✓ | ✓ | ✓ | ✓ | reference | ✓ | ✓ | web gate | VERIFIED |
| EXCEL-017 Data Validation | ✓ | ✓ | ✓ | ✓ | reference | ✓ | ✓ | web gate | VERIFIED |

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

### EXCEL-025 Named Ranges

Required:

- workbook scope;
- worksheet scope;
- create/edit/delete;
- formulas consuming names;
- print-area/print-title semantics where supported;
- save/reopen;
- browser + production E2E.

### EXCEL-026 View / Page Setup

Required:

- freeze rows/columns;
- gridline visibility where persisted by canonical model;
- formula-view state where persisted;
- print orientation/margins/page setup supported by the canonical engine;
- save/reopen;
- browser + production E2E.

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
