# GenOffice Excel Parity — Architecture Lock

Status: FROZEN
Scope: Excel renderer/workbook parity initiative across web and Electron.
Authority: This document is the controlling architectural contract for the Excel parity sprint.

## 1. Mission

GenOffice Excel is a Microsoft Excel-class spreadsheet application with equivalent desktop and web experiences. The target is not a browser-friendly approximation. The target is a GenOffice-native clone of the Microsoft Office spreadsheet experience, including the workspace shell, editing behavior, workbook semantics, feature commands, persistence, and recoverable document state.

The desktop and web applications are two hosts of one spreadsheet semantic system. Rendering is not authoritative for workbook semantics.

## 2. Layer order

```text
User Intent / UI
    → semantic workbook command / mutation
    → workbook domain / canonical contracts
    → canonical XLSX gateway / engine
    → OOXML bytes
    → host-specific transport
```

A lower layer may consume a higher-layer contract but may not redefine it. Browser UI must never become the source of truth for XLSX semantics.

## 3. Canonical ownership

- `packages/xlsx-gateway` owns canonical XLSX read/write semantics.
- `packages/contractor-core` owns the web HTTP contract and runtime validation for the office API.
- `apps/sheets` owns Electron-specific host integration and desktop chrome.
- `apps/web` owns web-specific host integration and web chrome.
- Univer is the shared in-session spreadsheet interaction engine, not the persistence authority.

Web code must not implement a second XLSX serializer, ZIP writer, OOXML parser, relationship parser, or desktop engine.

## 4. Desktop/web parity invariant

For every accepted feature:

1. equivalent user intent must result in equivalent workbook semantics;
2. equivalent workbook input must render equivalent semantics;
3. save/reopen must preserve the feature;
4. desktop and web must use the same canonical engine semantics wherever practical;
5. transport-specific differences must remain behind adapters.

A feature is not complete merely because it renders in Univer.

## 5. Formula fidelity invariant

Formula text written or rewritten in the browser must remain semantically synchronized with workbook structural operations. Recalculation echoes must never overwrite journaled formulas with literals. Relative references may change when Excel semantics require them to; absolute references must remain fixed.

Any change to formula mutation capture, structural-operation replay, or save ordering requires the formula regression gate to remain green.

## 6. Identity

Workbook, worksheet, cell, mutation, and structural operation identity must never depend on array position alone. Stable worksheet identifiers and canonical coordinates are required.

## 7. Mutation and journal model

User mutations are collected into canonical save-plan families. Declarative state families are snapshotted from the live Univer model where that is safer than mutation replay. Load-time imports are suppressed from undo/journal history.

The canonical engine remains responsible for composing structural operations, cell edits, styles, filters, validation, page setup, and other supported mutation families.

## 8. Import semantics

If a gateway feature can write a workbook feature but cannot read it back into the snapshot, the web editor must not pretend that feature parity is complete. Read/import, in-session rendering, mutation capture, save, and reopen verification are separate acceptance concerns.

Unsupported or unrepresentable OOXML must fail closed and preserve original bytes rather than silently degrade content.

## 9. Web purity

`apps/web/src` must not import:

- Electron APIs
- Node filesystem/process APIs
- `jszip`
- raw OOXML construction/parsing logic
- desktop renderer internals
- `apps/sheets/src` or `apps/docs/src`

Browser code communicates through typed host/API contracts.

## 10. Frozen desktop safety

The Electron renderer is the reference implementation for parity, but web parity work must not weaken desktop behavior. Changes to `apps/sheets/src/renderer`, shared contracts, or preload require explicit scope authorization and desktop regression evidence.

Pure extraction into a future shared package is permitted only when behavior is demonstrably unchanged.

## 11. Acceptance authority

An implementation agent's report is not evidence by itself. Work-item completion requires repository state, executable tests, architecture checks, deterministic fixtures, and required host/deployed evidence to satisfy the verification matrix.

## 12. Architecture changes

Any change to a frozen invariant requires an Architecture Change Request containing:

- motivation
- affected invariants
- alternatives considered
- compatibility impact
- migration plan
- verification impact
- explicit architect approval

Until approved, the architecture change must not be implemented.

## 13. Workflow

The sprint follows the WorkflowOS lifecycle:

```text
DRAFT
→ READY
→ ASSIGNED
→ IMPLEMENTING
→ PR_OPEN
→ VERIFYING
→ ARCHITECT_REVIEW
→ APPROVED
→ MERGED
→ VERIFIED
```

`VERIFICATION_FAILED` returns to `IMPLEMENTING`.
`CHANGES_REQUESTED` returns to `IMPLEMENTING`.
`ARCHITECTURE_CHANGE_REQUIRED` requires architecture review before implementation continues.

There is one active implementation PR per work item.
