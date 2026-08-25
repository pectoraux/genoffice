# GenOffice Excel Parity — Dependency Graph

## Legend

`→` means the left item must be VERIFIED before the right item becomes READY.

## Foundation

```text
EXCEL-001 Workspace Shell
        ↓
EXCEL-002 Workbook Open/Save
        ↓
EXCEL-003 Formula Fidelity
        ↓
EXCEL-005 Structural Operations
```

## Feature-family enablement

```text
EXCEL-002 + EXCEL-004
        ↓
EXCEL-010 Formatting

EXCEL-002 + EXCEL-005 + EXCEL-003
        ↓
EXCEL-011 Row/Column Operations
        ↓
EXCEL-012 Formula-Bar Fidelity
```

```text
EXCEL-001
  ↓
EXCEL-014 Workspace Shell
  ↓
EXCEL-015 Core Ribbon Mutation Persistence
```

```text
EXCEL-015
  ↓
EXCEL-016 Data Filter
  ↓
EXCEL-017 Data Validation
```

## Remaining feature sequence

```text
EXCEL-015
 ├──→ EXCEL-018 Remove Duplicates
 ├──→ EXCEL-019 Notes
 │       ↓
 │   EXCEL-020 Protection
 │
 ├──→ EXCEL-021 Tables
 │       ↓
 │   EXCEL-022 Images / Drawings
 │       ↓
 │   EXCEL-023 Charts
 │
 ├──→ EXCEL-024 Conditional Formatting
 ├──→ EXCEL-025 Named Ranges
 ├──→ EXCEL-026 View / Page Setup expansion
 └──→ EXCEL-027 Advanced Formatting
```

```text
EXCEL-011 + EXCEL-015
        ↓
EXCEL-028 Autosave / Recovery
```

```text
EXCEL-019 + EXCEL-021 + EXCEL-024
        ↓
EXCEL-029 Undo / Redo / Shared Journal
```

```text
EXCEL-014
        ↓
EXCEL-030 Theme / Locale
```

```text
EXCEL-018..027
        ↓
EXCEL-031 Ribbon Completion
        ↓
EXCEL-032 Dialog Parity
```

```text
EXCEL-029 + EXCEL-031 + EXCEL-032
        ↓
EXCEL-033 Shared Sheets Editor Core
        ↓
EXCEL-034 AI Panel Parity
        ↓
EXCEL-035 Large Workbook / Streaming
```

```text
EXCEL-018..035
        ↓
EXCEL-036 Final Parity Acceptance
```

## Parallelizable streams

### Data semantics stream

```text
Filter → Data Validation → Remove Duplicates → Tables → CF → Notes/Protection
```

### Visual stream

```text
Formatting → Advanced Formatting → Images → Charts → Ribbon → Dialogs
```

### Lifecycle stream

```text
Structural Ops → Formula Fidelity → Journal/Undo → Autosave/Recovery
```

### Host parity stream

```text
Workspace Shell → Theme/Locale → Ribbon Completion → Dialogs → Shared Core → AI
```

These streams may be implemented by separate agents only when they do not modify the same canonical contract simultaneously without an explicit coordination work item.

## Contract collision rule

The following areas are high-collision and require architect coordination when changed by parallel agents:

- `packages/xlsx-gateway/src/gateway/xlsx-gateway.ts`
- `packages/xlsx-gateway/src/gateway/*.ts`
- `packages/contractor-core/src/api/office-routes.ts`
- `apps/web/src/api/office-client.ts`
- `apps/web/src/screens/ExcelEditor.tsx`
- save-plan / mutation journal contracts

Parallel implementation is encouraged in isolated branches, but only one integration PR should change a given canonical save family at a time.
