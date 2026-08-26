# GenOffice Excel Parity — Architecture

## 1. Architectural objective

The Excel initiative provides one canonical workbook semantic model with two host experiences:

- Electron desktop: the current `apps/sheets` implementation and native host integration.
- Web: `apps/web` running in a browser against the typed office HTTP boundary.

The web application is not a separate spreadsheet implementation. It is a host of the same workbook semantics.

## 2. System topology

```text
                    ┌─────────────────────┐
                    │  Excel User Intent   │
                    └──────────┬──────────┘
                               │
                 ┌─────────────┴─────────────┐
                 │                           │
          Electron Host                Web Host
          apps/sheets                 apps/web
                 │                           │
          Univer + shell              Univer + shell
                 │                           │
          Desktop adapter              HTTP adapter
                 │                           │
          └─────────────┬─────────────┘
                        │
              Canonical workbook contract
                        │
                xlsx-gateway / engine
                        │
                     XLSX bytes
```

## 3. Canonical engine boundary

`packages/xlsx-gateway` is the workbook semantic authority. It owns:

- XLSX parsing
- cell values and formulas
- styles
- merges
- row/column structure
- filters
- data validation
- page setup/freeze panes
- hyperlinks
- notes/comments
- conditional formatting
- tables
- visual/image additions
- chart mutations
- defined names
- protection
- other accepted workbook mutation families

The web host must pass canonical state and mutation families to this engine rather than reproducing these algorithms.

## 4. Web transport boundary

The current office HTTP surface is intentionally small:

```text
POST /api/office/workbooks/open
POST /api/office/workbooks/save
POST /api/office/documents/open
POST /api/office/documents/save
```

The Excel roadmap extends the existing workbook open/save contract rather than creating one HTTP endpoint per ribbon button. Feature families are represented as typed snapshot fields and save-plan fields.

The wire layer owns runtime validation but does not own workbook semantics.

## 5. Renderer architecture

Univer provides in-session editing, selection, formula bar behavior, worksheet interaction, and native spreadsheet widgets.

GenOffice owns the host shell around Univer:

- title bar
- ribbon
- name box
- formula bar integration
- status bar
- theme
- dialogs
- host actions
- command journaling
- dirty state
- save/open lifecycle
- feature availability

Univer mutations must be normalized into canonical save families.

## 6. Snapshot model

Workbook open produces a canonical snapshot suitable for browser rendering. The snapshot may grow over time to include:

- cells
- styles
- merges
- row/column dimensions
- hidden state
- freeze panes
- filters
- data validation
- conditional formatting
- tables
- notes
- hyperlinks
- images/media metadata
- charts
- named ranges
- page setup
- pivots
- sparklines
- protection
- other accepted feature metadata

Advanced fields may be made lazy/opt-in if payload size becomes a material web constraint.

## 7. Save model

Save sends:

```text
original XLSX bytes
+ canonical save plan
    ├── cell edits
    ├── structural operations
    ├── style deltas
    ├── page setup
    ├── filter states
    ├── data validation states
    ├── note states
    ├── table additions
    ├── visual additions
    ├── chart edits
    ├── conditional-formatting states
    ├── defined-name state
    └── other accepted families
```

The canonical engine applies the plan and emits XLSX bytes. The browser never generates `styles.xml`, worksheet XML, relationship XML, ZIP archives, or media package structure.

## 8. Declarative vs mutation replay

Use declarative live-model snapshots for complex state families where mutation replay is fragile:

- filters
- data validation
- conditional formatting
- page setup
- tables
- charts
- visuals
- notes
- protection
- defined names

Use mutation journals for atomic edits and structural changes where the mutation stream is stable:

- cell edits
- style deltas
- row/column insert/delete
- merge/unmerge

Every family requires an explicit decision documented in the work item.

## 9. Formula and structural synchronization

Structural operations and formula rewrites are one semantic transaction.

For operations such as sort, insert/delete rows, insert/delete columns, the acceptance suite must prove:

- relative references follow moved cells according to Excel semantics;
- absolute references stay anchored;
- the browser's journal contains the rewritten formula text when applicable;
- the gateway performs structural permutation without independently rewriting formula semantics beyond its defined responsibility;
- save/reopen recomputes to the same semantic result.

## 10. Desktop reference boundary

`apps/sheets` is used as behavioral and visual reference. The parity sprint may inspect desktop code but must not copy Electron-only transport or filesystem logic into browser packages.

When desktop logic is transport-independent, it can later be extracted into a shared package after the web command surface stabilizes.

## 11. Future shared package

A future `packages/sheets-editor-core` may contain:

- Univer bootstrap
- locale/theme helpers
- formula-fidelity patches
- journal
- suppression
- selection formatting
- ribbon components
- dialogs
- host-independent import/collect transforms

Extraction is intentionally deferred until behavior is stable enough to avoid prematurely freezing incorrect abstractions.

## 12. Performance

The web implementation currently uses full workbook snapshots. The initial web target is ordinary office workbooks within the existing upload/payload ceiling.

If real-world workbook sizes exceed the practical threshold, the architecture permits a session/range API and viewport streaming without changing workbook semantics.

## 13. Security

All HTTP input is runtime validated. Browser clients validate server responses. Unsupported workbook features fail closed rather than silently disappearing.

No raw filesystem paths, Electron objects, Node APIs, or arbitrary OOXML are exposed across the browser boundary.

## 14. Visual parity target

The target is a GenOffice-native clone of the Microsoft Excel workspace, not a generic Univer embedding. The web shell must converge on the desktop shell's:

- title bar
- ribbon structure
- formula/name box area
- sheet tabs
- status bar
- zoom
- dialogs
- command availability
- theme behavior
- task flows

Feature parity requires both visual and semantic parity.
