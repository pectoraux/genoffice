# GenOffice Project — Dependency Graph

## Foundation DAG

```text
PROJECT-001
   ├── PROJECT-002 ──┐
   │                 ├── PROJECT-004 ──┐
   │                 └── PROJECT-005 ──┼── PROJECT-006
   └── PROJECT-003 ────────────────────┘
```

## Full roadmap dependency intent

```text
001 → 002 → 003
002 → 004
002 → 005
004 + 005 + 002 → 006
006 → 007 → 008 → 009
002 + 004 + 006 → 010 → 011 → 013
006 → 012 → 013
002 + 003 + 006 → 014 → 015 → 016
014 → 017 → 018 → 019
015 + 018 → 020
002 + 003 + 006 → 021 → 022 → 023 → 024
004 + 022 → 025
012 + 022 → 026
021 + 022 → 027
021 + 022 → 028
021 + 023 → 029 → 030 → 031
022 + 023 → 032
010 + 011 + 022 → 033
005 + 021 → 034
009 + 022 → 035
002 + 003 + 022 → 036 → 037
009 + 036 → 038
008 + 009 + 011 → 039 → 040
007 + 008 → 041
007 + 014 → 042
010 + 013 → 043
004 + 010 → 044
013 + 043 → 045
027 + 028 + 029 + 030 + 031 → 046
014 + 015 + 016 + 018 + 019 + 020 → 047
046 + 047 → 048 → 049
```

A work item cannot be authorized until all direct dependencies are objectively accepted. A dependency change requires synchronized updates to `work-items.md`, this file, and `architecture-lock.md` when the change affects a frozen invariant.

## Package dependency edges (PROJECT-014)

The `.gproj` adapter lives in `@genoffice/project-file`. Its static package dependencies are:

```text
@genoffice/project-file → @genoffice/project-engine → @genoffice/project-contracts
@genoffice/project-file → @genoffice/project-contracts (direct, for types + brand helpers)
```

`@genoffice/project-file` does NOT depend on `@genoffice/project-scheduling` (the scheduling package is a logical work-item dependency `006 → 014`, not a static package import — the file format stores canonical input only; derived scheduling state is re-computed by the scheduling engine on load). The adapter has NO React, Electron, Node, browser, HTTP, or MPP/MSPDI imports (architecture-lock §13). The `project-foundation.yml` CI workflow greps `packages/project-file` for forbidden imports alongside `project-contracts`, `project-engine`, and `project-scheduling`.
