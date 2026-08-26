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

## Package dependency edges (PROJECT-015)

PROJECT-015 (MSPDI import) reuses the accepted PROJECT-014 `@genoffice/project-file` package boundary — the MSPDI adapter lives beside the `.gproj` adapter under `packages/project-file/src/mspdi/**`. Its static package dependencies are unchanged:

```text
@genoffice/project-file → @genoffice/project-engine → @genoffice/project-contracts
@genoffice/project-file → @genoffice/project-contracts (direct, for types + brand helpers)
```

`@genoffice/project-file` STILL does NOT statically depend on `@genoffice/project-scheduling` at the package level — the MSPDI importer delegates semantic validation to `validateProjectDocument` (`@genoffice/project-engine`) and produces a canonical `ProjectDocument`; the scheduling engine runs on that document at the host/test layer (the PROJECT-015 test suite imports `schedule` from `@genoffice/project-scheduling` to prove imported documents are schedulable, but the package itself does not). The MSPDI adapter ships a pure-TypeScript XML tokenizer (`src/mspdi/xml-parser.ts`) — NO external XML library, NO `DOMParser`, NO Node `fs` (architecture-lock §13). No MSPDI export surface exists (PROJECT-016 is unauthorized); the work-item dependency `014 → 015 → 016` is preserved.


## Package dependency edges (PROJECT-016)

PROJECT-016 (MSPDI export) reuses the accepted package boundary — the exporter lives in `packages/project-file/src/mspdi/**` beside the importer, behind the same `MspdiFileAdapter`/`ProjectFileAdapter` surface (now `inspect` + `import` + `export`). Static package dependencies are unchanged:

```text
@genoffice/project-file → @genoffice/project-engine → @genoffice/project-contracts
@genoffice/project-file → @genoffice/project-contracts (direct, for types + brand helpers)
```

`@genoffice/project-file` STILL does NOT statically depend on `@genoffice/project-scheduling` — the exporter delegates validation to `validateProjectDocument` (`@genoffice/project-engine`) and never consults `DerivedSchedule` or runs a scheduler (a static source guard in the test suite asserts the absence of the scheduling import). The MSPDI XML writer is a pure-TypeScript serializer (`src/mspdi/xml-writer.ts`) — NO external XML library, NO `DOMParser`/`XMLSerializer`, NO Node `fs` (architecture-lock §13). The PROJECT-015 test-suite precedent applies: tests import `schedule`/`resolveCalendar` from `@genoffice/project-scheduling` to prove round-tripped documents schedule identically, but the package itself does not. The work-item chain `014 → 015 → 016` is complete; `016 → 017` (MPP feasibility) remains the next edge.
