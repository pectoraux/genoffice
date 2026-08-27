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

## Package dependency edges (PROJECT-017)

PROJECT-017 (MPP adapter feasibility) is an investigation — it introduces NO new package and NO new package edge. The added artifacts are `spec/project/mpp-feasibility.md` (the report deliverable) and `packages/project-file/tests/mpp-feasibility.test.ts` (a discipline suite importing only `vitest` plus `?raw` module sources — the report, `packages/project-file/package.json`, and `spec/project/architecture-lock.md`; no runtime import of any kind, no `node:`/`fs`/network access, satisfying the same CI boundary grep as every other project-file test).

Static package dependencies remain exactly as accepted since PROJECT-014:

```text
@genoffice/project-file → @genoffice/project-engine → @genoffice/project-contracts
@genoffice/project-file → @genoffice/project-contracts (direct, for types + brand helpers)
```

The spike that grounds the report ran entirely outside the repository (MPXJ 16.7.0 + OpenJDK 21 in a disposable `/tmp` workspace); nothing entered any package. The work-item chain `014 → 015 → 016 → 017` is complete; `017 → 018` (MPP import) remains the next edge — contingent on the Principal Architect accepting the report's recommended strategy (externalized MPXJ sidecar + foundation-level normalization-only adapter contract; see `spec/project/mpp-feasibility.md` §15/§16). MPP export (PROJECT-019) is recommended for rescoping to a diagnostic-only deliverable per report §17.

## Package dependency edges (PROJECT-018)

PROJECT-018 (MPP import) adds ONE new workspace package and changes no existing edge. The foundation four-package graph is untouched:

```text
@genoffice/project-file → @genoffice/project-engine → @genoffice/project-contracts
@genoffice/project-file → @genoffice/project-contracts (direct, for types + brand helpers)
```

The new HOST package `@genoffice/project-mpp-host` (NOT a foundation package — it is the sanctioned location for process code, exactly like `@genoffice/xlsx-gateway`/`platform-electron` are for the Sheets sidecar):

```text
@genoffice/project-mpp-host → @genoffice/project-file → @genoffice/project-engine → @genoffice/project-contracts
@genoffice/project-mpp-host → @genoffice/project-contracts (types)
@genoffice/project-mpp-host → @genoffice/project-scheduling (host-side schedule() of the imported document)
```

The dependency direction is host → foundation ONLY (a static architecture test asserts that `packages/project-file/src/mpp/**` never references the host package and carries no process imports — the CI foundation boundary grep also still covers the whole of `project-file`). The host package uses Node `child_process`/`fs` by design (one-shot MPXJ sidecar launcher); it has NO renderer, Electron, React, or HTTP imports (static guard). Tests import `vitest` + the packages above + the pinned external sidecar artifacts (`.sidecar-deps/`, gitignored, fetched by `scripts/fetch-sidecar-deps.mjs` with SHA-256 verification). The Project CI gate runs setup-java → fetch → typecheck → test for this package (19 steps total). The work-item chain `014 → 015 → 016 → 017 → 018` is complete; `018 → 019` (MPP export) is BLOCKED pending formal rescoping by the Principal Architect — the feasibility report recommends closing it as not-feasible with PROJECT-016's MSPDI export as the sanctioned interchange output.

## Package dependency edges (PROJECT-019)

PROJECT-019 (MPP export strategy / rescope — 019A) is an investigation, exactly like PROJECT-017: it introduces NO new package, NO new package edge, and NO production code. The added artifacts are `spec/project/mpp-export-strategy.md` (the decision-record deliverable) and `packages/project-file/tests/mpp-export-strategy.test.ts` (a discipline suite importing only `vitest` plus `?raw` module sources — the report, `packages/project-file/package.json`, the mpp/mspdi adapter sources, and `spec/project/architecture-lock.md`; no runtime import of any kind, satisfying the same CI boundary grep as every other project-file test).

Static package dependencies remain exactly as accepted since PROJECT-018:

```text
@genoffice/project-file → @genoffice/project-engine → @genoffice/project-contracts
@genoffice/project-file → @genoffice/project-contracts (direct, for types + brand helpers)
@genoffice/project-mpp-host → @genoffice/project-file (+ engine, contracts, scheduling — host-side import pipeline)
```

The work-item chain `014 → 015 → 016 → 017 → 018 → 019` completes with 019A as a DECISION RECORD: MPP export is deferred (outcome E — `.gproj` + MSPDI are the supported write formats; MPP stays import-only). The `019 → 047` edge survives unchanged in shape: PROJECT-047's golden-file compatibility suite charters import/export regression-blocking over the SUPPORTED formats. A hypothetical future 019B (commercial-SDK export sidecar) would add a NEW host-level package beside `project-mpp-host` — a proposed architecture addition documented in the report §14, NOT implemented, and requiring Principal Architect authorization plus host network-isolation hardening beyond Linux before any code exists.

## Package dependency edges (PROJECT-020)

PROJECT-020 (import compatibility diagnostics) adds NO new package and NO new package edge — the compatibility layer lives inside `@genoffice/project-file` (`src/compatibility/**`) and consumes only the accepted packages:

```text
@genoffice/project-file → @genoffice/project-engine → @genoffice/project-contracts
@genoffice/project-file → @genoffice/project-contracts (direct, for types + brand helpers)
@genoffice/project-mpp-host → @genoffice/project-file (+ engine, contracts, scheduling — host-side import pipeline)
```

`@genoffice/project-file` STILL does not statically depend on `@genoffice/project-scheduling`: the compatibility pipelines take the scheduler as an INJECTED, structurally-typed runner (`CompatibilityScheduleRunner` — `schedule` from the scheduling package is assignable without importing it), preserving the accepted edge and its static guards (a discipline test asserts the compatibility sources contain no scheduling-package reference, no `node:`/React/Electron imports, no `localeCompare`, and no clock/randomness). The host package's `importMppFromFileWithCompatibility` composes the full production pipeline with `buildCompatibilityReport` (host → foundation direction only). The work-item edge `015 + 018 → 020` completes; `020 → 047` carries forward with PROJECT-047's golden-file suite building on the canonical compatibility contract this increment defines.

## Package dependency edges (PROJECT-021)

PROJECT-021 (shared renderer core) implements the reserved `@genoffice/project-renderer-core` package boundary (architecture-lock §3). It adds NO new workspace package (the placeholder package exists since PROJECT-001) and changes no existing edge:

```text
@genoffice/project-renderer-core → @genoffice/project-contracts (types + brand helpers)
@genoffice/project-renderer-core → @genoffice/project-engine (applyProjectCommand, hierarchy utilities)
```

The work-item dependency `002 + 003 + 006 → 021` is honored exactly like the accepted project-file precedent: `DerivedSchedule` is a CONTRACTS type (the renderer core consumes schedule values typed by contracts), and the scheduling FUNCTION is injected as a structurally-typed `ScheduleRunner` — the package does NOT statically depend on `@genoffice/project-scheduling` (the strongest form of architecture-lock §3's "no scheduling authority": the package is structurally incapable of scheduling; every schedule value it holds was produced by the injected authority). Tests import `schedule` from the scheduling package at the TEST layer only, proving session/schedule integration with the real authority while the package dependency stays two-edged (asserted by the discipline suite's raw-source scan AND the extended CI foundation boundary grep, which now also covers `packages/project-renderer-core`). The renderer core equally never imports `@genoffice/project-file` / `@genoffice/project-mpp-host` (host-specific file/transport stays outside the renderer — R-009; documents arrive already-canonical). CI grows 19 → 21 steps (`Typecheck project-renderer-core`, `Test project-renderer-core`). The work-item chain `021 → 022 → 023 → 024` opens; complete views remain unauthorized until acceptance of this increment.

## Package dependency edges (PROJECT-022)

PROJECT-022 (Gantt / task grid / timeline views) adds NO new workspace package and NO new package edge — the view models live inside the accepted `@genoffice/project-renderer-core` (`src/views/**`, the PROJECT-020 `src/compatibility/**` in-package precedent), consuming only what the package already depends on:

```text
@genoffice/project-renderer-core → @genoffice/project-contracts (types + brand helpers)
@genoffice/project-renderer-core → @genoffice/project-engine (applyProjectCommand, hierarchy utilities)
```

The views consume the package's OWN accepted layers (the `projectDocumentView` projection, the `TimelineViewport`/`buildTimeAxis` timeline math) plus the canonical contracts types (`ProjectDocument`, `TaskSchedule`, `Dependency`, `ProjectTable`); they never import the scheduling package (no scheduling authority — lock §3/§6), never the file/host packages (R-009), and never React/Electron/Node/DOM APIs (lock §13 — the discipline suite's raw-source scan now globs `src/**/*.ts`, covering the views subfolder, and additionally guards the views against pixel/DOM layout APIs). CI is unchanged: the 21-step foundation gate already typechecks/tests the package and its boundary grep covers the whole package recursively. The work-item edge `021 → 022` completes; `022 → 023`, `022 → 025`, `022 → 026`, `022 → 027`, `022 → 028` open on acceptance.
