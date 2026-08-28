# GenOffice Project — Dependency Graph

## Canonical direct-dependency graph

Each line lists one work item (in PROJECT-number order — the authoritative roadmap order) and its complete set of DIRECT dependencies, exactly as recorded in the `Depends on` column of `work-items.md`. The former foundation-tree and chain-shorthand blocks are superseded by this explicit form (roadmap reconciliation increment — same direct dependencies, no roadmap reorder; the deterministic consistency guard `packages/project-file/tests/spec-consistency.test.ts` enforces the correspondence between the two documents).

```text
PROJECT-001 ← none
PROJECT-002 ← 001
PROJECT-003 ← 001, 002
PROJECT-004 ← 001, 002
PROJECT-005 ← 001, 002
PROJECT-006 ← 002, 004, 005
PROJECT-007 ← 002, 006
PROJECT-008 ← 006
PROJECT-009 ← 006
PROJECT-010 ← 002, 004, 006
PROJECT-011 ← 010
PROJECT-012 ← 006
PROJECT-013 ← 011, 012
PROJECT-014 ← 002, 003
PROJECT-015 ← 014
PROJECT-016 ← 014, 015
PROJECT-017 ← 014, 015
PROJECT-018 ← 017
PROJECT-019 ← 017
PROJECT-020 ← 015, 018
PROJECT-021 ← 002, 003, 006
PROJECT-022 ← 021
PROJECT-023 ← 022
PROJECT-024 ← 023, 005
PROJECT-025 ← 022, 004
PROJECT-026 ← 022, 012
PROJECT-027 ← 021, 022
PROJECT-028 ← 021, 022
PROJECT-029 ← 021, 023
PROJECT-030 ← 021, 023
PROJECT-031 ← 022, 023, 029
PROJECT-032 ← 022, 023
PROJECT-033 ← 010, 011, 022
PROJECT-034 ← 005, 021
PROJECT-035 ← 009, 022
PROJECT-036 ← 002, 003, 022
PROJECT-037 ← 021, 036
PROJECT-038 ← 009, 036
PROJECT-039 ← 008, 009, 011
PROJECT-040 ← 039
PROJECT-041 ← 007, 008
PROJECT-042 ← 007, 014
PROJECT-043 ← 010, 013
PROJECT-044 ← 004, 010
PROJECT-045 ← 013, 043
PROJECT-046 ← 027, 028, 029, 030, 031
PROJECT-047 ← 014, 015, 016, 018, 019, 020
PROJECT-048 ← 046, 047
PROJECT-049 ← 046, 047, 048
```

## Accepted frontier

The objectively accepted state and the next authorized product increment (the authorization gate below applies):

```text
Accepted: PROJECT-001..PROJECT-027 (PROJECT-019 completed as the 019A rescope decision record)
Next authorized: PROJECT-028
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

## Package dependency edges (PROJECT-023)

PROJECT-023 (Selection / editing) adds NO new workspace package and NO new package edge. The engine command additions (`SetTaskDuration`, `SetTaskFinish`) live inside the accepted `@genoffice/project-engine` — two members of the FROZEN command union whose implementation PROJECT-021 documented as deferred to this increment; the contracts package, the command union, and every existing edge are unchanged:

```text
@genoffice/project-renderer-core → @genoffice/project-contracts (types + brand helpers)
@genoffice/project-renderer-core → @genoffice/project-engine (applyProjectCommand, hierarchy utilities)
@genoffice/project-scheduling → @genoffice/project-engine (document validation + leveler registration, unchanged)
```

The editing surface lives INSIDE `@genoffice/project-renderer-core` (`src/editing.ts` — the pure edit model; `src/edit-flow.ts` — the session commit flow; the reducer/state/intent extensions), consuming only what the package already depends on. The editing modules never import the scheduling package (no scheduling authority — lock §3/§6; the draft reads the projection's by-reference schedule join passed as arguments), never the file/host packages (R-009), and never React/Electron/Node/DOM APIs (lock §13 — the discipline suite's raw-source scan covers them and now additionally guards the editing modules against scheduling-derived state, lock §11). CI is unchanged: the 21-step foundation gate already typechecks/tests both touched packages and its boundary grep covers them recursively. The work-item edge `023 → 024` opens on acceptance.

## Package dependency edges (PROJECT-024)

PROJECT-024 (Dependency editing) adds NO new workspace package and NO new package edge. The four engine command implementations (`AddDependency`, `RemoveDependency`, `ChangeDependencyType`, `ChangeLag`) live inside the accepted `@genoffice/project-engine` — the remaining members of the FROZEN command union whose implementation PROJECT-021 documented as deferred to this increment; the contracts package, the command union, and every existing edge are unchanged:

```text
@genoffice/project-renderer-core → @genoffice/project-contracts (types + brand helpers)
@genoffice/project-renderer-core → @genoffice/project-engine (applyProjectCommand, hierarchy utilities)
@genoffice/project-scheduling → @genoffice/project-engine (document validation + leveler registration, unchanged)
```

The dependency-editing surface lives INSIDE `@genoffice/project-renderer-core` (`src/dependency-editing.ts` — the pure edit model mirroring `src/editing.ts`; the builder additions in `src/commands.ts`; the `commitDependencyEditThroughSession` addition in `src/edit-flow.ts`; the reducer/state/intent extensions; the link reflection in `src/views/dependencies.ts` threaded through `src/views/timeline.ts` from `buildGanttView`'s existing `state` parameter), consuming only what the package already depends on. The dependency modules never import the scheduling package (no scheduling authority — lock §3/§6; the scheduling effects are asserted at the TEST layer against the injected REAL scheduler, the accepted precedent), never the file/host packages (R-009), and never React/Electron/Node/DOM APIs (lock §13 — the discipline suite's raw-source scan covers them and the lock §11 scheduling-free scan now includes `dependency-editing.ts`). CI is unchanged: the existing foundation gate already typechecks/tests both touched packages and its boundary grep covers them recursively. The work-item edge `023 → 024` completes on acceptance; `024 → 025` (calendar visualization) remains gated on the dependency graph's ordering.

## Package dependency edges (PROJECT-025)

PROJECT-025 (Calendar visualization) adds NO new workspace package and NO new package edge. The entire calendar projection lives INSIDE `@genoffice/project-renderer-core` (`src/calendar.ts` — the injected `CalendarWorkingTimeQuery` structural type, the catalog echo, the band classification, the surface builder; `src/views/timeline.ts` — the additive background + per-row surfaces and `ProjectRowCalendar`; `src/views/gantt-view.ts` — the `CalendarViewInput` threading); the contracts package, the command union, the scheduling package, and every existing edge are unchanged:

```text
@genoffice/project-renderer-core → @genoffice/project-contracts (types + brand helpers)
@genoffice/project-renderer-core → @genoffice/project-engine (applyProjectCommand, hierarchy utilities)
@genoffice/project-scheduling → @genoffice/project-engine (document validation + leveler registration, unchanged)
```

The calendar modules never import the scheduling package in production code (no scheduling authority — lock §3/§6; there is NO second calendar engine: every working-time answer comes from the injected `CalendarWorkingTimeQuery`, satisfied by the host-side canonical binding `workingIntervals ∘ resolveCalendar` — the `ScheduleRunner` injection precedent — and asserted at the TEST layer against the REAL scheduling package, the accepted precedent), never the file/host packages (R-009), and never React/Electron/Node/DOM APIs (lock §13 — the discipline suite's raw-source scan covers them, and the NEW calendar-engine guard scans `calendar.ts` for date-decomposition/exception-keying markers so no calendar evaluation can be re-implemented there). The per-build surface memo never persists across renders (lock §11). CI is unchanged: the existing foundation gate already typechecks/tests the package and its boundary grep covers it recursively. The work-item edge `004 + 022 → 025` completes on acceptance; `022 → 026`, `022 → 027`, `022 → 028` remain open.

## Package dependency edges (PROJECT-026)

PROJECT-026 (Critical-path/resource visualization) adds NO new workspace package and NO new package edge. The canonical side lives INSIDE the accepted `@genoffice/project-scheduling` (the NEW shared canonical allocation kernel `src/allocation-kernel.ts` — the ONE implementation of the resource demand/capacity semantics: the demand-contribution rule, the demand-interval construction, the resource-calendar resolution, the availability-capacity resolution, the calendar-aware segmentation, and the over-allocation predicate — consumed by TWO projections of the same authority: the NEW `src/allocation.ts` with the pure read-only `resourceAllocations(document, schedule)` time-phased tiling export, and the accepted `src/leveling.ts` whose `levelResources()` collapses the kernel's consecutive over-allocated runs into its conflict windows with byte-identical PROJECT-013 observable behavior — the single-authority correction from the Principal Architect's review of PR #28, so no second capacity-engine semantics can drift inside the scheduling package itself; the scheduling architecture suite fails any module that re-implements the kernel primitives or any consumer that stops consuming the kernel; `schedule()` and every other module are unchanged); the visualization side lives INSIDE `@genoffice/project-renderer-core` (`src/critical-path.ts` — the pure DerivedSchedule critical/float echo with the slack-bar geometry and the both-endpoints critical-link classification; `src/resources.ts` — the injected `ResourceAllocationQuery` structural type, the verbatim-clipped utilization band projection, and the 025-mirroring degradation boundary; `src/views/timeline.ts` — the additive `criticalPath` + `resourceUtilization` surfaces; `src/views/gantt-view.ts` — the `ResourceViewInput` threading); the contracts package, the command union, the engine package, and every existing edge are unchanged:

```text
@genoffice/project-renderer-core → @genoffice/project-contracts (types + brand helpers)
@genoffice/project-renderer-core → @genoffice/project-engine (applyProjectCommand, hierarchy utilities)
@genoffice/project-scheduling → @genoffice/project-engine (document validation + leveler registration, unchanged)
```

The PROJECT-026 modules never import the scheduling package in production code (no scheduling authority — lock §3/§6; there is NO second CPM engine and NO second capacity engine: every critical/float value is the verbatim `TaskSchedule` echo of the projection's by-reference schedule join — the PROJECT-012 rule — and every demand/capacity/over-allocation value comes from the INJECTED `ResourceAllocationQuery`, satisfied by the host-side canonical binding `resourceAllocations` — the `ScheduleRunner`/`CalendarWorkingTimeQuery` injection precedents — and asserted at the TEST layer against the REAL scheduling package, the accepted precedent), never the file/host packages (R-009), and never React/Electron/Node/DOM APIs (lock §13 — the discipline suite's raw-source scan covers them, and the NEW no-second-engine guards scan `critical-path.ts` for CPM/float computation markers and `resources.ts` for capacity-semantics markers so no critical-path, slack, or resource-capacity algorithm can be re-implemented there). CI is unchanged: the existing foundation gate typechecks/tests both touched packages and its boundary grep covers them recursively. The work-item edge `022 + 012 → 026` completes on acceptance; `022 → 027`, `022 → 028` remain open.

## Package dependency edges (PROJECT-027)

PROJECT-027 (Electron shell) adds ONE new workspace package — the desktop app surface `apps/project` (`@genoffice/project-desktop`), sanctioned by architecture-lock §3 ("Project desktop host code will be isolated under a Project-specific app surface"; an APP, not a foundation package — the `apps/*` workspace root, exactly like `apps/docs`/`apps/sheets` on the office side). It changes NO existing edge; the six-package Project foundation graph is untouched:

```text
@genoffice/project-desktop → @genoffice/project-renderer-core (the shared renderer binding)
@genoffice/project-desktop → @genoffice/project-engine (validation + command application, through the renderer-core session)
@genoffice/project-desktop → @genoffice/project-contracts (types + brand helpers)
@genoffice/project-desktop → @genoffice/project-scheduling (the canonical authorities, INJECTED: schedule/resolveCalendar∘workingIntervals/resourceAllocations)
@genoffice/project-desktop → @genoffice/project-file (the canonical .gproj + MSPDI adapters: open/save)
```

The host's own layering is enforced structurally, not by convention: the MAIN process (`src/main/**`) and PRELOAD (`src/preload/**`) import NO `@genoffice/*` package of any kind — they are native transport only (window/lifecycle, native menu with UNREGISTERED accelerators, dialogs, size-capped raw byte reads/writes through the ONE canonical bounded-read helper `src/main/bounded-read.ts` — both read surfaces (picker + argv/second-instance) share it, `NativeReadResult` values whose error variant carries no bytes, so the renderer can never receive uncapped file contents; the single-instance lock is keyed on the userData path, which is therefore installed BEFORE `requestSingleInstanceLock()` — isolated profiles carry isolated locks; the close-guard handshake); every `ProjectDocument`/byte semantic crossing happens in the renderer process. The renderer's `src/renderer/bindings.ts` is the ONE module importing the scheduling package — the three canonical injection seams the accepted increments define (`ScheduleRunner` PROJECT-021, `CalendarWorkingTimeQuery` PROJECT-025, `ResourceAllocationQuery` PROJECT-026) — and the desktop architecture suite (`apps/project/tests/unit/architecture.test.ts`) fails the build if any other host module imports scheduling, re-implements a working-time/capacity primitive, computes dates, or hand-builds a `ProjectCommand` literal (commands flow exclusively through the renderer-core builders + edit flows), or if either corrected transport invariant regresses (a read bypassing the bounded helper, or the userData path installed after the lock request — both guards negative-validated). The work-item edge `021 + 022 → 027` completes on acceptance; `021 + 022 → 028` (web shell) remains open and unauthorized. MPP import through the Java sidecar (`project-mpp-host`) is deliberately NOT wired into the shell yet — the work item's file surface is the canonical file adapter (`.gproj` open/save + MSPDI interchange through `@genoffice/project-file`); the MPP pipeline is a later host integration.

CI: the Project foundation gate grows 25 → 27 steps (`Typecheck project-desktop`, `Test project-desktop` — the unit/architecture suite) plus the NEW `desktop-e2e` job (ubuntu-22.04, the repo-wide Electron E2E platform pinning: Electron runtime download → `electron-vite build` → AppArmor sysctl best-effort → `xvfb-run` over `npm run test:e2e -w @genoffice/project-desktop`) — the first Project CI stage that runs the product as an application rather than as packages.

## Package dependency edges (PROJECT-028)

PROJECT-028 (Web shell) adds TWO new workspace packages and changes no existing foundation edge. The six-package Project foundation graph is untouched:

```text
@genoffice/project-host → @genoffice/project-renderer-core (the shared renderer binding's own dependency)
@genoffice/project-host → @genoffice/project-contracts (types + brand helpers)
@genoffice/project-host → @genoffice/project-scheduling (the canonical authorities, INJECTED — bindings.ts remains the single import site)
@genoffice/project-host → @genoffice/project-file (the canonical .gproj + MSPDI adapters: open/save)

@genoffice/project-desktop → @genoffice/project-host (the shared host binding — controller/DOM/translation/document flows/scheduling bindings/stylesheet/bridge contract)
@genoffice/project-desktop → @genoffice/project-contracts + @genoffice/project-file + @genoffice/project-renderer-core (test/E2E-layer imports only; the main/preload/IPC surfaces import NO @genoffice/* package, exactly as accepted at PROJECT-027)

@genoffice/project-web → @genoffice/project-host (the SAME shared host binding)
@genoffice/project-web → @genoffice/project-contracts + @genoffice/project-file + @genoffice/project-renderer-core (test/E2E-layer imports only)
```

`@genoffice/project-host` is the shared host binding layer (a HOST-layer package, NOT a foundation package — DOM APIs are its job; Electron/Node APIs are forbidden in it, asserted by its own discipline suite; the office-suite `renderer-bridge`/`platform` precedent for shared host-binding code): the PROJECT-027 desktop renderer binding moved there whole (the controller, the DOM layer, the translation tables, the document flows, the scheduling bindings, the deterministic stylesheet) plus the host-neutral bridge CONTRACT (`src/bridge.ts`: `ProjectHostBridge`, `MenuCommandId`, `MENU_COMMAND_IDS`, `NativeReadResult`, `OpenFileSelection`, `DiscardChoice`, `HostAppInfo`, `PROJECT_FILE_FILTERS` — extracted from the desktop's `src/shared/ipc.ts`, which keeps its accepted self-contained shape so the Electron main-process bundle still imports no `@genoffice/*` package; the two contract definitions are pinned structurally identical at compile time by the desktop architecture suite — the PROJECT-020 structural-typing precedent). The moved controller's public names became their host-neutral truth (`createProjectApp`, `ProjectHostApp`, `ProjectHostBridge`, `HostAppInfo`).

`@genoffice/project-web` is the web app surface (sanctioned by architecture-lock §3's web rule: "Project web code will use shared renderer contracts and host transport; it will not import Electron or Node APIs" — an APP, not a foundation package, the `apps/*` workspace root mirroring the desktop surface): the web transport bridge (`src/web-bridge.ts` — the ONE bounded web read `readCapped` with the size check before any read and `NativeReadResult` values whose error variant carries no bytes, the picker over a real file input, Blob-download saves, the three-button DOM discard dialog, the beforeunload close guard PURELY SYNCHRONOUS over the dirty probe — never initiating the controller's async close handshake (the in-app close request `requestClose` is that handshake's one web firing path), drag-and-drop external opens staged + served through `readFile`), the DOM menu bar (`src/menu.ts` — the shared `MENU_COMMAND_IDS` vocabulary, activation forwarded through the bridge's menu-command path, accelerators displayed but never executed), and the web entry. The browser-side implementation imports NO Node/Electron API — enforced at three layers: the source-level discipline suite, the Vite build (no Node/Electron polyfills — an accidental import fails the build), and the CI bundle scan over `apps/project-web/dist/` before the E2E boots the bundle in real chromium.

The desktop host's own layering is unchanged from PROJECT-027 (main/preload native transport only, both correction-round transport invariants intact — the bounded-read guards and the userData-before-lock ordering are still asserted, still negative-validated by the desktop suite; the desktop renderer surface is now the ENTRY only, importing the shared binding — a new desktop-suite guard pins that the shell never re-implements the controller/DOM/translation surface). The desktop E2E battery (24/24) is the regression proof that the extraction is behavior-transparent. The work-item edge `021 + 022 → 028` completes on acceptance; `021 + 023 → 029` (shared Project ribbon) remains gated on 028's acceptance. MPP import through the Java sidecar remains deliberately unwired in BOTH shells (the work item's file surface is the canonical file adapter; the browser has no sidecar).

CI: the Project foundation gate's foundation job grows by four steps (`Typecheck project-host`, `Test project-host`, `Typecheck project-web`, `Test project-web` — 26 steps total) plus the NEW `web-e2e` job (ubuntu-latest: `npm install --ignore-scripts` → `npx playwright install --with-deps chromium` → `vite build` → the static Node/Electron bundle scan over `apps/project-web/dist/` → `npm run test:e2e -w @genoffice/project-web` driving the real built bundle through `vite preview` in headless chromium) — the second Project CI stage running the product as an application, and the first in a browser.
