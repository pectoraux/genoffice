# GenOffice Project — Architecture Lock

Status: FROZEN
Scope: PROJECT-001 through PROJECT-006 foundation
Authority: This document is the controlling architectural contract for the Project initiative.

## 1. Mission

GenOffice Project is a Microsoft Project-class planning application with one host-neutral semantic model shared by desktop and web. The canonical product model is independent of Electron, React, browser APIs, HTTP, and filesystem APIs.

## 2. Layer order

Intent/UI → semantic ProjectCommand → Project Engine → ProjectDocument → Scheduling Engine → deterministic DerivedSchedule → host-specific transport/file integrations.

A lower layer may consume contracts from a higher layer but may never redefine a higher-layer concept. Rendering is never authoritative for scheduling, identity, persistence, or Project semantics.

## 3. Canonical package boundaries

- `packages/project-contracts`: stable domain and integration contracts; no host dependencies.
- `packages/project-engine`: domain validation, semantic commands, journal model; no UI/host dependencies.
- `packages/project-scheduling`: calendar primitives, dependency graph, deterministic scheduling; no UI/host dependencies.
- `packages/project-file`: file adapter boundary until PROJECT-014+; no renderer/parser leakage.
- `packages/project-renderer-core`: shared renderer boundary until PROJECT-021+; no scheduling authority.
- Project desktop host code will be isolated under a Project-specific app surface.
- Project web code will use shared renderer contracts and host transport; it will not import Electron or Node APIs.

## 4. Identity

Every Project entity has stable identity. Array position is never identity. UIDs are persistent source/interoperability identifiers; local IDs are canonical application identities.

## 5. Canonical time model

Foundation scheduling uses ISO-8601 UTC timestamps and integer working-minute durations. Calendar periods are represented as minute offsets inside a day. This removes host locale drift from deterministic scheduling. Time-zone/DST fidelity is an advanced-calendar concern and requires an architecture-change proposal before changing this invariant.

## 6. Scheduling authority

`schedule(projectDocument, schedulingOptions)` is the sole authoritative scheduling operation. It is a pure deterministic function from canonical inputs to `DerivedSchedule` or explicit diagnostics. It does not mutate the source ProjectDocument.

## 7. Dependency semantics

Supported relationship types are FS, SS, FF, and SF with integer working-minute lag/lead. Self-links, missing references, and cycles are invalid.

## 8. Calendar semantics

Calendars explicitly define weekly working periods and date exceptions. Calendar inheritance is resolved before scheduling. The scheduling engine provides deterministic `isWorking`, `addWorkingTime`, `subtractWorkingTime`, and `workingDuration` primitives.

## 9. Commands and history

UI mutations must be represented as semantic `ProjectCommand` values. Undo/redo is command-history based. A renderer may not invent Project semantics from raw state mutation.

## 10. Files

The canonical model is not MPP/MSPDI. File adapters translate to/from the canonical model. File parsing and XML construction never live in React.

## 11. Renderer rule

Renderer state may cache projections of canonical state, but may not own authoritative task dates, dependencies, critical path, float, resource leveling results, or persisted Project semantics.

## 12. Architecture changes

Any change to a frozen invariant requires a recorded architecture-change proposal containing motivation, affected invariants, alternatives, compatibility impact, migration plan, verification impact, and explicit authority approval. Until accepted, the proposal is not implemented.

## 13. Forbidden dependencies

Foundation packages must not import React/React DOM, Electron, Node filesystem/process APIs, browser globals, HTTP clients/server route modules, or Excel renderer packages.

Foundation semantic/runtime packages (`project-contracts`, `project-engine`, `project-scheduling`, `project-renderer-core`) must not import external MSPDI/MPP parser implementations.

`packages/project-file` is the sanctioned file-adapter boundary and may contain format-specific parser/serializer implementations.

No `project-engine`, `project-scheduling`, `project-renderer-core`, or host package may directly import format-specific parser internals.

File-format implementations remain behind the `project-file` adapter boundary.

(Clarified by ACR-001 — `spec/project/architecture-changes/ACR-001-project-file-adapter-boundary.md` — the recorded architecture-change proposal under §12; no prohibition is weakened.)

## 14. Completion authority

Agent claims are not evidence. A work item is complete only when repository state, automated tests, fixtures, static architecture checks, and required host/file evidence satisfy its verification matrix entry.
