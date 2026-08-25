# GenOffice Project — Architecture

## Runtime topology

```text
                 Project Renderer Core
                          │
            ┌─────────────┴─────────────┐
            │                           │
       Electron Host                Web Host
            │                           │
      Native transport             HTTP transport
            │                           │
            └─────────────┬─────────────┘
                          │
                    Project Engine
                          │
           ┌──────────────┼──────────────┐
           │              │              │
      Domain Engine   Scheduling      File Engine
                          │              │
           └──────────────┴──────────────┘
                          │
              ProjectDocument canonical model
```

## Package responsibilities

### `@genoffice/project-contracts`

Defines stable, host-neutral identities, entities, enums, Project commands, schedule outputs, file-plan metadata, and diagnostics.

### `@genoffice/project-engine`

Owns document validation, semantic command types/journal behavior, and future domain mutations. It is the semantic boundary between UI intent and canonical state.

### `@genoffice/project-scheduling`

Owns deterministic calendar calculations, dependency validation/topology, and the scheduling/critical-path pass. It consumes contracts and emits derived schedule state without mutating the document.

### `@genoffice/project-file`

Owns file-adapter interfaces and, in later increments, `.gproj`, MSPDI, and MPP translators. It never becomes renderer state. PROJECT-001 reserves the boundary; PROJECT-014+ implements formats.

### `@genoffice/project-renderer-core`

Owns shared Project view composition after the semantic foundation passes. Its purpose is projection and interaction, not schedule calculation. PROJECT-001 reserves the boundary; PROJECT-021+ implements it.

## Domain graph

```text
ProjectDocument
 ├── properties
 ├── tasks ────────┐
 ├── dependencies ─┤
 ├── calendars ────┤──→ scheduling engine → DerivedSchedule
 ├── resources ────┤
 ├── assignments ──┘
 ├── baselines
 └── view metadata
```

## Scheduling pipeline

1. Validate identities/references and the dependency graph.
2. Resolve inherited calendars.
3. Build a stable topological order using canonical task IDs as deterministic tie-breakers.
4. Execute the forward pass for early dates.
5. Derive summary-task dates from direct children.
6. Execute the backward pass from project finish.
7. Derive total/free float and criticality.
8. Emit immutable `DerivedSchedule` and diagnostics; never mutate the source document.

## Determinism contract

Foundation scheduling uses ISO-8601 UTC timestamps and integer working-minute durations. The same serialized ProjectDocument and options must produce equivalent derived values independent of locale, wall clock, or host runtime.

## Current foundation scope

PROJECT-001..006 intentionally stop before renderer implementation, file-format implementation, resource leveling, and advanced Project behavior. Placeholder package boundaries exist only to prevent later architectural drift.
