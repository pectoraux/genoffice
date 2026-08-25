# GenOffice Project — Requirements

## R-001 Canonical project model
The product SHALL maintain one canonical `ProjectDocument` independent of host/runtime and shall give every entity stable identity.

## R-002 Task semantics
Tasks SHALL represent hierarchy, scheduling mode, dates, duration, constraints, deadlines, priority, calendar, progress, work/cost fields, baselines, custom fields, parent identity, notes, and derived scheduling fields.

## R-003 Dependencies
The engine SHALL support FS, SS, FF, and SF relationships with positive/negative lag and SHALL reject cycles, self dependencies, missing task references, and missing calendar references.

## R-004 Calendar determinism
The engine SHALL support weekly working periods, working hours, holidays/exceptions, inherited calendars, task calendars, and deterministic working-time arithmetic.

## R-005 Scheduling determinism
Given identical canonical inputs, scheduling SHALL return identical task dates, dependency effects, early/late dates, criticality, and float on every run.

## R-006 Semantic commands
User intent SHALL be represented by semantic `ProjectCommand` values. The history model SHALL be capable of deterministic undo/redo without encoding Project semantics in the renderer.

## R-007 Resource-ready model
The canonical model SHALL contain resources and assignments even though resource calculations are deferred beyond PROJECT-006.

## R-008 File interoperability boundary
The architecture SHALL separate canonical state from `.gproj`, MSPDI, and MPP adapters. Unsupported imported features SHALL eventually produce explicit diagnostics rather than silent loss.

## R-009 Shared renderer boundary
Desktop and web SHALL use the same Project renderer contracts and semantic engine. Host-specific APIs SHALL remain outside the renderer.

## R-010 Verification authority
Each feature SHALL have objective acceptance evidence. UI appearance alone SHALL never establish completion. Save/reopen and canonical-engine re-evaluation are required for persisted features once file support exists.

## Foundation acceptance — PROJECT-001..006
A fixture containing tasks, calendars, dependencies, and deterministic scheduling inputs SHALL be validated and scheduled repeatedly with identical results. The result SHALL expose task start/finish, early/late dates, total/free float, and criticality, and the dependency engine SHALL reject invalid graphs.
