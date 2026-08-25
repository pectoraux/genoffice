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

## PROJECT-007 canonical semantic clarifications

These clarifications record the canonical decisions required to implement task/WBS/summary semantics (PROJECT-007). They refine R-001, R-002, and R-006 without altering any frozen invariant; no architecture-change proposal is required.

- **Sibling ordering representation**: the relative order of tasks inside `ProjectDocument.tasks` is the canonical sibling ordering. Array position expresses ordering only and is never identity (identity is `TaskId`) and never hierarchy (hierarchy is `parentTaskId`). This is the smallest architecture-compatible representation: no ordering field is added to the frozen Task model, and because ordering is positional an invalid sibling ordering cannot be represented.
- **WBS derivation**: WBS codes are derived deterministically by depth-first traversal of the canonical hierarchy, numbering siblings from 1 at each level (for example `1`, `1.1`, `1.2`, `2`, `2.1`) to arbitrary depth. WBS is a derived hierarchy representation and never identity.
- **Outline level derivation**: `outlineLevel` equals hierarchy depth (root = 1, child = parent + 1) and is recomputed on every accepted hierarchy mutation. Documents with outline levels inconsistent with depth are invalid.
- **Summary semantics**: the `summary` flag is a derived semantic value — a task is a summary iff it has at least one child. Summary dates and duration are computed by the scheduling engine roll-up and are never authoritative in renderer state or in document-level duration fields.
- **Task UID uniqueness**: task `uid` values are unique within a document because they are persistent interoperability identifiers; local `TaskId` remains the canonical application identity.
- **CreateTask**: inserts the task as the last child of `parentTaskId`, or as the last root task when no parent is given. Derived fields (`outlineLevel`, `wbs`, `summary`) are recomputed by the engine; a created task is always a leaf.
- **DeleteTask**: deleting a task deletes its entire descendant subtree (the Microsoft Project outline-deletion behavior) and removes dependencies, assignments, and baseline snapshots that reference deleted tasks, so every accepted mutation leaves the document valid.
- **IndentTask**: accepted only when `parentTaskId` is the task's immediately preceding sibling; the task becomes that sibling's last child. Indenting beneath a descendant, a missing target, the first task with no preceding sibling, or any other target is rejected with a diagnostic.
- **OutdentTask**: the task (with its own subtree) becomes the next sibling of its former parent. Outdenting a root task is rejected. A former parent with no remaining children reverts to a leaf.
- **Command application**: `applyProjectCommand` is a pure deterministic function. Rejected commands return the input document unchanged with diagnostics; accepted commands leave the document in canonical outline order with recomputed derived hierarchy fields. The PROJECT-006 invariant that dependencies between a summary task and its own descendants are invalid is preserved across hierarchy mutations.

## PROJECT-008 canonical semantic clarifications

These clarifications record the canonical decisions required to implement constraints, deadlines, and progress (PROJECT-008). They refine R-002 and R-006 without altering any frozen invariant; no architecture-change proposal is required.

### Constraint set

The canonical constraint set is exactly eight types — `asSoonAsPossible`, `asLateAsPossible`, `startNoEarlierThan`, `startNoLaterThan`, `mustStartOn`, `finishNoEarlierThan`, `finishNoLaterThan`, `mustFinishOn` — with explicit deterministic semantics. No aliases exist and no constraint is silently reinterpreted as another.

### Constraint requirements

- The six date-bounded types (`startNoEarlierThan`, `startNoLaterThan`, `mustStartOn`, `finishNoEarlierThan`, `finishNoLaterThan`, `mustFinishOn`) SHALL carry a valid `constraintDate`. A missing or malformed date is rejected with `MISSING_CONSTRAINT_DATE` / `INVALID_DATE`.
- `asSoonAsPossible` and `asLateAsPossible` SHALL NOT carry a `constraintDate`. A date supplied alongside either is rejected with `CONSTRAINT_DATE_NOT_ALLOWED` rather than silently honored or dropped.

### Constraint scheduling semantics (per type)

- `asSoonAsPossible`: the default no-constraint mode. Early dates come from dependencies/project start; scheduled dates equal early dates.
- `asLateAsPossible`: forward pass is dependency-driven; scheduled dates are pulled to the late window (bounded only by the project finish and successors). Carries no constraintDate.
- `startNoEarlierThan`: the forward pass pushes the early start to `max(candidate, constraintDate)`; the backward pass is unconstrained by the date. Scheduled dates equal early dates. Differs from MSO because it leaves total slack intact.
- `startNoLaterThan`: the forward pass is dependency-driven; the backward pull bounds the late start to `min(lateStart, constraintDate)`. Scheduled dates are pulled to the late window.
- `mustStartOn`: pins both early and late start to the constraintDate. This collapses total slack to zero (the task becomes critical), which is what separates MSO from SNET.
- `finishNoEarlierThan`: the forward pass pushes the early finish to `>= constraintDate` (adjusting the start backward across working time); the backward pass is unconstrained. Scheduled dates equal early dates. Differs from MFO because it leaves total slack intact.
- `finishNoLaterThan`: the forward pass is dependency-driven; the backward pull bounds the late finish to `min(lateFinish, constraintDate)`. Scheduled dates are pulled to the late window.
- `mustFinishOn`: pins both early and late finish to the constraintDate. This collapses total slack to zero (the task becomes critical), which is what separates MFO from FNET.

### Constraint interactions

- With dependencies: a constraint is applied after dependency-driven candidate dates. A hard constraint (MSO/MFO) MAY honor its date even when a dependency would place the task elsewhere; the engine resolves this deterministically and surfaces the conflict via the resulting dates and slack rather than by rejecting the schedule.
- With duration/milestones: zero-duration tasks honor the same constraint semantics; a milestone's start and finish coincide at the constraint instant.
- With summaries: constraints on a summary task do not override the child-rolled-up dates; summary scheduled dates remain the roll-up of their children.
- With the project finish: late dates are bounded by the project finish so that total slack always measures "slip without extending the project."

### Deadlines

A deadline is NOT a scheduling constraint. Setting a deadline on a task SHALL NOT move that task's dates. The scheduling engine derives deterministic deadline state for downstream reporting:

- `deadline`: the echoed deadline instant.
- `deadlineVariance`: signed working-minute span from `scheduledFinish` to the deadline, computed in the task's resolved calendar. Positive when the task finishes before the deadline (ahead/on time); negative when the task finishes after the deadline (missed); zero when they coincide.
- `deadlineMissed`: true when `scheduledFinish` is strictly after the deadline.

Deadlines apply identically to leaf tasks, milestones (compared against the milestone's single scheduled instant), and summary tasks (compared against the rolled-up finish).

### Progress model

- `percentComplete` is an integer-equivalent percentage in `[0, 100]`. Values outside the range, non-finite values, or values on summary tasks set via `SetPercentComplete` are rejected. The engine does not treat `percentComplete` as a renderer-only display field; it is a canonical input to derived progress state.
- `physicalPercentComplete` is an optional leaf-only percentage in `[0, 100]`; it is echoed (not aggregated) and is never authoritative for summary progress.
- The canonical model distinguishes planned duration, accomplished work (`actualDuration`), and remaining work (`remainingDuration`) without implementing the full resource/work engine (PROJECT-011). For a leaf task `actualDuration = round(duration * percentComplete / 100)` and `remainingDuration = duration - actualDuration`.

### Progress derivation and status precedence

The engine derives a `status` of `notStarted`, `inProgress`, or `complete` with explicit precedence (no silent behavior):

1. `percentComplete >= 100` → `complete` (even when the scheduled finish is in the future).
2. `0 < percentComplete < 100` → `inProgress`.
3. `percentComplete == 0`:
   - a project `statusDate` is set AND has reached/passed the scheduled start → `inProgress` (work should have begun but is not yet reported complete);
   - otherwise → `notStarted`.

Milestones are zero-duration binary events: `complete` at `percentComplete == 100`, otherwise `notStarted` (the in-progress window is empty for a zero-duration task).

### Summary progress

A summary task's progress is a duration-weighted roll-up of its subtree, NOT a copy of any single child. The engine sums each child's derived `actualDuration` and `remainingDuration` and computes `percentComplete = round(actual / (actual + remaining) * 100)` (or `100` when all children are complete and the subtree work is zero, else `0`). Summary status follows the rolled-up percentage: `complete` at 100, `notStarted` at 0 only when every child is `notStarted`, otherwise `inProgress`. Resource-weighted summary progress is explicitly deferred to PROJECT-011 and is not implemented here. A summary's stored `percentComplete` is never authoritative; `SetPercentComplete` on a summary is rejected with `SUMMARY_PROGRESS_NOT_SETTABLE`.

### Status date

Status evaluation uses `ProjectProperties.statusDate` only. Wall-clock "today", local timezone, and locale never enter the scheduling engine. The same serialized `ProjectDocument` plus scheduling options SHALL produce byte-identical `DerivedSchedule` output across repeated executions.

### Semantic commands

PROJECT-008 implements the `SetConstraint`, `SetPercentComplete`, and `SetDeadline` commands compatibly within the frozen `ProjectCommand` model. Every accepted command is deterministic, preserves stable identity, leaves the document valid, produces canonical diagnostics, and preserves unrelated hierarchy state. Every rejected command leaves the input document unchanged and returns deterministic diagnostics. `SetDeadline` was added as a compatible union member because the frozen contract had no deadline command; it does not alter any existing command shape.
