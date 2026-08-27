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

## PROJECT-009 canonical semantic clarifications

These clarifications record the canonical decisions required to implement baselines and status-date (PROJECT-009). They refine R-002, R-005, and R-006 without altering any frozen invariant; no architecture-change proposal is required.

### Baseline model

A baseline is an immutable snapshot of task `start`, `finish`, `duration`, `work`, and `cost` captured at a point in time. The frozen `Baseline` type already carried `id`, `name`, `capturedAt`, and `taskSnapshots`; PROJECT-009 makes the snapshot authoritative and adds the comparison projection. A document MAY carry multiple independent baselines (the `ProjectDocument.baselines` array), and every `Task.baseline` array is the canonical reverse index of which baselines track that task.

### Baseline capture

`CreateBaseline` carries a fully-formed `Baseline` value: the command stores it as-is and lets document validation enforce the canonical rules. No scheduling is performed inside the command mutator — building the snapshot from the current `DerivedSchedule` is the scheduling package's job (`captureBaseline`). This separation keeps the command pure and deterministic (the same document + the same `CreateBaseline` command always produce the same resulting document bytes). The mutator adds the baseline id to every task that has a snapshot, preserving the bidirectional invariant `task.baseline ⊆ document.baselines`.

### Baseline immutability through hierarchy mutations

Baselines are keyed by stable `TaskId`, and `IndentTask`/`OutdentTask`/`RenameTask` never change `TaskId`, so snapshots survive those mutations unchanged. `DeleteTask` already prunes dangling snapshots for deleted tasks (the Microsoft Project outline-deletion behavior). No other command mutates baseline state, so every accepted hierarchy mutation leaves the document's baseline state valid and the captured dates immutable.

### Baseline comparison (variance with explicit sign convention)

`compareBaseline` projects the current `DerivedSchedule` against a single baseline's immutable snapshots and emits per-task `BaselineVariance`. The sign convention is explicit and mirrors the Microsoft Project "Variance" table:

- `startVariance` and `finishVariance` are signed working-minute spans computed in the task's resolved calendar as `signedWorkingDuration(baseline, current)`: **positive when the current date is later than the baseline (the task slipped past its planned date)**; **negative when the current date is earlier (ahead of plan)**; zero when the dates coincide.
- `durationVariance` is `currentDuration - baselineDuration` (plain signed working-minutes): positive when the current task is longer than planned, negative when shorter, zero when equal.

Both `startVariance` and `finishVariance` are `undefined` when either the baseline snapshot or the current schedule lacks the corresponding date (a baseline captured before a task was scheduled, or a summary whose baseline has no finish). `durationVariance` is always defined because duration is always present in both the snapshot and the derived schedule. Tasks that the baseline did not snapshot are omitted from the comparison (a baseline only reports variance for tasks it captured).

### Status-date semantics

Baseline capture and comparison are deterministic and use `ProjectProperties.statusDate` only. Wall-clock "today", local timezone, and locale never enter the baseline engine. `captureBaseline` defaults the `capturedAt` instant to `ProjectProperties.statusDate` when an explicit override is not supplied; it never falls back to `Date.now()` (that would break byte-identical determinism). A baseline cannot be captured without a deterministic instant: when neither an explicit `capturedAt` nor a project status date is set, `captureBaseline` returns `undefined` so the caller surfaces a clean diagnostic instead of silently inventing a timestamp. The same serialized `ProjectDocument` plus the same baseline options SHALL produce byte-identical `BaselineComparison` output across repeated executions.

### Semantic command

PROJECT-009 implements the `CreateBaseline` command compatibly within the frozen `ProjectCommand` model (the union member already existed; the executor now accepts it). Every accepted command is deterministic, preserves stable identity, leaves the document valid, produces canonical diagnostics, and preserves unrelated hierarchy and baseline state. Every rejected command leaves the input document unchanged and returns deterministic diagnostics. No inverse is provided for `CreateBaseline`: a future `DeleteBaseline` command (PROJECT-038 territory: multiple-baseline management) is required to undo a capture cleanly, mirroring the existing `OutdentTask` precedent of leaving an undefined inverse when a single inverse command cannot restore prior state.

## PROJECT-010 canonical semantic clarifications

These clarifications record the canonical decisions required to implement resources and calendars (PROJECT-010). They refine R-004, R-005, R-006, and R-007 without altering any frozen invariant; no architecture-change proposal is required.

### Resource model

The frozen `Resource` contract already carried `id`, `uid`, `name`, `kind`, `maxUnits`, `standardRate`, `overtimeRate`, `costPerUse`, `calendarId`, and `availability`. PROJECT-010 makes these scheduling inputs authoritative and validates them. `ResourceId` is the canonical identity (never array position); `uid` is a unique persistent interoperability identifier mirroring the Task `uid` rule. A resource name is required (a resource with no name carries no scheduling meaning and is rejected with `MISSING_RESOURCE_NAME`).

### Resource type semantics

The canonical model supports exactly three resource kinds with explicit, documented semantics:

- **Work resources** represent people/equipment capacity in units. `maxUnits`, `standardRate`, `overtimeRate`, `costPerUse`, and `calendarId` are all meaningful scheduling inputs.
- **Material resources** represent consumable materials. `standardRate` and `costPerUse` are meaningful; the engine does not invent work-capacity semantics for materials (`maxUnits` is not a work-capacity bound).
- **Cost resources** represent pure cost categories. They must not silently become work-capacity resources: the scheduling engine treats a cost resource as carrying zero work capacity regardless of its stored `maxUnits`. `costPerUse` is the meaningful field.

`AssignmentSchedule.resourceType` echoes the kind so downstream layers never reinterpret a cost/material resource as a work-capacity resource.

### Numeric validation

Numeric scheduling inputs are never silently coerced. A non-finite or negative `maxUnits` is rejected with `INVALID_MAX_UNITS`; a negative `standardRate`/`overtimeRate` is rejected with `INVALID_RATE`; a negative `costPerUse` is rejected with `INVALID_COST_PER_USE`; a non-finite/negative assignment `units` is rejected with `INVALID_ASSIGNMENT_UNITS`. An invalid resource kind is rejected with `INVALID_RESOURCE_KIND`.

### Availability windows

A resource may carry availability windows (`{ start, finish?, units }`). `start` is always required; `finish`, when present, must be strictly after `start` (a zero-length or inverted window is rejected with `INVALID_AVAILABILITY_RANGE`, not silently dropped). `units` must be finite and non-negative (`INVALID_AVAILABILITY_UNITS`). Overlapping windows are NOT silently merged — the document accepts overlaps so callers can model shift patterns, and resolution ordering is deterministic in the scheduling engine (sorted by start). Resource availability is independent of array order.

### Calendar precedence

The canonical calendar precedence is explicit, deterministic, and never renderer-driven:

1. **Task calendar** — `Task.calendarId` governs task scheduling (the PROJECT-006 frozen behavior, unchanged). `TaskSchedule.resolvedCalendarId` exposes `task.calendarId ?? properties.defaultCalendarId`.
2. **Resource calendar** — `Resource.calendarId` is a scheduling INPUT exposed on `AssignmentSchedule.resolvedCalendarId` as `resource.calendarId ?? properties.defaultCalendarId`. It is independent of the task's resolved calendar.
3. **Project default calendar** — `ProjectProperties.defaultCalendarId` is the fallback for both.

PROJECT-010 establishes resource calendars as resolvable scheduling inputs. It does NOT authorize rewriting the accepted task-calendar scheduling semantics from PROJECT-006: resource calendars do not move task dates in this increment. The smallest deterministic integration exposes the resolved resource calendar id on the derived `AssignmentSchedule` and the resolved task calendar id on `TaskSchedule.resolvedCalendarId`. Calendar intersection and resource-calendar-driven task-date movement are PROJECT-011 work/cost territory.

The existing calendar engine (`resolveCalendar`, `isWorking`, `addWorkingTime`, `subtractWorkingTime`, `workingDuration`) is reused; no second `ResourceCalendarEngine` is created. Resource calendar inheritance and exceptions are resolved by the same accepted primitives.

### Assignments

The frozen `Assignment` contract already carried `id`, `taskId`, `resourceId`, `units`, and work/cost fields. PROJECT-010 makes the assignment reference authoritative and validates it. `AssignmentId` is the canonical identity (never array position). A duplicate `AssignmentId` is rejected with `DUPLICATE_ASSIGNMENT_ID`. A missing task/resource reference is rejected with `MISSING_TASK_REFERENCE`/`MISSING_RESOURCE_REFERENCE`. Two assignment rows cannot silently shadow the same task/resource relationship: a duplicate `(taskId, resourceId)` pair is rejected with `DUPLICATE_ASSIGNMENT_PAIR`.

PROJECT-010 does NOT authorize full assignment work/cost calculations (those are PROJECT-011). The `Assignment` carries work/cost fields in the frozen contract, but this increment only establishes the valid scheduling-input relationship. `AssignmentSchedule` projects the resolved scheduling inputs (calendar id, max units, type, units) without computing work or cost.

### Derived schedule extension

`DerivedSchedule` is extended with an optional `assignmentSchedules: Record<AssignmentId, AssignmentSchedule>`, built deterministically (sorted by `AssignmentId`) so the same serialized `ProjectDocument` + options always produce byte-identical schedule bytes independent of input array order. `TaskSchedule` is extended with an optional `resolvedCalendarId`. Both additions are optional and backward-compatible: existing PROJECT-006..009 consumers that do not read them are unaffected.

### Semantic commands

PROJECT-010 implements the `AssignResource` and `UnassignResource` commands compatibly within the frozen `ProjectCommand` model (both union members already existed; the executor now accepts them). Every accepted command is deterministic, preserves stable identity, leaves the document valid, produces canonical diagnostics, and preserves unrelated hierarchy/baseline state. Every rejected command leaves the input document unchanged and returns deterministic diagnostics.

- `AssignResource` carries a fully-formed `Assignment`; the mutator stores it and lets the post-mutation validator enforce canonical rules. Early rejection for duplicate id, missing task/resource references, and a duplicate task/resource pair keeps the diagnostic surface clean. The inverse is `UnassignResource` for the new assignment id.
- `UnassignResource` removes an assignment by `AssignmentId`; the removed assignment is captured so the inverse (`AssignResource`) restores the exact prior relationship. Removing an assignment never moves task dates (PROJECT-010 does not resource-level).

No resource creation/update commands are added: the frozen command union lacks them, and resources are document-level state set directly (consistent with the frozen architecture and the existing fixture pattern). `LevelResources` remains in the union but is NOT implemented by PROJECT-010 (it is PROJECT-013 territory).

### Assignment validity through mutations

Assignments are keyed by stable `TaskId`. `IndentTask`/`OutdentTask`/`RenameTask` never change `TaskId`, so assignments survive those mutations unchanged. `DeleteTask` already prunes assignments referencing deleted tasks (the Microsoft Project outline-deletion behavior), so every accepted hierarchy mutation leaves the document's assignment state valid.

### Resource calendar deterministic resolution

`resolveResourceCalendarId(document, resourceId)` is the single canonical boundary for resource calendar resolution, returning `resource.calendarId ?? properties.defaultCalendarId` (or `undefined` for a missing resource). Tests and downstream layers use this helper rather than re-deriving (and potentially diverging from) the engine's calendar choice.

### PROJECT-011 / PROJECT-013 boundary

PROJECT-010 establishes the resource model, resource calendars, assignment references, calendar resolution, and resource scheduling inputs. It does NOT implement assignment work, actual work, remaining work, resource cost, assignment cost, work/cost calculations (PROJECT-011), or resource leveling (PROJECT-013). The engine does not move tasks to resolve overload, generate leveling commands, or alter dates because of resource contention. Over-allocation detection is deferred until the model has the required work units.

## PROJECT-011 canonical semantic clarifications

These clarifications record the canonical decisions required to implement assignments, work, and cost (PROJECT-011). They refine R-002, R-005, R-006, and R-007 without altering any frozen invariant; no architecture-change proposal is required.

### Assignment units semantics

For work resources, `Assignment.units` is a capacity fraction with explicit canonical semantics: `1.0` = 100% allocation, `0.5` = 50%, `2.0` = 200% (two full units of capacity). Units are NEVER silently mixed with percent-complete, hours, or working-minutes. The engine rejects a non-finite or negative units value with `INVALID_ASSIGNMENT_UNITS` (established in PROJECT-010). A zero-unit assignment is permitted (the model allows `units = 0`); its derived work is 0.

### Work calculation

The canonical work formula for a work resource is:

```
assignment.work = task.duration × assignment.units   (WorkingMinutes)
```

The task duration is the already-scheduled duration in the task's resolved calendar (the PROJECT-006 accepted schedule, unchanged by PROJECT-011). The resource calendar is a scheduling INPUT but does NOT move task dates or change the work formula in this increment — resource-calendar-driven task-date movement is leveling (PROJECT-013). Work is computed against the accepted schedule; leveling remains PROJECT-013.

Task work for a leaf task is the sum of its assignment work across all assignments on that task. A task with no assignments has work 0 (no resources means no resource work). The document-level `Task.work` field is NOT authoritative for derived schedule output; the derived value lives in `TaskSchedule.work`.

### Actual and remaining work

```
assignment.actualWork = round(assignment.work × task.percentComplete / 100)
assignment.remainingWork = assignment.work − assignment.actualWork
```

`percentComplete` is the authoritative progress input (PROJECT-008 precedence preserved). The status date does NOT override `percentComplete` for work calculations — it affects the derived `status` field (notStarted/inProgress/complete) per PROJECT-008, but `actualWork` is derived purely from `percentComplete`. This is deterministic: the same serialized document + options always produce byte-identical derived work bytes.

### Material resource semantics

Material resources are NOT work-capacity resources. Material `units` represent a consumable quantity (e.g. 10 tons of concrete), NEVER person-hours. The canonical material cost is:

```
assignment.cost = assignment.units × resource.standardRate + resource.costPerUse
```

Material resources contribute zero work: `work = actualWork = remainingWork = 0`.

### Cost resource semantics

Cost resources are pure cost categories. They MUST NOT contribute work and MUST NOT consume `maxUnits` as work capacity. The `Assignment.cost` field is the authoritative cost input for a cost resource. The canonical cost-resource cost is:

```
assignment.cost = assignment.cost   (the authoritative document value)
assignment.actualCost = round(assignment.cost × task.percentComplete / 100)
assignment.remainingCost = assignment.cost − assignment.actualCost
```

Cost resources contribute zero work: `work = actualWork = remainingWork = 0`. The engine does not manufacture hourly labor cost for a cost resource.

### Rates

The accepted `Resource` rates are `standardRate`, `overtimeRate`, and `costPerUse`. Canonical rate semantics:

- **Work resources**: `standardRateCost = (assignment.work / 60) × resource.standardRate` (work is in minutes; rate is per hour). `costPerUse` is a flat per-assignment cost added to the rate-based cost. `assignment.cost = standardRateCost + overtimeCost + costPerUse`.
- **Overtime cost is deferred**: the frozen `Assignment` contract has no `overtimeWork` input field, so the engine cannot determine which portion of work is overtime. `overtimeCost = 0` in PROJECT-011. This is a documented deferred limitation, not a guess. A future work item that adds an `overtimeWork` input to the `Assignment` contract will enable overtime cost calculation; until then the `overtimeRate` is echoed on `AssignmentSchedule` but never applied to cost.
- **Material resources**: `standardRate` is a per-unit rate. `cost = units × standardRate + costPerUse`.
- **Cost resources**: `costPerUse` is not applied (the `Assignment.cost` field is the authoritative cost). `standardRate`/`overtimeRate` are echoed but not used for cost derivation.

The scheduler never silently ignores a populated rate: it is echoed on `AssignmentSchedule` so downstream layers never re-derive it.

### Multiple assignments and aggregate demand

A task can have multiple assignments. Task work is the deterministic sum of assignment work. Aggregate demand exceeding `maxUnits` (over-allocation, e.g. two assignments at 1.0 each on a resource with `maxUnits = 1`) is a data/scheduling condition, NOT a trigger for date mutation. The engine does NOT level (PROJECT-013 territory). Over-allocation is reported through the derived work values; the caller can detect it by comparing aggregate assignment units against `maxUnits`.

### Calendar interaction

Work and cost calculations respect canonical working-time inputs. The accepted calendar engine (`resolveCalendar`, `isWorking`, `addWorkingTime`, `subtractWorkingTime`, `workingDuration`) is reused; no second resource-calendar engine is created. Resource calendar resolution from PROJECT-010 remains authoritative.

When `task calendar != resource calendar`, assignment work is calculated using the TASK's scheduled duration (`task.duration × units`). The resource calendar is exposed on `AssignmentSchedule.resolvedCalendarId` but does not change the work formula in PROJECT-011. This is deterministic. The engine does NOT modify the accepted PROJECT-006 task scheduling dates to fit the resource calendar (leveling is PROJECT-013).

### Status date / actuals interaction

`ProjectProperties.statusDate` is used for deterministic current-date evaluation. The engine NEVER uses `Date.now()`, the system clock, local timezone, or browser time. `percentComplete` remains the authoritative progress input (PROJECT-008 precedence preserved). The status date's role is to derive the task `status` (notStarted/inProgress/complete) per PROJECT-008:

- `percentComplete >= 100` → complete (actualWork = work, remainingWork = 0)
- `0 < percentComplete < 100` → inProgress (actualWork = round(work × percent / 100))
- `percentComplete == 0`:
  - statusDate set AND has reached/passed scheduled start → status `inProgress`, but `actualWork = 0` (no work reported complete)
  - otherwise → status `notStarted`, `actualWork = 0`

The status date does NOT silently overwrite `percentComplete` or invent progress. `actualWork` is derived purely from `percentComplete`.

### Summary task work/cost roll-up

Summary task work/cost is a deterministic aggregation from descendants, NOT renderer state. A summary's work = sum of direct children's derived work (rolled up recursively). A summary's cost = sum of direct children's derived cost. Nested summaries roll up correctly because the engine processes deepest-first (children before parents). A summary's own assignments (if any) are NOT included in the roll-up — assignments belong on leaf tasks, and the roll-up is from children only. This is NOT resource-weighted; it is a plain sum.

### Derived schedule extension

`AssignmentSchedule` is extended with optional derived work/cost fields: `work`, `actualWork`, `remainingWork` (branded `WorkingMinutes`), `cost`, `actualCost`, `remainingCost` (plain `number`), plus echoed `standardRate`, `overtimeRate`, `costPerUse`. `TaskSchedule` is extended with optional `work`, `actualWork`, `remainingWork`, `cost`, `actualCost`, `remainingCost`. Both additions are optional and backward-compatible: existing PROJECT-006..010 consumers that do not read them are unaffected. All derived values are deterministic.

### Baseline integration

PROJECT-009 established immutable baseline snapshots. PROJECT-011 makes current derived work/cost values available so baseline comparisons use them: `captureBaseline` now prefers the DERIVED `TaskSchedule.work`/`TaskSchedule.cost` over the document-level `Task.work`/`Task.cost`, falling back to the document field when no derived value exists (backward-compatible with PROJECT-009 consumers). Historical baseline snapshots are never mutated automatically; the existing baseline immutability is preserved. Baseline comparison continues to use the captured snapshot values.

### Semantic commands

PROJECT-011 adds the `SetAssignmentUnits` command compatibly within the frozen `ProjectCommand` model (a new union member, mirroring the PROJECT-008 `SetDeadline` precedent). The command changes the `units` field on an existing assignment without an Unassign+Assign pair that would lose the stable `AssignmentId`. Every accepted command is deterministic, preserves stable identity, validates ranges (`INVALID_ASSIGNMENT_UNITS` for non-finite/negative), leaves the document valid, produces canonical diagnostics, and leaves the original document unchanged on rejection. The inverse restores the previous units value so undo/redo is deterministic.

`SetTaskWork` and `SetActualWork` are NOT added because work and actualWork are purely derived values (the scheduler computes them from `task.duration × units` and `percentComplete`). The frozen command union does not need commands to store derived values — derived work/cost belongs in the scheduler.

### Validation

PROJECT-011 extends validation for work/cost inputs. Work fields (`work`, `actualWork`, `remainingWork`) must be finite and non-negative (`INVALID_WORK`, `INVALID_ACTUAL_WORK`, `INVALID_REMAINING_WORK`). Cost fields (`cost`, `actualCost`, `remainingCost`) must be finite and non-negative (`INVALID_COST`, `INVALID_ACTUAL_COST`, `INVALID_REMAINING_COST`). The "exceeds" direction is a hard rejection: `actualWork > work` → `INCONSISTENT_WORK`; `actualCost > cost` → `INCONSISTENT_COST`. The sum invariant (`actual + remaining = total`) is NOT enforced at the document level because `actualWork`/`actualCost` are derived values the scheduler recomputes — a document may carry stale derived fields that the next `schedule()` call overwrites. Only the "exceeds" direction is a real corruption.

### Determinism

The same canonical `ProjectDocument` + scheduling options MUST produce byte-identical `DerivedSchedule` output. The engine does not use current system time, random numbers, locale-sensitive sorting, or array position as identity. Assignment schedules are built from assignments sorted by `AssignmentId`; task schedules are assembled from tasks sorted by `TaskId`; summary roll-up processes summaries deepest-first with canonical `TaskId` tie-breaking. JSON serialization round-trips preserve derived work/cost values unchanged.

## PROJECT-012 canonical semantic clarifications

These clarifications record the canonical decisions required to strengthen and prove the critical-path and float model across edge cases deferred from PROJECT-006. They refine R-005 (scheduling determinism) and the PROJECT-006 critical-path/float contract without altering any frozen invariant; no architecture-change proposal is required. PROJECT-012 is an edge-case CORRECTNESS INCREMENT on top of the accepted CPM implementation — the scheduler is not rewritten wholesale.

### Critical path and float authority

Critical-path and float calculations belong exclusively in the scheduling engine (`packages/project-scheduling`). The canonical `schedule(projectDocument, schedulingOptions)` function remains the sole authoritative source of `earlyStart`, `earlyFinish`, `lateStart`, `lateFinish`, `totalSlack`, `freeSlack`, and `critical`. Renderer, host, UI, and selector code MUST NOT compute authoritative critical path or float. No second CPM engine is created.

### Critical flag definition

The canonical critical flag is derived from the accepted float model:

```
critical := totalSlack <= 0
```

This definition is preserved exactly from PROJECT-006. A task with zero total slack is critical. A task with negative total slack (a hard constraint makes the schedule impossible relative to the project finish) is also critical. A task with positive total slack is non-critical. The engine does NOT add UI-defined "critical" state and does NOT select only one "primary" critical path — every legitimately critical task on every critical chain is marked critical.

### Total slack formula

The canonical total slack is the SIGNED working-time distance between earlyFinish and lateFinish, measured in the task's resolved calendar:

```
totalSlack = signedWorkingDuration(calendar, earlyFinish, lateFinish)
```

This is the accepted PROJECT-006 formula, preserved exactly. Working-time arithmetic (not wall-clock minutes) is used so that a task whose earlyFinish is Monday 17:00 and lateFinish is Tuesday 09:00 has zero total slack (there are 0 working minutes in the Monday evening gap). Negative total slack arises when a hard constraint (MFO, MSO) pulls a predecessor's lateFinish earlier than its earlyFinish; the signed working-duration returns a negative value, and `critical := totalSlack <= 0` marks the task critical.

### Free slack formula (relationship-aware)

Free slack is relationship-aware: the distinct equations for FS, SS, FF, and SF are preserved exactly from PROJECT-006. Free slack measures how far the task can slip before the earliest date of a successor is affected. For a task with multiple successors, free slack is the minimum across all successor links. For a task with no successors, free slack equals total slack (the task can slip until the project finish without affecting anything).

The anchor (start vs finish) differs by relationship type: FS and FF anchor on the task's earlyFinish; SS and SF anchor on the task's earlyStart. The bound is computed by subtracting the lag (working-time) from the successor's relevant early date. Negative free slack arises when a hard constraint on a successor pulls the bound earlier than the task's early anchor.

### Forward pass (dependency-type semantics)

The forward pass computes earlyStart and earlyFinish for every dependency type with explicit, distinct equations:

- **FS** (finish-to-start): `successor.earlyStart = addWorkingTime(calendar, predecessor.earlyFinish, lag)`
- **SS** (start-to-start): `successor.earlyStart = addWorkingTime(calendar, predecessor.earlyStart, lag)`
- **FF** (finish-to-finish): `successor.earlyStart = subtractWorkingTime(calendar, addWorkingTime(calendar, predecessor.earlyFinish, lag), successorDuration)` — the successor starts late enough that its finish lands at predecessor.earlyFinish + lag.
- **SF** (start-to-finish): `successor.earlyStart = subtractWorkingTime(calendar, addWorkingTime(calendar, predecessor.earlyStart, lag), successorDuration)` — the successor starts late enough that its finish lands at predecessor.earlyStart + lag.

Positive lag delays the successor; negative lag (lead) overlaps the successor with the predecessor; zero lag applies no offset. The engine does NOT use one generic equation where the relationship semantics differ. Lag is measured in working minutes in the successor's task calendar (the forward pass iterates the successor and uses `calendarFor(successor)`).

### Backward pass (dependency-type semantics)

The backward pass computes lateStart and lateFinish for every dependency type with explicit, distinct equations:

- **FS**: `predecessor.lateFinish = subtractWorkingTime(calendar, successor.lateStart, lag)`
- **FF**: `predecessor.lateFinish = subtractWorkingTime(calendar, successor.lateFinish, lag)`
- **SS**: `predecessor.lateFinish = addWorkingTime(calendar, subtractWorkingTime(calendar, successor.lateStart, lag), predecessorDuration)` — the predecessor's lateFinish is the start that lets it finish at successor.lateStart − lag.
- **SF**: `predecessor.lateFinish = addWorkingTime(calendar, subtractWorkingTime(calendar, successor.lateFinish, lag), predecessorDuration)` — the predecessor's lateFinish is the start that lets it finish at successor.lateFinish − lag.

Late dates are bounded by the project finish so that slack always measures "slip without extending the project". Start-domain bounds (SS/SF) alone can be looser than the project finish; the engine takes the earlier of the dependency bound and the project finish. Late-date propagation is deterministic: the backward pass iterates the reverse of the canonical topological order, so unrelated project branches never produce order-dependent answers. Lag is measured in working minutes in the predecessor's task calendar (the backward pass iterates the predecessor and uses `calendarFor(predecessor)`).

### Multiple critical paths

When two or more equal-length critical chains converge, every task on every critical chain is marked critical. The engine does NOT select one "primary" critical path and does NOT break ties by array order. Tasks with `totalSlack == 0` are critical regardless of which parallel chain they belong to. The canonical `TaskId` ordering (locale-free, code-unit comparison) ensures the same input always produces the same set of critical tasks.

### Negative slack

Negative total slack arises canonically when a hard constraint (MFO or MSO) on a successor pins the successor's late dates such that the predecessor's lateFinish precedes its earlyFinish. The engine does NOT clamp negative slack to zero — the signed working-duration returns the actual negative value, and `critical := totalSlack <= 0` marks the task critical. This is the canonical "impossible schedule" signal: the project cannot meet the hard constraint given the dependency chain. The engine produces the derived negative-slack values; the caller (reporting/UI layer) interprets them.

### Milestones

Zero-duration milestones are full participants in the CPM. A milestone's earlyFinish equals its earlyStart (zero working minutes). Milestones do not create divide-by-zero or calendar-arithmetic bugs — the working-time primitives handle zero-duration correctly. A milestone on the critical chain is critical; a milestone with float is non-critical. Milestones with lag, hard constraints, multiple predecessors, and multiple successors are all handled by the same relationship-aware forward/backward pass.

### Summary tasks

Summary criticality and float derive from the canonical scheduling engine, not from renderer state. A summary's early dates are the min/max of its children's early dates; its late dates are capped by its children's late dates (the tighter of the dependency-imposed bound and the children's late envelope). A summary is critical when `totalSlack <= 0`, computed from its rolled-up early/late dates. The engine does NOT simply copy the first or last child value and does NOT introduce renderer-specific summary logic. Nested summaries roll up correctly because the engine processes summaries deepest-first (children before parents) with canonical `TaskId` tie-breaking.

### Constraint interaction

All accepted PROJECT-008 constraint semantics are preserved: ASAP, ALAP, SNET, SNLT, MSO, FNET, FNLT, MFO. Hard constraints (MSO, MFO) pin both early and late dates, making the task critical. Soft constraints (SNET, FNET) push early dates later only, preserving slack. Late-scheduled constraints (ALAP, SNLT, FNLT) pull the scheduled start/finish to the late window bounded by the constraint. PROJECT-012 does NOT alter the accepted constraint semantics; it verifies critical path/float behavior under each constraint type with exact expected-value tests.

### Calendar interaction

Critical path and float use working-time arithmetic (the accepted `signedWorkingDuration`, `workingDuration`, `addWorkingTime`, `subtractWorkingTime` primitives), never wall-clock minute differences. When predecessor and successor have different task calendars, the forward pass measures lag in the successor's calendar and the backward pass measures lag in the predecessor's calendar — this is the accepted MS Project behavior. Calendar boundaries (weekends, holidays, split-day calendars, inherited calendars, nonworking boundaries) are handled by the same accepted primitives. A task spanning a weekend has its float measured in working minutes, not wall-clock hours.

### Disconnected branches

Tasks not connected to the main chain (isolated root tasks, independent parallel branches, isolated milestones, isolated summary subtrees) have deterministic total slack and criticality. A disconnected task's lateFinish defaults to the project finish, so its total slack is the working-time distance from its earlyFinish to the project finish. A disconnected task IS critical when it is the project finish (its earlyFinish equals the project finish, giving zero total slack). A disconnected task is NOT critical merely because it has no successor — the engine does not mark every leaf critical by default. The project finish is the latest leaf earlyFinish, so only the task(s) on the longest path are critical.

### Deterministic task ordering

The canonical topological order uses `TaskId` comparison (locale-free, code-unit comparison) as the deterministic tie-breaker. The same serialized `ProjectDocument` + options always produces the same topological order, the same forward/backward pass results, and byte-identical `DerivedSchedule` output. The engine does NOT use `localeCompare`, object/property iteration order, or array position as semantic ordering. Reversed task arrays, reversed dependency arrays, and reordered parallel branches all produce byte-identical schedule bytes.

### Edge-case coverage

PROJECT-012 proves the canonical CPM model across: multiple critical paths (two and three equal-length chains), converging paths (diamond, fan-in), diverging paths (fan-out), zero-duration milestones (critical, noncritical, with predecessor/successor, with lag, with hard constraints), mixed dependency types (FS/SS/FF/SF in one graph), lag/lead (positive lag, negative lead, zero lag), calendar boundaries (different task calendars, weekend, holiday, split-day, inherited, nonworking boundary), constraints (SNET, SNLT, MSO, FNET, FNLT, MFO, ALAP), summary-task hierarchies (critical-only children, critical+noncritical children, nested summaries, summary with dependency boundaries), disconnected/noncritical branches, dependency chains with different calendars, near-critical tasks (one-day and one-hour total slack), and negative/zero/positive float.

### PROJECT-013 boundary

PROJECT-012 does NOT implement `LevelResources`, resource leveling heuristics, resource-driven task movement, resource optimization, or over-allocation resolution. PROJECT-013 (resource leveling) is blocked until PROJECT-012 is independently accepted. PROJECT-012 may consume PROJECT-011 work/cost inputs only where useful for tests but does NOT change resource scheduling semantics.

## PROJECT-013 — Resource leveling canonical semantic clarifications

PROJECT-013 implements deterministic resource leveling as an ENGINE operation that produces semantic commands. The canonical operation is:

```text
levelResources(projectDocument, options) → LevelingResult
      → proposedCommands (SetTaskStart[])
      → applyProjectCommand (LevelResources applies the batch)
      → schedule(projectDocumentAfterLeveling)
      → DerivedSchedule
```

The leveler is a pure deterministic function in `packages/project-scheduling`. It detects work-resource over-allocation across the current derived schedule, deterministically selects which eligible task to delay to resolve each conflict, and emits semantic `SetTaskStart` commands that — when applied through the canonical `applyProjectCommand` and re-scheduled — move whole tasks later in time so no resource is asked to work above its `maxUnits` at any working instant. The leveler MUST NOT mutate the input document; the scheduler remains the sole scheduling authority.

### Architecture — frozen command + dependency injection

The frozen `LevelResources` command shape (`{ type: 'LevelResources'; taskIds?: TaskId[] }`) carries only the scope filter. The full `LevelingOptions` (date window, critical/priority/deadline policy) are passed by callers via the pure `levelResources()` function; the `LevelResources` engine command uses the documented default policy.

The engine package is a lower architectural layer than the scheduling package (scheduling → engine for document validation; never engine → scheduling statically). To preserve the layer boundary, the engine exposes a leveler slot (`registerLeveler(fn)`) that the scheduling package registers at module load. When the host imports the scheduling package (which every host that schedules must do), the slot is populated, and the engine's `LevelResources` command dispatch can call the leveler without a circular static import. If no leveler has been registered (a host that imports only the engine package and never the scheduling package), the `LevelResources` command is rejected deterministically with `LEVELING_NOT_AVAILABLE`.

### SetTaskStart — supporting command

PROJECT-013 also implements the `SetTaskStart` command dispatch (the command was in the frozen union but previously fell through to `UNSUPPORTED_COMMAND`). `SetTaskStart` sets the `task.start` field — the candidate earliest start the scheduler uses. The leveler's `LevelingResult.proposedCommands` are `SetTaskStart` values; exposing `SetTaskStart` dispatch lets hosts apply proposed delays one-by-one through the canonical `applyProjectCommand` path (honoring the `LevelingResult → semantic commands → applyProjectCommand` architecture) rather than only as a black-box `LevelResources` batch. The mutator rejects a missing task and a malformed date. The inverse restores the previous start; when the task had no previous start, no inverse is emitted (undo requires a host snapshot, mirroring the `CreateBaseline` precedent).

### Leveling policy (canonical defaults)

- **Scope**: `taskIds` restricts leveling to the named subset (by `TaskId`). When undefined or empty, every auto-scheduled leaf task in the document is in scope. Out-of-scope tasks still contribute to demand (they are the immovable side of a conflict); only in-scope tasks can be delayed.
- **Date window**: `levelingDateWindow` (`{ start?, finish? }`) restricts over-allocation detection to assignments whose scheduled window overlaps the window. When undefined, the entire project span is considered.
- **Manual tasks**: ALWAYS protected. `task.manualScheduled === true` is never delayed; the leveler treats it as immovable and emits `LEVELING_PROTECTED_MANUAL` when it is the only resolvable side.
- **Critical tasks**: `respectCritical` (default `false`) protects critical tasks. When true, a critical task is never delayed; the leveler picks the non-critical side. When both sides are critical and protection is on, the leveler emits `LEVELING_PROTECTED_CRITICAL`. The default (false) means critical tasks ARE levelable and may extend the project — leveling may produce negative slack, observable in the re-scheduled `DerivedSchedule`.
- **Priority**: `respectPriority` (default `true`) uses `task.priority` to order conflict resolution. Higher priority is kept in place; lower priority is delayed first (mirrors the MS Project "higher priority = harder to move" convention).
- **Tie-breaking** (deterministic, locale-free): when priority is equal (or `respectPriority` is off), the task with the earlier `scheduledStart` is kept; when starts are equal, the task with the lexicographically smaller `TaskId` is kept. The leveler NEVER uses `Date.now()`, `Math.random()`, `localeCompare`, or array position as ordering identity.
- **Constraints**: hard constraints (`mustStartOn`, `mustFinishOn`) make a task immovable. Soft constraints are respected as floors/ceilings on the delayed start/finish: `startNoEarlierThan` clamps the new start up to the SNET date; `startNoLaterThan` and `finishNoLaterThan` reject delays that would push the task past the ceiling (`LEVELING_CONSTRAINT_CONFLICT`); `finishNoEarlierThan` is always satisfied when delaying (later finish); `asLateAsPossible` allows delaying (negative slack may result).
- **Deadlines**: a deadline is NOT a constraint and is never mutated. Leveling may produce a deadline miss; the re-scheduled `DerivedSchedule` exposes `deadlineVariance`/`deadlineMissed` faithfully. `respectDeadlines` (default `false`) — when true, the leveler refuses to delay a task past its deadline and emits `LEVELING_DEADLINE_CONFLICT`.
- **Milestones**: zero-duration milestones have no work demand and are never levelable for capacity. They are skipped.
- **Summaries**: summary tasks are never directly delayed. Conflicts are always attributed to leaf tasks. A summary's rolled-up dates reflect its children's movement after re-scheduling.
- **Splitting**: NOT supported by the frozen `Task` model (a task has a single contiguous `[start, finish]` window). Leveling moves whole tasks only. A single assignment whose `units` exceed `resource.maxUnits` cannot be resolved by moving the task and emits `LEVELING_INCOMPLETE` (splitting is deferred to PROJECT-045).
- **Negative slack**: leveling may produce negative slack (delaying a critical task extends the project). The leveler does not clamp slack; the re-scheduled `DerivedSchedule` reports it faithfully.
- **Identity preservation**: the leveler NEVER changes `TaskId`, `DependencyId`, `ResourceId`, `AssignmentId`, or any baseline snapshot. Baselines are immutable; only the current schedule's `task.start` candidates move.

### Over-allocation detection

The leveler detects over-allocation with a segment-based sweep per work resource. For each work resource, it collects assignment intervals `{ start, finish, units }` from the derived schedule (only assignments on tasks that contribute to demand — leaf tasks with non-zero duration, including manual tasks; milestones and zero-duration tasks have empty windows and are skipped; summaries-with-children are skipped because their own assignments would double-count rolled-up children). The sweep collects EVERY boundary timestamp at which either demand, effective capacity, OR the resource's working status can change: assignment `start`/`finish` endpoints, availability-window `start`/`finish` endpoints, AND resource-calendar working-period `start`/`finish` endpoints (bounded to the assignment span). Between two consecutive boundaries the active assignment set, the effective capacity, AND the resource's working status are all constant, so one evaluation per segment is both sufficient and complete. The effective capacity on a segment is the resource's `maxUnits` OR the tightest covering availability window's `units` (availability windows define the resource's max units over time, MS Project semantics), evaluated at the segment midpoint. The effective capacity does NOT consider the resource calendar — the calendar gates WHETHER the resource can work, not HOW MUCH capacity it has while working.

### Resource-calendar-aware demand (capacity concepts)

Three distinct concepts govern a work resource's capacity, and the leveler keeps them explicitly separate:

- **Resource calendar** determines WHEN the resource can perform work (its resolved working periods, from `resource.calendarId ?? properties.defaultCalendarId`). Over-allocation is evaluated only where the resource is actually available to perform work: during a segment where the resource's calendar says it is NOT working, the resource supplies no work capacity, so the demand against capacity on that segment is ZERO — there is no over-allocation there (the resource is not being asked to work above capacity; it is not working at all). Zeroing demand on non-working segments clips conflict windows to the resource's working periods: any open conflict is closed at the working→non-working transition, so a reported conflict window never spans a non-working interval. Two task-calendar windows that overlap on a day the resource does NOT work therefore produce NO over-allocation (no false positive), and a conflict is clipped to the intersection of the assignment window with the resource's working periods.
- **Availability windows** determine HOW MUCH capacity the resource has WHILE it is working. An availability window `{ start, finish?, units }` overrides `maxUnits` during its span (the tightest covering window wins). Availability windows do NOT change whether the resource works — they only cap the units it can supply while working.
- **`maxUnits`** is the default capacity when no availability window covers the instant.

Omitting the resource calendar from detection would report a FALSE over-allocation whenever two task-calendar windows overlap on a non-working resource day, even though the resource supplies no work that day. Omitting availability-window boundaries would miss over-allocations that arise ONLY from a mid-assignment capacity drop (the assignment endpoints alone do not bracket the conflict). Including both — plus the resource-calendar working-period boundaries — is correctness-critical. The maximal window where combined `units` exceed effective capacity AND the resource is working is the reported conflict window. Conflicts are sorted deterministically (by `resourceId`, then window start, then the sorted `TaskId` set of the conflicting sides).

### Conflict identity (signature)

The leveler deduplicates reported over-allocations by a conflict signature that includes the conflict WINDOW identity, not just the resource/tasks/assignments: `resourceId | sorted taskIds | sorted assignmentIds | window.start (epoch ms) | window.finish (epoch ms)`. The same set of assignments can legitimately produce MULTIPLE distinct conflict windows (e.g. a capacity drop, then recovery, then a second drop). A signature without the window would collapse those into one `reportedSignatures` entry, and the final unresolved-pass would overwrite the first conflict's `resolved` state with the second's. Including the window (normalized to epoch milliseconds so the signature is independent of ISO string formatting) guarantees distinct windows are reported as distinct over-allocations.

### Conflict resolution

For each conflict, the leveler partitions sides into delayable vs protected (manual, summary, milestone, out-of-scope, hard-constrained, or critical-when-`respectCritical`). If no side is delayable, it emits the most specific diagnostic (`LEVELING_PROTECTED_CRITICAL` / `LEVELING_PROTECTED_MANUAL` / `LEVELING_CONSTRAINT_CONFLICT` / `LEVELING_NO_ELIGIBLE_TASK`). Otherwise it picks the side with the largest keep-score (lowest priority, latest start, largest `TaskId`) to delay. The new start is the latest `scheduledFinish` of the OTHER sides, advanced to the next working instant in the RESOURCE's resolved calendar (so the resource is never asked to work outside its own calendar; the task's own calendar is then re-applied by `schedule()` when computing the scheduled start from the pinned `task.start` candidate). The new start is then validated against the delayed task's soft constraints (SNET clamp, SNLT/FNLT ceiling) and deadline policy. If validation fails, the leveler emits `LEVELING_CONSTRAINT_CONFLICT` or `LEVELING_DEADLINE_CONFLICT` and leaves that conflict unresolved.

The leveler iterates: propose a delay → apply to the working copy → re-schedule → re-detect → repeat until no conflicts remain or no eligible sides. A task delayed multiple times across iterations is deduplicated in `proposedCommands` (keep the LAST `SetTaskStart` per task); the full audit trail remains in `actions`.

### Determinism contract

Given the same serialized `ProjectDocument` and `LevelingOptions`, the leveler produces byte-identical `proposedCommands`, `actions`, `overallocations`, and `diagnostics`. Reversed task arrays, reversed assignment arrays, reversed resource arrays, and serialized round-trips (JSON parse) all produce the same output. Repeated leveling runs are byte-identical. Leveling an already-leveled document is a no-op (`LEVELING_NO_OVERALLOCATION`). The leveler never depends on wall-clock time, `Math.random`, `localeCompare`, or array position as identity.

### Impossible / incomplete leveling

When an over-allocation cannot be resolved without violating protected constraints (or without splitting, which is deferred to PROJECT-045), the leveler does NOT silently return a partially modified document and claim success. It applies every resolvable delay and surfaces the remaining conflicts as diagnostics (`LEVELING_INCOMPLETE`, `LEVELING_CONSTRAINT_CONFLICT`, `LEVELING_NO_ELIGIBLE_TASK`, `LEVELING_PROTECTED_CRITICAL`, `LEVELING_PROTECTED_MANUAL`, `LEVELING_DEADLINE_CONFLICT`). The `LevelResources` engine command is NON-atomic for incomplete leveling: it accepts and applies every resolvable delay, and the diagnostics make the incompleteness explicit (never silent). The `overallocations` list flags each conflict as `resolved: true` or `resolved: false` so downstream layers can see exactly what was and was not eliminated.

### Baseline + identity protection

Leveling never mutates baseline snapshots, baseline IDs, captured dates, or historical work/cost. After leveling, the current `DerivedSchedule` may differ from the baseline; that difference is observable through existing baseline comparison (`compareBaseline`). Leveling never mutates `TaskId`, `DependencyId`, `ResourceId`, `AssignmentId`, or dependency structure — the leveled document's dependency graph remains acyclic and the same set of identities is preserved.

### PROJECT-045 boundary

PROJECT-013 is the first resource-leveling implementation. It is NOT the final advanced leveling system. PROJECT-045 (advanced resource leveling) may add: task splitting, resource-pool-aware leveling, advanced priority/ordering rules, effort-driven task reshaping, and other advanced constraints. The PROJECT-013 frozen `LevelingOptions` and `LevelingResult` shapes are extensible (additional optional fields can be added without breaking the contract), but the canonical `levelResources(document, options) → LevelingResult` operation and the `LevelResources → proposedCommands → applyProjectCommand → schedule` architecture are frozen.

## PROJECT-014 — Native `.gproj` format

The canonical GenOffice Project persistence format is `.gproj`. It is a deterministic, versioned, self-describing JSON envelope wrapping a canonical `ProjectDocument` payload. It is NOT an MPP wrapper and NOT MSPDI XML (those are PROJECT-015..019). The internal model remains `ProjectDocument`; the file adapter is responsible for deterministic serialization/deserialization only.

### Architecture

```text
ProjectDocument → ProjectFileAdapter → .gproj
.gproj → ProjectFileAdapter → ProjectDocument
```

The `ProjectFileAdapter` contract (`inspect` / `import` / `export` over `Uint8Array`) is defined in `@genoffice/project-file` and implemented by the canonical `gprojFileAdapter` singleton. The adapter depends ONLY on `@genoffice/project-contracts` (types + brand helpers) and `@genoffice/project-engine` (the canonical `validateProjectDocument`). It has NO React, Electron, Node, browser, HTTP, or MPP/MSPDI dependencies (architecture-lock §13). Serialization logic never lives in React, the Electron renderer, browser code, Gantt components, the Project ribbon, or dialogs.

### Envelope

Every `.gproj` file is a JSON object with this canonical shape (keys emitted in alphabetical order by the canonical serializer):

```json
{
  "document":      { ... ProjectDocument ... },
  "format":        "gproj",
  "formatVersion": 1,
  "metadata":      { "format": "gproj", "version": "1" }
}
```

- `format` is the magic identifier. Anything other than `"gproj"` is rejected as `INVALID_GPROJ`.
- `formatVersion` is the FILE-format version (an integer). The current version is `1`. Bumped when the envelope or schema evolves.
- `metadata` is a `ProjectFileMetadata` block (`{ format, version, sourceName? }`). `metadata.format` echoes `format`; `metadata.version` echoes `formatVersion` as a string. `sourceName` is host-supplied (the filename on disk) and is NOT stored in the file bytes — it is passed in via the `metadata` parameter to `inspect`/`import`.
- `document` is the canonical `ProjectDocument` payload. It is the SOLE authoritative persisted state.

`ProjectDocument.schemaVersion` (the payload's own schema marker, currently `1`) is distinct from `formatVersion` (the envelope's version). The payload schema version is owned by `@genoffice/project-contracts`; the envelope format version is owned by `@genoffice/project-file`.

### Versioning

The adapter implements explicit format versioning:

- **Current format version**: `1`.
- **Supported read versions**: `[1]`.
- **Unsupported-version behavior**: a `formatVersion` not in the supported set is rejected with an `UNSUPPORTED_GPROJ_VERSION` error diagnostic. The parser does NOT silently read a future format using the current code. The returned document is the canonical empty document (not a corrupted partial).
- **Future-version behavior**: when a future format version is introduced, the supported-read set is extended and the parser gains a migration path; until then, future versions fail deterministically.

### Serialization

Serialization is canonical. Equivalent `ProjectDocument` values produce byte-identical `.gproj` files. The serializer:

- Preserves array order verbatim (task arrays, dependency arrays, notes, availability windows, calendar exceptions — order IS semantically meaningful and is part of the canonical document form). It does NOT sort away meaningful task order.
- Sorts object keys by Unicode code point (NOT `localeCompare`, which is forbidden) so two semantically-equivalent objects with differently-ordered keys serialize to identical bytes. Key order in a `Record<...>` map is NOT semantic; array order IS semantic.
- Uses a fixed 2-space indent and a single trailing newline so bytes are stable across runtimes.
- Encodes the JSON text as UTF-8 via a pure-TypeScript encoder (host-neutral — no `Buffer`, no `TextEncoder` runtime dependency).
- Emits NO wall-clock timestamps, NO random UUIDs, NO `Date.now()`, NO `Math.random`. Stable IDs come from the document.

Canonical ordering is defined for: tasks, resources, assignments, dependencies, calendars, baselines, custom fields, views, tables, filters, groups (all preserve array order — the order is part of the canonical document form). Identity remains `TaskId`, `ResourceId`, `AssignmentId`, `DependencyId`, `CalendarId`, `BaselineId`, `CustomFieldId`, `ProjectViewId`, `ProjectTableId`, `ProjectFilterId`, `ProjectGroupId` (all branded strings, never array position).

### Deserialization

Deserialization pipeline:

1. **Parse safely** — UTF-8 decode + `JSON.parse` (no `eval`, no `Function`, no reviver, no prototype deserialization, no arbitrary constructors). Malformed JSON → `INVALID_GPROJ`.
2. **Validate envelope** — `format === "gproj"` (else `INVALID_GPROJ`); `formatVersion` ∈ supported set (else `UNSUPPORTED_GPROJ_VERSION`); `document` present and an object (else `SCHEMA_INVALID`); `schemaVersion === 1` (else `SCHEMA_INVALID`).
3. **Validate schema** — every required field present + correctly typed; branded identity fields are non-empty strings; enum fields are in their union; `Record<...>` map keys are filtered against prototype-pollution hazards (`__proto__`, `constructor`, `prototype`). Missing required field → `MISSING_REQUIRED_FIELD`; wrong type → the entity-specific code (`INVALID_IDENTITY` / `INVALID_TASK` / `INVALID_RESOURCE` / `INVALID_ASSIGNMENT` / `INVALID_CALENDAR` / `INVALID_BASELINE`).
4. **Construct canonical** — promote raw strings/numbers to branded values through the single canonical promotion point (`asTaskId` / `asISODateTime` / `asWorkingMinutes` / … from `@genoffice/project-contracts`). Malformed entities are DROPPED (partial recovery, explicitly represented by an error diagnostic); the document is otherwise constructed from the valid entities.
5. **Validate document** — delegate to the engine's canonical `validateProjectDocument` (no duplicated diagnostic system). The engine's codes (`DUPLICATE_TASK_ID`, `MISSING_TASK_REFERENCE`, `CALENDAR_PERIOD_MALFORMED`, `CALENDAR_CYCLE`, `MISSING_BASE_CALENDAR`, …) are surfaced as error-level `ImportDiagnostic` entries (they are already valid `ImportDiagnostic.code` strings).
6. **Return diagnostics** — every dropped entity or invalid reference is surfaced as an error-level diagnostic; nothing is silently discarded. A `GPROJ_READ` info diagnostic notes the format version that was read.
7. **Malformed → fail** — file-level errors (bad JSON, wrong magic, unsupported version) return an empty document + a single error diagnostic. Entity-level errors drop the entity + emit an error diagnostic (partial recovery, explicitly represented).

A malformed `.gproj` file fails deterministically. The adapter never produces a partially valid `ProjectDocument` and claims success — the diagnostic contract explicitly represents partial recovery (dropped entities are enumerated).

### Round-trip invariant

The mandatory acceptance gate:

```text
document → serialize → deserialize → serialize  ≡  first serialization (byte-identical)
```

and the semantic invariant:

```text
document → serialize → deserialize → canonicalize  ≡  document (semantic-identical)
```

Both are asserted by the PROJECT-014 test suite for every golden fixture.

### Derived state (NOT persisted)

The native format does NOT persist authoritative derived state. The following remain DERIVED from the canonical document (re-computed by the scheduling engine):

- `earlyStart`, `earlyFinish`, `lateStart`, `lateFinish`, `totalSlack`, `freeSlack`, `critical`
- `scheduledStart`, `scheduledFinish`
- `deadlineVariance`, `deadlineMissed`
- `status` (derived progress status)
- `actualDuration`, `remainingDuration`
- `resolvedCalendarId` (task + assignment)
- `work`, `actualWork`, `remainingWork`, `cost`, `actualCost`, `remainingCost` (derived task + assignment values)
- `taskSchedules`, `assignmentSchedules`, `projectStart`, `projectFinish` (the `DerivedSchedule`)
- `BaselineVariance`, `BaselineComparison` (derived baseline comparison)
- `LevelingResult`, `LevelingAction`, `LevelingOverallocation` (resource-leveling output)

The `.gproj` file stores ONLY the canonical input/state (`ProjectDocument`), not duplicated derived output. Canonical scheduling state remains derivable from `ProjectDocument`.

### Command / journal state (NOT persisted)

The command journal history (undo/redo) is NOT part of the canonical `.gproj` format. The `JournalEntry` / `ProjectCommandResult` model is a host-specific UI/runtime concern. The native format persists the canonical `ProjectDocument`, not host-specific UI history. This is a documented limitation: undo/redo across a save/reopen requires a host snapshot; the native format does not carry one. A future increment may introduce a persisted journal if the architecture-lock is amended; until then, the `.gproj` format stores the canonical document only.

### Baselines

Baseline snapshots are persisted EXACTLY:

- baseline IDs preserved
- snapshot task IDs (the `Record<TaskId, …>` keys) preserved
- captured dates (`capturedAt`) preserved
- baseline work/cost/duration preserved
- multiple baselines preserved
- ordering deterministic (array order preserved)

The adapter does NOT recompute historical baseline data during load. Baseline snapshots are immutable input; the derived `BaselineVariance` / `BaselineComparison` are re-computed by the scheduling engine.

### Calendars

Calendars are persisted with: calendar IDs, inheritance (`baseCalendarId`), working periods (`workingWeek` keyed by day 0-6), exceptions (date + periods), and references. Calendar semantics round-trip exactly. The adapter does NOT normalize away meaningful distinctions. Malformed calendar periods (non-integer bounds, out-of-range `00:00-24:00`, empty intervals, overlapping periods) and malformed working-week day keys (not 0-6) produce `INVALID_CALENDAR` diagnostics.

### Resources / assignments

Persisted: resource identity/type, `maxUnits`, rates (`standardRate`, `overtimeRate`, `costPerUse`), calendars, availability windows; assignment identity, task/resource references, units, work/cost. The adapter does NOT serialize renderer-specific resource state. The adapter does NOT serialize leveling actions as authoritative state (those are derived).

### Custom fields / views

Persisted: canonical `CustomField` definitions (id, name, type) and per-task `customFields` values; host-independent `ProjectView` / `ProjectTable` / `ProjectFilter` / `ProjectGroup` definitions. The adapter does NOT persist renderer implementation details. The following MUST NOT leak into `.gproj`: React component state, DOM geometry, Electron `BrowserWindow` state, scroll position, canvas coordinates (unless canonicalized by the model), native menu state.

### File metadata

The adapter uses the existing `ProjectFileMetadata` contract. Exact semantics:

- `filename` / `sourceName`: host context (the filename on disk). NOT stored in the file bytes. The host passes it in via the `metadata` parameter to `inspect`/`import`; the adapter reports it back in the returned `ProjectFileMetadata`.
- `format`: `"gproj"` (the magic identifier, also the `ProjectSavePlan['format']` discriminant).
- `version`: the file's `formatVersion` as a string (e.g. `"1"`).
- `modified` information: NOT stored in the file (wall-clock timestamps are forbidden in deterministic serialization). For deterministic tests, explicit metadata/time values are supplied.

### Save plan

The adapter is runtime-independent. The `ProjectSavePlan` contract (`{ format, path?, document }`) is the host-side save-request abstraction; the file adapter implements native serialization behind it without creating host-specific paths.

### Error / diagnostics model

The adapter uses the existing `ImportDiagnostic` contract (`{ code: string, severity, message, entityId? }`). The `code` field is a plain `string` (NOT a frozen union), so the adapter extends it with the smallest compatible set of codes (no duplicated diagnostic system):

- `INVALID_GPROJ` — the file is not valid JSON, the root is not an object, or the envelope `format` is not `"gproj"`.
- `UNSUPPORTED_GPROJ_VERSION` — the envelope `formatVersion` is not in the supported read set.
- `SCHEMA_INVALID` — the envelope or document structure has the wrong type.
- `MISSING_REQUIRED_FIELD` — a required field on an entity is absent.
- `INVALID_IDENTITY` — a branded identity field has the wrong primitive type or is empty.
- `INVALID_REFERENCE` — a reference field points to a non-existent entity.
- `INVALID_CALENDAR` — a calendar has a malformed working-week day key, exception, or period.
- `INVALID_BASELINE` — a baseline has a malformed `capturedAt` or snapshot.
- `INVALID_ASSIGNMENT` — an assignment has non-numeric work/cost/units or a malformed id.
- `INVALID_TASK` — a task has a non-numeric duration/priority/percentComplete or an unknown enum.
- `INVALID_RESOURCE` — a resource has a non-numeric rate/maxUnits or an unknown kind.

The engine's canonical `validateProjectDocument` diagnostics (e.g. `DUPLICATE_TASK_ID`, `MISSING_TASK_REFERENCE`, `MISSING_BASE_CALENDAR`, `CALENDAR_PERIOD_MALFORMED`, `CALENDAR_CYCLE`) are surfaced as additional error-level `ImportDiagnostic` entries (passed through verbatim — they are already valid `code` strings).

### File security

The parser safely handles: malformed JSON (rejected), deeply nested invalid structures (depth limit `64`), oversized collections (input byte limit `100 MiB`), invalid primitive types (rejected with entity-specific codes), unexpected fields (ignored safely — no crash, no execution), and prototype-pollution payloads (`__proto__`, `constructor`, `prototype` keys are filtered before the document is constructed). The parser does NOT execute file content. The parser does NOT deserialize arbitrary classes. `JSON.parse` is used with NO reviver (no code execution path).

### Determinism

Proven by the test suite:

1. Same `ProjectDocument` → identical bytes.
2. Reordered semantically-equivalent collections (e.g. `Record<CustomFieldId, …>` keys reordered) → identical bytes (object keys are sorted by Unicode code point).
3. Round-trip → identical canonical semantics.
4. `serialize → deserialize → serialize → identical bytes` (byte-identity invariant, asserted for every golden).

The serializer uses NO `Date.now()`, NO random UUID generation, NO `localeCompare`, NO object-insertion-order for semantic ordering. Stable IDs come from the document.

## PROJECT-015 — MSPDI XML import

MSPDI (Microsoft Project Data Interchange) is an XML format. The importer maps MSPDI XML into the existing canonical `ProjectDocument` and delegates semantic validation to the accepted engine. The adapter does NOT become the canonical model; MSPDI-specific identity is NOT preserved as GenOffice identity; no MSPDI XML is written into React/browser code. Import-only as delivered (MSPDI export is PROJECT-016, delivered after this increment; MPP is PROJECT-017–019).

### Architecture

```text
MSPDI XML → MSPDIAdapter → ProjectDocument → validateProjectDocument → schedule
```

The MSPDI adapter lives in `@genoffice/project-file` beside the accepted `.gproj` adapter (PROJECT-014). It reuses the existing `ProjectFileAdapter` boundary, `ImportDiagnostic` contract, brand-promotion helpers, and `validateProjectDocument` pipeline — it does NOT introduce a second file-adapter abstraction. It exposed `inspect` + `import` only as delivered (PROJECT-016 later added `export` behind the same boundary). The adapter depends ONLY on `@genoffice/project-contracts` and `@genoffice/project-engine`; it has NO React, Electron, Node, browser, HTTP, or external-XML-library dependencies (the XML parser is a pure-TypeScript tokenizer shipped in the package — the XML analog of the `.gproj` UTF-8 codec).

### Identity mapping

Canonical identity is branded-string identity; MSPDI UID is the persistent interoperability identifier (architecture-lock §4). The mapping is deterministic, stable within an import, collision-safe across entity families (distinct prefixes), and reproducible across repeated imports of identical XML. NO random IDs, NO `Date.now()`, NO host seed.

| MSPDI source                                                    | Canonical field                | Mapping                                       |
| --------------------------------------------------------------- | ------------------------------ | --------------------------------------------- |
| `<Task><UID>`                                                   | `task.uid`                     | verbatim (number)                             |
| `<Task><UID>`                                                   | `task.id`                      | `asTaskId('t'+uid)`                           |
| `<Resource><UID>`                                               | `resource.uid` / `resource.id` | verbatim / `asResourceId('r'+uid)`            |
| `<Assignment><UID>`                                             | `assignment.id`                | `asAssignmentId('a'+uid)`                     |
| `<Calendar><UID>`                                               | `calendar.id`                  | `asCalendarId('c'+uid)`                       |
| `<Calendar><BaseCalendarUID>`                                   | `calendar.baseCalendarId`      | `asCalendarId('c'+uid)`                       |
| `<Task><CalendarUID>` / `<Resource><CalendarUID>`               | `task`/`resource.calendarId`   | `asCalendarId('c'+uid)`                       |
| Predecessor link (succ, pred, type)                             | `dependency.id`                | `asDependencyId('d-'+succ+'-'+pred+'-'+type)` |
| Baseline slot index (0=Baseline, 1=Baseline1, …, 10=Baseline10) | `baseline.id`                  | `asBaselineId('b'+index)`                     |

Identity is NOT WBS. `<OutlineNumber>` reconstructs `parentTaskId` only (canonical identity is the branded id, not the outline code). The canonical hierarchy engine remains authoritative for `outlineLevel`/`summary` consistency.

### Task import

Supported MSPDI task fields map into the canonical `Task`: `id`, `uid`, `wbs` (from `<OutlineNumber>`), `outlineLevel`, `name`, `taskType` (MSPDI `<Type>` 0=fixedUnits/1=fixedDuration/2=fixedWork), `summary`, `milestone`, `manualScheduled`/`autoScheduled` (from `<Manual>`), `start`/`finish`, `duration`, `constraintType`/`constraintDate`, `deadline`, `priority`, `calendarId`, `percentComplete`, `work`/`remainingWork`/`actualWork`, `cost`/`actualCost`/`remainingCost`, `parentTaskId` (reconstructed), `baseline[]` (task-level `<Baseline>` slots → top-level baselines), `customFields` (per-task `<ExtendedAttribute>` values), `notes` (`<Notes>` text). Unsupported fields are NOT blindly copied: they either map deterministically to an existing canonical field, are preserved via the explicit `customFields` extension, or produce a diagnostic. Never silently discarded.

### WBS / hierarchy

`<OutlineNumber>` (e.g. `1`, `1.1`, `1.1.2`) deterministically implies `parentTaskId` (the parent outline is the current with the last `.…` suffix removed, looked up in the WBS→TaskId map). After import, `parentTaskId`/`outlineLevel`/`wbs`/`summary` are internally consistent; the canonical hierarchy engine validates them (`INCONSISTENT_OUTLINE_LEVEL`, `INCONSISTENT_SUMMARY_FLAG`, `MISSING_PARENT`, `PARENT_CYCLE`). A dangling parent outline emits `INVALID_MSPDI_REFERENCE` and leaves `parentTaskId` unset. Malformed MSPDI hierarchy never produces an invalid `ProjectDocument`.

### Dependencies

MSPDI `<PredecessorLink>` (stored on the successor) maps to `Dependency`: `predecessorId`/`successorId` via the task map, `type` (`<Type>` 0=FS/1=FF/2=SS/3=SF), `lagMinutes` (see Date/duration semantics). Dangling predecessor/task/resource references are dropped with `INVALID_MSPDI_REFERENCE` (partial recovery — the document stays valid). Cycles/self-links are detected by the canonical dependency validation (`SELF_DEPENDENCY`, `DEPENDENCY_CYCLE`).

### Calendars

MSPDI `<Calendar>` maps to the accepted `Calendar`: `id`, `name`, `baseCalendarId` (inheritance link preserved — NOT flattened), `workingWeek` (`<WeekDay>` `<DayType>` 1=Sunday..7=Saturday → canonical key `DayType-1`; `<WorkingTimes>` → `CalendarPeriod` via `HH:MM:SS` → whole-minute offsets — see Date / duration semantics), `exceptions` (`<Exception>` `<Start>` date → single-date exception; recurring/yearly `<Type>` and multi-day `Start≠Finish` map to a single-date exception with an `UNSUPPORTED_MSPDI_FEATURE` warning — the canonical model has no recurring-exception representation). Resource calendars and task calendars remain distinct. Malformed calendars produce `INVALID_MSPDI_CALENDAR` and never silently become the default calendar.

### Resources / assignments

MSPDI `<Resource>` maps to `Resource`: `id`, `uid`, `name`, `kind` (`<Type>` 1=work/2=material/3=cost), `maxUnits`, `standardRate`/`overtimeRate`/`costPerUse`, `calendarId`, `availability` (`<AvailabilityPeriods>` → `{start, finish?, units}`). MSPDI `<Assignment>` maps to `Assignment`: `id`, `taskId`, `resourceId`, `units`, `work`/`actualWork`/`remainingWork`, `cost`/`actualCost`/`remainingCost`. All references are validated; dangling `TaskUID`/`ResourceUID` drop the assignment with `INVALID_MSPDI_REFERENCE`. Resource leveling is NOT performed during import; the scheduler is not asked to work around resource conflicts at import time.

### Baselines

MSPDI per-task `<Baseline>`/`<Baseline1>`..`<Baseline10>` elements are collected into top-level `Baseline` entities keyed by slot index. Each baseline preserves task-snapshot identity (`{start?, finish?, duration, work, cost}`), `capturedAt` (MSPDI carries no per-baseline captured date — a deterministic fallback is used: `<LastSaved>` → `<CreationDate>` → project `<StartDate>`), and `name` (`Baseline`/`Baseline N`). MSPDI baseline semantics that exceed the current `Baseline` contract emit explicit diagnostics; baseline information is never silently discarded.

### Constraints / deadlines / progress

MSPDI `<ConstraintType>` (0–7) maps to the eight canonical constraint types: 0=ASAP, 1=ALAP, 2=MSO, 3=MFO, 4=SNET, 5=FNET, 6=SNLT, 7=FNLT. A date-bounded constraint without a valid `<ConstraintDate>` emits `INVALID_MSPDI_CONSTRAINT`. A hard MSPDI constraint is NEVER silently approximated as a soft constraint. An out-of-enum `<ConstraintType>` emits `UNSUPPORTED_MSPDI_FEATURE` and is dropped. Deadlines (`<Deadline>`) are mapped separately from constraints. Progress (`<PercentComplete>`) maps to PROJECT-008 progress semantics.

### Date / duration semantics

MSPDI `<Duration>`/`<Work>` are ISO-8601 durations (`PT8H0M0S`); the time-part (H/M/S) is converted to integer `WorkingMinutes` (hours×60 + minutes); any non-zero sub-minute seconds remainder emits `INVALID_MSPDI_DURATION` (no silent rounding — canonical `WorkingMinutes` is integer). Date-part components (`P1D`/`P1W`/`P1M`/`P1Y`) represent elapsed/calendar time with no faithful working-minute conversion at import time and emit `UNSUPPORTED_MSPDI_FEATURE`.

MSPDI `<LinkLag>` is stored in **tenths of the unit declared by `<LinkLagFormat>`**, and every supported working lag unit applies its own explicit conversion to integer `lagMinutes` (PROJECT-015 correction round 1 — this table is authoritative):

| `LinkLagFormat` | unit           | conversion to `lagMinutes`                                     |
| --------------- | -------------- | -------------------------------------------------------------- |
| 1               | working minute | `LinkLag / 10` (non-multiple-of-10 → `INVALID_MSPDI_DURATION`) |
| 3               | working hour   | `LinkLag / 10 × 60` (always whole minutes)                     |
| 5               | working day    | `LinkLag / 10 × MinutesPerDay`                                 |
| 7               | working week   | `LinkLag / 10 × MinutesPerWeek`                                |
| 9               | working month  | `LinkLag / 10 × DaysPerMonth × MinutesPerDay`                  |

The day/week/month factors are the project-level conversion settings declared by the MSPDI itself (`<MinutesPerDay>`, `<MinutesPerWeek>`, `<DaysPerMonth>` on the `<Project>` root). When the file declares no factor, the MSPDI default settings apply (480 / 2400 / 20 — the documented 8-hour-day, 40-hour-week, 20-day-month Microsoft Project defaults); using the format default is the MSPDI-defined semantics, not an approximation, so no diagnostic is emitted for an absent declaration. Factor validation is lazy with respect to the lag formats actually present (correction round 2): a declared factor is validated only when a dependency carrying a lag format that uses it is encountered — `MinutesPerDay` for day (5) and month (9) lags, `MinutesPerWeek` for week (7) lags, `DaysPerMonth` + `MinutesPerDay` for month (9) lags. Minute (1) and hour (3) lags are factor-independent and never trigger factor validation, so malformed declarations that no present lag format uses produce no diagnostic and never poison an otherwise valid import. A malformed declared factor that a present lag format uses (non-positive or non-integer) emits `INVALID_MSPDI` naming the declaration — at most once per declaration, regardless of how many dependencies use it — and the affected lag converts with the documented default (a declared value is never silently approximated). Any conversion that does not yield a whole minute emits `INVALID_MSPDI_DURATION` and defaults the lag to 0 (dependency retained — never silently rounded, never silently dropped). The factors are import-time conversion parameters only — the canonical `ProjectDocument` stores the resulting integer `WorkingMinutes`, never the factors themselves.

Elapsed `LinkLagFormat` values (2/4/6/8/10) emit `UNSUPPORTED_MSPDI_FEATURE` (elapsed lag traverses calendar time including non-working time — no faithful working-minute representation at import time) and default the lag to 0 with the dependency retained. Percentage lags (`LinkLagFormat` 35) emit `UNSUPPORTED_MSPDI_FEATURE` (a percentage lag is a fraction of the predecessor's working duration, which is schedule state the adapter does not compute at import time) and default the lag to 0 with the dependency retained. Unknown format codes emit `INVALID_MSPDI_DURATION`.

Calendar working-time boundaries are whole-minute: MSPDI `<FromTime>`/`<ToTime>` (`HH:MM:SS`) must resolve to integer minute offsets; a time carrying non-zero seconds emits `INVALID_MSPDI_CALENDAR` and the `WorkingTime` period is dropped (never silently rounded). Missing `<FromTime>`/`<ToTime>` and empty/inverted periods likewise emit `INVALID_MSPDI_CALENDAR` and are dropped — a malformed period is never silently skipped.

MSPDI date-times are normalized to canonical UTC `ISODateTime` (`YYYY-MM-DDTHH:MM:SS.000Z`): naive dates are interpreted as UTC (never the system timezone), `Z` is preserved, explicit `±HH:MM` offsets are converted to UTC. Malformed dates emit `INVALID_MSPDI_DATE`. No host locale, no system timezone as a semantic input.

### XML parsing

The importer uses a safe pure-TypeScript XML parser (no `DOMParser`, no Node `fs`, no external library — architecture-lock §13). Safety: no `eval`/`Function`/reviver/arbitrary-constructor; `<!DOCTYPE>` is rejected (no DTD subset processing, no external-entity resolution); only the five built-in named entities (`&lt; &gt; &amp; &quot; &apos;`) and numeric (`&#NN;`/`&#xNN;`) entities are resolved, any other named entity is rejected; a decoded-text-size cap guards against billion-laughs/quadratic expansion; a depth cap guards against pathologically nested XML and bounds the recursive-descent stack; a byte-size cap is enforced before decode. Malformed XML throws → `INVALID_MSPDI`.

### Diagnostics

The importer reuses the existing `ImportDiagnostic` model and adds MSPDI-specific codes only where necessary: `INVALID_MSPDI`, `UNSUPPORTED_MSPDI_VERSION` (`<SaveVersion>` not in the supported set — no silent forward-read), `UNSUPPORTED_MSPDI_FEATURE` (warning — lossy but valid mapping), `INVALID_MSPDI_REFERENCE`, `INVALID_MSPDI_DATE`, `INVALID_MSPDI_DURATION`, `INVALID_MSPDI_CALENDAR`, `INVALID_MSPDI_RESOURCE`, `INVALID_MSPDI_CONSTRAINT`, `MISSING_MSPDI_FIELD`, `MSPDI_READ` (info). Every dropped or approximated semantic feature is named in diagnostics. The engine's `validateProjectDocument` diagnostics are surfaced verbatim as error-level entries (severity `error`, no `entityId`) — the adapter does NOT invent a parallel semantic validator.

### Round-trip / canonicalization

Import-only. The acceptance invariant: `MSPDI XML → import → ProjectDocument → validate → schedule`. The resulting `ProjectDocument` is deterministic: repeated imports of byte-identical MSPDI produce byte-identical `ProjectDocument` bytes (verifiable via `serializeGproj`). Semantically-equivalent XML element-order variations produce identical canonical output (every field is extracted by name, never by position). XML field order that IS semantically meaningful (task arrays, assignment arrays, dependency arrays) is preserved verbatim — the importer does NOT sort away meaningful order.

### Determinism

Proven by the test suite: (1) same MSPDI bytes → same `ProjectDocument` bytes; (2) equivalent element-order variation → identical canonical `ProjectDocument`; (3) every valid golden imports with zero error-level diagnostics and passes `validateProjectDocument` + `schedule` deterministically. The importer uses NO `Date.now()`, NO random IDs, NO host locale, NO system timezone.

## PROJECT-016 — MSPDI XML export

Deterministic export from the canonical `ProjectDocument` to MSPDI XML, behind the same `ProjectFileAdapter` boundary as the `.gproj` adapter (PROJECT-014) and the MSPDI importer (PROJECT-015). The canonical `ProjectDocument` remains the source of truth; MSPDI-specific XML structures never become canonical domain state. No second file-adapter abstraction; no renderer/Electron/browser/Node code (the XML writer is the pure-TypeScript `src/mspdi/xml-writer.ts`, the serialization analog of the accepted pure-TS parser).

### Architecture

```text
ProjectDocument → MSPDIAdapter.export → MSPDI XML
```

`exportMspdi(document)` returns `{ bytes, diagnostics }`. The adapter (`mspdiFileAdapter`) now exposes `inspect` + `import` + `export` — the same host-neutral contract shape as `gprojFileAdapter`.

### Round-trip semantics

The primary acceptance invariant: `ProjectDocument → exportMspdi → importMspdi → ProjectDocument` must be **semantically equivalent** for all supported PROJECT-014/015 fields (properties, tasks, hierarchy, identity, dependencies, lagMinutes, constraints, deadlines, progress, calendars, resources, assignments, baselines, custom fields), and the re-derived `DerivedSchedule` must match after re-scheduling. Byte-identical MSPDI XML is NOT required by the format (many equivalent serializations exist); what is required is that the SAME document always serializes to byte-identical **canonical** XML. For import-convention documents (identities following the accepted PROJECT-015 mapping), the round-trip is byte-identical at the canonical-document level (provable via `serializeGproj`).

Refusal: a document that fails `validateProjectDocument` is NOT exported — zero bytes, one `INVALID_MSPDI_EXPORT` error, and the engine's diagnostics surfaced verbatim as error-level entries. The exporter never serializes semantically invalid canonical state.

### Identity mapping (reverse of PROJECT-015)

| Canonical source                  | MSPDI field         | Mapping                                                                                                                                   |
| --------------------------------- | ------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `Task.uid` / `Resource.uid`       | `<UID>`             | verbatim (non-negative integers; otherwise synthesized with `INVALID_MSPDI_EXPORT`)                                                       |
| `Calendar.id` matching `c<uid>`   | `<Calendar><UID>`   | parsed uid                                                                                                                                |
| `Calendar.id` otherwise           | `<Calendar><UID>`   | smallest unused non-negative integer (deterministic; `UNREPRESENTABLE_MSPDI_VALUE` warning — id remaps consistently, references included) |
| `Assignment.id` matching `a<uid>` | `<Assignment><UID>` | parsed uid / synthesized as above                                                                                                         |
| `Baseline.id` matching `b<slot>`  | baseline slot index | parsed slot / synthesized as above (slots 0..10)                                                                                          |

`TaskId`/`ResourceId` are NEVER exported as MSPDI UIDs (architecture-lock §4). No random IDs, no clock reads, no array position as identity: synthesis is a pure function of the document content (ids parsed in code-point order take their uid; the rest take the smallest unused integers in canonical order).

### Lag export representation

The deterministic canonical representation is **working minutes**: `LinkLagFormat = 1` with `LinkLag = lagMinutes × 10` (the exact inverse of the accepted import conversion for format 1). A lag whose `LinkLag` would fall outside the safe-integer range is exported as lag 0 with an `INVALID_MSPDI_EXPORT_LAG` **error** (the dependency is retained — lag semantics are never silently changed or rounded). Project-level factor declarations (`<MinutesPerDay>`/`<MinutesPerWeek>`/`<DaysPerMonth>`) are NOT emitted: the canonical document stores no factors (they are import-time conversion parameters only), minute-format lags are factor-independent, and values are never invented.

### WBS / hierarchy / task order

Tasks are exported in **hierarchical-DFS order** (parents before children; sibling order = the canonical task array order). `OutlineNumber` and `WBS` are a **deterministic projection** of `parentTaskId` + sibling order (`1`, `1.1`, `1.2`, …) — exactly the form the accepted importer reconstructs parents from, so the exported hierarchy imports back into the same canonical `parentTaskId` relationships. WBS is not canonical identity. A canonical `task.wbs` that differs from the derived outline is replaced by the derived value with an `UNREPRESENTABLE_MSPDI_VALUE` warning (hierarchy relationships stay exact); a task array order that is not DFS (a child preceding its parent) is canonicalized to DFS order with an `UNSUPPORTED_MSPDI_EXPORT_FEATURE` warning (sibling order preserved). Task sibling order IS semantically meaningful — reordering siblings legitimately changes the XML.

### Deterministic XML ordering

Explicit ordering everywhere; no `localeCompare` (code-point comparisons), no randomness, no clock reads, no host timezone, no object-insertion-order dependence (record keys re-sorted):

- Project root children: `SaveVersion`, `UID` (canonical `properties.id`; warning when empty), `Name`, `StartDate`, `FinishDate`?, `StatusDate`?, `LastSaved`? (baseline capturedAt carrier, only when baselines exist — never invented), then `Calendars`?, `ExtendedAttributes`?, `Tasks`?, `Resources`?, `Assignments`? (containers only when non-empty).
- Calendars: default calendar first (marked `IsBaseCalendarDefault`), then ascending UID; `WeekDays` always all seven (`DayType` 1..7, `DayWorking` derived from period presence); exceptions sorted by date; working periods sorted by start.
- Tasks: DFS order (above); within a task the fixed field order is `UID, ID, Name, WBS, OutlineNumber, OutlineLevel, Summary, Milestone, Manual, Type, ConstraintType?, ConstraintDate?, Deadline?, Priority, CalendarUID?, Start?, Finish?, Duration, Work, RemainingWork, ActualWork, Cost, ActualCost, RemainingCost, PercentComplete, PhysicalPercentComplete?, Notes?, PredecessorLink*, baseline slots (ascending), ExtendedAttribute*` (custom-field values sorted by field id).
- Resources: ascending UID; availability periods sorted by start.
- Assignments: sorted by (task UID, resource UID) — the semantic assignment key.
- Dependencies: sorted by (successor UID, predecessor UID, link type); emitted as `PredecessorLink` elements on the successor.
- Custom-field definitions: sorted by `FieldID`; baselines: ascending slot.

Reordered-but-equivalent canonical inputs (calendars, resources, assignments, dependencies, custom fields, exceptions, availability periods) produce **identical** XML bytes; task sibling order changes legitimately produce different bytes.

### Calendars

`BaseCalendarUID` is preserved (inheritance is NOT flattened). The accepted importer materializes all seven weekday keys, so a canonical derived calendar with a partial `workingWeek` (absent keys inherit the base chain under `resolveCalendar`) is **materialized**: every weekday key emits the child's own periods when present, otherwise the nearest ancestor's periods for that key (never-working when no ancestor defines it). This is the sanctioned normalization — the resolved semantics are exactly recoverable (the round-tripped calendar re-resolves identically) and it is disclosed with an `MSPDI_EXPORT_NORMALIZED` info diagnostic. It is a pure record walk over `baseCalendarId` links, not a scheduling computation. A working period ending at 24:00 (`endMinute` 1440) is emitted as `24:00:00` with an `UNREPRESENTABLE_MSPDI_VALUE` warning (the importer's HH:MM:SS rule drops it on re-import). Exceptions export `<Start>`/`<Finish>` as the date at `T00:00:00` with the exception's working periods (recurring/multi-day shapes do not exist canonically).

### Derived state policy

The exporter never consults `DerivedSchedule` and never runs a scheduler (static source guard: no scheduling-package import, no derived-schedule references). The only derived projections are format-mandated and documented: the outline number projection and the weekday materialization (both above). Neither computes dates or durations.

### Baselines

Baseline snapshots export into MSPDI baseline slots (`Baseline`, `Baseline1`..`Baseline10`) with task snapshot `Start`/`Finish`/`Duration`/`Work`/`Cost`; slot ordering is ascending and deterministic. `capturedAt` has no per-baseline MSPDI representation: a uniform capturedAt round-trips exactly through the `<LastSaved>` carrier; divergent values are carried by the first (lowest-slot) baseline's capturedAt with `UNREPRESENTABLE_MSPDI_VALUE` warnings for the rest. `task.baseline` (the derived reverse index) is reconstructed empty by the accepted importer: a non-empty index is disclosed with an `UNSUPPORTED_MSPDI_EXPORT_FEATURE` warning; a listing with no matching snapshot is dropped with a warning. No unsupported baseline metadata is invented.

### Custom fields

Definitions export as `ExtendedAttribute` (`FieldID` = canonical id, `Alias` = name, `Type` = the canonical type literal `text|number|boolean|date`); values export per task (number → decimal text, boolean → `true`/`false`, null → `FieldID` with no `Value`). Field type is never silently changed. Two importer-side re-parse limitations are disclosed with `UNREPRESENTABLE_MSPDI_VALUE` warnings rather than silently accepted: string values that re-parse as number/boolean, and non-finite numeric values.

### Unsupported-feature policy (no silent loss)

Every dropped, re-projected, or non-reconstructible piece of canonical state is named by a diagnostic. Export codes: `INVALID_MSPDI_EXPORT` (refusal; also uid synthesis), `INVALID_MSPDI_EXPORT_LAG` (unrepresentable lag — exported 0 with error), `UNREPRESENTABLE_MSPDI_VALUE` (identity remap, 24:00 period ends, divergent capturedAt, empty names, non-integer/negative durations, wbs re-projection, numeric-looking string values), `UNSUPPORTED_MSPDI_EXPORT_FEATURE` (emitted-but-not-reconstructed state: `physicalPercentComplete`, multiple notes collapsed into the single `<Notes>` field, non-empty `task.baseline` index, view/table/filter/group definitions, non-DFS task order), `MSPDI_EXPORT_NORMALIZED` (weekday materialization info), `MSPDI_WRITTEN` (info). Views/tables/filters/groups have no MSPDI representation and are not exported (one warning with the count).

### XML representation / security

Pure-TypeScript writer (`src/mspdi/xml-writer.ts`): no `eval`, no `Function`, no arbitrary constructors, no browser DOM APIs, no Node APIs, no external XML libraries. Explicit escaping in text contexts (`&` `<` `>` and CR as `&#xD;`) and attribute contexts (additionally `"` `'`, tab/LF as numeric references); element/attribute names validated against the XML name production so malformed markup is impossible by construction. Deterministic physical layout: XML declaration, LF endings, two-space indentation, one element per line, inline text leaves. The five built-in named entities are exactly the set the accepted parser resolves, so exported text round-trips losslessly.

### External compatibility evidence boundary

PROJECT-016 establishes (1) internal round-trip correctness (export → accepted importer → semantically equivalent canonical document) and (2) MSPDI structural correctness of the emitted XML (namespace, root element, `SaveVersion` 16, valid element/attribute structure, correct enumeration codes). It does NOT claim verified Microsoft Project compatibility: no real Microsoft-Project-generated MSPDI corpus was available in the environment, and interop details that only real Microsoft consumers can validate (e.g. the `<Exception>` element shape — the exporter emits `<Start>`/`<Finish>` direct children per the accepted PROJECT-015 importer contract rather than the Microsoft `TimePeriod` wrapper, project-level `<UID>`, string-typed `<Type>` on `ExtendedAttribute` definitions) are explicitly **unproven claims**. External validation with real Microsoft Project files is deferred to the compatibility work (PROJECT-020 / PROJECT-047).

### Determinism proof

(1) The same document exports to byte-identical XML across repeated calls (asserted three times and across all valid goldens). (2) Reordered semantically-equivalent non-identity collections produce identical bytes. (3) Representative goldens (E01 full document, E04 lag encoding) are asserted against hand-embedded canonical XML so the writer cannot drift silently. (4) The exporter/writer source carries no clock, no randomness, and no locale-aware comparison (static source guard).

## PROJECT-017 — MPP adapter feasibility (investigation)

PROJECT-017 is an investigation, not a feature increment: the deliverable is the 20-section feasibility report `spec/project/mpp-feasibility.md` (mandated section set, one-of-four decision line) plus the discipline suite `packages/project-file/tests/mpp-feasibility.test.ts`. NO production MPP parser/writer code, NO new runtime dependency, NO adapter/scheduling behavior change was introduced; the grounding spike ran entirely outside the repository (disposable `/tmp` workspace — MPXJ 16.7.0 + OpenJDK 21, 7 real corpus files from the MPXJ LGPL test distribution covering MPP8/MPP9/MPP14), and its commands/outputs are recorded in the PROJECT-017 worklog entry (report [S17]).

Feasibility decision (the report's authoritative closing line): **FEASIBLE — MSPDI/INTERMEDIARY ADAPTER** — MPP import only, via an externalized MPXJ (LGPL) sidecar process feeding the accepted PROJECT-015 importer, with five mechanical adapter-owned normalizations (N1–N5: strip `-1` sentinel references; filter the hidden `UID 0`/`OutlineLevel 0` placeholder task (and the analogous placeholder resource); rewrite midnight-wrapping `WorkingTime` periods to `endMinute 1440`; pre-filter `-65535` "unassigned" assignments with an expected-loss diagnostic). MPP export is NOT feasible under the current architecture, licensing constraints, and ecosystem: no open-source MPP writer exists; the only programmatic writers are commercial Aspose.Tasks (US$1,797+/yr/developer OEM) or installed-Microsoft-Project COM automation — both rejected. PROJECT-016's MSPDI export remains the sanctioned Microsoft interchange output.

Constraints the report locks for any future PROJECT-018/019 (nothing in PROJECT-017 authorizes either):

- `architecture-lock.md` is untouched and stays satisfied by construction: foundation packages would carry only the adapter contract + pure-TypeScript MSPDI normalization layer; the MPXJ conversion process is a host-level sidecar outside all foundation packages (§14 of the report).
- Fidelity is tiered A–D per canonical feature (§9); known intermediary losses are exactly the accepted MSPDI boundary losses plus the diagnosed N1–N5 normalizations — no new silent loss.
- Determinism: repeat-run re-conversion of the corpus is data-byte stable except the non-semantic `<CurrentDate>` save timestamp (ignored by the accepted importer); PROJECT-018 must prove repeat-run byte-equality at the canonical-document level (§12).
- Security posture: MPP is untrusted binary input; the sidecar process boundary (no network, `-Xmx` cap, hard timeout, single-file I/O) is a precondition for any PROJECT-018 authorization (§8).
- Corpus model: pinned external download of the MPXJ LGPL test corpus at test time; no wholesale copy into the repository (§11).

## PROJECT-018 — MPP import (externalized MPXJ sidecar)

Implements the PROJECT-017-approved strategy verbatim: `MPP bytes → host-managed MPXJ conversion sidecar → MSPDI XML → the accepted PROJECT-015 importer → ProjectDocument → validateProjectDocument → schedule()`. No MPP parser exists in TypeScript; no MPP parsing exists inside foundation packages; the semantic authority remains the accepted importer and the engine. MPP export remains unauthorized (PROJECT-019 requires formal rescoping).

### Architecture (two layers, one dependency direction)

- **Foundation layer (`packages/project-file/src/mpp/**`, pure TypeScript, zero `node:` imports — CI boundary grep extends over it like all foundation source):** the adapter contract types (staged diagnostics; `MppSidecarFrame`; `MppConversionOutcome`; `MppImportResult`), the five N1–N5 normalizations (`normalize.ts` — parse via the accepted `parseXml`, transform a cloned tree, re-serialize deterministically via a compact `serializeXmlNode` that reuses the PROJECT-016 writer's escape primitives), the diagnostic codes, and `importMppFromMspdi(outcome)` (normalize → `importMspdi` → `validateProjectDocument`).
- **Host layer (`packages/project-mpp-host`, a NEW workspace package — NOT a foundation package):** `MppSidecarLauncher` (one-shot process model), the protocol frame validator, `importMppFromFile`/`importMppFromBytes` (full pipeline incl. `schedule()`), the Java sidecar source, the pinned-download tooling, and the tests. It may use Node child-process/fs APIs by design; it depends on `project-contracts`, `project-file`, and `project-scheduling` (host → foundation direction only, asserted by a static architecture test).

### Sidecar protocol (v1)

Invocation is a DIRECT argument array (never a shell string; `shell` is never enabled; paths travel as argv entries — no shell-injection surface): `java -Djava.awt.headless=true -Xmx512m -cp <mpxj.jar:lib/*> MppSidecar.java <input.mpp> <output.mspdi> <requestId>`. Under the default `networkIsolation: 'required'` policy the launcher wraps this array one level deeper — `unshare --net --map-root-user java …` (still a direct argv, no shell; the wrapper executes the JVM in place, so timeout/kill and exit-code semantics are unchanged). The Java sidecar is a single-file source-launcher program (runs on a JRE; no javac, exactly as the PROJECT-017 spike proved). Stdout carries EXACTLY ONE JSON status line (the frame; last non-empty line authoritative — JVM banner tolerance); the MSPDI payload travels ONLY via the output file, so sidecar stdout can never contaminate MSPDI parsing. Stderr is diagnostic noise (capped, attached to failures only). Exit 0 = the frame is the truth (`ok:true`, or `ok:false` with a protocol code such as `UNSUPPORTED_MPP_FORMAT`); exit ≠ 0 = unexpected failure → `MPP_SIDECAR_EXIT` with the stderr tail. The launcher enforces a wall-clock timeout with SIGTERM→SIGKILL escalation, stdout/stderr caps, an unreadable-input check (`MPP_INPUT_UNREADABLE` — a missing, permission-denied, or otherwise unreadable path is a READABILITY failure, deliberately distinct from any size failure, with the underlying OS reason such as ENOENT/EACCES preserved in the message), an input-size limit (`MPP_MAX_INPUT_BYTES` 100 MiB, checked before any spawn), and an output-size limit (`MPP_MAX_MSPDI_OUTPUT_BYTES` 100 MiB, aligned with the accepted importer's cap). A `commandBuilder` test seam exists solely so the failure-injection suite can drive the real process-management logic with a deterministic stand-in; production callers never set it.

One-shot process model: every conversion spawns a fresh JVM — the strongest failure containment (a hostile/corrupt MPP cannot poison a long-lived process; `-Xmx` applies per conversion). The ~1–3 s JVM start cost is the price of isolation; resident-pool amortization is deferred to PROJECT-048. The sidecar source performs no application-level network access and executes no code from the input file (static source guard), and — independently of that source-level posture — the launcher enforces OS-level network isolation on the process boundary (see the security model below).

### MPP version support

MPP8 (Project 98), MPP9 (Project 2000/2002, incl. Project-2003 down-level saves), MPP12 (Project 2003/2007), MPP14 (Project 2010 and all later versions — the format version remains MPP14). Coverage is proven by real corpus files of every version (below). Claims are anchored to the tested MPP FORMAT versions, not to exact Microsoft product versions. Unsupported/unrecognized inputs (including password-protected non-MPP9 files and corrupt containers) produce the deterministic `MPP_UNSUPPORTED_FORMAT` refusal.

### Normalizations N1–N5 (exactly the PROJECT-017 definitions)

N1 strips `-1` sentinel `CalendarUID` references on tasks/resources ("inherit the default"); N2 strips the `-1` sentinel `BaseCalendarUID` on calendars (the same-pass root-cause twin, separately diagnosed for traceability); N3 filters the hidden placeholder task (`UID 0`/`OutlineLevel 0`/`WBS "0"` — the MPP summary artifact) and the analogous null-name placeholder resource; N4 rewrites `WorkingTime` `ToTime 00:00:00` ("until midnight", the Microsoft/MPXJ convention) to the ISO-8601 day-end `24:00:00` (canonical `endMinute: 1440`); N5 drops "unassigned" assignments (`ResourceUID -65535`) with a warning. Every rewrite/drop emits a stage-`'normalization'` diagnostic naming the affected entity; N1–N4 are severity `info` (lossless rewrites), N5 is `warning` (expected loss). Nothing else is normalized — no unapproved hidden transformations exist.

### Documented PROJECT-015 minimal correction (24:00:00)

The real corpus revealed an MPP-origin gap: MPXJ emits "until midnight" periods as `ToTime 00:00:00`, and no lossless XML-level rewrite exists because every textual hour is capped at 23:59 by the accepted `mspdiTimeToMinutes`. Per the delegation protocol this was NOT done silently: the minimal compatible correction accepts exactly `24:00:00` → minute 1440 (the ISO 8601 / XSD `xsd:time` legal day-end expression) in `mspdiTimeToMinutes`; every other out-of-range value (`25:00:00`, `24:00:01`, `24:30:00`) remains rejected. The one pre-existing test line that listed `24:00:00` as malformed was updated to its nearest malformed neighbors, and a dedicated boundary test was added. No other PROJECT-015 behavior changed (the full 015 suite passes untouched).

### Diagnostics provenance (staged, never flattened)

`MppDiagnostic = ImportDiagnostic & { stage: 'sidecar' | 'normalization' | 'mspdi' | 'canonical' | 'scheduling' }`. Sidecar/protocol codes (`MPP_SIDECAR_UNAVAILABLE`, `MPP_SIDECAR_TIMEOUT`, `MPP_SIDECAR_EXIT`, `MPP_SIDECAR_RESPONSE_INVALID`, `MPP_INPUT_UNREADABLE`, `MPP_INPUT_TOO_LARGE`, `MPP_OUTPUT_TOO_LARGE`, `MPP_UNSUPPORTED_FORMAT`, `MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE`) and normalization codes (five) are defined by the foundation; MSPDI codes remain the accepted PROJECT-015 set; canonical codes remain the engine set; scheduling codes remain the scheduling-engine set. The caller can always tell whether a feature was imported exactly, normalized (which N), unsupported, rejected, or dropped with a diagnostic — and whether an input-side failure means "cannot be read at all" (`MPP_INPUT_UNREADABLE`) rather than "too large" (`MPP_INPUT_TOO_LARGE`).

### Error atomicity

A fatal CONVERSION failure (sidecar error/timeout/crash/invalid frame, the fail-closed network-isolation refusal `MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE`, an unreadable or oversized input, malformed MSPDI, or a canonical-validation rejection) yields `emptyProjectDocument()` — never a partially authoritative document — with every stage's diagnostics preserved and the input file never mutated (the pipeline writes only to a per-import temp workspace that is deterministically removed; the result holds the bytes it needs in memory). Recoverable importer-level errors keep the accepted PROJECT-015 semantics (recovered document + error diagnostics; any error-level diagnostic means "not usable"). A scheduling failure keeps the valid canonical document and returns the schedule in its rejected shape — scheduling is derived state, not conversion.

### Determinism

The same MPP bytes + the same pinned sidecar version produce the same canonical `ProjectDocument`: proven by repeated full-pipeline runs over all 8 corpus files (`serializeGproj` byte-identical, `schedule()` JSON-identical, diagnostics identical). Canonical determinism is measured AFTER MSPDI import because the sidecar's `<CurrentDate>` save stamp is non-semantic (the accepted importer ignores it entirely — re-verified for this increment).

### Real corpus (external, checksummed, LGPL-clean)

8 real files pinned to MPXJ commit `abdbf6ef85654e3eff35c11c5e76cf08da842dce` (tag v16.7.0) — `DurationTest8.mpp` (MPP8), `SubprojectA-9.mpp` + `task-baselines-project2003-mpp9.mpp` (MPP9), `mpp12relations.mpp` (MPP12), `mpp14relations.mpp` + `mpp14assignmentcustom.mpp` + `ResourceIdAndUniqueId-project2010-mpp14.mpp` (MPP14), `sample.mpp` (legacy pre-2003 container, richest file: all four dependency types, midnight working periods, calendar inheritance). The corpus manifest (`packages/project-mpp-host/corpus/corpus-manifest.json`) records per file: filename, format, SHA-256, source path/provenance, expected N1–N5 counts, expected errors, and the expected canonical shape (entity counts + projectFinish); `scripts/fetch-sidecar-deps.mjs` downloads and SHA-256-verifies everything into the gitignored `.sidecar-deps/` workspace (the MPXJ 16.7.0 distribution — SHA-256 `2a149f3a…` — plus the corpus). NOTHING LGPL enters the repository. Goldens I01–I12 are external-file-driven (they assert the manifest contract against the real files — no fabricated MPP bytes).

### Licensing / distribution

MPXJ 16.7.0 (GNU LGPL) is downloaded at development/CI time, never committed. Today GenOffice distributes no MPXJ artifacts, so no LGPL redistribution duty is triggered by the repository. The obligations that a future BUNDLED release must discharge (license text, source offer, aggregate notice, relinkability via the process boundary) are documented in `packages/project-mpp-host/LICENSE-THIRD-PARTY.md`. The version is pinned in two synchronized places (`MPXJ_PINNED_VERSION` in the foundation contract + the manifest) with a test keeping them in lockstep; the npm license gate is untouched (zero npm dependencies added). Rejected dependency routes (the @byteink/mppjs AOT wrapper, Aspose, CheerpJ) stay rejected per PROJECT-017 §18.

### Security model

MPP files are untrusted binary input. Containment: the sidecar is a separate headless JVM with a hard `-Xmx` cap, a wall-clock timeout, direct-argv invocation (no shell), single-file I/O limited to the caller-provided input/output paths inside an isolated per-import temp workspace, deterministic cleanup, and size caps on input and output. The accepted PROJECT-015 parser's existing hardening (byte caps, depth caps, entity guards) governs the MSPDI that crosses the boundary.

**Network isolation — enforced at the OS level, fail-closed (architect-review correction; replaces the former "known limitation" posture).** "No network access" is a runtime invariant of the process boundary, not a property of today's source code: under the default `networkIsolation: 'required'` policy the launcher wraps every sidecar process in a fresh kernel-enforced network-isolated context — on Linux a user namespace (current user mapped to root, which is what makes unprivileged namespace creation possible; NO extra filesystem privileges — the kernel uid is unchanged) plus a network namespace whose only interface is a DOWN loopback (no routes, no DNS, no connectable sockets) — via `unshare --net --map-root-user` as a direct argument array (the wrapper executes the JVM in place, so process management is unchanged). A host that cannot provide the mechanism FAILS CLOSED with `MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE` (a sidecar-stage error, atomic: empty document, no sidecar process started); `'off'` is an explicit operator opt-out for local development, never a silent fallback. Runtime enforcement is verified by a real test: a process launched through the wrapper cannot connect even to a live loopback listener in the parent namespace (`ENETUNREACH`), while the identical unwrapped control connects — proving both the isolation and the validity of the test apparatus. The Java source additionally performs no application-level network access and executes no code from the input file (static guard). Non-Linux hosts currently have no OS mechanism implemented here: under the default policy they fail closed; running unisolated there requires the explicit `'off'` operator decision (a future host-hardening work item may add mechanisms for those platforms). The real-corpus e2e suite runs the production `'required'` posture wherever the host provides the mechanism and records the posture in its output.

### Performance evidence (operational viability, not the PROJECT-048 budget)

Representative full-pipeline wall times on the CI/dev machine (JVM start dominant): every corpus file converts + imports + schedules in ≈1.2–2.5 s; the largest fixture (56 tasks + 55 assignments) ≈2.5 s with a 49 501-byte `.gproj`; the richest (sample.mpp, 22 tasks/4 calendars) ≈2.1 s with a 23 770-byte `.gproj`; MSPDI intermediates are tens-of-KB scale. Peak memory is bounded by the `-Xmx512m` JVM cap plus the node-side buffers. This proves normal-project operational viability; large-scale budgets remain PROJECT-048.

### External compatibility evidence boundary

Established: internal correctness (N1–N5 + round-trip + determinism + atomicity over 8 real files across MPP8/9/12/14, all Microsoft-Project-generated corpus artifacts from the MPXJ test distribution). NOT established: compatibility with any specific Microsoft Project product version's files beyond the tested corpus files, enterprise/server fields, password-protected modern MPP files, or `.mpt` templates — explicitly unproven claims.

### CI

`project-foundation.yml` (the Project gate) grows from 15 to 19 steps: `Setup Java (MPXJ sidecar)` (temurin 21) → `Fetch MPP sidecar dependencies` (pinned, SHA-256-verified download) → `Typecheck project-mpp-host` → `Test project-mpp-host` (the full 89-test suite incl. the real-corpus e2e under the enforced network-isolation posture and the network-isolation evidence suite). The four foundation packages' steps and the boundary grep are unchanged (the new host package is deliberately OUTSIDE the foundation boundary grep — it is the sanctioned location for process code).

## PROJECT-019 — MPP export strategy / rescope (investigation; no production code)

PROJECT-019 as originally defined ("MPP export; compatibility suite; exporter emits supported MPP semantics") is formally RESCOPED to **019A — the strategy/product-decision increment**: the deliverable is `spec/project/mpp-export-strategy.md` (19 mandated sections + a single closing strategy-decision line) plus synchronized updates to `requirements.md`, `work-items.md`, `dependency-graph.md`, `verification-matrix.md` and the discipline suite `packages/project-file/tests/mpp-export-strategy.test.ts`. **Zero production code, zero new runtime dependencies, zero changes to accepted PROJECT-014…018 semantics.** No MPP writer, no MPP export API, no commercial SDK dependency, no COM bridge exists anywhere in the repository after this increment.

Strategy decision (the report's authoritative closing line): **MPP EXPORT DEFERRED — MSPDI/.gproj IS THE SUPPORTED WRITE PATH** (outcome E of the brief's A–E set). MPP remains import-only (PROJECT-018). `.gproj` (PROJECT-014) and MSPDI (PROJECT-016) are the sanctioned write formats. The documented user path to a true `.mpp` is the user-installed-Microsoft-Project workflow: export MSPDI (Microsoft's designated interchange format), open in Project, save-as MPP.

Requirements this decision locks:

- **Evidence currency**: every external export claim was re-verified against live primary sources for this increment (report §19 [E1]–[E20]): MPXJ still has no MPP writer (FAQ verbatim + write-howto surface + latest release 16.7.0); Microsoft designates XML as the Project interchange format, Project desktop cannot save the 2000–2003 formats, and Project-for-the-web has no MPP export; Microsoft does not recommend or support server-side Office automation (KB 257757); Aspose.Tasks saves MPP only under a paid license (OEM US$2,397 + from US$1,797/yr) with format-version-preservation behavior and no documented determinism contract.
- **Determinism invariant (unchanged and extended to the export question)**: the same canonical `ProjectDocument` must produce deterministic output. This is PROVEN for `.gproj` and MSPDI (accepted increments) and is UNPROVEN for every MPP-writer candidate (Aspose: closed-source, undocumented; COM: Project stamps save timestamps; no suppression/normalization/post-processing mechanism exists for any of them — post-processing would itself require an MPP writer). No future increment may claim deterministic MPP output without a licensed empirical proof using the round-trip oracle: `ProjectDocument → MSPDI → converter → MPP → PROJECT-018 import pipeline → ProjectDocument′` with canonical byte-equality asserted across repeated runs (report §8).
- **Conditional reopening (019B, NOT AUTHORIZED)**: the commercial-SDK path may be opened ONLY by an explicit business decision, and requires: license procurement, an EULA/redistribution legal review, a licensed feasibility spike proving the oracle determinism and the §7 fidelity-tier estimates, host network-isolation hardening beyond Linux (an Aspose sidecar would fail closed on Windows/macOS under the current `networkIsolation: 'required'` policy), and Principal Architect acceptance of a new host-level export-sidecar package. Metered Aspose licensing (which contacts Aspose's metering service) is disqualified by the offline + isolation requirements.
- **Proposed-only follow-up (019C, NOT IMPLEMENTED)**: an `MPP_EXPORT_UNSUPPORTED` diagnostic family with actionable save-as-MSPDI guidance at the adapter-contract level — the first production-code deliverable of the rescoped PROJECT-019, contingent on the Principal Architect accepting this rescope.
- **Dependency-graph meaning**: the work-item chain `014 → 015 → 016 → 017 → 018 → 019` completes with 019A (a decision record, exactly as 017 was); the `019 → 047` edge survives with PROJECT-047's compatibility suite chartering import/export regression-blocking over the SUPPORTED formats (`.gproj`, MSPDI; MPP import-only).
- **BLOCKED-until-acceptance rule**: this rescope proposes, it does not decide — PROJECT-019 stays blocked until the Principal Architect accepts the 019A decision record; nothing in it authorizes 019B or 019C.

## PROJECT-020 — Import compatibility diagnostics

Builds the canonical compatibility-diagnostics layer: given an imported project from `.gproj`, MSPDI, or MPP, produce a deterministic, provenance-aware `CompatibilityReport` that explains exactly what was preserved, normalized, approximated, rejected, or discarded. The layer lives at the project-file/import boundary (`packages/project-file/src/compatibility/**`) — never in React/Electron/web code; there is no second canonical model, no renderer-owned diagnostics, and no new parser (the accepted import pipelines are consumed verbatim through `importGprojWithCompatibility` / `importMspdiWithCompatibility` / `importMppWithCompatibility`).

### Diagnostic model

`CompatibilityDiagnostic` extends the accepted `ImportDiagnostic` contract structurally (existing callers are unaffected): every diagnostic carries `code` / `severity` / `message` / optional `entityId` plus the compatibility semantics `format` (gproj/mspdi/mpp), `stage`, `loss` (the data-loss classification), `recoverability`, and a best-effort `entityType` derived from the deterministic identity prefixes (`t`/`r`/`a`/`c`/`b`/`d-`/numeric custom-field ids). No `field` or `sourceLocation` properties exist — the parsers do not track them, and the layer never invents provenance it does not have.

**Stages** (deterministic names; the accepted `MppDiagnosticStage` values are preserved verbatim, with `'gproj'` added for the native read stage): `gproj → canonical → scheduling` (`.gproj`), `mspdi → canonical → scheduling` (MSPDI), `sidecar → normalization → mspdi → canonical → scheduling` (MPP). Format-specific failures are never collapsed into a generic "import failed" — a sidecar refusal, an N-code normalization, an importer rejection, a canonical validation failure, and a scheduling failure are all distinguishable stages.

**Severity rules** (unchanged from the accepted adapters, now formalized): `info` = mechanical normalization that preserves semantics (N1–N4, read sentinels); `warning` = semantic loss or supported-with-degradation (N5, unsupported features, the two provenance additions below); `error` = the document (or an entity) cannot be safely imported/validated. Unsupported optional fields are never fatal; true semantic loss is never `info`.

**Recoverability** (the ladder, never "best effort"): `fatal` — no authoritative ProjectDocument exists (file-level failure, sidecar failure, the MPP pipeline's atomic canonical rejection); `partial` — permits partial structural import only (entity/feature dropped, or an invalid document still constructed and returned by the `.gproj`/MSPDI adapters); `canonical` — the canonical document is valid, schedulable, and save-eligible (warnings, and scheduling-stage errors: a derived-state failure never invalidates the document); `preserved` — fully preserved semantics (informational bookkeeping).

**Data-loss classification** (every issue maps to exactly one): `none` (bookkeeping, e.g. read sentinels), `normalized` (mechanical rewrites that preserve semantics — N1–N4), `approximated` (a canonical value derived from the best available source — baseline `capturedAt` fallback), `dropped` (parseable source data with no canonical representation — N5 unassigned assignments, MSPDI `PhysicalPercentComplete`), `unsupported` (features with no faithful representation — elapsed durations, percentage lags, unsupported versions/formats), `invalid` (malformed/unresolvable input that was rejected). N5 is classified `dropped`, NOT `normalized`: its code, message, and warning severity all state a drop, and hiding true semantic loss behind a lossless label would violate the severity rules — this is the one deliberate deviation from the brief's "N1–N5: NORMALIZED" example, justified by the brief's own actual-behavior and no-hidden-loss rules.

**Classification knowledge base**: `COMPATIBILITY_CODE_CLASSIFICATIONS` is a total, deterministic table over every known code — the PROJECT-014 `.gproj` family, the PROJECT-015 MSPDI import family (incl. the two provenance additions below), the PROJECT-018 N1–N5 + sidecar families, the complete canonical engine validation code set, and the scheduling-engine failure codes. Unknown codes (forward compatibility) fall back deterministically (the format's import stage; severity-derived loss). A lockstep test reads the engine + scheduling sources and asserts every code they can emit has a table entry.

### Aggregation determinism

`buildCompatibilityReport` is a pure function: the same `{format, sourceVersion, diagnostics, schedulingDiagnostics}` always produce the same report — same diagnostics, same canonical ordering (`stage → severity → code → entityType → entityId → message`, severity ordered error/warning/info, with a stable tie-break on the producer's original order so source order survives only among otherwise-identical keys), same counts, same status/authority/save-eligibility. No dependence on object iteration order, `localeCompare` (plain code-unit comparison), filesystem enumeration, randomness, or timestamps.

**De-duplication policy (deliberate: none in the aggregation layer).** Identical diagnostics are distinct, countable occurrences whose multiplicity carries information (five N4 midnight rewrites in `sample.mpp` are five diagnostics; collapsing them would undercount `normalizedCount`). Declaration-level uniqueness is a PRODUCER contract, already in force in the accepted pipeline (the lazy lag-factor validation emits ONE diagnostic per malformed declaration regardless of how many dependencies use it — proven by test); entity-scoped diagnostics are distinct by their `entityId`. Accidental de-duplication of meaningful distinct warnings is therefore impossible by construction. Aggregation cost is O(n log n) (one sort) with O(1) table lookups — no quadratic de-duplication scan exists.

### Status dimensions, atomicity, save eligibility

The report's `status` separates three independent dimensions, never a single boolean: `import` (`success` / `success-with-warnings` / `success-with-errors` / `failure`), `validation` (`success` / `failure` / `not-attempted`), `scheduling` (`success` / `failure` / `not-attempted`). `import: 'failure'` means NO authoritative document: for `.gproj`/MSPDI, the read sentinel (`GPROJ_READ`/`MSPDI_READ`) is absent (every fatal path returns before it is emitted); for MPP, a sidecar-stage error, a fatally failed MSPDI import, or the canonical rejection (atomic empty-document). Entity-level errors are `success-with-errors` (a document WAS constructed, degraded and diagnosed). `validation: 'failure'` ⇔ canonical-stage error diagnostics; `not-attempted` when the import failed first (the canonical-stage error that appears after a fatally malformed sidecar MSPDI output is the accepted pipeline's validation of the ATOMIC EMPTY document — an atomicity artifact the report records honestly). `scheduling` is only attempted over an authoritative document (the host pipeline's trivial scheduling of the atomic empty document after a fatal conversion is NOT an attempt).

`authoritative = import ≠ 'failure' ∧ validation = 'success'`. `saveEligibility = 'allowed'` iff authoritative: a canonically valid document may be saved with its degradation recorded in the report (never presented as lossless) — this matches the accepted corpus precedent (`mpp14assignmentcustom.mpp` imports with four entity-level `INVALID_MSPDI_DURATION` errors, remains canonically valid, and is saved/reopened in the PROJECT-018 I12 golden). Canonical validation errors prohibit save; fatal failures have nothing to save. The layer ADVISES (the report field); the accepted `serializeGproj` API is unchanged.

### MSPDI provenance additions (the only parser-side changes)

Two known PROJECT-015 limitations previously produced NO diagnostic (silent behavior gaps); exposing their provenance is the minimal change the brief sanctions ("unless required to expose accurate diagnostic provenance"):

- `MSPDI_PHYSICAL_PERCENT_COMPLETE_DROPPED` (warning, `dropped`): a task carrying a non-zero `<PhysicalPercentComplete>` that the canonical import does not reconstruct (the accepted importer never read it; PROJECT-016's exporter emits it for fidelity and warns about the same round-trip gap). Zero values are not diagnosed (nothing is lost — the canonical default).
- `MSPDI_BASELINE_CAPTURED_AT_APPROXIMATED` (warning, `approximated`): MSPDI carries no per-baseline captured date, so each created baseline slot's canonical `capturedAt` is the documented deterministic fallback (`<LastSaved>` → `<CreationDate>` → the project `<StartDate>` → the epoch default) — one warning per created baseline naming the fallback source actually used.

Both are additive: no accepted code is renamed, no mapping changes, no accepted test asserted exact full diagnostic arrays (verified: every existing suite asserts error-level filters or code-presence), and the full accepted suites pass unchanged.

### MPP source-version provenance (sidecar frame completion)

The sidecar protocol contract has always declared the success frame's optional `format` field ("Detected source format, e.g. `"MPP14"`"); the Java sidecar now populates it via MPXJ (`FileType` + `MppFileType`): `"MPP" + getMppFileType()` for MPP inputs. This is the byte-true container generation, NOT the filename or product provenance — the corpus itself proves the difference (`SubprojectA-9.mpp` is an MPP9-era name but an MPP14 container; the corpus manifest records both the provenance `format` label and the new `detectedFormat`, which the compatibility e2e asserts). The field is the compatibility report's honest `sourceVersion` for MPP imports; the protocol version stays 1 (additive optional field; the fake-sidecar fixtures without it remain valid).

### Scheduling injection (no new package edge)

`@genoffice/project-file` still does NOT statically depend on `@genoffice/project-scheduling` (the accepted dependency edge since PROJECT-014, asserted by the existing static guards). The scheduler is INJECTED as a structurally-typed runner (`CompatibilityScheduleRunner`); hosts/tests pass `schedule` directly (its `DerivedSchedule` is structurally assignable — the accepted test-layer precedent, now productized behind the pipeline boundary). The host package additionally exposes `importMppFromFileWithCompatibility` (full production pipeline + report; `MppFullImportResult` gains the additive `sourceFormat`).

### Scope exclusions

The import compatibility layer covers IMPORT pipelines only. The PROJECT-016 export code family (`INVALID_MSPDI_EXPORT`, `UNREPRESENTABLE_MSPDI_VALUE`, `UNSUPPORTED_MSPDI_EXPORT_FEATURE`, …) remains the export contract and is out of the import report's scope; the export-side limitations named by the brief (multi-note collapse, views/tables/filters/groups with no MSPDI representation, physicalPercentComplete round-trip) are covered by the accepted export diagnostics and documented in the PROJECT-016 section. No MPP export exists (PROJECT-019A decision record).

### Security and performance

Reports add no new sensitive data of their own: no environment variables, credentials, tokens, host paths, or process internals enter any report field the layer constructs (a security test asserts this over representative reports). Producer messages are passed through verbatim (deterministic provenance) — the host launcher's `MPP_INPUT_UNREADABLE` message intentionally carries the OS reason and input path (operator-facing provenance, unchanged accepted behavior). Aggregation is O(n log n) in the diagnostic count with O(1) classification lookups; large-project performance budgets remain PROJECT-048.

### CI

No workflow step change: the new tests run inside the existing `project-foundation.yml` steps (`Test project-file` now includes the 47-test compatibility suite; `Test project-mpp-host` includes the 9-test real-corpus compatibility e2e under the enforced network-isolation posture wherever the host provides the mechanism). One stability change to the host package's vitest config: `fileParallelism: false` — every host test file spawns real JVM sidecar conversions, and running files in parallel on a shared CI runner multiplies JVM contention (the root cause of the documented transient per-test timeouts in the PROJECT-019 CI round and the first PROJECT-020 round); sequential files keep at most one conversion pipeline on the runner at any time, and the compatibility e2e carries an explicit 120 s per-test timeout for its two-pass determinism runs.

## PROJECT-021 — Shared renderer core

Establishes the host-independent Project renderer boundary (`@genoffice/project-renderer-core`): the shared projection and control layer that desktop (Electron) and web hosts both consume, on top of the accepted canonical domain (`@genoffice/project-contracts`), the semantic engine (`@genoffice/project-engine`), and the derived schedule (`DerivedSchedule` from the scheduling authority). The renderer core is a projection/control layer, NOT a second source of Project semantics (architecture-lock §2/§11): every scheduling value it renders is a verbatim join of the scheduling engine's output, and every mutation it produces is a semantic `ProjectCommand` applied through the engine (lock §9). Complete views are out of scope (PROJECT-022+); this increment is the boundary they are built on.

### Package boundary and forbidden surfaces

Static dependencies are exactly `@genoffice/project-contracts` + `@genoffice/project-engine` (asserted by a discipline test and enforced by the extended CI foundation boundary grep, which now covers `packages/project-renderer-core` incl. its tests). The renderer core NEVER imports React/React DOM, Electron, Node/browser/HTTP modules (lock §13), and — the strongest form of "no scheduling authority" (lock §3/§6) — it does NOT statically depend on `@genoffice/project-scheduling` either: the scheduler is INJECTED as a structurally-typed `ScheduleRunner` (`(document) => DerivedSchedule`), the accepted PROJECT-020 compatibility-pipeline precedent. Hosts/tests pass `schedule` from the scheduling package (structurally assignable; test-layer import precedent); the package itself is structurally incapable of scheduling. File/host transports (`project-file`, `project-mpp-host`) are equally forbidden — host-specific APIs remain outside the renderer (R-009): loading/saving happens in hosts through the file boundary, and the renderer core receives an already-canonical `ProjectDocument`.

### View state, intents, and the reducer

`ProjectViewState` is the immutable, JSON-serializable interaction state: typed per-entity selections (ordered task selection with shift-extend `anchor` and keyboard `focus`, dependency selection, resource selection), the collapsed-summary set, the timeline viewport (`{start, finish}` ISO instants — the ONLY time values in view state; no task dates, float, criticality, or leveling results ever live here, lock §11), and the active canonical view/table/filter/group references (ids into the document-declared `ProjectView`/`ProjectTable`/`ProjectFilter`/`ProjectGroup` definitions — references are resolved at projection time; filter/group EXPRESSION evaluation is NOT part of this increment). Hosts translate platform events into the closed `ProjectViewIntent` union (select/toggle/extend, selectTasks, clearSelection, toggleCollapse/setCollapsed/collapseAll/expandAll, setViewport/scaleViewport/fitViewport, setActiveView/Table/Filter/Group) and dispatch them through the pure `reduceViewState(state, intent, {document, schedule})`. Determinism contract: the same `(state, intent, context)` always yields the same next state — no wall clock, no randomness, no `localeCompare` (discipline-scanned). Intents referencing entities that do not exist in the live document are deterministic no-ops (the state reference is returned unchanged); every reduction is reconciled against the document (`reconcileViewState`), so dead references are dropped immediately — hosts call the same `reconcileViewState` after any document replacement (command, undo/redo, file load).

Documented deterministic view rules: extend-selection replaces with the canonical outline-order range between the anchor and the target (both directions); the task-selection `anchorId`/`focusId` are always members of the surviving `taskIds` when present (the `TaskSelection` invariant): the reducer maintains it by construction, and `reconcileViewState` validates anchor/focus against the SURVIVING SELECTION, not mere document existence — a live-but-unselected anchor/focus (e.g. from a malformed externally restored state) is dropped, so the invariant holds after every reduction AND after external state restoration; the collapsed set contains ONLY summary task ids (the invariant `collapsed ⊆ summary TaskIds` — collapse is a summary-tree operation, a leaf has no subtree to hide): collapse intents referencing leaf or unknown tasks are deterministic no-ops (the state reference is returned unchanged), and `reconcileViewState` prunes non-summary entries — a collapsed summary whose subtree was deleted (the engine recomputes `summary`) or leaf ids restored from persisted host state — so the invariant holds after every reduction AND after external state restoration (hosts restoring a persisted view state run `reconcileViewState` once before first use); collapseAll collapses exactly the summary tasks; the viewport span is clamped to `[1 minute, 100 years]`; scaling keeps the focus instant stationary (default focus = window midpoint); `fitViewport` pads the project window 2% per side with a one-day minimum total span, preferring the derived schedule's project window, then the canonical properties window, then a 30-day default from the properties start.

### Projection

`projectDocumentView(document, schedule, state)` produces the immutable `ProjectViewProjection` both hosts render: visible `ProjectTaskRow`s in canonical outline order (a task is hidden iff an ancestor is collapsed), the project window (derived schedule's `projectStart`/`projectFinish` preferred, canonical properties as fallback — never invented), the schedule's diagnostics echoed verbatim, and the resolved active view definitions. Each row is a verbatim join: canonical Task fields (identity, WBS, outline level, name, flags, priority, duration, percent-complete, constraint, deadline), `resourceNames` projected from the assignment array in document order (deduplicated, first occurrence kept — a pure structural projection), the collapse flag, and — the scheduling join — `row.schedule`, the EXACT `TaskSchedule` object from `DerivedSchedule.taskSchedules[taskId]`, joined BY REFERENCE (asserted by test). The projection layer owns no scheduling semantics: a task without a schedule entry carries `schedule: undefined` (no scheduler wired, schedule failure, unschedulable task) and dates are never invented, defaulted, or degraded. Projection is pure and deterministic (3×-repeat byte-identity test), never mutates its inputs, and is linear in task count.

### Timeline math

Pure, locale-free viewport and axis math shared by both hosts (renderer parity): `viewportFraction`/`viewportInstant` (instant ↔ fraction in `[0, 1]` over the visible window — hosts multiply by pixel width; the core never sees a pixel), `scaleViewport` (factor scaling around a focus instant, integer-millisecond arithmetic, span clamps), `fitViewport`, and `buildTimeAxis` (contiguous `[start, finish)` bands at day/week/month level chosen deterministically from the span: day < 93 days, week < 3 years, month beyond; weeks Monday-aligned UTC, months calendar-aligned UTC; bands carry instants and level but NO formatted label — label formatting is a host/locale concern, keeping the core locale-free). All arithmetic is UTC epoch-milliseconds over ISO instants (the canonical time model, lock §5); calendar working-time bands belong to PROJECT-025 and will be projected from document calendars, never re-derived here.

### Structural command builders

The shared translation for outline gestures whose canonical mapping must not be re-implemented per host: `nextTaskIdentity` (deterministic `t{n}` local-id allocation — smallest n greater than every existing `t{n}` — plus `max(uid)+1`), `buildCreateTaskCommand` (identity + canonical creation defaults — fixedDuration, 480 working minutes, priority 500; every derived field is recomputed by the engine on acceptance — with parent wiring from the explicit insert position `TaskInsertPosition`, a discriminated union expressing EXACTLY the positions the frozen `CreateTask` command can execute: `{kind: 'lastRoot'}` appends as the last root task, `{kind: 'lastChildOf', parentId}` appends as the LAST child of that parent — PROJECT-007 acceptance semantics; the command union has no insertion index, so an immediately-after-anchor row insert is NOT expressible and the renderer core does not invent a second insertion model — a nonexistent `parentId` surfaces as the engine's `MISSING_PARENT` rejection, keeping the engine the single validation authority), `buildCreateTaskInSiblingGroupCommand` (the outline gesture "create a new task in this row's sibling group" mapped ONCE in the shared core: the anchor's parent — or the root level for a root anchor — with the new task appended as the group's LAST member, i.e. after the anchor's LAST sibling, never immediately after the anchor; `undefined` for an unknown anchor — a disabled gesture, the same contract as indent/outdent), `buildIndentCommand` (target parent computed with the ENGINE's own `previousSiblingOf`, so the builder and the PROJECT-007 acceptance rule can never disagree; `undefined` when no preceding sibling exists — a disabled gesture, never an invented parent), `buildOutdentCommand` (`undefined` for roots), and `buildDeleteSelectionCommands` (ancestors absorb their selected descendants — outline-deletion semantics; deletions emitted in reverse canonical outline order). Field-level edits (rename, duration, constraints, percent complete) are 1:1 `ProjectCommand` mappings hosts construct directly from the contracts; their interaction surfaces are PROJECT-023/024 scope. The renderer core never mutates a document — builders only produce command values.

### Renderer session (command-to-renderer integration)

`ProjectRendererSession` is the immutable control layer binding the semantic engine to a host view: the current canonical `ProjectDocument`, the `DerivedSchedule` of that exact document (present iff a scheduler is wired), the command history, and monotonic counters (`revision` per document change; `commandSeq` per ACCEPTED command, driving deterministic `c{n}` command ids). `applyRendererCommand` delegates to the canonical `applyProjectCommand` verbatim: rejected commands return the SAME session reference with the engine's diagnostics surfaced (hosts render toasts from them); accepted commands journal the entry (command, result, and the immutable before/after document/schedule snapshots), clear the redo future, re-schedule through the injected runner, and bump the revision. Because `applyProjectCommand` never mutates its input (accepted commands produce new documents), the snapshots are stable references.

Undo/redo is SNAPSHOT-based: undo restores the entry's `before` document AND its exact derived schedule; redo restores the `after` snapshot — byte-identical restoration, total over EVERY accepted command including the deliberately non-invertible `OutdentTask`, `CreateBaseline`, and `LevelResources`, for which the engine documents that undo "requires a host-level document snapshot" (the renderer core IS that shared host layer; the engine's per-command `inverse` values remain available on journal entries for hosts wanting command-level undo semantics). This is still command-history-based undo (lock §9): the history stack drives which engine-produced state is restored; the session never constructs documents itself and never invents semantics from raw state mutation. The engine now implements the ENTIRE frozen command union (`SetTaskDuration`/`SetTaskFinish` landed with PROJECT-023; the four dependency commands landed with PROJECT-024 — PROJECT-021 documented the earlier subset state and the deferrals); the session applies whatever the engine accepts, and the `UNSUPPORTED_COMMAND` diagnostic is now the runtime safety net against malformed command VALUES arriving through untyped boundaries (the pre-024 suites used the then-unimplemented `AddDependency` as their example; the 023/024 suites updated those tests to malformed-value payloads) — no renderer-side command semantics exist.

### Persistence boundary

View state and sessions are host-owned UI state: hosts may persist them in their own workspace preferences, but the `.gproj` adapter never serializes renderer state (the accepted PROJECT-014 rule — no React component state, DOM geometry, scroll positions, or viewport windows in `.gproj`). The injected `scheduleRunner` is a function and intentionally not serializable state; hosts re-inject it on restore. Canonical view definitions (`ProjectView`/`ProjectTable`/`ProjectFilter`/`ProjectGroup`) remain document state — the renderer references them by id and never writes them.

### CI

`project-foundation.yml` extends from 19 to 21 steps: the foundation boundary grep now includes `packages/project-renderer-core` (src AND tests — the discipline suite itself uses only vitest + `?raw` module sources, no `node:` imports), and `Typecheck project-renderer-core` + `Test project-renderer-core` run after the project-file steps (before the Java setup). The renderer-core vitest suite imports `@genoffice/project-scheduling` at the TEST layer only (the accepted project-file precedent) to prove session/schedule integration with the real scheduling authority. One CI-stability measure ships with this increment's correction round (the documented PROJECT-019/020 transient class): `e2e-real-corpus.test.ts` — whose every test spawns a real JVM sidecar conversion under the enforced isolation wrapper — now carries explicit per-test budgets (30 s single-pass, 120 s for the two-pass I12 determinism runs), the exact discipline the PROJECT-020 CI round applied to `e2e-compatibility.test.ts`; the file's FIRST cold conversion exceeded the 5 s vitest default on slow shared pull_request runners while the identical suite passed on the push runner and locally (twice, at 5015/5017 ms — runner variance, not a code defect; no production code changed).

## PROJECT-022 — Gantt / task grid / timeline views

Establishes the five Gantt view surfaces — `ProjectTaskGrid`, `ProjectTimeline`, `ProjectGanttBars`, `ProjectDependencies`, `ProjectMilestones` — as HOST-INDEPENDENT view models built on the accepted PROJECT-021 projection/state/session layer: `buildGanttView(document, projection, state, layout)` composes the synchronized, virtualized surface both hosts render. The views are pure projection and interaction geometry over `ProjectDocument` + `DerivedSchedule` + `ProjectViewState`; no scheduling logic enters the view layer (lock §11): every date a view model carries is a verbatim echo of the projection's by-reference schedule join, and no geometry is ever invented for a row without the schedule value it needs.

### Boundary and geometry discipline

The view models live INSIDE `@genoffice/project-renderer-core` (`src/views/**`, the accepted PROJECT-020 `src/compatibility/**` in-package precedent): no new package, no new package edge, static dependencies stay exactly `@genoffice/project-contracts` + `@genoffice/project-engine`, and no CI workflow change is needed — the 21-step foundation gate already typechecks/tests the package and its boundary grep covers the whole package (the discipline suite's raw-source scan now globs `src/**/*.ts`, covering the views subfolder). All geometry is expressed in FRACTION space (x ∈ [0, 1] across the timeline viewport) and ABSOLUTE visible-row-index space (y): hosts multiply by pixels, the core never sees a pixel or a DOM API (discipline-scanned), and grid cells carry STRUCTURED unformatted values — durations as working minutes, dates as ISO instants, predecessors as typed link records — so label/number/date formatting stays a host/locale concern (the locale-free rule). Layout inputs (`firstRow`, `visibleRows`, `overscan`) are HOST scroll state passed as plain arguments per render — logical rows, never pixels, and deliberately NOT `ProjectViewState` (ephemeral host layout is not persisted interaction state; the PROJECT-014/021 persistence boundary is unchanged).

### Virtualization and synchronization

`buildRowWindow(rowCount, {firstRow, visibleRows, overscan})` computes the ONE clamped inclusive row window both panes render from (`lastIndex < firstIndex` is the canonical empty window; all inputs are clamped deterministically). Synchronization is BY CONSTRUCTION: `buildGanttView` computes the window once and both panes — grid rows, timeline rows, bars, milestones, and dependency links — address rows by the same absolute indices over that same window, so the panes cannot drift (the "synchronized virtualized core surface" acceptance). Column virtualization is deliberately OUT of scope: the canonical `ProjectTable.columns` carries field-name strings and NO width data, so there is no canonical column geometry to virtualize against — horizontal column scrolling is a pure host layout concern. Time virtualization is the accepted viewport (`TimelineViewport` + `buildTimeAxis`), unchanged.

### Task grid (ProjectTaskGrid)

Columns resolve from the ACTIVE canonical table (the projection's `activeTable`, i.e. the state's `activeTableId`): the table's field-name strings (the `.gproj` convention: `name`, `duration`, `start`, `finish`, …) map onto the supported field set (`rowNumber`, `taskName`, `duration`, `start`, `finish`, `percentComplete`, `predecessors`, `resourceNames`, `wbs`, `outlineLevel`, `priority`, `uid`); an unrecognized field name maps to an `unsupported` cell carrying the raw string (never a crash, never invented data). When no table is active — or the active table declares no columns — the documented DEFAULT set applies (`rowNumber, name, duration, start, finish, predecessors, resourceNames`, the Microsoft Project Entry-like set). `rowNumber` is the task's CANONICAL 1-based document position: stable under collapse (hidden rows keep their numbers, exactly the MS Project ID-column behavior — never a recomputed visible-sequence number). Scheduling-value cells follow the schedule-first precedence the projection's project window established: duration/percent cells carry `schedule.duration`/`schedule.percentComplete` (the derived, rolled-up echoes) when a schedule exists and the canonical task field otherwise; start/finish cells carry `schedule.scheduledStart`/`scheduledFinish` and are `empty` when the row has no schedule — dates are never invented. `taskName` cells carry the name plus the structural indent/expander data (outlineLevel, summary, milestone, collapsed). `predecessors` cells list the row's incoming dependencies in document order as structured links (dependency id, predecessor TaskId + uid for MS-style ID display, type, lagMinutes). Grid rows join the projection rows BY REFERENCE.

### Gantt bars, milestones, and dependency links

`ProjectGanttBars`: a row carries bar geometry iff its schedule has BOTH `scheduledStart` and `scheduledFinish` AND `finish > start` — positions are the clamped viewport fractions with explicit `startsBefore`/`finishesAfter` clipping flags; summary rows render `kind: 'summary'` from the same date source (the rolled-up window); the progress point is a documented linear interpolation over the RAW (unclamped) span — `start + (percent/100) × (finish − start)` with the percent source `schedule.percentComplete ?? task.percentComplete` clamped to [0, 100] — a rendering convention, not a scheduling computation. `ProjectMilestones`: a row carries diamond geometry iff its schedule has a `scheduledStart` AND the row is milestone-like (the `milestone` flag, OR zero task duration, OR a zero-span schedule window); the diamond instant is `scheduledStart`; instants outside the viewport clamp with `beforeViewport`/`afterViewport` flags. The two rules are ORTHOGONAL: a flagged milestone with a real schedule span keeps BOTH a bar and a diamond (hosts render both). `ProjectDependencies`: links are pure visibility projection — the documented edge anchors per type (FS finish→start, SS start→start, FF finish→finish, SF start→finish), endpoint resolution to the task's OWN row when visible and in-window, to its NEAREST VISIBLE ANCESTOR's row when hidden by collapse (the collapsed summary's rolled-up edge; Microsoft Project behavior), and OMISSION when an endpoint is scrolled out of the window, when both endpoints resolve to the same row (a link entirely inside one collapsed subtree), or when the resolved row lacks the needed schedule instant — never an invented position. The canonical endpoint TaskIds stay on the link record even after ancestor resolution. The route is the deterministic four-point elbow `[(xFrom, yFrom), (xm, yFrom), (xm, yTo), (xTo, yTo)]` with `xm = (xFrom + xTo) / 2` — hosts render the polyline; link hit-testing needs a pixel tolerance against that polyline and stays a host concern.

### ProjectTimeline and hit testing

`buildTimeline` composes the right pane over the shared window: the viewport echoed by reference, the deterministic `buildTimeAxis` bands at the span-derived level, the in-window rows (projection rows joined by reference), and the three geometry surfaces above; an unparseable/degenerate viewport yields an EMPTY model rather than invented values. `hitTestGantt(timeline, {rowIndex, fraction}, tolerance?)` is the geometry→entity inverse for pointer interaction: milestones resolve before bars (the smaller target wins at overlapping positions), the tolerance widens both hit rectangles by a fraction-space epsilon (hosts pass `pixelTolerance / timelineWidth`; 0 = exact containment). What hosts DO with a hit — the selection intents (PROJECT-023) and the task-cell editing activation — follows the documented two-step pointer contract; PROJECT-024 threads the interaction state into the link surface the same way (`buildTimeline`'s optional `state` parameter joins `selected`/`editingField` onto the dependency links).

### Scope discipline

Out of scope and NOT implemented: any host UI/Electron/web code (PROJECT-027/028 — the dependency graph places both shells AFTER 022; the browser/desktop E2E verification for this surface therefore lands with the shells, when there is a DOM to drive — this increment's evidence is the component + architecture-discipline suite over the package, the accepted PROJECT-021 precedent, plus a test-layer integration run against the REAL scheduling package), column-width virtualization (no canonical widths), editing surfaces (023/024), calendar working-time bands (025 — will be projected from document calendars, never re-derived), critical/resource visualization (026), baseline/tracking bars (035), split views, and any scheduling computation anywhere in the view layer. Determinism contract: every builder is pure — the same inputs always produce byte-identical view models (3×-repeat tested), no wall clock, no randomness, no `localeCompare` (discipline-scanned), inputs never mutated, projection rows and schedules joined by reference.

## PROJECT-023 — Selection / editing

Establishes semantic task editing through the existing command/session pipeline — the acceptance target is NOT a visually editable grid but the full chain `user gesture → renderer-core selection/edit intent → semantic ProjectCommand → project-engine → ProjectDocument → canonical scheduler → projection refresh`. The renderer remains a projection and interaction layer (lock §9/§11): it NEVER mutates a canonical task — every edit leaves the renderer as a `ProjectCommand` value and reaches the document only through the engine, exactly like every other mutation.

### Engine: the deferred `SetTaskDuration` / `SetTaskFinish` commands

PROJECT-021 documented that the engine implemented a subset of the frozen command union with `SetTaskDuration`/`SetTaskFinish` deferred to this increment; both are now implemented (the union itself is unchanged — frozen):

- `SetTaskDuration` sets the working `duration` on a LEAF task (one of the scheduler's two leaf inputs; the derived finish is `start + duration` in working time). Rejections: `MISSING_TASK`; `SUMMARY_DURATION_NOT_SETTABLE` (the summary duration is a derived roll-up of the subtree — the accepted `SetPercentComplete` precedent: accepting the edit would store a value every derivation overwrites, an invisible edit); `INVALID_DURATION` for anything outside the canonical working-minute domain — non-negative INTEGER minutes (the domain the MSPDI interchange round-trip documents; fractional or negative values are already diagnosed as `UNREPRESENTABLE` at the file boundary, so the engine never admits them into the document). Zero is valid (milestones are zero-duration). The inverse always exists (duration is a required field).
- `SetTaskFinish` sets the STORED `finish` on a LEAF task. Honest semantics: the stored finish is an interchange echo (the MSPDI importer/exporter round-trips it verbatim), NOT a scheduling input — the scheduler derives the authoritative finish from the leaf's `start` + `duration` and never reads the stored field. Editing the SCHEDULED finish therefore goes through start/duration edits; this command pins the stored field exactly like `SetTaskStart` pins the stored start, and the engine performs no working-time arithmetic (computing a duration from a finish would require the task's resolved calendar — the scheduler is the sole scheduling authority). Rejections: `MISSING_TASK`; `SUMMARY_FINISH_NOT_SETTABLE` (the summary finish is a roll-up of the children's finishes); `INVALID_DATE` (the `SetTaskStart` validation). The inverse is emitted only when a previous finish existed — the frozen command shape requires a string payload, so no command CLEARS the field; undo of a first-ever pin requires a host snapshot (the `SetTaskStart` precedent, satisfied by the renderer session's snapshot undo).

### Editing model (which cells edit, what the draft is)

The editable fields are the four PROJECT-022 grid fields `taskName`/`duration`/`start`/`finish` (`EditableTaskField` — the 022 field ids). Editability is a projection-layer fact derived from the canonical task: `taskName` on EVERY row; `duration`/`start`/`finish` on LEAF rows only (a summary's scheduling values are derived roll-ups). Field scope discipline: percent-complete, constraint, and deadline edits remain direct command construction per the PROJECT-021 rule (their interactive surfaces are later increments); dependency editing landed with PROJECT-024 (its own editing model mirrors this one — see the 024 section); milestone-ness is not command-editable at all (no command in the frozen union flips the `milestone` flag — a duration edit on a flagged milestone keeps bar+diamond orthogonality, the documented 022 rule).

The active cell edit is `TaskEditing` — `{taskId, field, draft}` — carried as the ADDITIVE optional `editing` slice of `ProjectViewState` (one intent union, one reducer, one state; JSON-serializable like the rest). At most one edit is live (beginning a new edit replaces an active one); activating an edit (`beginTaskEdit`) is a deterministic no-op for an unknown task or a non-editable field, and otherwise SELECTS the edited row (the Microsoft Project cell-edit gesture: the selection becomes exactly that row) with anchor and focus on it. `reconcileViewState` drops an edit whose task no longer exists, so the state is self-consistent after any document replacement.

The draft protocol is CANONICAL TEXT, initialized from the DISPLAYED cell value (the 022 schedule-first precedence — the user edits what the grid shows): the name verbatim; the duration as decimal working minutes of `schedule.duration ?? task.duration`; start/finish as the schedule's scheduled instants, EMPTY when the cell is empty (dates are never invented). Display formatting stays a host/locale concern (the locale-free rule); the draft is protocol, not presentation.

### Commit translation and the single validation authority

`commitTaskEdit(document, editing)` is the pure draft→command translation with a closed outcome union: `apply` (the command), `noChange` (the built payload equals the task's CURRENT stored field — nothing dispatched, so no no-op command is ever journaled), `invalid` (a syntax/structural failure: `missingTask` or `unparseableDuration`), or `none` (no active edit). The layer split is strict: the renderer owns TEXT SYNTAX only — a duration draft must match the canonical decimal rule (optional minus, digits, optional `.`-fraction; no whitespace, exponents, or locale separators), while negativity/fractionality PARSE and become commands the ENGINE rejects (`INVALID_DURATION`); names and instants pass through VERBATIM (the engine's `INVALID_DATE` is the single date validator). The `noChange` comparison is against the STORED field, not the displayed cell — a start edit initialized from the derived `scheduledStart` of a dependency-delayed task still produces a real command (it pins the candidate start to the displayed instant).

### Keyboard navigation

`moveTaskFocus` (direction `up`/`down`/`first`/`last`, optional `extend`) walks the task focus through the VISIBLE row order — the canonical outline order minus subtrees of collapsed summaries (the projection's visibility rule), so focus never lands on a hidden row. With nothing focused, any direction bootstraps by selecting the first visible row; moves past the first/last visible row are reference-equal no-ops (clamped). `extend` applies the ACCEPTED 021 shift-extend rule verbatim — the selection becomes the canonical outline-order range from the anchor (or the current focus when no anchor exists) to the target; no new selection semantics are introduced, and the anchor/focus selection invariant holds after every step.

### Selection/edit reflection in the projection (view integration)

The projection reflects the interaction state: every `ProjectTaskRow` carries `selected` (membership in `state.tasks.taskIds`), `focused` (`state.tasks.focusId`), and — present iff the active cell edit targets the row — `editingField` (`state.editing.field`). These are pure by-value echoes of the view state joined onto the row exactly the way the collapse flag always was (lock §11-clean: no scheduling value is involved), so grid and timeline rows — which carry the projection rows BY REFERENCE — inherit the flags into every PROJECT-022 view model, and `buildGanttView` output is selection-aware with no second join. The edit DRAFT is deliberately NOT projection state: it is live user input en route to a command and stays on `state.editing.draft`; the projection reflects the edit TARGET only (which row, which field).

The documented hidden-selection policy: HIDING IS NOT DESELECTING. A task selected and then hidden by a collapsed ancestor keeps its selection in the view state (the visibility rule projects rows; the selection is TaskId-based and never keyed on visibility, row number, or DOM position); it is simply absent from `rows` while hidden and re-projects as a selected row when expanded. The same holds for members of an outline-order range that cross a collapsed subtree: the accepted 021 extend rule selects them in state, the visible projection flags the visible members, and the hidden members re-project on expand.

Host pointer interaction is a two-step contract over existing pieces — the renderer core provides the geometry→entity inverse (`hitTestGantt` → `{kind, taskId}`) and the standard selection intents; the host maps the hit to `selectTask` (or `beginTaskEdit` for an activation) exactly as it maps keyboard events. The integration suite proves the wiring deterministically at the core level: grid row → `selectTask` → reflected flags in the rebuilt grid AND timeline panes (the synchronized-surface guarantee), bar hit → `selectTask` → the selected row, milestone hit → `selectTask` → the selected row, and the collapsed-row retention/expand round-trip. No geometry algorithm changed to support selection (the 022 geometry is untouched; only the row's state echo was added).

### The commit flow (command/session integration)

`commitTaskEditThroughSession(session, state)` composes the accepted machinery in one deterministic call — pure translation, `applyRendererCommand` (engine application, journaling, rescheduling through the injected runner), then `reduceViewState(endTaskEdit, …)` against the session's current document (editor end + view-state reconciliation): selection preservation after accepted AND rejected mutations is BY CONSTRUCTION. Behavior contract: an ACCEPTED command journals the entry (`c{n}`), replaces document+schedule, and reconciles the state (surviving selection kept, dead references pruned); a REJECTED command returns the SAME session reference with the engine's diagnostics surfaced verbatim, the editor still ends (the Microsoft Project revert-cell behavior — the host renders a toast from the diagnostics), and the state reconciles against the unchanged document; `noChange`/`invalid` dispatch no command at all; `none` is a pure no-op (same session AND state references). Undo/redo is the accepted session snapshot machinery — total over every committed edit (including the no-inverse first-ever `SetTaskFinish` pin), restoring the exact prior document AND its exact derived schedule.

### Scope discipline and evidence

Out of scope and NOT implemented at 023: dependency editing (delivered by PROJECT-024 — the next section; the 023 cycle left the frozen `AddDependency`/`RemoveDependency`/`ChangeDependencyType`/`ChangeLag` commands unimplemented engine-side and used `AddDependency` as its UNSUPPORTED_COMMAND example — both facts are now historical), percent-complete/constraint/deadline interactive surfaces, any column/cell cursor between columns (host layout — the 022 no-canonical-widths rule), advanced dialogs (the work-item's explicit exclusion), clipboard/undo-menu UI, and any host UI/Electron/web code (PROJECT-027/028 — the shells the dependency graph places after this surface; the browser/desktop E2E lands with them, when a DOM exists to drive). The work-item's "E2E + command audit" verification is satisfied by the command audit over the REAL scheduling package at the test layer (the accepted 021/022 precedent): the full pipeline test drives gesture → intent → command → engine → document → scheduler → projection refresh and asserts the semantic effect (the derived finish moves exactly as the scheduler derives it). The evidence also carries the golden scenario battery S01–S12 (single/multi/range/collapsed-hierarchy/keyboard-range selection; rename with stable identity and recomputed schedule; percent completion 0/partial/100/invalid through the engine authority; the rejected edit leaving everything unchanged; edit + reschedule; undo/redo restoring document + schedule + selection; delete-selected reconciliation; and the 3× byte-identical repeated interaction sequence) plus the integration battery (projection reflection, grid/bar/milestone → selection, selection preservation through create/indent/outdent, and the UNSUPPORTED_COMMAND path with the document and selection untouched) — each golden asserts the complete observable tuple: selection taskIds, anchorId, focusId, the edit target where applicable, the ProjectDocument, the DerivedSchedule, and the projection state. Determinism contract: the reducer, the translation, and the flow are pure — the same inputs always produce byte-identical outcomes (3×-repeat tested), no wall clock, no randomness, no `localeCompare` (discipline-scanned), inputs never mutated; the discipline suite now also scans the editing modules for scheduling-derived state (lock §11).

## PROJECT-024 — Dependency editing

Establishes semantic dependency editing through the existing command/session pipeline over the canonical dependency graph — the strict semantic path `dependency gesture → renderer intent/builder → semantic ProjectCommand → project-engine → ProjectDocument → schedule() → projection/link refresh`. The renderer remains a projection and interaction layer (lock §9/§11): it NEVER mutates a dependency and NEVER invents dependency semantics — cycle topology, reference validity, link uniqueness, and the summary↔descendant rule are the ENGINE's authority; the renderer only builds command VALUES (identity allocation, defaults, gesture guards) and translates drafts.

### Engine: the four frozen dependency commands

PROJECT-021 documented the dependency commands as deferred engine work; all four are now implemented (the union itself is unchanged — frozen):

- `AddDependency` appends the link in canonical document order. Early rejections (single crisp diagnostics — the same codes the document validator reports, with the post-command validator as the safety net): `DUPLICATE_DEPENDENCY_ID` (id exists); `MISSING_TASK_REFERENCE` (either endpoint is not a task); `SELF_DEPENDENCY` (predecessor === successor); `INVALID_LAG` (non-integer minutes); `DUPLICATE_DEPENDENCY_LINK` (another dependency already owns the `(predecessor, successor, type)` key); `SUMMARY_DEPENDENCY` (either endpoint is an ancestor of the other — the roll-up fixpoint rule, checked by parent-chain walk in both directions); `DEPENDENCY_CYCLE` (adding predecessor→successor cycles iff the successor already reaches the predecessor over the existing edges — a deterministic breadth-first reachability check over the raw edge graph, the single-edge incremental form of the validator's Kahn check). `affectedTaskIds` = both endpoints; the inverse is `RemoveDependency` of the new id.
- `RemoveDependency` filters the link by identity. `MISSING_DEPENDENCY` when absent; `affectedTaskIds` = the removed link's endpoints; the inverse carries the FULL record (id, endpoints, type, lag) so undo restores the link exactly without a host snapshot.
- `ChangeDependencyType` re-keys the link. `MISSING_DEPENDENCY`; `INVALID_DEPENDENCY_TYPE` for a payload outside FS/SS/FF/SF (the runtime domain guard — the union type cannot stop a malformed value at a boundary); `DUPLICATE_DEPENDENCY_LINK` when another dependency already owns the target key (a same-type change keeps the key and is an accepted idempotent write — the `SetTaskDuration` precedent; the renderer's `noChange` check keeps no-ops from ever being dispatched). Type changes never alter the edge topology, so no cycle check applies. The inverse is the previous type.
- `ChangeLag` replaces the integer working-minute lag (negative = lead — the canonical scheduler domain). `MISSING_DEPENDENCY`; `INVALID_LAG` for non-integer values. The inverse is the previous lag.

Dependencies carry NO derived fields, so every accepted mutation is a pure array edit (append / filter / map); the scheduler remains the sole scheduling authority — the derived dates move exactly as `schedule()` derives them (FS/SS/FF/SF forward-pass equations with working-time lag in the predecessor's calendar; SF/FF bound the successor's required start and are floored by the project start).

### Dependency command builders (the gesture layer)

The shared translation for the link gestures (the PROJECT-021 builder discipline): `nextDependencyIdentity` (deterministic `d{n}` local-id allocation — smallest n greater than every existing `d{n}`; ids that do not match the pattern, e.g. the MSPDI importer's `d-{succ}-{pred}-{type}`, are ignored, so allocation never collides with imported links and never renumbers them), `buildAddDependencyCommand` (identity + the canonical creation defaults — `FS` with zero lag, the Microsoft Project link defaults — with optional explicit type/lag; `undefined` for the gestures the engine would ALWAYS reject on endpoint structure alone: unknown predecessor/successor or a self-referencing link — the disabled-gesture contract of indent/outdent, never an invented command), and `buildRemoveDependencySelectionCommands` (one `RemoveDependency` per EXISTING selected id in canonical document order; unknown ids dropped). Cycle risk, duplicate links, and the summary rule are NOT builder concerns — those engine diagnostics surface through the session like every other rejection.

### Dependency editing model

The dependency-editing model mirrors the 023 task model: the editable fields are `type` and `lag` (`EditableDependencyField` — both editable on EVERY dependency; there is no summary/leaf distinction in the dependency domain). The active edit is `DependencyEditing` — `{dependencyId, field, draft}` — the ADDITIVE optional `dependencyEditing` slice of `ProjectViewState`. The SINGLE-EDITOR rule holds across both families: at most one editor is live (beginning a task edit ends an active dependency edit and vice versa); activating a dependency edit (`beginDependencyEdit`) is a deterministic no-op for an unknown dependency and otherwise SELECTS the edited link (the 023 cell-edit gesture's dependency analog). `reconcileViewState` drops a dependency edit whose link no longer exists.

The draft protocol is CANONICAL TEXT initialized from the stored values: the type code verbatim (`FS`/`SS`/`FF`/`SF` — protocol, not a localized label; display formatting stays a host concern), the lag as decimal minutes. `commitDependencyEdit(document, editing)` is the pure translation with the closed outcome union (`apply`/`noChange`/`invalid`/`none`): a type draft must be exactly one of the four codes (`unparseableDependencyType` otherwise); a lag draft must match the canonical decimal rule (`unparseableLag` otherwise) while fractional values PARSE and become commands the ENGINE rejects (`INVALID_LAG` — the single-validation-authority split, the `SetTaskDuration` precedent); `noChange` compares against the stored field.

`commitDependencyEditThroughSession(session, state)` composes the accepted machinery exactly like the task flow — pure translation, `applyRendererCommand`, editor end, view-state reconciliation — with the same behavior contract: ACCEPTED commands journal and refresh (an accepted lag/type edit keeps the edited link selected — it survives); REJECTED commands return the SAME session reference with verbatim diagnostics; `noChange`/`invalid` dispatch nothing; `none` is a pure no-op. Undo/redo is the accepted snapshot machinery (total: `AddDependency`/`RemoveDependency` round-trip through their full-record inverses byte-identically).

### Link interaction-state reflection (view integration)

The dependency-link surface reflects the interaction state the same way the projection rows do: every `ProjectDependencyLink` carries `selected` (membership in `state.dependencies`) and — present iff the active dependency edit targets the link — `editingField` (`state.dependencyEditing.field`). The reflection is a pure by-value echo joined in `buildDependencies` through `buildTimeline`'s ADDITIVE optional `state` parameter (`buildGanttView` passes the state it already receives; callers that omit it get `selected: false` and no `editingField` with IDENTICAL geometry — reflection is never a geometry input, and the 022 link tests pass unchanged). The edit DRAFT stays live state; only the edit TARGET projects. A link whose row is scrolled out of the window produces no link (the 022 visibility rule); its selection stays in the view state and re-projects when it scrolls back in.

### Scope discipline and evidence

Out of scope and NOT implemented: link drag-and-drop creation geometry (a host pointer concern — the shared builder IS the creation gesture's canonical mapping; hosts map their pointer events to it exactly as they map keyboard events to intents), the predecessor-column grid cells' interactive surface (the 022 grid projects `ProjectGridPredecessorLink` read-only; editing a predecessor CELL resolves to the same four commands), resource leveling (the work-item's explicit exclusion — `LevelResources` already exists and is untouched), and any host UI/Electron/web code (PROJECT-027/028 — the browser/desktop E2E lands with the shells, the accepted 021/022/023 precedent). The work-item's "E2E + schedule verification" is satisfied by the command audit + component suite over the REAL scheduling package at the test layer: the full pipeline tests drive gesture → builder/intent → command → engine → document → scheduler → link-surface refresh and assert the semantic scheduling effect (the derived start moves exactly as the scheduler derives it, for every dependency type and for lag/lead). The evidence carries the golden scenario battery D01–D12 (create FS/FS+lag/FS+lead/SS/FF/SF; lag edit; type edit; the rejected edit leaving everything unchanged; the cycle rejection; undo/redo restoring document + schedule + selection; and the 3× byte-identical repeated interaction sequence) plus the integration battery (the full semantic path; link-surface refresh for create/retype/remove with the edge-anchor re-route; the reflection echo and its geometry-neutrality; cycle/self-reference/invalid-reference rejection with the document and selection untouched; selection preservation through create/remove/edit; and deterministic 3× replays) — each golden asserts the complete observable tuple: dependency selection, task selection, the edit target where applicable, the ProjectDocument, the DerivedSchedule, and the link surface. Determinism contract: the builders, the reducer, the translation, and the flow are pure — the same inputs always produce byte-identical outcomes (3×-repeat tested), no wall clock, no randomness, no `localeCompare` (discipline-scanned), inputs never mutated; the discipline suite's lock §11 scheduling-free scan covers `dependency-editing.ts`, and the public-surface lockstep carries the PROJECT-024 exports.

## PROJECT-025 — Calendar visualization

Establishes the calendar visualization surface over canonical calendar semantics the Project domain already produces — the authorized semantic path verbatim: `ProjectDocument calendars + canonical scheduling semantics → project-renderer-core calendar projection → Timeline / Gantt visualization`. The renderer is a projection layer only: there is NO second calendar engine in `project-renderer-core`. Every "which instants are working time" answer comes from the INJECTED canonical working-time query (the `ScheduleRunner` injection precedent — lock §3/§6 keep the package statically scheduling-free); weekday resolution, exception matching, inheritance resolution, and period validation remain the scheduling engine's authority (lock §8), and the engine's own degradation boundary (`schedule()` catches calendar errors into diagnostics) is mirrored exactly.

### The injected canonical working-time query

`CalendarWorkingTimeQuery` is a structural function type — `(calendars, calendarId, start, finish) → readonly CalendarWorkingInterval[]` — satisfied by the canonical binding hosts write once:

```
(calendars, calendarId, start, finish) =>
  workingIntervals(resolveCalendar({ calendars }, calendarId), start, finish)
```

The query receives the document's calendar array (the canonical book data) so the binding stays document-independent and pure. The renderer core never imports `@genoffice/project-scheduling` (the existing architecture discipline scan); the binding is host code (and, at the test layer, the REAL package — the accepted project-file/scheduler precedent). `CalendarViewInput` — `{workingTime?, calendarId?}` — is the per-render input threading the query (and an optional background-calendar choice, defaulting to the document's `defaultCalendarId`) into `buildGanttView`/`buildTimeline`; like the layout inputs it is ephemeral render state, never persisted `ProjectViewState`.

### The calendar projection (`src/calendar.ts`)

Three pure surfaces:

- **The catalog** (`buildCalendarCatalog`) — a pure structural echo of `ProjectDocument.calendars` in document order: verbatim `Calendar` references, `name`, `baseCalendarId`, the declared working weekdays (0=Sunday..6=Saturday — the canonical `getUTCDay` convention; a weekday declared with an empty period list is a non-working day), declared minutes per weekday (sums of the declared period bounds — a display aggregation of the DECLARATION, never an evaluation over dates), `declaredWeeklyMinutes`, `exceptionCount`, and the direct task-reference count (`task.calendarId` equality; the RESOLVED calendar per task stays the scheduling authority's answer). The catalog needs NO evaluator — it is the document-only surface hosts render as the calendar view/legend.
- **The band classification** (`classifyCalendarBands`) — pure interval algebra over the evaluator's output: parse, clip to the window, drop empty spans, sort by start (then finish), merge overlapping/adjacent spans, complement the gaps into contiguous `[start, finish)` `ProjectCalendarBand`s classified working/non-working. No calendar semantics anywhere — the module never decomposes a date (no weekday/date-part extraction), never reads `workingWeek` for evaluation, never matches exception dates; the discipline suite scans for exactly those markers (without date decomposition no calendar evaluation is possible). Order-independent: the output never depends on the input intervals' order.
- **The surface builder** (`buildCalendarSurface`) — evaluates ONE calendar over ONE window through the query (exactly one call) and classifies its output. `calendarId` defaults to the document's `defaultCalendarId`; the calendar's `name` echoes when the id exists in the document. Degradation mirrors `schedule()` exactly: an evaluator error carrying a string `code` (structurally the canonical `CalendarError` family — missing id, inheritance cycle, malformed periods) degrades to `status: 'unresolvable'` with the diagnostic echoed verbatim (an `ImportDiagnostic` shape: code/severity/message — the SAME code the authority's own diagnostics carry for the same defect) and NO bands; an error WITHOUT a code (a host-binding bug) is re-thrown, exactly like `schedule()` re-throws non-calendar errors. A degenerate window (unparseable or empty span) has no bands and does not consult the evaluator. Bands are never invented.

### Timeline / Gantt integration

`ProjectTimeline` gains two ADDITIVE optional surfaces (present iff a working-time query was threaded AND the viewport is a real span; absent otherwise — never invented):

- `calendar` — the timeline background's `ProjectCalendarSurface` (the project default calendar unless `CalendarViewInput.calendarId` overrides — the classic Gantt non-working-time shading).
- `rowCalendars` — one `ProjectRowCalendar` per in-window row whose schedule carries `resolvedCalendarId`: `{rowIndex, taskId, calendarId, surface}`. The per-row calendar is the scheduling AUTHORITY's own echo (`TaskSchedule.resolvedCalendarId` — PROJECT-010 exposed it precisely so downstream layers read the deterministic choice without re-deriving it); rows without a resolved id (no scheduler wired, schedule failure, absent field) carry no surface. The Microsoft Project visualization semantics: each task row shades by ITS resolved calendar, so per-task calendar differences are visible row by row. Surfaces for identical calendar ids are computed once per build and shared by reference (a per-build memo over the pure query; nothing is cached across renders — lock §11). `rowCalendars` present-and-empty is real information (query threaded, no scheduled rows in window).

The calendar surfaces are ADDITIVE in the strict 023/024 sense: with or without them, `bands`/`rows`/`bars`/`milestones`/`links` (and the task grid) are byte-identical — projection never feeds geometry (tested explicitly). Hit-testing is untouched (bands are background shading, not pointer entities).

### Scope discipline and evidence

Out of scope and NOT implemented: calendar EDITING (no calendar commands exist in the frozen union — `SetCalendarWorkingWeek`/exception editing is "advanced calendars", the work item's related-work column, and requires its own authorization; the catalog and surfaces are read-only projections), time-zone/DST fidelity (lock §5 freezes the UTC minute-offset model; an advanced-calendar proposal would be required), a grid calendar column (the 022 grid surface is untouched this increment), resolved-calendar inspection beyond the evaluated bands (the catalog echoes definitions; resolution IS the evaluated surface — the authority's answer, never a renderer re-derivation), and any host UI/Electron/web code (PROJECT-027/028 — the browser/desktop E2E lands with the shells, the accepted 021–024 precedent; this increment's evidence is the component + architecture-discipline suites over the package plus test-layer integration against the REAL scheduling package). The work-item's "E2E + deterministic fixtures" is satisfied by: the unit battery (`tests/calendar.test.ts` — catalog echo/aggregations, band-classification algebra incl. contiguity/order-independence/clipping/merging, surface degradation incl. the degenerate-window no-query path and the uncoded-error re-throw); the integration battery (`tests/calendar-integration.test.ts` — the REAL `resolveCalendar`+`workingIntervals` binding; band EQUALITY with the authority's own output; weekend/holiday/extra-working-day/inheritance/exception-override semantics; the full session → projection → `buildGanttView` pipeline with per-row resolved calendars; the additive-surface geometry-neutrality; the absent-without-query degradation; the diagnostic-code EQUIVALENCE with `schedule()`'s calendar-error degradation for both a dangling default calendar and an inheritance cycle; end-to-end 3× determinism); and the golden battery C01–C12 (`tests/calendar-goldens.test.ts` — standard week; split periods; holiday exception; working-Sunday exception; inheritance merge; exception override; per-task calendars through the real schedule; the no-query absent surfaces with byte-identical geometry; the unresolvable-default degradation; viewport-zoom re-projection; session commands moving the schedule while the surfaces stay stable; and the 3× byte-identical mixed command/viewport sequence). Determinism contract: every function is pure — the same inputs always produce byte-identical outputs (3×-repeat tested), no wall clock, no randomness, no `localeCompare` (discipline-scanned), inputs never mutated; the discipline suite additionally scans `calendar.ts` for calendar-evaluation markers (date decomposition, exception date keying) — the structural no-second-engine guard — and the public-surface lockstep carries the PROJECT-025 exports.

## PROJECT-026 — Critical-path / resource visualization

The authorized semantic path (verbatim from the acceptance directive): the increment remains a **projection of canonical `DerivedSchedule` and resource state** — it must not introduce a new critical-path calculation, slack calculation, or resource-capacity algorithm in the renderer.

```
ProjectDocument
       +
DerivedSchedule
       +
canonical resource/assignment results
       ↓
project-renderer-core
       ↓
critical-path / resource visualizations
```

### Critical-path projection authority (the PROJECT-012 rule made structural)

PROJECT-012 already froze the authority: "Critical-path and float calculations belong exclusively in the scheduling engine. Renderer, host, UI, and selector code MUST NOT compute authoritative critical path or float. No second CPM engine is created." PROJECT-026 delivers the visualization side of that rule:

- `buildCriticalPath(document, projection, viewport, rowWindow)` projects the projection rows' by-reference `TaskSchedule` join into `ProjectCriticalPathSurface`: one `ProjectTaskFloat` per in-window row that carries a schedule — `critical`, `totalSlack`, and `freeSlack` echoed VERBATIM (signed working minutes, never clamped — negative slack is the authority's impossible-schedule signal and is passed through for the host's reporting surface), plus the slack-bar geometry.
- The slack bar is pure interval GEOMETRY over two canonical instants: the window `[scheduledFinish, lateFinish)` clamped to the viewport with the Gantt-bar edge-flag convention. The bar exists iff the authority's own `totalSlack` is POSITIVE (there is float to draw — the authority's answer, not a renderer calculation) and the instants form a real span. Zero-slack tasks carry no bar even when the authority's snapped late dates differ wall-clock (a zero-slack task's lateFinish can sit at the next working instant while the working-time distance is zero); negative slack carries no bar either. There is NO working-time arithmetic, NO dependency-graph traversal, and NO slack formula anywhere in the renderer module — the discipline suite scans for exactly those markers.
- A dependency link is classified critical iff BOTH of its canonical endpoint tasks carry `critical: true` — a pure two-boolean projection of the authority's task flags ("every legitimately critical task on every critical chain", PROJECT-012), NEVER a driving-path analysis: which links actually bind the chain is a slack analysis that stays with the scheduler. The convention is documented as exactly that.
- Degradation mirrors the accepted projection rules: rows without schedules carry no float (dates are never invented, lock §11); a degenerate viewport yields an EMPTY surface (the 022 rule); the timeline carries the surface iff the projection joined a schedule (`hasSchedule`) — never invented without one.

### The canonical resource-allocation authority (new scheduling export)

The time-phased resource answer is a CANONICAL scheduling result: `resourceAllocations(document, schedule)` in `packages/project-scheduling/src/allocation.ts` — a pure read-only projection of a GIVEN DerivedSchedule (never schedules, never mutates, never proposes commands; `schedule()` stays the sole authoritative scheduling operation, lock §6; leveling stays `levelResources`). The sweep MIRRORS the accepted leveler's demand/capacity semantics verbatim (the leveler itself is untouched): the demand-contribution rules (summary roll-ups, milestones, zero-duration tasks contribute nothing), the assignment demand intervals from the canonical scheduled windows, the capacity-window resolution (tightest covering window), the resource-calendar-aware demand zeroing (zero demand while the resource is not working — never a false over-allocation), and the segment boundary collection (assignment + capacity-window + working-period endpoints, midpoint evaluation). Where the leveler collapses consecutive over-allocated segments into conflict windows, the allocation emits EVERY segment of the tiling (including zero-demand non-working segments) so the projection receives the authority's complete time-phased answer; each segment carries `demandUnits`, `capacityUnits`, the canonical `overallocated` flag (the authority's own conflict predicate), and the sorted active `assignmentIds`. The equivalence is asserted by test: the union of consecutive over-allocated segments equals the leveler's `overallocations` windows (start/finish, peak demand, tightest capacity, assignment set). Determinism mirrors the leveler's contract: allocations sorted by `resourceId` (locale-free code-unit), segments ascending, assignment ids sorted — invariant under task/assignment/resource array reordering and serialization round-trips.

### The injected allocation query (no second capacity engine)

The renderer's resource visualization receives the canonical answer through the INJECTED `ResourceAllocationQuery` — a structural function type satisfied by binding `resourceAllocations` from `@genoffice/project-scheduling` (the accepted `ScheduleRunner` / `CalendarWorkingTimeQuery` injection precedents; the renderer package stays statically scheduling-free, lock §3/§6). The `ResourceViewInput` threads the query together with the CURRENT derived schedule (the session's own schedule object — the allocation must match the schedule the rest of the view projects, never a stale or re-derived one); it is a per-RENDER input like the layout and calendar inputs, never persisted `ProjectViewState`. `buildResourceUtilization(document, schedule, allocation, window)` consults the evaluator exactly once and CLIPS its segments to the viewport: no aggregation, no merging, no renderer-side classification — the demand, the capacity, and the authority's own over-allocation flag are echoed verbatim; the resource's display name is echoed from the document when the id exists. Degradation mirrors the PROJECT-025 `buildCalendarSurface` boundary exactly: a degenerate window → `ok` with no bands WITHOUT consulting the evaluator; a thrown error carrying a string code (the canonical `CalendarError` family) → `status: 'unresolvable'` with the diagnostic echoed verbatim and NO bands; an uncoded error (a host-binding bug) is re-thrown. (Through the REAL binding the scheduler's own validation is strictly stronger — a resource referencing a missing calendar degrades `schedule()` itself — so the coded-error boundary is the defensive mirror, unit-tested with a synthetic coded query, and the integration battery documents the scheduler-first degradation end-to-end.)

### Timeline integration

`ProjectTimeline` gains the ADDITIVE optional `criticalPath` (present iff the projection joined a schedule; EMPTY floats/links are real information) and `resourceUtilization` (present iff a `ResourceViewInput` was threaded; absent otherwise — never invented). `buildGanttView` threads the resource input as the calendar input's sibling (a new trailing optional parameter); the critical surface needs no input — it joins automatically from the projection's schedule. Additivity is the 023/024/025 contract: with or without either surface, every geometry surface (`bands`/`rows`/`bars`/`milestones`/`links`) and the task grid are byte-identical — projection never feeds geometry (tested explicitly).

### Scope discipline and evidence

Out of scope and NOT implemented: critical-path/resource GRID COLUMNS (the 022 grid surface is untouched this increment — the 025 precedent; the floats surface carries the values and hosts read `row.schedule` directly for cell display), resource leveling interactions (the `LevelResources` command and `levelResources` are untouched — the allocation is read-only), resource USAGE tables / time-phased work per assignment (the `assignmentSchedules` projection already carries per-assignment work; a usage-table surface is later scope), slack EDITING (no commands exist for it), any host UI/Electron/web code (PROJECT-027/028 — the browser/desktop E2E lands with the shells, the accepted 021–025 precedent; this increment's evidence is the component + architecture-discipline suites over the packages plus test-layer integration against the REAL scheduling package), and reports (the work item's related-work column — the negative-slack and over-allocation signals are echoed for the reporting layer to interpret, never interpreted here). The work-item's "E2E + schedule assertions" is satisfied by: the canonical battery (`packages/project-scheduling/tests/project-026.test.ts` — the demand tiling incl. the calendar-aware zero-demand nights, the over-allocation flag, the capacity-window segment splits, the skip rules, determinism/purity/reordering invariance, and the LEVELER CROSS-CHECK proving the mirrored sweep); the unit batteries (`tests/critical-path.test.ts` — the verbatim echo contract, the slack geometry incl. the snapped-late-date and negative-slack cases, the both-endpoints link classification, windowing, degradation, determinism; `tests/resource-utilization.test.ts` — the verbatim-clipped band contract, the name echo, the degenerate/coded/uncoded degradation, the exactly-once evaluator consultation, determinism/purity); the integration battery (`tests/critical-resource-integration.test.ts` — the REAL `resourceAllocations` binding; float EQUALITY with the authority's own TaskSchedule values on the chain/branch/MFO fixtures; band EQUALITY with the authority's output clipped; the over-allocation windows equaling the REAL leveler's conflict record; the command → schedule → surface re-projection with byte-identical undo restoration; the scheduler-first degradation for a broken resource calendar; the additive pipeline with BOTH the calendar and resource inputs threaded; end-to-end 3× determinism); and the golden battery R01–R12 (`tests/critical-resource-goldens.test.ts` — the critical chain; the float branch with its slack bar; the impossible schedule (negative slack echoed verbatim); the zero-slack milestone; the mixed link classification; the calendar-aware demand tiling; the over-allocation; the capacity-window drop; viewport re-projection; the command/undo/redo tuple restoration; the absent-input degradation with byte-identical geometry; and the 3× byte-identical mixed command/viewport sequence). Determinism contract: every function is pure — the same inputs always produce byte-identical outputs (3×-repeat tested), no wall clock, no randomness, no `localeCompare` (discipline-scanned), inputs never mutated; the discipline suite additionally scans `critical-path.ts` for CPM/float computation markers (working-time arithmetic, graph traversal) and `resources.ts` for capacity-semantics markers (the capacity fields, the capacity-window data, assignment unit reads, the calendar evaluators) — the structural no-second-engine guards — and the public-surface lockstep carries the PROJECT-026 exports.
