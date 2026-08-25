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
