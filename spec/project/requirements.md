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
