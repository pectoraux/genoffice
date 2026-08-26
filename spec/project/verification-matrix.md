# GenOffice Project — Verification Matrix

| Work item        | Semantic evidence                                  | Static/architecture evidence                      | Host evidence                         | File evidence                   | Completion gate                                                         |
| ---------------- | -------------------------------------------------- | ------------------------------------------------- | ------------------------------------- | ------------------------------- | ----------------------------------------------------------------------- |
| PROJECT-001      | contract compile                                   | package boundary scan; lock review                | N/A                                   | N/A                             | six specs committed and internally consistent                           |
| PROJECT-002      | domain validation tests                            | stable-ID architecture tests                      | N/A                                   | schema serialization fixture    | invalid identities/references rejected                                  |
| PROJECT-003      | command/journal tests                              | forbidden raw-mutation scan                       | N/A                                   | N/A                             | semantic command union + history invariants pass                        |
| PROJECT-004      | calendar golden suite                              | host-import scan                                  | N/A                                   | calendar fixture                | repeated working-time calculations match                                |
| PROJECT-005      | graph validation/topology suite                    | dependency determinism scan                       | N/A                                   | graph fixture                   | all link types + invalid graph rejection pass                           |
| PROJECT-006      | scheduling golden suite + repeat-run byte equality | no UI/host imports; deterministic iteration guard | N/A                                   | golden ProjectDocument fixtures | dates, dependencies, critical path, total/free float all match expected |
| PROJECT-007..013 | domain/scheduling goldens                          | package boundaries                                | later E2E                             | later save/reopen               | semantic acceptance before renderer work                                |
| PROJECT-014..020 | round-trip/import/export goldens                   | parser isolation checks                           | N/A                                   | golden files + diagnostics      | no silent loss                                                          |
| PROJECT-021..031 | command-to-renderer integration                    | renderer dependency checks                        | desktop + Playwright E2E              | save/reopen                     | semantic and renderer parity                                            |
| PROJECT-032..045 | view/scheduling/resource tests                     | view projections have no scheduling authority     | desktop/web E2E                       | view fixtures                   | each view matches canonical derived state                               |
| PROJECT-046      | cross-host scenario suite                          | package boundary scan                             | desktop/web parity                    | persisted scenario artifacts    | identical semantic outcomes                                             |
| PROJECT-047      | golden compatibility suite                         | adapter isolation                                 | optional host smoke                   | golden file corpus              | import/export regressions blocked                                       |
| PROJECT-048      | benchmark suite                                    | performance architecture assertions               | desktop/web performance runs          | representative project fixtures | agreed large-project budgets met                                        |
| PROJECT-049      | production E2E                                     | release architecture gate                         | desktop + web production verification | save/reopen + import/export     | release acceptance complete                                             |

## Foundation evidence requirements

PROJECT-006 must include at minimum:

- sequential dependency fixture
- FS/SS/FF/SF fixture
- lag fixture
- holiday/calendar fixture
- critical-path/float fixture
- cycle/self/missing-reference rejection tests
- repeated scheduling equality test

Agent narrative is never sufficient evidence without these artifacts.

## PROJECT-012 evidence requirements

PROJECT-012 (Critical path / float edge cases) must include at minimum:

- 18 golden fixtures CP01–CP18 covering: simple critical chain, two/three critical paths, diamond convergence, mixed dependency types, lag/lead, free-vs-total slack, negative slack, near-critical branch, critical/noncritical milestone, summary criticality, nested summary, constraint + critical path, calendar boundary, holiday, isolated branch, reordered deterministic input.
- additional required tests covering: fan-out, fan-in, FS lag, FS lead, SS lag, FF lag, SF lag, multiple successors, freeSlack < totalSlack, freeSlack = totalSlack, zero free slack, negative slack, one-day near-critical, one-hour near-critical, critical milestone, noncritical milestone, milestone with predecessor, summary critical roll-up, nested summary critical path, SNET/MSO/FNET/MFO constraint interaction, different task calendars, weekend boundary, holiday boundary, isolated independent task, repeated deterministic schedule, serialized round-trip byte equality.
- every applicable golden fixture asserts exact `taskId`, `earlyStart`, `earlyFinish`, `lateStart`, `lateFinish`, `totalSlack`, `freeSlack`, `critical`, `scheduledStart`, `scheduledFinish`, and `projectFinish` — not merely counts of critical tasks.
- byte-identical `DerivedSchedule` output under repeated runs, reversed task/dependency arrays, and reordered parallel branches.
- all accepted PROJECT-006 through PROJECT-011 tests remain green (no regression, no weakened goldens).

## PROJECT-013 evidence requirements

PROJECT-013 (Resource leveling) must include at minimum:

- 25 golden fixtures L01–L25 covering: simple two-task overload, three-task overload, 200% single-assignment demand, priority resolution, TaskId tie-break (equal priority), critical-task policy (respectCritical true and false), FS dependency propagation, SS dependency propagation, FF dependency propagation, SF dependency propagation, positive lag, negative lead, SNET protection, MSO protection, MFO protection, deadline interaction (respectDeadlines true and false), resource calendar restriction, availability window restriction, multiple resources on the same task, baseline preservation, impossible leveling (both MSO), repeated deterministic leveling, reordered input leveling (reversed task/assignment/resource arrays + serialized round-trip), material/cost resources excluded, mid-assignment capacity drop (L25 — over-allocation bounded only by availability-window transitions; no assignment event brackets the conflict).
- additional required tests covering: SNLT protection, FNET protection, FNLT protection, milestone behavior (zero-duration skipped), summary behavior (not directly levelable; rolls up), manual task behavior (protected; LEVELING_PROTECTED_MANUAL when both sides manual), already-non-overallocated resource (LEVELING_NO_OVERALLOCATION), leveling result reapplied through the canonical scheduler (individual SetTaskStart dispatch matches the LevelResources batch), stable-ID preservation (TaskId/ResourceId/AssignmentId/DependencyId/BaselineId unchanged), no new dependency cycles, scope filter (taskIds), scope-empty (LEVELING_SCOPE_EMPTY), different assignment units, one 200% assignment (LEVELING_INCOMPLETE), FS predecessor-delay propagation to successor, split-day task calendar, LevelingResult shape, schedule-after-leveling is canonical, multi-window conflict identity (the same assignments producing two distinct conflict windows are reported as two distinct over-allocations — they do NOT collapse into one entry whose `resolved` state the final pass would overwrite).
- every applicable golden fixture asserts exact `taskId`, `scheduledStart`, `scheduledFinish`, `proposedCommands` (the semantic `SetTaskStart` values), `affectedTaskIds`, `overallocations.peakDemand`, `overallocations.maxUnits`, `overallocations.resolved`, and (where applicable) `critical`, `totalSlack`, `deadline`, `deadlineMissed`, `deadlineVariance`, `constraintType`, `constraintDate`.
- byte-identical `LevelingResult` output (proposedCommands + actions + overallocations + diagnostics) under three repeated runs, reversed task arrays, reversed assignment arrays, reversed resource arrays, and serialized round-trip (JSON parse).
- idempotency: leveling an already-leveled document emits `LEVELING_NO_OVERALLOCATION` with empty `proposedCommands`.
- semantic-command verification: applying `proposedCommands` via individual `SetTaskStart` dispatch produces the same document + schedule as the `LevelResources` batch.
- impossible-leveling verification: when no task can be delayed (both MSO, or all manual), `proposedCommands` is empty, the document is unchanged, and the diagnostic (`LEVELING_CONSTRAINT_CONFLICT` / `LEVELING_PROTECTED_MANUAL` / `LEVELING_NO_ELIGIBLE_TASK`) is surfaced.
- baseline immutability: baseline snapshots (`start`, `finish`, `duration`, `work`, `cost`), baseline IDs, and `capturedAt` are unchanged after leveling; only the current schedule's `task.start` candidates move.
- identity preservation: `TaskId`, `ResourceId`, `AssignmentId`, `DependencyId`, and `BaselineId` sets are unchanged; the dependency graph remains acyclic (the leveled document still schedules cleanly).
- no-regression: all accepted PROJECT-006 through PROJECT-012 tests remain green (no weakened goldens, no changed CPM formulas).
- sweep segmentation: the over-allocation sweep includes availability-window boundaries (start/finish) in addition to assignment start/finish endpoints. A golden regression (L25) proves an over-allocation that arises ONLY from a mid-assignment capacity drop (no assignment event brackets the conflict) is detected and resolved.
- conflict-signature window identity: the conflict signature includes the conflict window start + finish (epoch-normalized). A multi-window regression proves the same assignments producing two distinct conflict windows are reported as two distinct over-allocations (no collapse, no `resolved`-state overwrite by the final pass).
- material/cost exclusion asserted exactly: no `|| true` tautology; the over-allocations list is asserted to contain only work-resource ids, and the exact resource-id set is asserted.
