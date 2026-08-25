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
