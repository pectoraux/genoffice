# GenOffice Project — Verification Matrix

| Work item | Semantic evidence | Static/architecture evidence | Host evidence | File evidence | Completion gate |
|---|---|---|---|---|---|
| PROJECT-001 | contract compile | package boundary scan; lock review | N/A | N/A | six specs committed and internally consistent |
| PROJECT-002 | domain validation tests | stable-ID architecture tests | N/A | schema serialization fixture | invalid identities/references rejected |
| PROJECT-003 | command/journal tests | forbidden raw-mutation scan | N/A | N/A | semantic command union + history invariants pass |
| PROJECT-004 | calendar golden suite | host-import scan | N/A | calendar fixture | repeated working-time calculations match |
| PROJECT-005 | graph validation/topology suite | dependency determinism scan | N/A | graph fixture | all link types + invalid graph rejection pass |
| PROJECT-006 | scheduling golden suite + repeat-run byte equality | no UI/host imports; deterministic iteration guard | N/A | golden ProjectDocument fixtures | dates, dependencies, critical path, total/free float all match expected |
| PROJECT-007..013 | domain/scheduling goldens | package boundaries | later E2E | later save/reopen | semantic acceptance before renderer work |
| PROJECT-014..020 | round-trip/import/export goldens | parser isolation checks | N/A | golden files + diagnostics | no silent loss |
| PROJECT-021..031 | command-to-renderer integration | renderer dependency checks | desktop + Playwright E2E | save/reopen | semantic and renderer parity |
| PROJECT-032..045 | view/scheduling/resource tests | view projections have no scheduling authority | desktop/web E2E | view fixtures | each view matches canonical derived state |
| PROJECT-046 | cross-host scenario suite | package boundary scan | desktop/web parity | persisted scenario artifacts | identical semantic outcomes |
| PROJECT-047 | golden compatibility suite | adapter isolation | optional host smoke | golden file corpus | import/export regressions blocked |
| PROJECT-048 | benchmark suite | performance architecture assertions | desktop/web performance runs | representative project fixtures | agreed large-project budgets met |
| PROJECT-049 | production E2E | release architecture gate | desktop + web production verification | save/reopen + import/export | release acceptance complete |

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
