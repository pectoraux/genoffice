/**
 * PROJECT-028 — web E2E fixtures: the SAME canonical fixture document the
 * desktop E2E battery drives (parity by construction): REAL canonical
 * documents exported through the REAL `.gproj` adapter, composed with the
 * shared host binding's own new-document template. Nothing here schedules
 * or serializes by hand — the bytes the browser opens are the adapter's
 * own deterministic output.
 */
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  asAssignmentId,
  asDependencyId,
  asResourceId,
  asTaskId,
  asWorkingMinutes,
} from '@genoffice/project-contracts'
import type { ProjectDocument } from '@genoffice/project-contracts'
import { gprojFileAdapter } from '@genoffice/project-file'
import { defaultNewTask } from '@genoffice/project-renderer-core'
import { newProjectDocument } from '@genoffice/project-host'

/**
 * The canonical E2E fixture document (identical to the desktop battery's):
 *
 * - `t1` Design (2 days), `t2` Build (4 days, FS after t1), `t3` Ship
 *   (milestone, FS after t2) — the critical chain;
 * - `t4` Review (1 day, independent) — float → a slack bar;
 * - resource `r1` Crew (maxUnits 1) assigned 100% to BOTH `t1` and `t4`
 *   (both start at the project start → the shared window is
 *   over-allocated: the authority's flag, exercised end-to-end);
 * - the standard Mon–Fri calendar (the host template's own).
 */
export function e2eFixtureDocument(): ProjectDocument {
  const base = newProjectDocument('E2E Build')
  const identity = (n: number) => ({ id: asTaskId(`t${n}`), uid: n })
  const task = (n: number, name: string, overrides: Record<string, unknown> = {}) => ({
    ...defaultNewTask(identity(n)),
    name,
    ...overrides,
  })
  return {
    ...base,
    properties: { ...base.properties, name: 'E2E Build' },
    tasks: [
      task(1, 'Design', { duration: asWorkingMinutes(960) }),
      task(2, 'Build', { duration: asWorkingMinutes(1920) }),
      task(3, 'Ship', { milestone: true, duration: asWorkingMinutes(0) }),
      task(4, 'Review', { duration: asWorkingMinutes(480) }),
    ],
    dependencies: [
      {
        id: asDependencyId('d1'),
        predecessorId: asTaskId('t1'),
        successorId: asTaskId('t2'),
        type: 'FS',
        lagMinutes: 0,
      },
      {
        id: asDependencyId('d2'),
        predecessorId: asTaskId('t2'),
        successorId: asTaskId('t3'),
        type: 'FS',
        lagMinutes: 0,
      },
    ],
    resources: [
      {
        id: asResourceId('r1'),
        uid: 1,
        name: 'Crew',
        kind: 'work',
        maxUnits: 1,
        standardRate: 0,
        overtimeRate: 0,
        costPerUse: 0,
        availability: [],
      },
    ],
    assignments: [
      {
        id: asAssignmentId('a1'),
        taskId: asTaskId('t1'),
        resourceId: asResourceId('r1'),
        units: 1,
        work: asWorkingMinutes(960),
        actualWork: asWorkingMinutes(0),
        remainingWork: asWorkingMinutes(960),
        cost: 0,
        actualCost: 0,
        remainingCost: 0,
      },
      {
        id: asAssignmentId('a2'),
        taskId: asTaskId('t4'),
        resourceId: asResourceId('r1'),
        units: 1,
        work: asWorkingMinutes(480),
        actualWork: asWorkingMinutes(0),
        remainingWork: asWorkingMinutes(480),
        cost: 0,
        actualCost: 0,
        remainingCost: 0,
      },
    ],
  }
}

/** The fixture as the adapter's own deterministic `.gproj` bytes. */
export function e2eFixtureBytes(): Uint8Array {
  const exported = gprojFileAdapter.export(e2eFixtureDocument())
  if (exported.diagnostics.length > 0) {
    throw new Error(`Fixture export produced diagnostics: ${JSON.stringify(exported.diagnostics)}`)
  }
  return exported.bytes
}

/** Writes the fixture as a real `.gproj` file; returns its absolute path. */
export async function writeE2EFixture(name = 'e2e-build.gproj'): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-project-web-fixture-'))
  const path = join(dir, name)
  await writeFile(path, e2eFixtureBytes())
  return path
}
