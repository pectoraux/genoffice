/**
 * PROJECT-027 — E2E fixtures: REAL canonical documents exported through the
 * REAL `.gproj` adapter (the new-document template comes from the shared
 * host binding since PROJECT-028). The fixture builder composes the host's own
 * new-document template with document data (contracts types + the
 * renderer-core default-task builder), schedules nothing itself, and the
 * bytes the app opens are the adapter's own deterministic output.
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
 * The canonical E2E fixture document:
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

/** Writes the fixture as a real `.gproj` file; returns its absolute path. */
export async function writeE2EFixture(name = 'e2e-build.gproj'): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-project-fixture-'))
  const path = join(dir, name)
  const exported = gprojFileAdapter.export(e2eFixtureDocument())
  if (exported.diagnostics.length > 0) {
    throw new Error(`Fixture export produced diagnostics: ${JSON.stringify(exported.diagnostics)}`)
  }
  await writeFile(path, exported.bytes)
  return path
}

/**
 * A VALID `.gproj` fixture padded to an EXACT byte size with JSON trailing
 * whitespace. The bounded-transport E2E uses this to make byte-size the
 * ONLY variable: a padded document would load if (and only if) its bytes
 * actually crossed the IPC boundary, so a refused load is proof the
 * transport cap fired — and an accepted load at exactly the cap is proof
 * the boundary read delivers.
 *
 * Validity proof (small-scale, cheap): the adapter carries its OWN input
 * cap (defense in depth — the same 100 MiB as the host transport cap), so
 * an oversized fixture cannot be proven by importing it whole. Instead a
 * padded PROBE (base document + the same trailing-space tail) is imported
 * through the REAL adapter — clean import proves trailing-whitespace
 * padding preserves validity — and the full-size buffer is constructed as
 * exactly that composition (base bytes + all-space tail, byte-wise fill).
 * Note the two caps fail with DIFFERENT messages ("File exceeds the … byte
 * limit" — transport; "Input exceeds maximum size (… bytes)" — adapter),
 * which is what lets the E2E assert WHICH layer refused the oversized read.
 */
export async function writePaddedE2EFixture(
  totalBytes: number,
  name = 'padded.gproj',
): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-project-fixture-'))
  const path = join(dir, name)
  const exported = gprojFileAdapter.export(e2eFixtureDocument())
  if (exported.diagnostics.length > 0) {
    throw new Error(`Fixture export produced diagnostics: ${JSON.stringify(exported.diagnostics)}`)
  }
  if (exported.bytes.byteLength > totalBytes) {
    throw new Error(
      `Fixture body (${exported.bytes.byteLength} bytes) exceeds the target size ${totalBytes}`,
    )
  }
  // PROBE: base document + a trailing-space tail imports cleanly through
  // the real adapter — trailing JSON whitespace preserves validity.
  const probe = new Uint8Array(exported.bytes.byteLength + 1024)
  probe.set(exported.bytes)
  probe.fill(0x20, exported.bytes.byteLength)
  const roundTrip = gprojFileAdapter.import(probe)
  if (roundTrip.diagnostics.some((diagnostic) => diagnostic.severity === 'error')) {
    throw new Error(
      'Trailing-whitespace padding no longer imports cleanly through the real adapter',
    )
  }
  // The full-size fixture is exactly that composition: the base bytes
  // followed by an all-space tail to the exact byte size.
  const padded = new Uint8Array(totalBytes)
  padded.set(exported.bytes)
  padded.fill(0x20, exported.bytes.byteLength)
  await writeFile(path, padded)
  return path
}
