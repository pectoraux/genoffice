/**
 * PROJECT-026 — the resource-visualization projection layer (unit battery).
 *
 * Pure-projection semantics only — NO scheduling package anywhere in this
 * file (the allocation query is a stub; the REAL canonical binding is
 * exercised in `critical-resource-integration.test.ts` and the golden
 * battery, the accepted test-layer precedent): the verbatim-clipped band
 * contract (the authority's segments echoed — demand, capacity, the
 * over-allocation flag, the assignment ids — with NO aggregation, NO
 * merging, NO renderer-side classification), the name echo, the degradation
 * contract (degenerate window, coded-error echo, uncoded-error re-throw),
 * and determinism/purity.
 */
import { describe, expect, it, vi } from 'vitest'
import type { DerivedSchedule, ProjectDocument } from '@genoffice/project-contracts'
import { asAssignmentId, asResourceId } from '@genoffice/project-contracts'
import {
  buildResourceUtilization,
  type ResourceAllocation,
  type ResourceAllocationQuery,
} from '../src/index.js'
import { makeDocument, makeResource } from './fixtures.js'

const VIEWPORT = { start: '2026-08-03T00:00:00.000Z', finish: '2026-08-05T00:00:00.000Z' } // Mon→Wed
const SCHEDULE: DerivedSchedule = { taskSchedules: {}, diagnostics: [] }

const segment = (
  start: string,
  finish: string,
  demandUnits: number,
  capacityUnits: number,
  overallocated: boolean,
  assignmentIds: string[] = ['a1'],
) => ({
  start,
  finish,
  demandUnits,
  capacityUnits,
  overallocated,
  assignmentIds: assignmentIds.map((value) => asAssignmentId(value)),
})

const allocationOf =
  (entries: ResourceAllocation[]): ResourceAllocationQuery =>
  () =>
    entries

describe('PROJECT-026 buildResourceUtilization — the verbatim-clipped band projection', () => {
  it("clips the authority's segments to the viewport and echoes every value verbatim", () => {
    const document = makeDocument({ resources: [makeResource({ id: 'r1', name: 'Builder' })] })
    const entry: ResourceAllocation = {
      resourceId: asResourceId('r1'),
      segments: [
        segment('2026-08-03T09:00:00.000Z', '2026-08-03T17:00:00.000Z', 1, 1, false),
        segment('2026-08-03T17:00:00.000Z', '2026-08-04T09:00:00.000Z', 0, 1, false),
        segment('2026-08-04T09:00:00.000Z', '2026-08-04T17:00:00.000Z', 1.6, 1, true, ['a1', 'a2']),
        // Fully outside the viewport (Wed) — dropped, never invented back.
        segment('2026-08-05T09:00:00.000Z', '2026-08-05T17:00:00.000Z', 1, 1, false),
      ],
    }
    const surface = buildResourceUtilization(document, SCHEDULE, allocationOf([entry]), VIEWPORT)
    expect(surface.status).toBe('ok')
    expect(surface.start).toBe(VIEWPORT.start)
    expect(surface.finish).toBe(VIEWPORT.finish)
    expect(surface.resources).toHaveLength(1)
    expect(surface.resources?.[0]?.name).toBe('Builder')
    expect(surface.resources?.[0]?.resourceId).toBe(asResourceId('r1'))
    // Three in-viewport bands; values echoed verbatim from the segments.
    expect(surface.resources?.[0]?.bands).toEqual([
      {
        resourceId: asResourceId('r1'),
        start: '2026-08-03T09:00:00.000Z',
        finish: '2026-08-03T17:00:00.000Z',
        demandUnits: 1,
        capacityUnits: 1,
        overallocated: false,
        assignmentIds: [asAssignmentId('a1')],
      },
      {
        resourceId: asResourceId('r1'),
        start: '2026-08-03T17:00:00.000Z',
        finish: '2026-08-04T09:00:00.000Z',
        demandUnits: 0,
        capacityUnits: 1,
        overallocated: false,
        assignmentIds: [asAssignmentId('a1')],
      },
      {
        resourceId: asResourceId('r1'),
        start: '2026-08-04T09:00:00.000Z',
        finish: '2026-08-04T17:00:00.000Z',
        demandUnits: 1.6,
        capacityUnits: 1,
        overallocated: true,
        assignmentIds: [asAssignmentId('a1'), asAssignmentId('a2')],
      },
    ])
  })

  it('clips segments straddling the viewport edges (partial bands)', () => {
    const document = makeDocument({ resources: [makeResource({ id: 'r1' })] })
    const entry: ResourceAllocation = {
      resourceId: asResourceId('r1'),
      segments: [
        // Starts BEFORE the viewport (Sunday) and ends inside it.
        segment('2026-08-02T09:00:00.000Z', '2026-08-03T12:00:00.000Z', 1, 1, false),
        // Ends AFTER the viewport (Wednesday) and starts inside it.
        segment('2026-08-04T12:00:00.000Z', '2026-08-06T12:00:00.000Z', 2, 1, true),
      ],
    }
    const surface = buildResourceUtilization(document, SCHEDULE, allocationOf([entry]), VIEWPORT)
    expect(surface.resources?.[0]?.bands).toEqual([
      {
        resourceId: asResourceId('r1'),
        start: '2026-08-03T00:00:00.000Z',
        finish: '2026-08-03T12:00:00.000Z',
        demandUnits: 1,
        capacityUnits: 1,
        overallocated: false,
        assignmentIds: [asAssignmentId('a1')],
      },
      {
        resourceId: asResourceId('r1'),
        start: '2026-08-04T12:00:00.000Z',
        finish: '2026-08-05T00:00:00.000Z',
        demandUnits: 2,
        capacityUnits: 1,
        overallocated: true,
        assignmentIds: [asAssignmentId('a1')],
      },
    ])
  })

  it('keeps an entry with an EMPTY band list when no segment intersects the viewport', () => {
    const document = makeDocument({ resources: [makeResource({ id: 'r1' })] })
    const entry: ResourceAllocation = {
      resourceId: asResourceId('r1'),
      segments: [segment('2026-09-01T09:00:00.000Z', '2026-09-01T17:00:00.000Z', 1, 1, false)],
    }
    const surface = buildResourceUtilization(document, SCHEDULE, allocationOf([entry]), VIEWPORT)
    expect(surface.status).toBe('ok')
    expect(surface.resources).toEqual([{ resourceId: asResourceId('r1'), name: 'r1', bands: [] }])
  })

  it("preserves the authority's entry order and drops the name echo for unknown ids", () => {
    const document = makeDocument({ resources: [makeResource({ id: 'r1' })] }) // r2 unknown
    const entries: ResourceAllocation[] = [
      {
        resourceId: asResourceId('r2'),
        segments: [segment('2026-08-03T09:00:00.000Z', '2026-08-03T17:00:00.000Z', 0.5, 1, false)],
      },
      {
        resourceId: asResourceId('r1'),
        segments: [segment('2026-08-03T09:00:00.000Z', '2026-08-03T17:00:00.000Z', 1, 1, false)],
      },
    ]
    const surface = buildResourceUtilization(document, SCHEDULE, allocationOf(entries), VIEWPORT)
    expect(surface.resources?.map((entry) => entry.resourceId as string)).toEqual(['r2', 'r1'])
    expect(surface.resources?.[0]?.name).toBeUndefined()
    expect(surface.resources?.[1]?.name).toBe('r1')
  })

  it('consults the evaluator exactly ONCE per build and never for a degenerate window', () => {
    const document = makeDocument({})
    const query = vi.fn(allocationOf([]))
    // Degenerate window: status ok, no bands, evaluator NOT consulted.
    const degenerate = buildResourceUtilization(document, SCHEDULE, query, {
      start: '2026-08-03T00:00:00.000Z',
      finish: '2026-08-03T00:00:00.000Z',
    })
    expect(degenerate).toEqual({
      status: 'ok',
      resources: [],
      start: '2026-08-03T00:00:00.000Z',
      finish: '2026-08-03T00:00:00.000Z',
    })
    expect(query).not.toHaveBeenCalled()
    const unparseable = buildResourceUtilization(document, SCHEDULE, query, {
      start: 'not-a-date',
      finish: '2026-08-04T00:00:00.000Z',
    })
    expect(unparseable.status).toBe('ok')
    expect(query).not.toHaveBeenCalled()
    // Real window: exactly one consultation.
    buildResourceUtilization(document, SCHEDULE, query, VIEWPORT)
    expect(query).toHaveBeenCalledTimes(1)
    expect(query).toHaveBeenCalledWith(document, SCHEDULE)
  })

  it('degrades a coded evaluator error to unresolvable with the diagnostic echoed verbatim', () => {
    const document = makeDocument({})
    const coded: ResourceAllocationQuery = () => {
      const error = new Error('Calendar chain has a cycle') as Error & { code: string }
      error.code = 'CALENDAR_CYCLE'
      throw error
    }
    const surface = buildResourceUtilization(document, SCHEDULE, coded, VIEWPORT)
    expect(surface.status).toBe('unresolvable')
    expect(surface.resources).toBeUndefined()
    expect(surface.diagnostic).toEqual({
      code: 'CALENDAR_CYCLE',
      severity: 'error',
      message: 'Calendar chain has a cycle',
    })
  })

  it('handles a coded non-Error throw (string message) and re-throws uncoded errors', () => {
    const document = makeDocument({})
    const codedNonError: ResourceAllocationQuery = () => {
      throw { code: 'LEVELING_INCOMPLETE' }
    }
    const surface = buildResourceUtilization(document, SCHEDULE, codedNonError, VIEWPORT)
    expect(surface.status).toBe('unresolvable')
    expect(surface.diagnostic).toEqual({
      code: 'LEVELING_INCOMPLETE',
      severity: 'error',
      message: String({ code: 'LEVELING_INCOMPLETE' }),
    })
    const uncoded: ResourceAllocationQuery = () => {
      throw new Error('host-binding bug')
    }
    expect(() => buildResourceUtilization(document, SCHEDULE, uncoded, VIEWPORT)).toThrow(
      'host-binding bug',
    )
  })

  it('is deterministic (3× byte-identical) and never mutates its inputs', () => {
    const document: ProjectDocument = makeDocument({
      resources: [makeResource({ id: 'r1', name: 'Builder' })],
    })
    const entry: ResourceAllocation = {
      resourceId: asResourceId('r1'),
      segments: [
        segment('2026-08-03T09:00:00.000Z', '2026-08-03T17:00:00.000Z', 1, 1, false),
        segment('2026-08-04T09:00:00.000Z', '2026-08-04T17:00:00.000Z', 1.5, 1, true, ['a1', 'a2']),
      ],
    }
    const query = allocationOf([entry])
    const documentBefore = JSON.stringify(document)
    const first = JSON.stringify(buildResourceUtilization(document, SCHEDULE, query, VIEWPORT))
    for (let i = 0; i < 2; i += 1) {
      expect(JSON.stringify(buildResourceUtilization(document, SCHEDULE, query, VIEWPORT))).toBe(
        first,
      )
    }
    expect(JSON.stringify(document)).toBe(documentBefore)
    expect(JSON.stringify(entry)).toBe(
      JSON.stringify({
        resourceId: asResourceId('r1'),
        segments: [
          segment('2026-08-03T09:00:00.000Z', '2026-08-03T17:00:00.000Z', 1, 1, false),
          segment('2026-08-04T09:00:00.000Z', '2026-08-04T17:00:00.000Z', 1.5, 1, true, [
            'a1',
            'a2',
          ]),
        ],
      }),
    )
  })
})
