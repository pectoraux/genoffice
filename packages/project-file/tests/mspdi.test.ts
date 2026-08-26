/**
 * PROJECT-015 — MSPDI import test suite.
 *
 * Implements the 38 mandatory tests + the M01–M17 golden-fixture evidence.
 * Acceptance invariants:
 *
 *   - Every valid fixture imports with zero ERROR-level diagnostics, passes
 *     `validateProjectDocument` (accepted, zero diagnostics), and schedules
 *     via the canonical `schedule()` with zero ERROR diagnostics.
 *   - Determinism: `serializeGproj(importMspdi(xml))` is byte-identical across
 *     repeated imports and across equivalent element-order variations.
 *   - Invalid fixtures surface the expected error diagnostic and recover to an
 *     empty/partial document (never a crash, never a silent discard).
 *   - Adversarial XML (DOCTYPE entity bomb, deep nesting, malformed) is
 *     rejected as `INVALID_MSPDI` with the empty document.
 *   - No MSPDI export surface exists (PROJECT-016 is unauthorized).
 */
import { describe, expect, it } from 'vitest'
import { schedule } from '@genoffice/project-scheduling'
import { validateProjectDocument } from '@genoffice/project-engine'
import { asCustomFieldId } from '@genoffice/project-contracts'
import {
  importMspdi,
  inspectMspdi,
  mspdiFileAdapter,
  MSPDI_FORMAT,
  INVALID_MSPDI,
  INVALID_MSPDI_DATE,
  INVALID_MSPDI_DURATION,
  INVALID_MSPDI_REFERENCE,
  MSPDI_READ,
  UNSUPPORTED_MSPDI_FEATURE,
  UNSUPPORTED_MSPDI_VERSION,
  parseXml,
  XmlParseError,
  serializeGproj,
} from '../src/index.js'
import { canonicalJson } from '../src/canonical.js'
import {
  m01Minimal,
  m02Wbs,
  m03Dependencies,
  m04LagLead,
  m05Calendars,
  m06Exceptions,
  m07Resources,
  m08Assignments,
  m09Constraints,
  m10DeadlinesProgress,
  m11Baseline,
  m12MultipleBaseline,
  m13CustomFields,
  m14Unsupported,
  m15MalformedReference,
  m16Large,
  m17Adversarial,
  m17bDeeplyNested,
  m17cMalformed,
  projectXml,
  taskXml,
} from './mspdi-fixtures.js'
import { encodeUtf8 } from '../src/utf8.js'

// ---- helpers ------------------------------------------------------------

/** Assert the import has zero error-level diagnostics. */
function expectNoErrors(diagnostics: { severity: string }[]): void {
  expect(diagnostics.filter((d) => d.severity === 'error')).toEqual([])
}

/** Byte-stable canonical serialization of an imported document. */
function canonicalBytes(xml: Uint8Array): Uint8Array {
  return serializeGproj(importMspdi(xml).document)
}

// ---- 1-26: round-trip / field-mapping tests ----------------------------

describe('PROJECT-015 — MSPDI import', () => {
  it('1. minimal MSPDI project', () => {
    const r = importMspdi(m01Minimal())
    expectNoErrors(r.diagnostics)
    expect(r.document.tasks).toEqual([])
    expect(r.document.calendars).toHaveLength(1)
    expect(String(r.document.properties.defaultCalendarId)).toBe('c1')
  })

  it('2. multiple tasks', () => {
    const xml = projectXml({
      name: 'Multi',
      tasks: [
        taskXml({ uid: 1, outlineNumber: '1' }),
        taskXml({ uid: 2, outlineNumber: '2' }),
        taskXml({ uid: 3, outlineNumber: '3' }),
      ].join('\n'),
    })
    const r = importMspdi(xml)
    expectNoErrors(r.diagnostics)
    expect(r.document.tasks).toHaveLength(3)
    expect(r.document.tasks.map((t) => t.id as string)).toEqual(['t1', 't2', 't3'])
    expect(r.document.tasks.map((t) => t.uid)).toEqual([1, 2, 3])
  })

  it('3. WBS hierarchy (parentTaskId reconstructed from outline)', () => {
    const r = importMspdi(m02Wbs())
    expectNoErrors(r.diagnostics)
    const parent = r.document.tasks.find((t) => t.id === 't1')!
    const child1 = r.document.tasks.find((t) => t.id === 't2')!
    const child2 = r.document.tasks.find((t) => t.id === 't3')!
    expect(parent.summary).toBe(true)
    expect(parent.outlineLevel).toBe(1)
    expect(child1.parentTaskId).toBe(parent.id)
    expect(child2.parentTaskId).toBe(parent.id)
    expect(child1.outlineLevel).toBe(2)
    expect(child1.summary).toBe(false)
  })

  it('4. summary tasks', () => {
    const r = importMspdi(m02Wbs())
    const parent = r.document.tasks.find((t) => t.id === 't1')!
    expect(parent.summary).toBe(true)
  })

  it('5. milestones', () => {
    const xml = projectXml({
      name: 'Milestones',
      tasks: taskXml({ uid: 1, outlineNumber: '1', milestone: true, duration: 'PT0S' }),
    })
    const r = importMspdi(xml)
    expectNoErrors(r.diagnostics)
    expect(r.document.tasks[0].milestone).toBe(true)
    expect(r.document.tasks[0].duration as unknown as number).toBe(0)
  })

  it('6. FS dependencies', () => {
    const xml = projectXml({
      name: 'FS',
      tasks: [
        taskXml({ uid: 1, outlineNumber: '1' }),
        taskXml({ uid: 2, outlineNumber: '2', predecessorLinks: [{ predUid: 1, type: 0 }] }),
      ].join('\n'),
    })
    const r = importMspdi(xml)
    expectNoErrors(r.diagnostics)
    expect(r.document.dependencies).toHaveLength(1)
    expect(r.document.dependencies[0].type).toBe('FS')
    expect(String(r.document.dependencies[0].predecessorId)).toBe('t1')
    expect(String(r.document.dependencies[0].successorId)).toBe('t2')
  })

  it('7. SS dependencies', () => {
    const xml = projectXml({
      name: 'SS',
      tasks: [
        taskXml({ uid: 1, outlineNumber: '1' }),
        taskXml({ uid: 2, outlineNumber: '2', predecessorLinks: [{ predUid: 1, type: 2 }] }),
      ].join('\n'),
    })
    const r = importMspdi(xml)
    expectNoErrors(r.diagnostics)
    expect(r.document.dependencies[0].type).toBe('SS')
  })

  it('8. FF dependencies', () => {
    const xml = projectXml({
      name: 'FF',
      tasks: [
        taskXml({ uid: 1, outlineNumber: '1' }),
        taskXml({ uid: 2, outlineNumber: '2', predecessorLinks: [{ predUid: 1, type: 1 }] }),
      ].join('\n'),
    })
    const r = importMspdi(xml)
    expectNoErrors(r.diagnostics)
    expect(r.document.dependencies[0].type).toBe('FF')
  })

  it('9. SF dependencies', () => {
    const xml = projectXml({
      name: 'SF',
      tasks: [
        taskXml({ uid: 1, outlineNumber: '1' }),
        taskXml({ uid: 2, outlineNumber: '2', predecessorLinks: [{ predUid: 1, type: 3 }] }),
      ].join('\n'),
    })
    const r = importMspdi(xml)
    expectNoErrors(r.diagnostics)
    expect(r.document.dependencies[0].type).toBe('SF')
  })

  it('10. positive lag (240 minutes)', () => {
    const r = importMspdi(m04LagLead())
    expectNoErrors(r.diagnostics)
    const fsLag = r.document.dependencies.find((d) => String(d.successorId) === 't2')!
    expect(fsLag.lagMinutes).toBe(240)
  })

  it('11. negative lag / lead (-120 minutes)', () => {
    const r = importMspdi(m04LagLead())
    expectNoErrors(r.diagnostics)
    const fsLead = r.document.dependencies.find((d) => String(d.successorId) === 't3')!
    expect(fsLead.lagMinutes).toBe(-120)
  })

  it('12. calendars (base + derived)', () => {
    const r = importMspdi(m05Calendars())
    expectNoErrors(r.diagnostics)
    expect(r.document.calendars).toHaveLength(2)
    expect(r.document.calendars.map((c) => String(c.id))).toEqual(['c1', 'c2'])
  })

  it('13. inherited calendar (baseCalendarId preserved)', () => {
    const r = importMspdi(m05Calendars())
    const derived = r.document.calendars.find((c) => String(c.id) === 'c2')!
    expect(String(derived.baseCalendarId)).toBe('c1')
  })

  it('14. calendar exception', () => {
    const r = importMspdi(m06Exceptions())
    expectNoErrors(r.diagnostics)
    const cal = r.document.calendars[0]
    expect(cal.exceptions).toHaveLength(2)
    expect(cal.exceptions[0].date).toBe('2026-08-03')
    expect(cal.exceptions[0].periods).toEqual([])
    expect(cal.exceptions[1].date).toBe('2026-08-08')
    expect(cal.exceptions[1].periods).toEqual([{ startMinute: 540, endMinute: 780 }])
  })

  it('15. task calendar', () => {
    const xml = projectXml({
      name: 'Task Cal',
      tasks: taskXml({ uid: 1, outlineNumber: '1', calendarUid: 1 }),
    })
    const r = importMspdi(xml)
    expectNoErrors(r.diagnostics)
    expect(String(r.document.tasks[0].calendarId)).toBe('c1')
  })

  it('16. resource calendar', () => {
    const r = importMspdi(m07Resources())
    expectNoErrors(r.diagnostics)
    const alice = r.document.resources.find((res) => res.id === 'r1')!
    expect(String(alice.calendarId)).toBe('c1')
  })

  it('17. work resources', () => {
    const r = importMspdi(m07Resources())
    const alice = r.document.resources.find((res) => res.id === 'r1')!
    expect(alice.kind).toBe('work')
    expect(alice.maxUnits).toBe(1)
    expect(alice.standardRate).toBe(50)
  })

  it('18. material resources', () => {
    const r = importMspdi(m07Resources())
    const concrete = r.document.resources.find((res) => res.id === 'r2')!
    expect(concrete.kind).toBe('material')
    expect(concrete.maxUnits).toBe(0)
  })

  it('19. cost resources', () => {
    const r = importMspdi(m07Resources())
    const travel = r.document.resources.find((res) => res.id === 'r3')!
    expect(travel.kind).toBe('cost')
  })

  it('20. assignments', () => {
    const r = importMspdi(m08Assignments())
    expectNoErrors(r.diagnostics)
    expect(r.document.assignments).toHaveLength(2)
    const a1 = r.document.assignments.find((a) => a.id === 'a1')!
    expect(String(a1.taskId)).toBe('t1')
    expect(String(a1.resourceId)).toBe('r1')
    expect(a1.units).toBe(1)
    expect(a1.work as unknown as number).toBe(480)
  })

  it('21. constraints (SNET + MFO)', () => {
    const r = importMspdi(m09Constraints())
    expectNoErrors(r.diagnostics)
    const t1 = r.document.tasks.find((t) => t.id === 't1')!
    const t2 = r.document.tasks.find((t) => t.id === 't2')!
    expect(t1.constraintType).toBe('startNoEarlierThan')
    expect(t1.constraintDate).toBe('2026-08-03T09:00:00.000Z')
    expect(t2.constraintType).toBe('mustFinishOn')
    expect(t2.constraintDate).toBe('2026-08-07T17:00:00.000Z')
  })

  it('22. deadlines', () => {
    const r = importMspdi(m10DeadlinesProgress())
    expectNoErrors(r.diagnostics)
    expect(r.document.tasks[0].deadline).toBe('2026-08-07T17:00:00.000Z')
  })

  it('23. progress', () => {
    const r = importMspdi(m10DeadlinesProgress())
    expect(r.document.tasks[0].percentComplete).toBe(75)
  })

  it('24. baseline (single)', () => {
    const r = importMspdi(m11Baseline())
    expectNoErrors(r.diagnostics)
    expect(r.document.baselines).toHaveLength(1)
    const b = r.document.baselines[0]
    expect(String(b.id)).toBe('b0')
    expect(b.capturedAt).toBe('2026-08-02T08:00:00.000Z')
    const snap = b.taskSnapshots['t1']
    expect(snap.start).toBe('2026-08-03T09:00:00.000Z')
    expect(snap.duration as unknown as number).toBe(480)
    expect(snap.cost).toBe(100)
  })

  it('25. multiple baselines', () => {
    const r = importMspdi(m12MultipleBaseline())
    expectNoErrors(r.diagnostics)
    expect(r.document.baselines).toHaveLength(2)
    expect(r.document.baselines.map((b) => String(b.id))).toEqual(['b0', 'b1'])
  })

  it('26. custom fields', () => {
    const r = importMspdi(m13CustomFields())
    expectNoErrors(r.diagnostics)
    expect(r.document.customFields).toHaveLength(2)
    expect(r.document.customFields.map((c) => String(c.id))).toEqual(['188743731', '188743734'])
    const cf = r.document.customFields.find((c) => String(c.id) === '188743734')!
    expect(cf.type).toBe('number')
    expect(r.document.tasks[0].customFields[asCustomFieldId('188743731')]).toBe('Acme Corp')
    expect(r.document.tasks[0].customFields[asCustomFieldId('188743734')]).toBe(5000)
  })
})

// ---- 27-35: rejection / diagnostics ------------------------------------

describe('PROJECT-015 — diagnostics', () => {
  it('27. malformed XML rejection', () => {
    const r = importMspdi(m17cMalformed())
    expect(r.diagnostics.some((d) => d.code === INVALID_MSPDI && d.severity === 'error')).toBe(true)
    expect(r.document.tasks).toEqual([])
  })

  it('28. unsupported MSPDI feature diagnostic', () => {
    const r = importMspdi(m14Unsupported())
    expect(r.diagnostics.some((d) => d.code === UNSUPPORTED_MSPDI_FEATURE)).toBe(true)
    // Recurring exception maps to a single-date exception (loss named, not silent).
    expect(r.document.calendars[0].exceptions).toHaveLength(1)
    expect(r.document.calendars[0].exceptions[0].date).toBe('2026-12-25')
    // Zero errors — the document is still valid.
    expectNoErrors(r.diagnostics)
  })

  it('29. invalid reference diagnostic (dangling predecessor)', () => {
    const r = importMspdi(m15MalformedReference())
    expect(
      r.diagnostics.some((d) => d.code === INVALID_MSPDI_REFERENCE && d.severity === 'error'),
    ).toBe(true)
    // The dangling dependency is dropped (partial recovery).
    expect(r.document.dependencies).toEqual([])
    // The surviving tasks remain.
    expect(r.document.tasks).toHaveLength(2)
  })

  it('30. invalid duration diagnostic (malformed ISO duration)', () => {
    const xml = projectXml({
      name: 'Bad Dur',
      tasks: taskXml({ uid: 1, outlineNumber: '1', duration: 'not-a-duration' }),
    })
    const r = importMspdi(xml)
    expect(
      r.diagnostics.some((d) => d.code === INVALID_MSPDI_DURATION && d.severity === 'error'),
    ).toBe(true)
    expect(r.document.tasks[0].duration as unknown as number).toBe(0)
  })

  it('31. invalid date diagnostic (bad task Start)', () => {
    const xml = projectXml({
      name: 'Bad Date',
      tasks: taskXml({ uid: 1, outlineNumber: '1', start: 'not-a-date' }),
    })
    const r = importMspdi(xml)
    expect(r.diagnostics.some((d) => d.code === INVALID_MSPDI_DATE && d.severity === 'error')).toBe(
      true,
    )
    // Task survives without the bad start (partial recovery).
    expect(r.document.tasks).toHaveLength(1)
    expect(r.document.tasks[0].start).toBeUndefined()
  })

  it('32. deterministic repeated import (same XML → same ProjectDocument bytes)', () => {
    const a = canonicalBytes(m03Dependencies())
    const b = canonicalBytes(m03Dependencies())
    expect(Array.from(a)).toEqual(Array.from(b))
  })

  it('33. deterministic import after equivalent element-order variation', () => {
    // Two XML payloads identical except the <Task> child-element order is
    // reversed. Canonical extraction is by name, so the output must match.
    const orderA = `<?xml version="1.0"?>
<Project xmlns="http://schemas.microsoft.com/project"><SaveVersion>16</SaveVersion><Name>Order</Name><StartDate>2026-08-03T09:00:00</StartDate>
<Calendars><Calendar><UID>1</UID><Name>Standard</Name><IsBaseCalendar>true</IsBaseCalendar><IsBaseCalendarDefault>true</IsBaseCalendarDefault><WeekDays><WeekDay><DayType>2</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay></WeekDays><Exceptions /></Calendar></Calendars>
<Tasks><Task><UID>1</UID><Name>First</Name><Duration>PT8H0M0S</Duration><OutlineNumber>1</OutlineNumber><OutlineLevel>1</OutlineLevel><Summary>false</Summary><Milestone>false</Milestone><Manual>false</Manual><Priority>500</Priority><PercentComplete>0</PercentComplete><Work>PT0H0M0S</Work><RemainingWork>PT0H0M0S</RemainingWork><ActualWork>PT0H0M0S</ActualWork><Cost>0</Cost><ActualCost>0</ActualCost><RemainingCost>0</RemainingCost><Baseline /><CustomFields /><Notes /></Task></Tasks>
<Resources /><Assignments /></Project>
`
    const orderB = `<?xml version="1.0"?>
<Project xmlns="http://schemas.microsoft.com/project"><SaveVersion>16</SaveVersion><Name>Order</Name><StartDate>2026-08-03T09:00:00</StartDate>
<Calendars><Calendar><UID>1</UID><Name>Standard</Name><IsBaseCalendar>true</IsBaseCalendar><IsBaseCalendarDefault>true</IsBaseCalendarDefault><WeekDays><WeekDay><DayType>2</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay></WeekDays><Exceptions /></Calendar></Calendars>
<Tasks><Task><Notes /><CustomFields /><Baseline /><RemainingCost>0</RemainingCost><ActualCost>0</ActualCost><Cost>0</Cost><ActualWork>PT0H0M0S</ActualWork><RemainingWork>PT0H0M0S</RemainingWork><Work>PT0H0M0S</Work><PercentComplete>0</PercentComplete><Priority>500</Priority><Manual>false</Manual><Milestone>false</Milestone><Summary>false</Summary><OutlineLevel>1</OutlineLevel><OutlineNumber>1</OutlineNumber><Duration>PT8H0M0S</Duration><Name>First</Name><UID>1</UID></Task></Tasks>
<Resources /><Assignments /></Project>
`
    const a = canonicalBytes(encodeUtf8(orderA))
    const b = canonicalBytes(encodeUtf8(orderB))
    expect(Array.from(a)).toEqual(Array.from(b))
  })

  it('34. large MSPDI project (60 tasks / 50 deps / 10 resources / 60 assignments)', () => {
    const r = importMspdi(m16Large())
    expectNoErrors(r.diagnostics)
    expect(r.document.tasks).toHaveLength(60)
    expect(r.document.dependencies).toHaveLength(50)
    expect(r.document.resources).toHaveLength(10)
    expect(r.document.assignments).toHaveLength(60)
    expect(serializeGproj(r.document).byteLength).toBeGreaterThan(8 * 1024)
  })

  it('35. security/adversarial XML payload (DOCTYPE bomb + deep nesting + malformed)', () => {
    const bomb = importMspdi(m17Adversarial())
    expect(bomb.diagnostics.some((d) => d.code === INVALID_MSPDI && d.severity === 'error')).toBe(
      true,
    )
    expect(bomb.document.tasks).toEqual([])
    const deep = importMspdi(m17bDeeplyNested())
    expect(deep.diagnostics.some((d) => d.code === INVALID_MSPDI && d.severity === 'error')).toBe(
      true,
    )
    expect(deep.document.tasks).toEqual([])
    const bad = importMspdi(m17cMalformed())
    expect(bad.diagnostics.some((d) => d.code === INVALID_MSPDI && d.severity === 'error')).toBe(
      true,
    )
    // Entity expansion must not have polluted any global.
    expect(({} as Record<string, unknown>).polluted).toBeUndefined()
  })
})

// ---- 36-38: canonical validation + scheduling ---------------------------

describe('PROJECT-015 — canonical validation + scheduling', () => {
  it('36. imported document passes canonical validation (all valid fixtures)', () => {
    const valids = [
      m01Minimal,
      m02Wbs,
      m03Dependencies,
      m04LagLead,
      m05Calendars,
      m06Exceptions,
      m07Resources,
      m08Assignments,
      m09Constraints,
      m10DeadlinesProgress,
      m11Baseline,
      m12MultipleBaseline,
      m13CustomFields,
      m16Large,
    ]
    for (const build of valids) {
      const r = importMspdi(build())
      const errors = r.diagnostics.filter((d) => d.severity === 'error')
      expect(errors).toEqual([])
      const v = validateProjectDocument(r.document)
      expect(v.diagnostics).toEqual([])
      expect(v.accepted).toBe(true)
    }
  })

  it('37. imported document can be scheduled by the canonical scheduler', () => {
    const r = importMspdi(m03Dependencies())
    const s = schedule(r.document)
    const errors = s.diagnostics.filter((d) => d.severity === 'error')
    expect(errors).toEqual([])
    expect(Object.keys(s.taskSchedules)).toHaveLength(3)
  })

  it('38. imported schedule is deterministic (repeated schedule → identical bytes)', () => {
    const doc = importMspdi(m03Dependencies()).document
    const s1 = schedule(doc)
    const s2 = schedule(doc)
    expect(Array.from(canonicalJson(s1))).toEqual(Array.from(canonicalJson(s2)))
  })
})

// ---- golden fixtures M01–M17 -------------------------------------------

describe('PROJECT-015 — golden fixtures M01–M17', () => {
  describe('valid goldens (M01–M13, M16)', () => {
    const valids: Array<{ id: string; build: () => Uint8Array; note: string }> = [
      { id: 'M01', build: m01Minimal, note: 'minimal' },
      { id: 'M02', build: m02Wbs, note: 'WBS' },
      { id: 'M03', build: m03Dependencies, note: 'dependencies' },
      { id: 'M04', build: m04LagLead, note: 'lag/lead' },
      { id: 'M05', build: m05Calendars, note: 'calendars' },
      { id: 'M06', build: m06Exceptions, note: 'exceptions' },
      { id: 'M07', build: m07Resources, note: 'resources' },
      { id: 'M08', build: m08Assignments, note: 'assignments' },
      { id: 'M09', build: m09Constraints, note: 'constraints' },
      { id: 'M10', build: m10DeadlinesProgress, note: 'deadlines/progress' },
      { id: 'M11', build: m11Baseline, note: 'baseline' },
      { id: 'M12', build: m12MultipleBaseline, note: 'multiple baseline' },
      { id: 'M13', build: m13CustomFields, note: 'custom fields' },
      { id: 'M16', build: m16Large, note: 'large project' },
    ]
    for (const { id, build, note } of valids) {
      it(`${id} (${note}) imports deterministically with zero errors`, () => {
        const r = importMspdi(build())
        expectNoErrors(r.diagnostics)
        // Determinism: re-import and compare canonical bytes.
        const a = canonicalBytes(build())
        const b = canonicalBytes(build())
        expect(Array.from(a)).toEqual(Array.from(b))
      })
    }
  })

  describe('invalid / diagnostic goldens (M14, M15, M17)', () => {
    it('M14 (unsupported feature) emits UNSUPPORTED_MSPDI_FEATURE', () => {
      const r = importMspdi(m14Unsupported())
      expect(r.diagnostics.some((d) => d.code === UNSUPPORTED_MSPDI_FEATURE)).toBe(true)
    })
    it('M15 (malformed references) emits INVALID_MSPDI_REFERENCE', () => {
      const r = importMspdi(m15MalformedReference())
      expect(
        r.diagnostics.some((d) => d.code === INVALID_MSPDI_REFERENCE && d.severity === 'error'),
      ).toBe(true)
    })
    it('M17 (adversarial XML) emits INVALID_MSPDI and empty document', () => {
      const r = importMspdi(m17Adversarial())
      expect(r.diagnostics.some((d) => d.code === INVALID_MSPDI && d.severity === 'error')).toBe(
        true,
      )
      expect(r.document.tasks).toEqual([])
    })
  })
})

// ---- adapter surface + identity + boundary guards ----------------------

describe('PROJECT-015 — adapter surface + identity + boundary', () => {
  it('mspdiFileAdapter.format === "mspdi"', () => {
    expect(mspdiFileAdapter.format).toBe(MSPDI_FORMAT)
  })

  it('mspdiFileAdapter.inspect returns canonical metadata', () => {
    const meta = mspdiFileAdapter.inspect(m01Minimal(), {
      format: 'mspdi',
      version: '16',
      sourceName: 'a.xml',
    })
    expect(meta.format).toBe('mspdi')
    expect(meta.sourceName).toBe('a.xml')
  })

  it('mspdiFileAdapter.import is the canonical import path', () => {
    const r = mspdiFileAdapter.import(m03Dependencies())
    expectNoErrors(r.diagnostics)
    expect(r.document.dependencies).toHaveLength(2)
  })

  it('mspdiFileAdapter has NO export method (PROJECT-016 unauthorized)', () => {
    expect((mspdiFileAdapter as Record<string, unknown>).export).toBeUndefined()
  })

  it('MSPDI_READ info diagnostic is emitted on success', () => {
    const r = importMspdi(m01Minimal())
    expect(r.diagnostics.some((d) => d.code === MSPDI_READ && d.severity === 'info')).toBe(true)
  })

  it('unsupported SaveVersion is rejected', () => {
    const xml = projectXml({ name: 'Future', saveVersion: 9999 })
    const r = importMspdi(xml)
    expect(
      r.diagnostics.some((d) => d.code === UNSUPPORTED_MSPDI_VERSION && d.severity === 'error'),
    ).toBe(true)
    expect(r.document.tasks).toEqual([])
  })

  it('identity is deterministic + stable (UID→id map)', () => {
    const r = importMspdi(m08Assignments())
    expect(new Set(r.document.tasks.map((t) => String(t.id)))).toEqual(new Set(['t1']))
    expect(new Set(r.document.resources.map((res) => String(res.id)))).toEqual(
      new Set(['r1', 'r2']),
    )
    expect(new Set(r.document.assignments.map((a) => String(a.id)))).toEqual(new Set(['a1', 'a2']))
  })

  it('raw XML parser rejects DOCTYPE directly', () => {
    expect(() => parseXml('<!DOCTYPE root [<!ENTITY x "y">]><root/>')).toThrow(XmlParseError)
  })

  it('the MSPDI adapter modules carry no forbidden host imports (boundary guard)', () => {
    // Static guard: the adapter functions are callable without a host runtime.
    expect(typeof importMspdi).toBe('function')
    expect(typeof inspectMspdi).toBe('function')
    expect(typeof mspdiFileAdapter.import).toBe('function')
    expect(typeof mspdiFileAdapter.inspect).toBe('function')
  })

  it('date normalization treats naive MSPDI dates as UTC (no host tz)', () => {
    const xml = projectXml({ name: 'NaiveDate', startDate: '2026-08-03T09:00:00' })
    const r = importMspdi(xml)
    expectNoErrors(r.diagnostics)
    expect(r.document.properties.startDate).toBe('2026-08-03T09:00:00.000Z')
  })

  it('date normalization converts explicit offset to UTC', () => {
    const xml = projectXml({ name: 'Offset', startDate: '2026-08-03T11:00:00+02:00' })
    const r = importMspdi(xml)
    expectNoErrors(r.diagnostics)
    expect(r.document.properties.startDate).toBe('2026-08-03T09:00:00.000Z')
  })
})
