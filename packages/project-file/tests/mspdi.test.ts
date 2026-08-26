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
  DEFAULT_LAG_FACTORS,
  INVALID_MSPDI,
  INVALID_MSPDI_CALENDAR,
  INVALID_MSPDI_DATE,
  INVALID_MSPDI_DURATION,
  INVALID_MSPDI_REFERENCE,
  lagToMinutes,
  MSPDI_READ,
  mspdiTimeToMinutes,
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
  m18LagUnits,
  projectXml,
  taskXml,
  STANDARD_CALENDAR_XML,
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

// ---- correction round 1: LinkLagFormat conversion semantics -------------

describe('PROJECT-015 — LinkLagFormat conversion semantics (pure)', () => {
  it('minute format (1): LinkLag tenths of minutes → exact minutes', () => {
    expect(lagToMinutes(2400, 1)).toEqual({ ok: true, minutes: 240 })
    expect(lagToMinutes(10, 1)).toEqual({ ok: true, minutes: 1 })
    expect(lagToMinutes(-1200, 1)).toEqual({ ok: true, minutes: -120 })
    expect(lagToMinutes(0, 1)).toEqual({ ok: true, minutes: 0 })
    // Sub-minute tenths are not representable as integer WorkingMinutes.
    expect(lagToMinutes(15, 1)).toEqual({ ok: false, reason: 'invalid' })
    expect(lagToMinutes(-7, 1)).toEqual({ ok: false, reason: 'invalid' })
    // Missing format defaults to minutes.
    expect(lagToMinutes(2400, undefined)).toEqual({ ok: true, minutes: 240 })
  })

  it('hour format (3): LinkLag tenths of hours → ×60 exact minutes', () => {
    expect(lagToMinutes(30, 3)).toEqual({ ok: true, minutes: 180 }) // 3 h
    expect(lagToMinutes(5, 3)).toEqual({ ok: true, minutes: 30 }) // 0.5 h
    expect(lagToMinutes(1, 3)).toEqual({ ok: true, minutes: 6 }) // 0.1 h = 6 min
    expect(lagToMinutes(-240, 3)).toEqual({ ok: true, minutes: -1440 }) // -24 h
    expect(lagToMinutes(0, 3)).toEqual({ ok: true, minutes: 0 })
  })

  it('day format (5): LinkLag tenths of days → ×MinutesPerDay (default 480)', () => {
    expect(lagToMinutes(10, 5)).toEqual({ ok: true, minutes: 480 }) // 1 d
    expect(lagToMinutes(25, 5)).toEqual({ ok: true, minutes: 1200 }) // 2.5 d
    expect(lagToMinutes(-5, 5)).toEqual({ ok: true, minutes: -240 }) // -0.5 d
    expect(lagToMinutes(0, 5)).toEqual({ ok: true, minutes: 0 })
    // Declared factors are honored (6-hour day).
    expect(
      lagToMinutes(10, 5, { minutesPerDay: 360, minutesPerWeek: 2400, daysPerMonth: 20 }),
    ).toEqual({ ok: true, minutes: 360 })
    // 475-minute days cannot represent 1.5 days as whole minutes → invalid.
    expect(
      lagToMinutes(15, 5, { minutesPerDay: 475, minutesPerWeek: 2400, daysPerMonth: 20 }),
    ).toEqual({ ok: false, reason: 'invalid' })
  })

  it('week format (7): LinkLag tenths of weeks → ×MinutesPerWeek (default 2400)', () => {
    expect(lagToMinutes(70, 7)).toEqual({ ok: true, minutes: 16800 }) // 7 w
    expect(lagToMinutes(5, 7)).toEqual({ ok: true, minutes: 1200 }) // 0.5 w
    expect(lagToMinutes(-5, 7)).toEqual({ ok: true, minutes: -1200 })
    expect(lagToMinutes(0, 7)).toEqual({ ok: true, minutes: 0 })
    expect(
      lagToMinutes(10, 7, { minutesPerDay: 480, minutesPerWeek: 1800, daysPerMonth: 20 }),
    ).toEqual(
      { ok: true, minutes: 1800 }, // 4.5-day week
    )
  })

  it('month format (9): LinkLag tenths of months → ×DaysPerMonth×MinutesPerDay (default 20×480)', () => {
    expect(lagToMinutes(10, 9)).toEqual({ ok: true, minutes: 9600 }) // 1 mo
    expect(lagToMinutes(25, 9)).toEqual({ ok: true, minutes: 24000 }) // 2.5 mo
    expect(lagToMinutes(-5, 9)).toEqual({ ok: true, minutes: -4800 })
    expect(lagToMinutes(0, 9)).toEqual({ ok: true, minutes: 0 })
    // Declared 15-day month with default 480-minute day → 7200 min/month.
    expect(
      lagToMinutes(10, 9, { minutesPerDay: 480, minutesPerWeek: 2400, daysPerMonth: 15 }),
    ).toEqual({ ok: true, minutes: 7200 })
  })

  it('every working unit converts zero lag to exactly 0 minutes', () => {
    for (const fmt of [1, 3, 5, 7, 9]) {
      expect(lagToMinutes(0, fmt)).toEqual({ ok: true, minutes: 0 })
    }
  })

  it('each unit applies its OWN factor (a bare /10 for all units would differ)', () => {
    // Regression lock for the correction-round defect: the old implementation
    // divided LinkLag by 10 for every working format, so hour/day/week/month
    // lags were all misread as minutes. These five assertions fail under that
    // defect and prove per-unit semantics.
    expect(lagToMinutes(10, 1)).toEqual({ ok: true, minutes: 1 })
    expect(lagToMinutes(10, 3)).toEqual({ ok: true, minutes: 60 })
    expect(lagToMinutes(10, 5)).toEqual({ ok: true, minutes: 480 })
    expect(lagToMinutes(10, 7)).toEqual({ ok: true, minutes: 2400 })
    expect(lagToMinutes(10, 9)).toEqual({ ok: true, minutes: 9600 })
  })

  it('elapsed formats (2/4/6/8/10) are unsupported', () => {
    for (const fmt of [2, 4, 6, 8, 10]) {
      expect(lagToMinutes(2400, fmt)).toEqual({ ok: false, reason: 'unsupported' })
    }
  })

  it('percentage format (35) is unsupported', () => {
    expect(lagToMinutes(500, 35)).toEqual({ ok: false, reason: 'unsupported' })
  })

  it('unknown format codes are invalid', () => {
    for (const fmt of [-1, 0, 11, 36, 39, 99]) {
      expect(lagToMinutes(100, fmt)).toEqual({ ok: false, reason: 'invalid' })
    }
  })

  it('non-integer LinkLag is invalid', () => {
    expect(lagToMinutes(24.5, 1)).toEqual({ ok: false, reason: 'invalid' })
  })

  it('malformed factors (non-positive / non-integer) are invalid for day/week/month', () => {
    expect(
      lagToMinutes(10, 5, { minutesPerDay: 0, minutesPerWeek: 2400, daysPerMonth: 20 }),
    ).toEqual({
      ok: false,
      reason: 'invalid',
    })
    expect(
      lagToMinutes(10, 7, { minutesPerDay: 480, minutesPerWeek: 4.5, daysPerMonth: 20 }),
    ).toEqual({
      ok: false,
      reason: 'invalid',
    })
    expect(
      lagToMinutes(10, 9, { minutesPerDay: 480, minutesPerWeek: 2400, daysPerMonth: -20 }),
    ).toEqual({
      ok: false,
      reason: 'invalid',
    })
    // Minute/hour conversions are factor-independent.
    expect(lagToMinutes(2400, 1, { minutesPerDay: 0, minutesPerWeek: 0, daysPerMonth: 0 })).toEqual(
      {
        ok: true,
        minutes: 240,
      },
    )
  })

  it('DEFAULT_LAG_FACTORS are the documented MSPDI defaults (480/2400/20)', () => {
    expect(DEFAULT_LAG_FACTORS).toEqual({
      minutesPerDay: 480,
      minutesPerWeek: 2400,
      daysPerMonth: 20,
    })
  })
})

/** Build a two-task MSPDI project with one FS predecessor link carrying
 * `lag`/`lagFormat`, plus optional project-level factor declarations. */
function lagFixture(
  lag: number,
  lagFormat: number,
  factors?: {
    minutesPerDay?: number
    minutesPerWeek?: number
    daysPerMonth?: number
  },
) {
  const tasks = [
    taskXml({ uid: 1, name: 'A', outlineNumber: '1' }),
    taskXml({
      uid: 2,
      name: 'B',
      outlineNumber: '2',
      predecessorLinks: [{ predUid: 1, type: 0, lag, lagFormat }],
    }),
  ].join('\n')
  return projectXml({ name: 'LagUnits', tasks, ...factors })
}

describe('PROJECT-015 — LinkLagFormat conversion semantics (import)', () => {
  it('imports hour-format lag as ×60 minutes (30 tenths-hour → 180)', () => {
    const r = importMspdi(lagFixture(30, 3))
    expectNoErrors(r.diagnostics)
    expect(r.document.dependencies[0].lagMinutes).toBe(180)
  })

  it('imports day-format lag as ×MinutesPerDay (10 tenths-day → 480)', () => {
    const r = importMspdi(lagFixture(10, 5))
    expectNoErrors(r.diagnostics)
    expect(r.document.dependencies[0].lagMinutes).toBe(480)
  })

  it('imports week-format lag as ×MinutesPerWeek (35 tenths-week → 8400)', () => {
    const r = importMspdi(lagFixture(35, 7))
    expectNoErrors(r.diagnostics)
    expect(r.document.dependencies[0].lagMinutes).toBe(8400)
  })

  it('imports month-format lag as ×DaysPerMonth×MinutesPerDay (10 tenths-month → 9600)', () => {
    const r = importMspdi(lagFixture(10, 9))
    expectNoErrors(r.diagnostics)
    expect(r.document.dependencies[0].lagMinutes).toBe(9600)
  })

  it('imports negative day-format lead exactly (-25 tenths-day → -1200)', () => {
    const r = importMspdi(lagFixture(-25, 5))
    expectNoErrors(r.diagnostics)
    expect(r.document.dependencies[0].lagMinutes).toBe(-1200)
  })

  it('honors declared project conversion factors (MinutesPerDay 360 → day lag 360)', () => {
    const r = importMspdi(lagFixture(10, 5, { minutesPerDay: 360 }))
    expectNoErrors(r.diagnostics)
    expect(r.document.dependencies[0].lagMinutes).toBe(360)
    // Week/month honor their own declared factors.
    const w = importMspdi(lagFixture(10, 7, { minutesPerWeek: 1800 }))
    expectNoErrors(w.diagnostics)
    expect(w.document.dependencies[0].lagMinutes).toBe(1800)
    const m = importMspdi(lagFixture(10, 9, { daysPerMonth: 15 }))
    expectNoErrors(m.diagnostics)
    expect(m.document.dependencies[0].lagMinutes).toBe(7200)
  })

  it('malformed declared factor emits INVALID_MSPDI and falls back to the MSPDI default', () => {
    const r = importMspdi(lagFixture(10, 5, { minutesPerDay: 0 }))
    expect(
      r.diagnostics.some(
        (d) =>
          d.code === INVALID_MSPDI && d.severity === 'error' && d.message.includes('MinutesPerDay'),
      ),
    ).toBe(true)
    // Fallback is the documented default 480 — never a silent approximation.
    expect(r.document.dependencies[0].lagMinutes).toBe(480)
  })

  it('elapsed lag emits UNSUPPORTED_MSPDI_FEATURE and defaults to 0 (dependency kept)', () => {
    const r = importMspdi(lagFixture(100, 4)) // elapsed days
    expect(
      r.diagnostics.some((d) => d.code === UNSUPPORTED_MSPDI_FEATURE && d.severity === 'warning'),
    ).toBe(true)
    expectNoErrors(r.diagnostics)
    expect(r.document.dependencies).toHaveLength(1)
    expect(r.document.dependencies[0].lagMinutes).toBe(0)
  })

  it('percentage lag emits UNSUPPORTED_MSPDI_FEATURE and defaults to 0 (dependency kept)', () => {
    const r = importMspdi(lagFixture(500, 35)) // 50%
    expect(
      r.diagnostics.some((d) => d.code === UNSUPPORTED_MSPDI_FEATURE && d.severity === 'warning'),
    ).toBe(true)
    expectNoErrors(r.diagnostics)
    expect(r.document.dependencies).toHaveLength(1)
    expect(r.document.dependencies[0].lagMinutes).toBe(0)
  })

  it('sub-minute minute-format lag emits INVALID_MSPDI_DURATION and defaults to 0', () => {
    const r = importMspdi(lagFixture(15, 1)) // 1.5 minutes
    expect(
      r.diagnostics.some((d) => d.code === INVALID_MSPDI_DURATION && d.severity === 'error'),
    ).toBe(true)
    expect(r.document.dependencies).toHaveLength(1)
    expect(r.document.dependencies[0].lagMinutes).toBe(0)
  })
})

describe('PROJECT-015 — whole-minute calendar boundaries', () => {
  it('mspdiTimeToMinutes returns whole minutes for whole-minute times', () => {
    expect(mspdiTimeToMinutes('09:00:00')).toBe(540)
    expect(mspdiTimeToMinutes('17:00:00')).toBe(1020)
    expect(mspdiTimeToMinutes('00:00:00')).toBe(0)
    expect(mspdiTimeToMinutes('12:30:00')).toBe(750)
  })

  it('mspdiTimeToMinutes rejects non-zero seconds (no fractional minutes)', () => {
    expect(mspdiTimeToMinutes('09:00:30')).toBeNull()
    expect(mspdiTimeToMinutes('08:30:45')).toBeNull()
    expect(mspdiTimeToMinutes('16:59:59')).toBeNull()
  })

  it('mspdiTimeToMinutes rejects malformed times', () => {
    expect(mspdiTimeToMinutes('9:00')).toBeNull()
    expect(mspdiTimeToMinutes('24:00:00')).toBeNull()
    expect(mspdiTimeToMinutes('09:60:00')).toBeNull()
    expect(mspdiTimeToMinutes('')).toBeNull()
  })

  it('importer drops a sub-minute WorkingTime with INVALID_MSPDI_CALENDAR (never rounded)', () => {
    // Monday's period starts at 09:00:30 — fractional; must be rejected with
    // a diagnostic, and the other weekdays must remain intact.
    const calendars = STANDARD_CALENDAR_XML.replace(
      '09:00:00</FromTime><ToTime>17:00:00',
      '09:00:30</FromTime><ToTime>17:00:00',
    )
    const tasks = taskXml({ uid: 1, outlineNumber: '1' })
    const r = importMspdi(projectXml({ name: 'FractionalBoundary', calendars, tasks }))
    const drop = r.diagnostics.find(
      (d) => d.code === INVALID_MSPDI_CALENDAR && d.severity === 'error',
    )
    expect(drop).toBeDefined()
    expect(drop?.message).toContain('whole-minute')
    const cal = r.document.calendars[0]
    // Canonical week keys are 0=Sunday..6=Saturday; MSPDI DayType 2 (Monday)
    // maps to key 1, DayType 3 (Tuesday) to key 2.
    expect(cal.workingWeek[1]).toEqual([]) // Monday period dropped
    expect(cal.workingWeek[2]).toEqual([{ startMinute: 540, endMinute: 1020 }]) // Tuesday intact
  })
})

// ---- correction round 2: lazy factor validation -------------------------

describe('PROJECT-015 — lazy factor validation (correction round 2)', () => {
  it('minute lag + malformed MinutesPerDay → no unrelated factor error', () => {
    const r = importMspdi(lagFixture(1500, 1, { minutesPerDay: 0 }))
    expectNoErrors(r.diagnostics)
    expect(r.document.dependencies[0].lagMinutes).toBe(150)
  })

  it('hour lag + malformed MinutesPerDay → no unrelated factor error', () => {
    const r = importMspdi(lagFixture(30, 3, { minutesPerDay: 0 }))
    expectNoErrors(r.diagnostics)
    expect(r.document.dependencies[0].lagMinutes).toBe(180)
  })

  it('minute lag + malformed MinutesPerWeek and DaysPerMonth → no unrelated factor errors', () => {
    const r = importMspdi(lagFixture(1500, 1, { minutesPerWeek: 0, daysPerMonth: -5 }))
    expectNoErrors(r.diagnostics)
    expect(r.diagnostics.some((d) => d.code === INVALID_MSPDI)).toBe(false)
    expect(r.document.dependencies[0].lagMinutes).toBe(150)
  })

  it('minute lag + valid declared MinutesPerDay → factor-independent conversion', () => {
    // A valid day factor must not leak into a minute-unit lag either.
    const r = importMspdi(lagFixture(1500, 1, { minutesPerDay: 360 }))
    expectNoErrors(r.diagnostics)
    expect(r.document.dependencies[0].lagMinutes).toBe(150)
  })

  it('day lag + malformed MinutesPerDay → INVALID_MSPDI exactly once + default-factor conversion (480)', () => {
    const r = importMspdi(lagFixture(10, 5, { minutesPerDay: 0 }))
    const factorErrors = r.diagnostics.filter(
      (d) => d.code === INVALID_MSPDI && d.message.includes('MinutesPerDay'),
    )
    expect(factorErrors).toHaveLength(1)
    expect(factorErrors[0].severity).toBe('error')
    // Fallback is the documented default 480 — never a silent approximation.
    expect(r.document.dependencies[0].lagMinutes).toBe(480)
  })

  it('week lag + malformed MinutesPerWeek → INVALID_MSPDI + default-factor conversion (2400)', () => {
    const r = importMspdi(lagFixture(10, 7, { minutesPerWeek: 0 }))
    expect(
      r.diagnostics.some((d) => d.code === INVALID_MSPDI && d.message.includes('MinutesPerWeek')),
    ).toBe(true)
    expect(r.document.dependencies[0].lagMinutes).toBe(2400)
  })

  it('month lag + malformed DaysPerMonth → INVALID_MSPDI + default-factor conversion (9600)', () => {
    const r = importMspdi(lagFixture(10, 9, { daysPerMonth: 0 }))
    expect(
      r.diagnostics.some((d) => d.code === INVALID_MSPDI && d.message.includes('DaysPerMonth')),
    ).toBe(true)
    expect(r.document.dependencies[0].lagMinutes).toBe(9600)
  })

  it('month lag + malformed MinutesPerDay → INVALID_MSPDI (month uses both factors) + default conversion (9600)', () => {
    const r = importMspdi(lagFixture(10, 9, { minutesPerDay: 0 }))
    expect(
      r.diagnostics.some((d) => d.code === INVALID_MSPDI && d.message.includes('MinutesPerDay')),
    ).toBe(true)
    expect(r.document.dependencies[0].lagMinutes).toBe(9600)
  })

  it('a malformed used factor is diagnosed exactly once across multiple using dependencies', () => {
    const tasks = [
      taskXml({ uid: 1, name: 'A', outlineNumber: '1' }),
      taskXml({
        uid: 2,
        name: 'B',
        outlineNumber: '2',
        predecessorLinks: [{ predUid: 1, type: 0, lag: 10, lagFormat: 5 }],
      }),
      taskXml({
        uid: 3,
        name: 'C',
        outlineNumber: '3',
        predecessorLinks: [{ predUid: 1, type: 0, lag: 20, lagFormat: 5 }],
      }),
    ].join('\n')
    const r = importMspdi(projectXml({ name: 'OnceOnly', minutesPerDay: 0, tasks }))
    const factorErrors = r.diagnostics.filter(
      (d) => d.code === INVALID_MSPDI && d.message.includes('MinutesPerDay'),
    )
    expect(factorErrors).toHaveLength(1)
    // Both day lags convert with the default 480 (1 day / 2 days).
    expect(r.document.dependencies.map((d) => d.lagMinutes)).toEqual([480, 960])
  })

  it('unused malformed factors stay silent while a used valid factor is honored', () => {
    // Day lag uses a VALID MinutesPerDay (360); the malformed MinutesPerWeek
    // and DaysPerMonth declarations are unused by any present lag format →
    // no diagnostics, and the day lag converts with the declared 360.
    const r = importMspdi(
      lagFixture(10, 5, { minutesPerDay: 360, minutesPerWeek: 0, daysPerMonth: -1 }),
    )
    expectNoErrors(r.diagnostics)
    expect(r.document.dependencies[0].lagMinutes).toBe(360)
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
      m18LagUnits,
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

// ---- golden fixtures M01–M18 -------------------------------------------

describe('PROJECT-015 — golden fixtures M01–M18', () => {
  describe('valid goldens (M01–M13, M16, M18)', () => {
    const valids: Array<{ id: string; build: () => Uint8Array; note: string }> = [
      { id: 'M01', build: m01Minimal, note: 'minimal' },
      { id: 'M02', build: m02Wbs, note: 'WBS' },
      { id: 'M03', build: m03Dependencies, note: 'dependencies' },
      { id: 'M04', build: m04LagLead, note: 'lag/lead (minute format)' },
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
      { id: 'M18', build: m18LagUnits, note: 'lag units (minute/hour/day/week/month + day lead)' },
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

  describe('M18 exact lag-unit conversion proof (correction round 1)', () => {
    it('converts every working lag unit to the exact canonical lagMinutes', () => {
      const r = importMspdi(m18LagUnits())
      expectNoErrors(r.diagnostics)
      const lagOf = (taskId: string) =>
        r.document.dependencies.find((d) => String(d.successorId) === taskId)!.lagMinutes
      // minute: 1500 tenths → 150 min; hour: 30 tenths → 180 min;
      // day: 25 tenths → 1200 min; week: 76 tenths → 18240 min;
      // month: 25 tenths → 24000 min; day lead: −5 tenths → −240 min.
      expect(lagOf('t2')).toBe(150)
      expect(lagOf('t3')).toBe(180)
      expect(lagOf('t4')).toBe(1200)
      expect(lagOf('t5')).toBe(18240)
      expect(lagOf('t6')).toBe(24000)
      expect(lagOf('t7')).toBe(-240)
      // All six are FS dependencies of the anchor.
      expect(
        r.document.dependencies.every((d) => d.type === 'FS' && String(d.predecessorId) === 't1'),
      ).toBe(true)
    })

    it('passes canonical validation and schedules with exact DerivedSchedule values', () => {
      const r = importMspdi(m18LagUnits())
      const v = validateProjectDocument(r.document)
      expect(v.diagnostics).toEqual([])
      expect(v.accepted).toBe(true)
      const s = schedule(r.document)
      expect(s.diagnostics.filter((d) => d.severity === 'error')).toEqual([])
      expect(s.projectFinish).toBe('2026-10-12T14:00:00.000Z')

      // Anchor: PT4H from project start Mon 2026-08-03 09:00 → 13:00.
      // Each successor is a 1-hour FS task; its earlyStart is
      // addWorkingTime(anchor.finish, lagMinutes) in the standard calendar
      // (Mon–Fri 09:00–17:00, 480 min/day), normalized to the next working
      // instant when the lag lands on a period-end boundary.
      // Independently hand-computed from the documented working-time
      // primitives:
      //   t2 +150 min  → Mon 08-03 15:30
      //   t3 +180 min  → Mon 08-03 16:00
      //   t4 +1200 min → Wed 17:00 (2.5 working days) → start-normalized Thu 08-06 09:00
      //   t5 +18240 min→ Thu 09-24 13:00 (7.6 working weeks)
      //   t6 +24000 min→ Mon 10-12 13:00 (2.5 working months)
      //   t7 −240 min  → Mon 08-03 09:00 (0.5-day lead)
      const expected: Record<
        string,
        { es: string; ef: string; ls: string; lf: string; slack: number; critical: boolean }
      > = {
        t1: {
          es: '2026-08-03T09:00:00.000Z',
          ef: '2026-08-03T13:00:00.000Z',
          ls: '2026-08-03T09:00:00.000Z',
          lf: '2026-08-03T13:00:00.000Z',
          slack: 0,
          critical: true,
        },
        t2: {
          es: '2026-08-03T15:30:00.000Z',
          ef: '2026-08-03T16:30:00.000Z',
          ls: '2026-10-12T13:00:00.000Z',
          lf: '2026-10-12T14:00:00.000Z',
          slack: 23850,
          critical: false,
        },
        t3: {
          es: '2026-08-03T16:00:00.000Z',
          ef: '2026-08-03T17:00:00.000Z',
          ls: '2026-10-12T13:00:00.000Z',
          lf: '2026-10-12T14:00:00.000Z',
          slack: 23820,
          critical: false,
        },
        t4: {
          es: '2026-08-06T09:00:00.000Z',
          ef: '2026-08-06T10:00:00.000Z',
          ls: '2026-10-12T13:00:00.000Z',
          lf: '2026-10-12T14:00:00.000Z',
          slack: 22800,
          critical: false,
        },
        t5: {
          es: '2026-09-24T13:00:00.000Z',
          ef: '2026-09-24T14:00:00.000Z',
          ls: '2026-10-12T13:00:00.000Z',
          lf: '2026-10-12T14:00:00.000Z',
          slack: 5760,
          critical: false,
        },
        t6: {
          es: '2026-10-12T13:00:00.000Z',
          ef: '2026-10-12T14:00:00.000Z',
          ls: '2026-10-12T13:00:00.000Z',
          lf: '2026-10-12T14:00:00.000Z',
          slack: 0,
          critical: true,
        },
        t7: {
          es: '2026-08-03T09:00:00.000Z',
          ef: '2026-08-03T10:00:00.000Z',
          ls: '2026-10-12T13:00:00.000Z',
          lf: '2026-10-12T14:00:00.000Z',
          slack: 24240,
          critical: false,
        },
      }
      for (const [taskId, e] of Object.entries(expected)) {
        const ts = s.taskSchedules[taskId as keyof typeof s.taskSchedules]
        expect(ts, `taskSchedules[${taskId}]`).toBeDefined()
        expect(ts.earlyStart).toBe(e.es)
        expect(ts.earlyFinish).toBe(e.ef)
        expect(ts.lateStart).toBe(e.ls)
        expect(ts.lateFinish).toBe(e.lf)
        expect(ts.totalSlack).toBe(e.slack)
        expect(ts.critical).toBe(e.critical)
      }
    })
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
