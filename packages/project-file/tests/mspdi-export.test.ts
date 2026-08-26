/**
 * PROJECT-016 — MSPDI export test suite.
 *
 * Implements the 38 mandatory test areas + the E01–E15 golden-fixture
 * evidence. Acceptance invariants:
 *
 *   - ROUND-TRIP: for every valid golden, `exportMspdi → importMspdi`
 *     reproduces a byte-identical canonical document (asserted via
 *     `serializeGproj`) — except golden-specific known-loss fields, which are
 *     asserted explicitly alongside their export warnings (no silent loss).
 *   - VALIDATE + SCHEDULE: every valid golden round-trips through
 *     `validateProjectDocument` (accepted) and the canonical `schedule()`
 *     with a byte-identical `DerivedSchedule`.
 *   - DETERMINISM: the same document exports to byte-identical XML across
 *     repeated calls; reordered semantically-equivalent non-identity
 *     collections (calendars, resources, assignments, dependencies, custom
 *     fields, exceptions, availability periods) produce identical bytes;
 *     task sibling order IS semantically meaningful and is preserved.
 *   - EXACT BYTES: representative goldens (E01 full document, E04 lag
 *     encoding) are asserted against hand-embedded canonical XML so the
 *     deterministic writer cannot drift silently.
 *   - NO SILENT LOSS: every unsupported/unrepresentable canonical state
 *     surfaces an explicit export diagnostic (E15 family); semantically
 *     invalid documents are REFUSED (zero bytes + error diagnostics).
 *   - NO AUTHORITY LEAKAGE: the exporter source carries no clock, no
 *     randomness, no `localeCompare`, and no scheduling-package import
 *     (DerivedSchedule is never consulted; no second scheduling algorithm).
 *
 * The tests import `schedule`/`resolveCalendar` from
 * `@genoffice/project-scheduling` at the TEST layer only (same precedent as
 * the accepted PROJECT-015 suite); the package itself stays
 * contracts+engine-only.
 */
import { describe, expect, it } from 'vitest'
import { resolveCalendar, schedule } from '@genoffice/project-scheduling'
import { validateProjectDocument } from '@genoffice/project-engine'
import type { ProjectDocument } from '@genoffice/project-contracts'
import {
  exportMspdi,
  importMspdi,
  mspdiFileAdapter,
  MSPDI_FORMAT,
  serializeGproj,
  parseXml,
  INVALID_MSPDI_EXPORT,
  INVALID_MSPDI_EXPORT_LAG,
  MSPDI_EXPORT_NORMALIZED,
  MSPDI_WRITTEN,
  UNSUPPORTED_MSPDI_EXPORT_FEATURE,
  UNREPRESENTABLE_MSPDI_VALUE,
} from '../src/index.js'
import { decodeUtf8 } from '../src/utf8.js'
import exporterSource from '../src/mspdi/exporter.ts?raw'
import writerSource from '../src/mspdi/xml-writer.ts?raw'
import {
  e01Minimal,
  e02Hierarchy,
  e03Dependencies,
  e04LagLead,
  e05Constraints,
  e06DeadlinesProgress,
  e07Calendars,
  e08Resources,
  e09Assignments,
  e10Baseline,
  e11MultipleBaselines,
  e12CustomFields,
  e13Comprehensive,
  e14Large,
  e15PartialInheritance,
  e15MidnightEnd,
  e15DivergentCapturedAt,
  e15NativeCalendarId,
  e15MultipleNotes,
  e15NumericStringCustomField,
  e15HugeLag,
  e15InvalidUid,
  e15zDuplicateTaskId,
  e15zDependencyCycle,
  e15zMissingCalendar,
  e15zNegativeDuration,
  e15zMissingConstraintDate,
  MONDAY,
} from './mspdi-export-fixtures.js'

// ---- helpers -------------------------------------------------------------

function expectNoErrors(diagnostics: Array<{ severity: string }>): void {
  expect(diagnostics.filter((d) => d.severity === 'error')).toEqual([])
}

function codesOf(result: { diagnostics: Array<{ code: string }> }): string[] {
  return result.diagnostics.map((d) => d.code)
}

function text(bytes: Uint8Array): string {
  return decodeUtf8(bytes)
}

/** Export → import; asserts zero import errors and returns the document. */
function roundTrip(document: ProjectDocument): ProjectDocument {
  const exported = exportMspdi(document)
  expectNoErrors(exported.diagnostics)
  expect(exported.bytes.length).toBeGreaterThan(0)
  const imported = importMspdi(exported.bytes)
  expectNoErrors(imported.diagnostics)
  return imported.document
}

// ---- the 38 required test areas ------------------------------------------

describe('PROJECT-016 — MSPDI export', () => {
  describe('1–5. minimal project, properties, hierarchy, summaries, milestones', () => {
    it('1. minimal project export — exact canonical XML bytes (E01)', () => {
      const xml = text(exportMspdi(e01Minimal()).bytes)
      const weekdays = (dayType: number, working: boolean): string => {
        const head = `        <WeekDay>\n          <DayType>${dayType}</DayType>\n          <DayWorking>${working}</DayWorking>\n`
        if (!working) return `${head}        </WeekDay>`
        return `${head}          <WorkingTimes>\n            <WorkingTime>\n              <FromTime>09:00:00</FromTime>\n              <ToTime>17:00:00</ToTime>\n            </WorkingTime>\n          </WorkingTimes>\n        </WeekDay>`
      }
      const expected = `<?xml version="1.0" encoding="utf-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <SaveVersion>16</SaveVersion>
  <UID>p1</UID>
  <Name>Minimal</Name>
  <StartDate>2026-08-03T09:00:00.000Z</StartDate>
  <Calendars>
    <Calendar>
      <UID>1</UID>
      <Name>Standard</Name>
      <IsBaseCalendar>true</IsBaseCalendar>
      <IsBaseCalendarDefault>true</IsBaseCalendarDefault>
      <WeekDays>
${weekdays(1, false)}
${weekdays(2, true)}
${weekdays(3, true)}
${weekdays(4, true)}
${weekdays(5, true)}
${weekdays(6, true)}
${weekdays(7, false)}
      </WeekDays>
    </Calendar>
  </Calendars>
  <Tasks>
    <Task>
      <UID>1</UID>
      <ID>1</ID>
      <Name>Only Task</Name>
      <WBS>1</WBS>
      <OutlineNumber>1</OutlineNumber>
      <OutlineLevel>1</OutlineLevel>
      <Summary>false</Summary>
      <Milestone>false</Milestone>
      <Manual>false</Manual>
      <Type>1</Type>
      <Priority>500</Priority>
      <Duration>PT8H0M0S</Duration>
      <Work>PT0H0M0S</Work>
      <RemainingWork>PT0H0M0S</RemainingWork>
      <ActualWork>PT0H0M0S</ActualWork>
      <Cost>0</Cost>
      <ActualCost>0</ActualCost>
      <RemainingCost>0</RemainingCost>
      <PercentComplete>0</PercentComplete>
    </Task>
  </Tasks>
</Project>
`
      expect(xml).toEqual(expected)
    })

    it('2. project properties round-trip (name, id, dates, default calendar)', () => {
      const doc = e13Comprehensive()
      const back = roundTrip(doc)
      expect(back.properties.id).toBe('comprehensive-1')
      expect(back.properties.name).toBe('Comprehensive & Co <Alpha>')
      expect(back.properties.startDate).toBe(MONDAY)
      expect(back.properties.finishDate).toBe(doc.properties.finishDate)
      expect(back.properties.statusDate).toBe(doc.properties.statusDate)
      expect(back.properties.defaultCalendarId).toBe(doc.properties.defaultCalendarId)
    })

    it('3. WBS hierarchy round-trips into identical parentTaskId relationships', () => {
      const back = roundTrip(e02Hierarchy())
      const parents = new Map(
        back.tasks.map((t) => [t.id as string, t.parentTaskId as string | undefined]),
      )
      expect(parents.get('t1')).toBeUndefined()
      expect(parents.get('t2')).toBe('t1')
      expect(parents.get('t3')).toBe('t1')
      expect(parents.get('t4')).toBeUndefined()
      expect(back.tasks.map((t) => t.wbs)).toEqual(['1', '1.1', '1.2', '2'])
    })

    it('4. summary tasks round-trip (summary flag from hierarchy)', () => {
      const back = roundTrip(e02Hierarchy())
      expect(back.tasks.map((t) => t.summary)).toEqual([true, false, false, false])
    })

    it('5. milestones round-trip (zero duration + milestone flag)', () => {
      const back = roundTrip(e06DeadlinesProgress())
      const milestone = back.tasks.find((t) => t.name === 'Milestone')!
      expect(milestone.milestone).toBe(true)
      expect(milestone.duration).toBe(0)
    })
  })

  describe('6–12. dependencies and lag/lead', () => {
    it('6–9. FS/SS/FF/SF dependencies round-trip with correct MSPDI type codes', () => {
      const doc = e03Dependencies()
      const xml = text(exportMspdi(doc).bytes)
      // Type codes: FS=0, FF=1, SS=2, SF=3 (inverse of the import map).
      expect(xml).toContain('<Type>0</Type>') // FS
      expect(xml).toContain('<Type>1</Type>') // FF
      expect(xml).toContain('<Type>2</Type>') // SS
      expect(xml).toContain('<Type>3</Type>') // SF
      const back = roundTrip(doc)
      expect(back.dependencies.map((d) => `${d.type}:${d.lagMinutes}`).sort()).toEqual([
        'FF:0',
        'FS:0',
        'SF:0',
        'SS:0',
      ])
    })

    it('10–12. positive lag, negative lead, zero lag — exact canonical encoding (E04)', () => {
      const xml = text(exportMspdi(e04LagLead()).bytes)
      // The deterministic canonical lag representation: working minutes
      // (LinkLagFormat = 1) with LinkLag = lagMinutes × 10.
      expect(xml).toContain(
        [
          '      <PredecessorLink>',
          '        <PredecessorUID>1</PredecessorUID>',
          '        <Type>0</Type>',
          '        <LinkLag>1500</LinkLag>',
          '        <LinkLagFormat>1</LinkLagFormat>',
          '      </PredecessorLink>',
        ].join('\n'),
      )
      expect(xml).toContain('<LinkLag>-2400</LinkLag>')
      expect(xml).toContain('<LinkLag>0</LinkLag>')
      expect(xml).toContain('<LinkLagFormat>1</LinkLagFormat>')
      const back = roundTrip(e04LagLead())
      expect(back.dependencies.map((d) => d.lagMinutes).sort((a, b) => a - b)).toEqual([
        -240, 0, 150,
      ])
    })
  })

  describe('13–15. constraints, deadlines, progress', () => {
    it('13. all eight constraint types round-trip with the MSPDI enumeration', () => {
      const doc = e05Constraints()
      const xml = text(exportMspdi(doc).bytes)
      for (let code = 0; code <= 7; code += 1) {
        expect(xml).toContain(`<ConstraintType>${code}</ConstraintType>`)
      }
      const back = roundTrip(doc)
      const types = back.tasks.map((t) => t.constraintType)
      expect(types).toEqual([
        'asSoonAsPossible',
        'asLateAsPossible',
        'startNoEarlierThan',
        'startNoLaterThan',
        'mustStartOn',
        'finishNoEarlierThan',
        'finishNoLaterThan',
        'mustFinishOn',
      ])
      // Date-bounded constraints keep their dates; ASAP/ALAP carry none.
      expect(back.tasks[2].constraintDate).toBe('2026-08-04T09:00:00.000Z')
      expect(back.tasks[0].constraintDate).toBeUndefined()
    })

    it('14. deadlines round-trip separately from constraints', () => {
      const back = roundTrip(e06DeadlinesProgress())
      const tracked = back.tasks.find((t) => t.name === 'Tracked')!
      expect(tracked.deadline).toBe('2026-08-07T09:00:00.000Z')
      expect(tracked.constraintType).toBeUndefined() // deadline ≠ constraint
      const xml = text(exportMspdi(e06DeadlinesProgress()).bytes)
      expect(xml).toContain('<Deadline>2026-08-07T09:00:00.000Z</Deadline>')
    })

    it('15. progress round-trips (percentComplete, work/cost actuals; physicalPercentComplete warns)', () => {
      const doc = e06DeadlinesProgress()
      const exported = exportMspdi(doc)
      expectNoErrors(exported.diagnostics)
      // The known-loss field is disclosed, not silent.
      expect(
        exported.diagnostics.some(
          (d) => d.code === UNSUPPORTED_MSPDI_EXPORT_FEATURE && d.entityId === 't1',
        ),
      ).toBe(true)
      const back = importMspdi(exported.bytes)
      expectNoErrors(back.diagnostics)
      // physicalPercentComplete is the ONE known-loss field of E06.
      const tracked = back.document.tasks.find((t) => t.name === 'Tracked')!
      expect(tracked.percentComplete).toBe(50)
      expect(tracked.work).toBe(480)
      expect(tracked.actualWork).toBe(240)
      expect(tracked.remainingWork).toBe(240)
      expect(tracked.cost).toBe(1000)
      expect(tracked.actualCost).toBe(500)
      expect(tracked.remainingCost).toBe(500)
      expect(tracked.physicalPercentComplete).toBeUndefined()
      const strippedOriginal = {
        ...doc,
        tasks: doc.tasks.map((t) =>
          t.physicalPercentComplete === undefined
            ? t
            : { ...t, physicalPercentComplete: undefined },
        ),
      }
      expect(serializeGproj(back.document)).toEqual(serializeGproj(strippedOriginal))
      // Schedule equivalence still holds.
      expect(JSON.stringify(schedule(back.document))).toEqual(
        JSON.stringify(schedule(strippedOriginal)),
      )
    })
  })

  describe('16–18. calendars, inheritance, exceptions', () => {
    it('16. calendars round-trip (week, default marker first, names)', () => {
      const doc = e07Calendars()
      const back = roundTrip(doc)
      expect(back.calendars).toHaveLength(2)
      expect(back.properties.defaultCalendarId).toBe('c1')
      expect(back.calendars.find((c) => c.id === 'c2')?.baseCalendarId).toBe('c1')
    })

    it('17. calendar inheritance round-trips (BaseCalendarUID preserved, not flattened)', () => {
      const doc = e07Calendars()
      const xml = text(exportMspdi(doc).bytes)
      const baseUidIndex = xml.indexOf('<BaseCalendarUID>')
      expect(baseUidIndex).toBeGreaterThan(-1)
      const back = roundTrip(doc)
      const derived = back.calendars.find((c) => c.id === 'c2')!
      expect(derived.baseCalendarId).toBe('c1')
      // The derived calendar keeps its own (overriding) week.
      expect(derived.workingWeek[2]).toEqual([{ startMinute: 960, endMinute: 1320 }])
    })

    it('17b. partial inherited workingWeek is materialized with exactly recoverable semantics', () => {
      const doc = e15PartialInheritance()
      const exported = exportMspdi(doc)
      expectNoErrors(exported.diagnostics)
      // The materialization is disclosed as a normalization note.
      expect(
        exported.diagnostics.some((d) => d.code === MSPDI_EXPORT_NORMALIZED && d.entityId === 'c2'),
      ).toBe(true)
      const back = importMspdi(exported.bytes)
      expectNoErrors(back.diagnostics)
      // The RESOLVED calendars are identical (resolveCalendar is the
      // canonical semantics — the exporter never flattens past equivalence).
      const original = resolveCalendar({ calendars: doc.calendars }, doc.calendars[1].id)
      const roundTripped = resolveCalendar(
        { calendars: back.document.calendars },
        back.document.calendars.find((c) => c.name === 'Tuesdays Only Override')!.id,
      )
      expect(JSON.stringify(roundTripped)).toEqual(JSON.stringify(original))
      // Tuesday is overridden by the child; the other weekdays inherit the
      // base's periods; Sunday/Saturday stay non-working.
      expect(original.workingWeek[2]).toEqual([{ startMinute: 600, endMinute: 840 }])
      expect(original.workingWeek[3]).toEqual([{ startMinute: 540, endMinute: 1020 }])
      expect(original.workingWeek[0]).toEqual([])
    })

    it('18. calendar exceptions round-trip (working + non-working)', () => {
      const back = roundTrip(e07Calendars())
      const standard = back.calendars.find((c) => c.id === 'c1')!
      expect(standard.exceptions).toEqual([{ date: '2026-12-25', periods: [] }])
      const derived = back.calendars.find((c) => c.id === 'c2')!
      expect(derived.exceptions).toEqual([
        { date: '2026-08-08', periods: [{ startMinute: 600, endMinute: 840 }] },
        { date: '2026-12-25', periods: [] },
      ])
    })
  })

  describe('19–22. resources and assignments', () => {
    it('19. work resources round-trip (rates, max units, calendar, availability)', () => {
      const back = roundTrip(e08Resources())
      const engineer = back.resources.find((r) => r.name === 'Engineer')!
      expect(engineer.kind).toBe('work')
      expect(engineer.maxUnits).toBe(1)
      expect(engineer.standardRate).toBe(50)
      expect(engineer.overtimeRate).toBe(75)
      expect(engineer.calendarId).toBe('c1')
      expect(engineer.availability).toEqual([
        { start: '2026-08-03T09:00:00.000Z', finish: '2026-12-31T17:00:00.000Z', units: 1 },
      ])
    })

    it('20. material resources round-trip', () => {
      const back = roundTrip(e08Resources())
      const concrete = back.resources.find((r) => r.name === 'Concrete')!
      expect(concrete.kind).toBe('material')
      expect(concrete.standardRate).toBe(120)
      expect(concrete.costPerUse).toBe(10)
    })

    it('21. cost resources round-trip', () => {
      const back = roundTrip(e08Resources())
      const legal = back.resources.find((r) => r.name === 'Legal Fee')!
      expect(legal.kind).toBe('cost')
      expect(legal.costPerUse).toBe(500)
    })

    it('22. assignments round-trip across all three resource kinds', () => {
      const back = roundTrip(e09Assignments())
      expect(back.assignments).toHaveLength(3)
      const byTask = new Map(back.assignments.map((a) => [a.taskId as string, a]))
      const a1 = byTask.get('t1')!
      expect(a1.resourceId).toBe('r1')
      expect(a1.units).toBe(1)
      expect(a1.work).toBe(480)
      const a2 = byTask.get('t2')!
      expect(a2.resourceId).toBe('r2')
      expect(a2.units).toBe(10)
      expect(a2.cost).toBe(1210)
      const a3 = byTask.get('t3')!
      expect(a3.resourceId).toBe('r3')
      expect(a3.cost).toBe(500)
    })
  })

  describe('23–25. baselines and custom fields', () => {
    it('23. baselines round-trip (snapshots, capturedAt via <LastSaved>)', () => {
      const doc = e10Baseline()
      const back = roundTrip(doc)
      expect(back.baselines).toHaveLength(1)
      const baseline = back.baselines[0]
      expect(baseline.id).toBe('b0')
      expect(baseline.name).toBe('Baseline')
      expect(baseline.capturedAt).toBe('2026-08-01T12:00:00.000Z')
      expect(baseline.taskSnapshots['t1']).toEqual({
        start: '2026-08-03T09:00:00.000Z',
        finish: '2026-08-04T09:00:00.000Z',
        duration: 480,
        work: 480,
        cost: 100,
      })
      expect(baseline.taskSnapshots['t2']).toEqual({ duration: 960, work: 0, cost: 0 })
    })

    it('24. multiple baselines round-trip (deterministic slot ordering)', () => {
      const doc = e11MultipleBaselines()
      const xml = text(exportMspdi(doc).bytes)
      expect(xml.indexOf('<Baseline>')).toBeGreaterThan(-1)
      expect(xml.indexOf('<Baseline>')).toBeLessThan(xml.indexOf('<Baseline1>'))
      const back = roundTrip(doc)
      expect(back.baselines.map((b) => b.id)).toEqual(['b0', 'b1'])
      expect(back.baselines[1].taskSnapshots['t1'].cost).toBe(120)
    })

    it('25. custom fields round-trip (definitions + values incl. null)', () => {
      const back = roundTrip(e12CustomFields())
      expect(back.customFields.map((f) => `${f.id}:${f.type}:${f.name}`)).toEqual([
        'cf1:text:Comment',
        'cf2:number:Score',
        'cf3:boolean:Flag',
        'cf4:date:Reviewed On',
        'cf5:text:Optional',
      ])
      const rich = back.tasks.find((t) => t.name === 'Rich')!
      expect(rich.customFields).toEqual({
        cf1: 'hello world',
        cf2: 42,
        cf3: true,
        cf4: '2026-08-03T00:00:00.000Z',
        cf5: null,
      })
      const sparse = back.tasks.find((t) => t.name === 'Sparse')!
      expect(sparse.customFields).toEqual({ cf1: 'second', cf2: -7.5, cf3: false })
    })
  })

  describe('26–29. determinism', () => {
    it('26. deterministic XML — three exports produce identical bytes', () => {
      const doc = e13Comprehensive()
      const a = exportMspdi(doc).bytes
      const b = exportMspdi(doc).bytes
      const c = exportMspdi(doc).bytes
      expect(b).toEqual(a)
      expect(c).toEqual(a)
    })

    it('27. repeated export across all goldens is byte-stable', () => {
      for (const doc of [
        e01Minimal(),
        e02Hierarchy(),
        e03Dependencies(),
        e04LagLead(),
        e05Constraints(),
        e07Calendars(),
        e08Resources(),
        e09Assignments(),
        e10Baseline(),
        e11MultipleBaselines(),
        e12CustomFields(),
        e14Large(),
      ]) {
        expect(exportMspdi(doc).bytes).toEqual(exportMspdi(doc).bytes)
      }
    })

    it('28. round-trip through the PROJECT-015 importer — byte-identical canonical documents', () => {
      for (const doc of [
        e01Minimal(),
        e02Hierarchy(),
        e03Dependencies(),
        e04LagLead(),
        e05Constraints(),
        e07Calendars(),
        e08Resources(),
        e09Assignments(),
        e10Baseline(),
        e11MultipleBaselines(),
        e12CustomFields(),
        e13Comprehensive(),
        e14Large(),
      ]) {
        const back = roundTrip(doc)
        expect(serializeGproj(back)).toEqual(serializeGproj(doc))
        expect(validateProjectDocument(back).accepted).toBe(true)
        expect(JSON.stringify(schedule(back))).toEqual(JSON.stringify(schedule(doc)))
      }
    })

    it('29. reordered canonical inputs produce identical canonical MSPDI bytes', () => {
      const doc = e13Comprehensive()
      const baseline = text(exportMspdi(doc).bytes)
      // Reverse every order-insignificant collection (identity collections
      // keep their ids; only ARRAY order changes).
      const shuffled: ProjectDocument = {
        ...doc,
        calendars: [...doc.calendars].reverse(),
        resources: [...doc.resources].reverse(),
        assignments: [...doc.assignments].reverse(),
        dependencies: [...doc.dependencies].reverse(),
        customFields: [...doc.customFields].reverse(),
      }
      expect(text(exportMspdi(shuffled).bytes)).toEqual(baseline)
      // Exceptions and availability periods are reorder-invariant too.
      const shuffledExceptions: ProjectDocument = {
        ...doc,
        calendars: doc.calendars.map((c) => ({
          ...c,
          exceptions: [...c.exceptions].reverse(),
        })),
      }
      expect(text(exportMspdi(shuffledExceptions).bytes)).toEqual(baseline)
      const shuffledAvailability: ProjectDocument = {
        ...doc,
        resources: doc.resources.map((r) => ({
          ...r,
          availability: [...r.availability].reverse(),
        })),
      }
      expect(text(exportMspdi(shuffledAvailability).bytes)).toEqual(baseline)
    })

    it('29b. task sibling order IS semantically meaningful — reordering changes the XML', () => {
      const doc = e02Hierarchy()
      const baseline = text(exportMspdi(doc).bytes)
      const reordered: ProjectDocument = {
        ...doc,
        // Swap the two children of t1: sibling order changes → different
        // outline numbers → different canonical XML (correct behavior).
        tasks: [doc.tasks[0], doc.tasks[2], doc.tasks[1], doc.tasks[3]],
      }
      const xml = text(exportMspdi(reordered).bytes)
      expect(xml).not.toEqual(baseline)
      const back = roundTrip(reordered)
      // DFS emission keeps parents first; the swapped siblings keep their new
      // relative order, so t3 now owns outline 1.1 and t2 owns 1.2.
      expect(back.tasks.map((t) => `${t.id}:${t.wbs}`)).toEqual([
        't1:1',
        't3:1.1',
        't2:1.2',
        't4:2',
      ])
    })
  })

  describe('30–31. diagnostics: unsupported, unrepresentable, refusal', () => {
    it('30a. divergent baseline capturedAt values are disclosed (single <LastSaved> carrier)', () => {
      const doc = e15DivergentCapturedAt()
      const r = exportMspdi(doc)
      expectNoErrors(r.diagnostics)
      expect(r.diagnostics.filter((d) => d.code === UNREPRESENTABLE_MSPDI_VALUE)).toHaveLength(1)
      const back = importMspdi(r.bytes)
      expect(back.diagnostics.filter((d) => d.severity === 'error')).toEqual([])
      // Both baselines re-import with the carrier (first-slot) capturedAt.
      expect(back.document.baselines.map((b) => b.capturedAt)).toEqual([MONDAY, MONDAY])
    })

    it('30b. a working period ending at 24:00 is disclosed as unrepresentable', () => {
      const r = exportMspdi(e15MidnightEnd())
      expectNoErrors(r.diagnostics)
      expect(
        r.diagnostics.some(
          (d) => d.code === UNREPRESENTABLE_MSPDI_VALUE && d.message.includes('24:00:00'),
        ),
      ).toBe(true)
      expect(text(r.bytes)).toContain('<ToTime>24:00:00</ToTime>')
    })

    it('30c. non-convention calendar identity is remapped consistently with a warning', () => {
      const doc = e15NativeCalendarId()
      const r = exportMspdi(doc)
      expectNoErrors(r.diagnostics)
      expect(
        r.diagnostics.some(
          (d) => d.code === UNREPRESENTABLE_MSPDI_VALUE && d.entityId === 'standard',
        ),
      ).toBe(true)
      const back = importMspdi(r.bytes)
      expect(back.diagnostics.filter((d) => d.severity === 'error')).toEqual([])
      // Every reference to 'standard' remapped consistently to 'c0'.
      expect(back.document.calendars[0].id).toBe('c0')
      expect(back.document.properties.defaultCalendarId).toBe('c0')
      expect(back.document.tasks[0].calendarId).toBe('c0')
    })

    it('30d. multiple notes collapse into the single <Notes> field with a warning', () => {
      const doc = e15MultipleNotes()
      const r = exportMspdi(doc)
      expectNoErrors(r.diagnostics)
      expect(
        r.diagnostics.some(
          (d) => d.code === UNSUPPORTED_MSPDI_EXPORT_FEATURE && d.entityId === 't1',
        ),
      ).toBe(true)
      const back = importMspdi(r.bytes)
      expect(back.document.tasks[0].notes).toEqual(['first note\nsecond note'])
    })

    it('30e. string custom-field values that re-parse as number/boolean are disclosed', () => {
      const doc = e15NumericStringCustomField()
      const r = exportMspdi(doc)
      expectNoErrors(r.diagnostics)
      expect(r.diagnostics.filter((d) => d.code === UNREPRESENTABLE_MSPDI_VALUE)).toHaveLength(2)
      const back = importMspdi(r.bytes)
      expect(back.document.tasks[0].customFields).toEqual({ cf1: 123, cf2: true })
    })

    it('30f. a canonical wbs that differs from the derived outline is disclosed', () => {
      const doc = e01Minimal()
      const odd = { ...doc, tasks: [{ ...doc.tasks[0], wbs: 'A.9' }] }
      const r = exportMspdi(odd)
      expectNoErrors(r.diagnostics)
      expect(
        r.diagnostics.some(
          (d) => d.code === UNREPRESENTABLE_MSPDI_VALUE && d.message.includes('A.9'),
        ),
      ).toBe(true)
      // The exported outline still reconstructs the SAME hierarchy (none).
      const back = importMspdi(r.bytes)
      expect(back.document.tasks[0].parentTaskId).toBeUndefined()
    })

    it('30g. views/tables/filters/groups have no MSPDI representation — disclosed', () => {
      const doc = e01Minimal()
      const withView = {
        ...doc,
        views: [{ id: 'v1' as never, name: 'Gantt', type: 'gantt' }],
        tables: [{ id: 'tb1' as never, name: 'Entry', columns: ['Name'] }],
      }
      const r = exportMspdi(withView as unknown as ProjectDocument)
      expectNoErrors(r.diagnostics)
      expect(
        r.diagnostics.some(
          (d) =>
            d.code === UNSUPPORTED_MSPDI_EXPORT_FEATURE &&
            d.message.includes('view/table/filter/group'),
        ),
      ).toBe(true)
      const back = importMspdi(r.bytes)
      expect(back.document.views).toEqual([])
    })

    it('30h. an unrepresentable lag is exported as 0 with an error (never silently changed)', () => {
      const doc = e15HugeLag()
      const r = exportMspdi(doc)
      expect(
        r.diagnostics.some((d) => d.code === INVALID_MSPDI_EXPORT_LAG && d.severity === 'error'),
      ).toBe(true)
      const back = importMspdi(r.bytes)
      expect(back.document.dependencies[0].lagMinutes).toBe(0)
      expect(back.document.dependencies[0].type).toBe('FS') // dependency retained
    })

    it('30i. invalid task/resource uids are synthesized deterministically with errors', () => {
      const doc = e15InvalidUid()
      const r = exportMspdi(doc)
      const errors = r.diagnostics.filter((d) => d.code === INVALID_MSPDI_EXPORT)
      expect(errors).toHaveLength(2)
      const back = importMspdi(r.bytes)
      expect(back.document.tasks.map((t) => t.uid).sort((a, b) => a - b)).toEqual([0, 1])
      expect(back.document.resources.map((x) => x.uid).sort((a, b) => a - b)).toEqual([0, 2])
      // Synthesis is deterministic (no random IDs).
      const again = importMspdi(exportMspdi(doc).bytes)
      expect(serializeGproj(again.document)).toEqual(serializeGproj(back.document))
    })

    it('30j. a non-hierarchical task array order is canonicalized to DFS with a warning', () => {
      const doc = e02Hierarchy()
      // Child t2 before parent t1 — engine-valid, but not DFS order.
      const odd = { ...doc, tasks: [doc.tasks[1], doc.tasks[0], doc.tasks[2], doc.tasks[3]] }
      const r = exportMspdi(odd)
      expectNoErrors(r.diagnostics)
      expect(
        r.diagnostics.some(
          (d) => d.code === UNSUPPORTED_MSPDI_EXPORT_FEATURE && d.message.includes('DFS'),
        ),
      ).toBe(true)
      const back = importMspdi(r.bytes)
      // Same hierarchy + same sibling order as the canonical DFS projection.
      expect(back.document.tasks.map((t) => `${t.id}:${t.wbs}:${t.parentTaskId ?? '-'}`)).toEqual([
        't1:1:-',
        't2:1.1:t1',
        't3:1.2:t1',
        't4:2:-',
      ])
    })

    it('30k. task.baseline reverse-index handling is disclosed (derived, reconstructed empty)', () => {
      const doc = e10Baseline()
      const populated = {
        ...doc,
        tasks: doc.tasks.map((t, i) => ({ ...t, baseline: i === 0 ? (['b0'] as never) : [] })),
      }
      const r = exportMspdi(populated)
      expectNoErrors(r.diagnostics)
      expect(
        r.diagnostics.some(
          (d) => d.code === UNSUPPORTED_MSPDI_EXPORT_FEATURE && d.entityId === 't1',
        ),
      ).toBe(true)
      const back = importMspdi(r.bytes)
      expect(back.document.tasks[0].baseline).toEqual([]) // reconstructed empty
      expect(back.document.baselines[0].taskSnapshots['t1'].cost).toBe(100) // data preserved
    })

    it('31. malformed canonical documents are REFUSED (zero bytes + errors)', () => {
      for (const [label, doc] of [
        ['duplicate task id', e15zDuplicateTaskId()],
        ['dependency cycle', e15zDependencyCycle()],
        ['missing calendar reference', e15zMissingCalendar()],
        ['negative duration', e15zNegativeDuration()],
        ['missing constraint date', e15zMissingConstraintDate()],
      ] as Array<[string, ProjectDocument]>) {
        const r = exportMspdi(doc)
        expect(r.bytes.length).toBe(0)
        expect(
          r.diagnostics.some((d) => d.code === INVALID_MSPDI_EXPORT && d.severity === 'error'),
        ).toBe(true)
        // The engine's own diagnostics are surfaced verbatim (error level).
        expect(r.diagnostics.filter((d) => d.severity === 'error').length).toBeGreaterThan(1)
        void label
      }
    })
  })

  describe('32–34. no random IDs, no current-time values, no DerivedSchedule authority', () => {
    it('32. no random IDs — identity synthesis is a pure function of the document', () => {
      const doc = e15NativeCalendarId()
      const a = exportMspdi(doc)
      const b = exportMspdi(doc)
      expect(b.bytes).toEqual(a.bytes)
      expect(JSON.stringify(b.diagnostics)).toEqual(JSON.stringify(a.diagnostics))
    })

    it('33. no current-time values — exporter/writer source has no clock', () => {
      // Static source guard: the export modules never read the clock, never
      // use randomness, and never consult the host locale.
      for (const source of [exporterSource, writerSource]) {
        expect(source).not.toContain('Date.now')
        expect(source).not.toContain('new Date(')
        expect(source).not.toContain('Math.random')
        expect(source).not.toContain('localeCompare')
      }
      // Behavioral proof: the exact E01 bytes (asserted in test 1) contain no
      // timestamp outside the document's own canonical values — the same
      // document exported at any time produces those exact bytes (test 26).
    })

    it('34. no export of DerivedSchedule authority — no scheduling import, no derived XML', () => {
      // Static guard: the exporter does not import the scheduling package
      // (the canonical scheduler stays the sole scheduling authority) and
      // does not reference derived-schedule state.
      expect(exporterSource).not.toContain('project-scheduling')
      expect(exporterSource).not.toContain('DerivedSchedule')
      expect(exporterSource).not.toContain('earlyStart')
      expect(exporterSource).not.toContain('totalSlack')
      expect(exporterSource).not.toContain('critical')
      // Behavioral proof: the exported XML carries no derived-schedule data.
      const xml = text(exportMspdi(e13Comprehensive()).bytes)
      expect(xml).not.toContain('EarlyStart')
      expect(xml).not.toContain('TotalSlack')
      expect(xml).not.toContain('Critical')
      expect(xml).not.toContain('ScheduledStart')
    })
  })

  describe('35–36. XML escaping and parser round-trip', () => {
    it('35. XML escaping — &, <, >, quotes and CR round-trip exactly', () => {
      const doc = e01Minimal()
      const gnarly = {
        ...doc,
        properties: { ...doc.properties, name: 'A & B <C> "D" \'E\'' },
        tasks: [
          {
            ...doc.tasks[0],
            name: 'Fish & <Chips> "x" \'y\'',
            notes: ['line one\r\nline two'],
          },
        ],
      }
      const r = exportMspdi(gnarly)
      expectNoErrors(r.diagnostics)
      const xml = text(r.bytes)
      expect(xml).toContain('A &amp; B &lt;C&gt; "D" \'E\'')
      expect(xml).toContain('Fish &amp; &lt;Chips&gt; "x" \'y\'')
      expect(xml).toContain('&#xD;')
      const back = importMspdi(r.bytes)
      expectNoErrors(back.diagnostics)
      expect(back.document.properties.name).toBe('A & B <C> "D" \'E\'')
      expect(back.document.tasks[0].name).toBe('Fish & <Chips> "x" \'y\'')
      expect(back.document.tasks[0].notes).toEqual(['line one\r\nline two'])
    })

    it('36. exported XML parses through the accepted PROJECT-015 parser', () => {
      const xml = exportMspdi(e13Comprehensive()).bytes
      const root = parseXml(xml)
      expect(root.name).toBe('Project')
      expect(root.attributes['xmlns']).toBe('http://schemas.microsoft.com/project')
      const tasks = root.children.find((c) => c.name === 'Tasks')!
      expect(tasks.children.filter((c) => c.name === 'Task')).toHaveLength(4)
      const calendars = root.children.find((c) => c.name === 'Calendars')!
      expect(calendars.children.filter((c) => c.name === 'Calendar')).toHaveLength(2)
      // Declared namespace + declaration are stable byte prefixes.
      expect(
        text(xml).startsWith(
          '<?xml version="1.0" encoding="utf-8"?>\n<Project xmlns="http://schemas.microsoft.com/project">',
        ),
      ).toBe(true)
    })
  })

  describe('37–38. large project and schedule equivalence', () => {
    it('37. large project export (60 tasks / 59 dependencies / 10 resources / 60 assignments)', () => {
      const doc = e14Large()
      const r = exportMspdi(doc)
      expectNoErrors(r.diagnostics)
      const back = roundTrip(doc)
      expect(back.tasks).toHaveLength(60)
      expect(back.dependencies).toHaveLength(59)
      expect(back.resources).toHaveLength(10)
      expect(back.assignments).toHaveLength(60)
      expect(serializeGproj(back)).toEqual(serializeGproj(doc))
      expect(text(r.bytes)).toContain('<ID>60</ID>')
    })

    it('38. export/import schedule equivalence across every valid golden', () => {
      for (const doc of [
        e01Minimal(),
        e02Hierarchy(),
        e03Dependencies(),
        e04LagLead(),
        e05Constraints(),
        e07Calendars(),
        e08Resources(),
        e09Assignments(),
        e10Baseline(),
        e11MultipleBaselines(),
        e12CustomFields(),
        e13Comprehensive(),
        e14Large(),
      ]) {
        const back = roundTrip(doc)
        const s1 = schedule(doc)
        const s2 = schedule(back)
        expect(JSON.stringify(s2)).toEqual(JSON.stringify(s1))
      }
    })
  })

  describe('adapter surface + diagnostics contract', () => {
    it('mspdiFileAdapter.export is the canonical export path', () => {
      const doc = e01Minimal()
      const r = mspdiFileAdapter.export(doc)
      expect(r.bytes.length).toBeGreaterThan(0)
      const back = mspdiFileAdapter.import(r.bytes)
      expectNoErrors(back.diagnostics)
      expect(serializeGproj(back.document)).toEqual(serializeGproj(doc))
    })

    it('mspdiFileAdapter.format is "mspdi"', () => {
      expect(mspdiFileAdapter.format).toBe(MSPDI_FORMAT)
    })

    it('MSPDI_WRITTEN info diagnostic is emitted on success', () => {
      const r = exportMspdi(e01Minimal())
      expect(r.diagnostics.some((d) => d.code === MSPDI_WRITTEN && d.severity === 'info')).toBe(
        true,
      )
    })

    it('export diagnostics use only PROJECT-016 export codes (plus verbatim engine codes on refusal)', () => {
      const allowed = new Set([
        'INVALID_MSPDI_EXPORT',
        'INVALID_MSPDI_EXPORT_LAG',
        'UNREPRESENTABLE_MSPDI_VALUE',
        'UNSUPPORTED_MSPDI_EXPORT_FEATURE',
        'MSPDI_EXPORT_NORMALIZED',
        'MSPDI_WRITTEN',
      ])
      for (const doc of [
        e01Minimal(),
        e15PartialInheritance(),
        e15MidnightEnd(),
        e15DivergentCapturedAt(),
        e15NativeCalendarId(),
        e15MultipleNotes(),
        e15NumericStringCustomField(),
        e15HugeLag(),
        e15InvalidUid(),
      ]) {
        for (const d of exportMspdi(doc).diagnostics) {
          expect(allowed.has(d.code)).toBe(true)
        }
      }
      // Refusal additionally surfaces the engine's own codes verbatim.
      const refusal = exportMspdi(e15zDependencyCycle()).diagnostics
      expect(refusal[0].code).toBe('INVALID_MSPDI_EXPORT')
      expect(refusal.some((d) => d.code === 'DEPENDENCY_CYCLE')).toBe(true)
      void codesOf
    })
  })
})
