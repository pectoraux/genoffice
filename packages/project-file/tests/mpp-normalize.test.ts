/**
 * PROJECT-018 — MPP foundation contract tests (pure TypeScript, no Java).
 *
 * Covers the five PROJECT-017-approved normalizations N1–N5 against
 * synthetic MPXJ-convention MSPDI payloads (the exact conventions observed
 * in the real corpus during the PROJECT-017 spike and the PROJECT-018
 * trial run), the staged-diagnostic contract of `importMppFromMspdi`, and
 * the error-atomicity rules. The real-corpus end-to-end evidence lives in
 * packages/project-mpp-host (which needs java + the pinned MPXJ dist).
 *
 * The negative control in every normalization test is the SAME payload
 * WITHOUT normalization: the accepted PROJECT-015 importer produces the
 * documented error there — proving the normalization removes a real,
 * observed failure mode rather than a hypothetical one.
 */
import { describe, expect, it } from 'vitest'
import { importMspdi, parseXml, emptyProjectDocument } from '../src/index'
import {
  MPP_NORMALIZED_BASE_CALENDAR_SENTINEL,
  MPP_NORMALIZED_MIDNIGHT_PERIOD,
  MPP_NORMALIZED_PLACEHOLDER_RECORD,
  MPP_NORMALIZED_SENTINEL_REFERENCE,
  MPP_DROPPED_UNASSIGNED_ASSIGNMENT,
  MPP_SIDECAR_UNAVAILABLE,
  MPP_DIAGNOSTIC_CODES,
  MPXJ_PINNED_VERSION,
  MPP_SUPPORTED_FORMAT_VERSIONS,
  MPP_MAX_INPUT_BYTES,
  MPP_MAX_MSPDI_OUTPUT_BYTES,
  importMppFromMspdi,
  normalizeMspdiForCanonicalImport,
  serializeXmlNode,
  type MppConversionOutcome,
} from '../src/index'
import { validateProjectDocument } from '@genoffice/project-engine'
import {
  STANDARD_CALENDAR_XML,
  assignmentXml,
  projectXml,
  resourceXml,
  taskXml,
} from './mspdi-fixtures'

/** A calendar whose BaseCalendarUID is the MPXJ `-1` sentinel. */
const SENTINEL_BASE_CALENDAR_XML = `
      <Calendar>
        <UID>2</UID>
        <Name>Nights</Name>
        <BaseCalendarUID>-1</BaseCalendarUID>
        <IsBaseCalendar>false</IsBaseCalendar>
        <WeekDays>
          <WeekDay><DayType>1</DayType><DayWorking>false</DayWorking></WeekDay>
          <WeekDay><DayType>2</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
        </WeekDays>
        <Exceptions />
      </Calendar>`

/** A weekday whose working period runs "until midnight" (MPXJ convention:
 * ToTime 00:00:00). */
const MIDNIGHT_CALENDAR_XML = `
      <Calendar>
        <UID>1</UID>
        <Name>Standard</Name>
        <IsBaseCalendar>true</IsBaseCalendar>
        <IsBaseCalendarDefault>true</IsBaseCalendarDefault>
        <WeekDays>
          <WeekDay><DayType>1</DayType><DayWorking>false</DayWorking></WeekDay>
          <WeekDay><DayType>2</DayType><DayWorking>true</DayWorking><WorkingTimes>
            <WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime>
            <WorkingTime><FromTime>23:00:00</FromTime><ToTime>00:00:00</ToTime></WorkingTime>
          </WorkingTimes></WeekDay>
          <WeekDay><DayType>3</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
          <WeekDay><DayType>4</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
          <WeekDay><DayType>5</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
          <WeekDay><DayType>6</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
          <WeekDay><DayType>7</DayType><DayWorking>false</DayWorking></WeekDay>
        </WeekDays>
        <Exceptions />
      </Calendar>`

function outcome(
  bytes: Uint8Array,
  sidecarDiagnostics: MppConversionOutcome['sidecarDiagnostics'] = [],
): MppConversionOutcome {
  return {
    mspdiBytes: bytes,
    frame: { version: 1, requestId: 'test', ok: true },
    sidecarDiagnostics,
  }
}

function codesOf(diags: ReadonlyArray<{ code: string }>): string[] {
  return [...new Set(diags.map((d) => d.code))]
}

describe('PROJECT-018 — N1: sentinel CalendarUID strip', () => {
  const bytes = projectXml({
    name: 'N1',
    tasks: taskXml({ uid: 1, outlineNumber: '1', calendarUid: -1 }),
  })

  it('the un-normalized payload fails in the accepted importer (negative control)', () => {
    const raw = importMspdi(bytes)
    expect(
      raw.diagnostics.some((d) => d.code === 'INVALID_MSPDI_REFERENCE' && d.severity === 'error'),
    ).toBe(true)
  })

  it('normalization strips the sentinel and diagnoses it (stage normalization, severity info)', () => {
    const result = normalizeMspdiForCanonicalImport(bytes)
    expect(result.diagnostics).toHaveLength(1)
    expect(result.diagnostics[0].code).toBe(MPP_NORMALIZED_SENTINEL_REFERENCE)
    expect(result.diagnostics[0].severity).toBe('info')
    expect(result.diagnostics[0].stage).toBe('normalization')
    expect(result.diagnostics[0].entityId).toBe('t1')
    // The rewritten XML no longer contains the sentinel:
    const text = new TextDecoder().decode(result.bytes)
    expect(text).not.toContain('<CalendarUID>-1</CalendarUID>')
  })

  it('the normalized payload imports with zero sentinel-reference errors', () => {
    const result = normalizeMspdiForCanonicalImport(bytes)
    const imported = importMspdi(result.bytes)
    expect(imported.diagnostics.filter((d) => d.code === 'INVALID_MSPDI_REFERENCE')).toEqual([])
    expect(imported.document.tasks).toHaveLength(1)
    expect(validateProjectDocument(imported.document).accepted).toBe(true)
  })

  it('resource sentinel CalendarUID is stripped too (with r-id entity)', () => {
    const withResource = projectXml({
      name: 'N1r',
      tasks: taskXml({ uid: 1, outlineNumber: '1' }),
      resources: resourceXml({ uid: 5, name: 'R5', calendarUid: -1 }),
    })
    const result = normalizeMspdiForCanonicalImport(withResource)
    const n1 = result.diagnostics.filter((d) => d.code === MPP_NORMALIZED_SENTINEL_REFERENCE)
    expect(n1).toHaveLength(1)
    expect(n1[0].entityId).toBe('r5')
  })
})

describe('PROJECT-018 — N2: sentinel BaseCalendarUID strip', () => {
  const bytes = projectXml({
    name: 'N2',
    calendars: `${STANDARD_CALENDAR_XML}${SENTINEL_BASE_CALENDAR_XML}`,
    tasks: taskXml({ uid: 1, outlineNumber: '1' }),
  })

  it('the un-normalized payload fails in the accepted importer (negative control)', () => {
    const raw = importMspdi(bytes)
    expect(raw.diagnostics.some((d) => d.code === 'MISSING_BASE_CALENDAR')).toBe(true)
  })

  it('normalization strips the sentinel and diagnoses it with the calendar entity', () => {
    const result = normalizeMspdiForCanonicalImport(bytes)
    const n2 = result.diagnostics.filter((d) => d.code === MPP_NORMALIZED_BASE_CALENDAR_SENTINEL)
    expect(n2).toHaveLength(1)
    expect(n2[0].severity).toBe('info')
    expect(n2[0].stage).toBe('normalization')
    expect(n2[0].entityId).toBe('c2')
    expect(new TextDecoder().decode(result.bytes)).not.toContain(
      '<BaseCalendarUID>-1</BaseCalendarUID>',
    )
  })

  it('the normalized payload imports without MISSING_BASE_CALENDAR', () => {
    const result = normalizeMspdiForCanonicalImport(bytes)
    const imported = importMspdi(result.bytes)
    expect(imported.diagnostics.filter((d) => d.code === 'MISSING_BASE_CALENDAR')).toEqual([])
    expect(validateProjectDocument(imported.document).accepted).toBe(true)
  })
})

describe('PROJECT-018 — N3: hidden placeholder records', () => {
  it('the MPP summary-artifact task (UID 0 / OutlineLevel 0 / WBS "0") is filtered', () => {
    const placeholder = `<Task><UID>0</UID><ID>0</ID><Name></Name><WBS>0</WBS><OutlineNumber>0</OutlineNumber><OutlineLevel>0</OutlineLevel><Summary>1</Summary><Milestone>0</Milestone></Task>`
    const bytes = projectXml({
      name: 'N3t',
      tasks: `${placeholder}${taskXml({ uid: 1, outlineNumber: '1' })}`,
    })
    // Negative control: the placeholder trips the accepted importer.
    expect(importMspdi(bytes).diagnostics.some((d) => d.code === 'INVALID_OUTLINE_LEVEL')).toBe(
      true,
    )
    const result = normalizeMspdiForCanonicalImport(bytes)
    const n3 = result.diagnostics.filter((d) => d.code === MPP_NORMALIZED_PLACEHOLDER_RECORD)
    expect(n3).toHaveLength(1)
    expect(n3[0].entityId).toBe('t0')
    expect(n3[0].severity).toBe('info')
    const imported = importMspdi(result.bytes)
    expect(imported.document.tasks).toHaveLength(1)
    expect(imported.diagnostics.filter((d) => d.code === 'INVALID_OUTLINE_LEVEL')).toEqual([])
    expect(validateProjectDocument(imported.document).accepted).toBe(true)
  })

  it('the analogous null-name placeholder resource is filtered', () => {
    const placeholder = `<Resource><UID>0</UID><ID>0</ID><Name></Name><Type>1</Type></Resource>`
    const bytes = projectXml({
      name: 'N3r',
      tasks: taskXml({ uid: 1, outlineNumber: '1' }),
      resources: `${placeholder}${resourceXml({ uid: 1, name: 'Real' })}`,
    })
    const result = normalizeMspdiForCanonicalImport(bytes)
    const n3 = result.diagnostics.filter((d) => d.code === MPP_NORMALIZED_PLACEHOLDER_RECORD)
    expect(n3).toHaveLength(1)
    expect(n3[0].entityId).toBe('r0')
    const imported = importMspdi(result.bytes)
    expect(imported.document.resources).toHaveLength(1)
    expect(imported.document.resources[0]?.name).toBe('Real')
  })
})

describe('PROJECT-018 — N4: until-midnight working period rewrite', () => {
  it('ToTime 00:00:00 is rewritten to the day-end 24:00:00 (canonical endMinute 1440)', () => {
    const bytes = projectXml({
      name: 'N4',
      calendars: MIDNIGHT_CALENDAR_XML,
      tasks: taskXml({ uid: 1, outlineNumber: '1' }),
    })
    // Negative control: the period is rejected as inverted by the importer.
    expect(
      importMspdi(bytes).diagnostics.some(
        (d) => d.code === 'INVALID_MSPDI_CALENDAR' && d.severity === 'error',
      ),
    ).toBe(true)

    const result = normalizeMspdiForCanonicalImport(bytes)
    const n4 = result.diagnostics.filter((d) => d.code === MPP_NORMALIZED_MIDNIGHT_PERIOD)
    expect(n4).toHaveLength(1)
    expect(n4[0].severity).toBe('info')
    expect(n4[0].stage).toBe('normalization')

    const imported = importMspdi(result.bytes)
    expect(imported.diagnostics.filter((d) => d.code === 'INVALID_MSPDI_CALENDAR')).toEqual([])
    // Tuesday (week key 1) carries the 23:00→24:00 period as {1380, 1440}:
    const tuesday = imported.document.calendars[0]?.workingWeek[1]
    expect(tuesday).toContainEqual({ startMinute: 1380, endMinute: 1440 })
    // ...and the ordinary 09:00–17:00 period survived untouched:
    expect(tuesday).toContainEqual({ startMinute: 540, endMinute: 1020 })
    expect(validateProjectDocument(imported.document).accepted).toBe(true)
  })

  it('a FromTime of 00:00:00 (start of day) is NOT rewritten', () => {
    const calendar = MIDNIGHT_CALENDAR_XML.replace(
      '<FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime>\n            <WorkingTime><FromTime>23:00:00</FromTime>',
      '<FromTime>00:00:00</FromTime><ToTime>03:00:00</ToTime></WorkingTime>\n            <WorkingTime><FromTime>23:00:00</FromTime>',
    )
    const bytes = projectXml({
      name: 'N4from',
      calendars: calendar,
      tasks: taskXml({ uid: 1, outlineNumber: '1' }),
    })
    const result = normalizeMspdiForCanonicalImport(bytes)
    // Only the ToTime-00:00:00 period was rewritten — the 00:00→03:00 period
    // is legitimate and untouched (exactly one N4 diagnostic).
    expect(
      result.diagnostics.filter((d) => d.code === MPP_NORMALIZED_MIDNIGHT_PERIOD),
    ).toHaveLength(1)
  })
})

describe('PROJECT-018 — N5: unassigned assignment drop', () => {
  const bytes = projectXml({
    name: 'N5',
    tasks: taskXml({ uid: 1, outlineNumber: '1' }),
    resources: resourceXml({ uid: 9, name: 'Crew' }),
    assignments:
      assignmentXml({ uid: 1, taskUid: 1, resourceUid: -65535 }) +
      assignmentXml({ uid: 2, taskUid: 1, resourceUid: 9 }),
  })

  it('the un-normalized payload fails in the accepted importer (negative control)', () => {
    expect(
      importMspdi(bytes).diagnostics.some(
        (d) => d.code === 'INVALID_MSPDI_REFERENCE' && d.severity === 'error',
      ),
    ).toBe(true)
  })

  it('normalization drops the placeholder assignment with an expected-loss WARNING', () => {
    const result = normalizeMspdiForCanonicalImport(bytes)
    const n5 = result.diagnostics.filter((d) => d.code === MPP_DROPPED_UNASSIGNED_ASSIGNMENT)
    expect(n5).toHaveLength(1)
    expect(n5[0].severity).toBe('warning')
    expect(n5[0].stage).toBe('normalization')
    expect(n5[0].entityId).toBe('t1')
    const imported = importMspdi(result.bytes)
    expect(imported.document.assignments).toHaveLength(1)
    expect(imported.document.assignments[0]?.resourceId).toBe('r9')
    expect(imported.diagnostics.filter((d) => d.code === 'INVALID_MSPDI_REFERENCE')).toEqual([])
  })
})

describe('PROJECT-018 — normalization invariants', () => {
  it('a payload needing no normalization is returned byte-identical with zero diagnostics', () => {
    const bytes = projectXml({ name: 'Clean', tasks: taskXml({ uid: 1, outlineNumber: '1' }) })
    const result = normalizeMspdiForCanonicalImport(bytes)
    expect(result.diagnostics).toEqual([])
    expect(result.bytes).toEqual(bytes)
  })

  it('malformed XML passes through unchanged (the importer reports it, stage mspdi)', () => {
    const garbage = new TextEncoder().encode('this is not xml at all <broken')
    const result = normalizeMspdiForCanonicalImport(garbage)
    expect(result.diagnostics).toEqual([])
    expect(result.bytes).toEqual(garbage)
    const imported = importMspdi(result.bytes)
    expect(imported.diagnostics.some((d) => d.code === 'INVALID_MSPDI')).toBe(true)
  })

  it('normalization is deterministic: identical input → identical output bytes', () => {
    const bytes = projectXml({
      name: 'Det',
      tasks: taskXml({ uid: 1, outlineNumber: '1', calendarUid: -1 }),
    })
    const a = normalizeMspdiForCanonicalImport(bytes)
    const b = normalizeMspdiForCanonicalImport(bytes)
    expect(a.bytes).toEqual(b.bytes)
    expect(a.diagnostics).toEqual(b.diagnostics)
  })

  it('serializeXmlNode round-trips through the accepted parser with escaping intact', () => {
    // NOTE: the fixture builders interpolate raw text — special characters
    // must be pre-escaped exactly as a real MSPDI writer would emit them.
    const escapedName = 'A &amp; B &lt;tag&gt; &quot;quoted&quot; &#39;single&#39;'
    const bytes = projectXml({
      name: 'Escape &amp; Round-Trip',
      tasks: taskXml({ uid: 1, outlineNumber: '1', name: escapedName }),
    })
    const parsed = parseXml(bytes)
    const text = `<?xml version="1.0" encoding="UTF-8"?>\n${serializeXmlNode(parsed)}`
    const reparsed = parseXml(new TextEncoder().encode(text))
    expect(serializeXmlNode(reparsed)).toBe(serializeXmlNode(parsed))
    // And the semantic content survived:
    const imported = importMspdi(new TextEncoder().encode(text))
    expect(imported.document.tasks[0]?.name).toBe('A & B <tag> "quoted" \'single\'')
  })
})

describe('PROJECT-018 — importMppFromMspdi staged contract', () => {
  it('a successful outcome stages normalization + mspdi + (accepted) canonical diagnostics', () => {
    const bytes = projectXml({
      name: 'Staged',
      tasks: taskXml({ uid: 1, outlineNumber: '1', calendarUid: -1 }),
    })
    const result = importMppFromMspdi(outcome(bytes))
    const stages = new Set(result.diagnostics.map((d) => d.stage))
    expect(stages.has('normalization')).toBe(true)
    expect(stages.has('mspdi')).toBe(true)
    expect(
      result.diagnostics.every(
        (d) => d.stage === 'normalization' || d.stage === 'mspdi' || d.stage === 'canonical',
      ),
    ).toBe(true)
    expect(result.diagnostics.every((d) => Object.hasOwn(d, 'severity'))).toBe(true)
    expect(validateProjectDocument(result.document).accepted).toBe(true)
  })

  it('sidecar-stage error diagnostics short-circuit to the empty document', () => {
    const result = importMppFromMspdi(
      outcome(projectXml({ tasks: taskXml({ uid: 1, outlineNumber: '1' }) }), [
        {
          code: MPP_SIDECAR_UNAVAILABLE,
          severity: 'error',
          message: 'java not found',
          stage: 'sidecar',
        },
      ]),
    )
    expect(result.document).toEqual(emptyProjectDocument())
    expect(result.diagnostics).toHaveLength(1)
    expect(result.diagnostics[0].stage).toBe('sidecar')
  })

  it('malformed sidecar MSPDI is an mspdi-stage INVALID_MSPDI error with the empty document', () => {
    const garbage = new TextEncoder().encode('<Project><Oops')
    const result = importMppFromMspdi(outcome(garbage))
    expect(result.document).toEqual(emptyProjectDocument())
    const invalid = result.diagnostics.filter((d) => d.code === 'INVALID_MSPDI')
    expect(invalid).toHaveLength(1)
    expect(invalid[0].stage).toBe('mspdi')
    expect(invalid[0].severity).toBe('error')
  })

  it('a canonical-validation rejection is atomic: empty document + canonical-stage error', () => {
    // Cycle: valid XML, imports, but the engine rejects the document.
    const tasks =
      taskXml({ uid: 1, outlineNumber: '1', predecessorLinks: [{ predUid: 2, type: 0 }] }) +
      taskXml({ uid: 2, outlineNumber: '2', predecessorLinks: [{ predUid: 1, type: 0 }] })
    const result = importMppFromMspdi(outcome(projectXml({ name: 'Cycle', tasks })))
    expect(result.document).toEqual(emptyProjectDocument())
    const canonical = result.diagnostics.filter((d) => d.stage === 'canonical')
    expect(canonical.length).toBeGreaterThan(0)
    expect(canonical.every((d) => d.severity === 'error')).toBe(true)
  })

  it('recoverable importer errors keep the accepted PROJECT-015 semantics (document returned)', () => {
    // A dangling NON-sentinel calendar reference is a recoverable importer
    // error (recovery to the default calendar) — the MPP layer must not
    // flatten or hide it, and must not treat it as atomic failure.
    const bytes = projectXml({
      name: 'Recover',
      tasks: taskXml({ uid: 1, outlineNumber: '1', calendarUid: 999 }),
    })
    const result = importMppFromMspdi(outcome(bytes))
    expect(result.document.tasks).toHaveLength(1)
    const recoverable = result.diagnostics.filter(
      (d) => d.code === 'INVALID_MSPDI_REFERENCE' && d.stage === 'mspdi',
    )
    expect(recoverable.length).toBeGreaterThan(0)
  })
})

describe('PROJECT-018 — foundation contract constants', () => {
  it('the diagnostic-code table is the documented union', () => {
    expect(MPP_DIAGNOSTIC_CODES).toHaveLength(12)
    expect([...MPP_DIAGNOSTIC_CODES]).toEqual(
      expect.arrayContaining([
        MPP_NORMALIZED_SENTINEL_REFERENCE,
        MPP_NORMALIZED_BASE_CALENDAR_SENTINEL,
        MPP_NORMALIZED_PLACEHOLDER_RECORD,
        MPP_NORMALIZED_MIDNIGHT_PERIOD,
        MPP_DROPPED_UNASSIGNED_ASSIGNMENT,
      ]),
    )
  })

  it('version support and limits are the documented values', () => {
    expect([...MPP_SUPPORTED_FORMAT_VERSIONS]).toEqual(['MPP8', 'MPP9', 'MPP12', 'MPP14'])
    expect(MPXJ_PINNED_VERSION).toBe('16.7.0')
    expect(MPP_MAX_INPUT_BYTES).toBe(100 * 1024 * 1024)
    expect(MPP_MAX_MSPDI_OUTPUT_BYTES).toBe(100 * 1024 * 1024)
  })

  it('every MPP diagnostic code round-trips through codesOf helper (smoke)', () => {
    // Guard against accidental duplicate codes in the table:
    expect(new Set(MPP_DIAGNOSTIC_CODES).size).toBe(MPP_DIAGNOSTIC_CODES.length)
    expect(codesOf([{ code: 'A' }, { code: 'A' }, { code: 'B' }])).toEqual(['A', 'B'])
  })
})
