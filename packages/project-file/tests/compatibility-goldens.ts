/**
 * PROJECT-020 — compatibility golden fixtures D01–D19.
 *
 * Stable byte payloads (and synthetic sidecar outcomes) exercising every
 * required compatibility surface across the three import paths. Each
 * golden is deterministic so report determinism is provable; the tests
 * assert the COMPLETE diagnostic objects (code, severity, stage, format,
 * loss, recoverability, entityType, entityId) — never just codes.
 *
 * D01 clean .gproj                     D11 MPP N3 (placeholder records)
 * D02 malformed .gproj                 D12 MPP N4 (midnight period)
 * D03 unsupported .gproj version       D13 MPP N5 (unassigned assignment)
 * D04 clean MSPDI                      D14 MPP unsupported version
 * D05 MSPDI unsupported feature        D15 sidecar failure
 * D06 MSPDI invalid references         D16 network-isolation failure
 * D07 MSPDI calendar degradation       D17 canonical validation failure
 * D08 clean MPP                        D18 scheduling failure (injected runner)
 * D09 MPP N1 (sentinel CalendarUID)    D19 composite compatibility report
 * D10 MPP N2 (sentinel BaseCalendarUID)
 */
import { encodeUtf8 } from '../src/utf8.js'
import { serializeGproj } from '../src/deserialize.js'
import type { MppConversionOutcome, MppDiagnostic, MppSidecarCounts } from '../src/mpp/types.js'
import { MPP_SIDECAR_PROTOCOL_VERSION } from '../src/mpp/types.js'
import {
  MPP_INPUT_TOO_LARGE,
  MPP_INPUT_UNREADABLE,
  MPP_SIDECAR_EXIT,
  MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE,
  MPP_SIDECAR_RESPONSE_INVALID,
  MPP_SIDECAR_TIMEOUT,
  MPP_SIDECAR_UNAVAILABLE,
  MPP_UNSUPPORTED_FORMAT,
} from '../src/mpp/diagnostics.js'
import { g02Wbs } from './fixtures.js'

// ---- shared XML builders --------------------------------------------------

/** A standard Mon–Fri 09:00–17:00 calendar (UID 1, base + default). */
export const STANDARD_CALENDAR = `
  <Calendar>
    <UID>1</UID>
    <Name>Standard</Name>
    <IsBaseCalendar>true</IsBaseCalendar>
    <IsBaseCalendarDefault>true</IsBaseCalendarDefault>
    <WeekDays>
      <WeekDay><DayType>1</DayType><DayWorking>false</DayWorking></WeekDay>
      <WeekDay><DayType>2</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
      <WeekDay><DayType>3</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
      <WeekDay><DayType>4</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
      <WeekDay><DayType>5</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
      <WeekDay><DayType>6</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
      <WeekDay><DayType>7</DayType><DayWorking>false</DayWorking></WeekDay>
    </WeekDays>
    <Exceptions />
  </Calendar>`

/** Wrap task/resource/assignment/calendar XML into a full MSPDI document. */
export function mspdiDocument(opts: {
  tasks?: string
  resources?: string
  assignments?: string
  calendars?: string
  saveVersion?: number
  lastSaved?: string
  minutesPerDay?: number
}): Uint8Array {
  const calendars = opts.calendars ?? STANDARD_CALENDAR
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <SaveVersion>${opts.saveVersion ?? 16}</SaveVersion>
  <Name>Compatibility Golden</Name>
  <ScheduleFromStart>true</ScheduleFromStart>
  <StartDate>2026-08-03T09:00:00</StartDate>
  ${opts.minutesPerDay !== undefined ? `<MinutesPerDay>${opts.minutesPerDay}</MinutesPerDay>` : ''}
  ${opts.lastSaved !== undefined ? `<LastSaved>${opts.lastSaved}</LastSaved>` : ''}
  <Calendars>${calendars}</Calendars>
  <Tasks>${opts.tasks ?? ''}</Tasks>
  <Resources>${opts.resources ?? ''}</Resources>
  <Assignments>${opts.assignments ?? ''}</Assignments>
</Project>`
  return encodeUtf8(xml)
}

/** One minimal `<Task>` element (single line, explicit fields). */
export function taskXml(uid: number, extra = ''): string {
  return `<Task><UID>${uid}</UID><ID>${uid}</ID><Name>Task ${uid}</Name><OutlineLevel>1</OutlineLevel><Summary>false</Summary><Milestone>false</Milestone><Manual>false</Manual><Duration>PT8H0M0S</Duration><Priority>500</Priority><PercentComplete>0</PercentComplete><Work>PT0H0M0S</Work><RemainingWork>PT8H0M0S</RemainingWork><ActualWork>PT0H0M0S</ActualWork><Cost>0</Cost><ActualCost>0</ActualCost><RemainingCost>0</RemainingCost><OutlineNumber>${uid}</OutlineNumber>${extra}</Task>`
}

// ---- D01–D03: .gproj -------------------------------------------------------

/** D01 — a clean `.gproj` (the accepted g02 WBS document: 3 tasks). */
export function d01CleanGproj(): Uint8Array {
  return serializeGproj(g02Wbs())
}

/** D02 — a malformed `.gproj` envelope (not JSON at all). */
export function d02MalformedGproj(): Uint8Array {
  return encodeUtf8('{"format":"gproj","formatVersion":1,"document": <not json')
}

/** D03 — a `.gproj` envelope with an unsupported format version. */
export function d03UnsupportedGprojVersion(): Uint8Array {
  return encodeUtf8(
    JSON.stringify({
      format: 'gproj',
      formatVersion: 99,
      document: {
        schemaVersion: 1,
        properties: {
          id: 'p',
          name: 'p',
          startDate: '2026-08-03T09:00:00.000Z',
          defaultCalendarId: 'standard',
        },
      },
    }),
  )
}

// ---- D04–D07, D17: MSPDI --------------------------------------------------

/** D04 — clean MSPDI (one task, standard calendar). */
export function d04CleanMspdi(): Uint8Array {
  return mspdiDocument({ tasks: taskXml(1) })
}

/** D05 — MSPDI unsupported feature: a percentage lag (format 35). */
export function d05UnsupportedFeatureMspdi(): Uint8Array {
  const t1 = taskXml(1)
  const t2 = taskXml(
    2,
    '<PredecessorLink><PredecessorUID>1</PredecessorUID><Type>0</Type><LinkLag>50</LinkLag><LinkLagFormat>35</LinkLagFormat></PredecessorLink>',
  )
  return mspdiDocument({ tasks: t1 + t2 })
}

/** D06 — MSPDI invalid reference: a PredecessorUID that resolves nowhere. */
export function d06InvalidReferenceMspdi(): Uint8Array {
  const t1 = taskXml(1)
  const t2 = taskXml(
    2,
    '<PredecessorLink><PredecessorUID>99</PredecessorUID><Type>0</Type></PredecessorLink>',
  )
  return mspdiDocument({ tasks: t1 + t2 })
}

/** D07 — MSPDI calendar degradation: a recurring calendar exception. */
export function d07CalendarDegradationMspdi(): Uint8Array {
  const calendar = STANDARD_CALENDAR.replace(
    '<Exceptions />',
    '<Exceptions><Exception><Start>2026-08-10T00:00:00</Start><Finish>2026-08-12T00:00:00</Finish><Type>3</Type><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>13:00:00</ToTime></WorkingTime></WorkingTimes></Exception></Exceptions>',
  )
  return mspdiDocument({ tasks: taskXml(1), calendars: calendar })
}

/** D17 — canonical validation failure: a dependency cycle (t1 ⇄ t2). */
export function d17CanonicalFailureMspdi(): Uint8Array {
  const t1 = taskXml(
    1,
    '<PredecessorLink><PredecessorUID>2</PredecessorUID><Type>0</Type></PredecessorLink>',
  )
  const t2 = taskXml(
    2,
    '<PredecessorLink><PredecessorUID>1</PredecessorUID><Type>0</Type></PredecessorLink>',
  )
  return mspdiDocument({ tasks: t1 + t2 })
}

/** A malformed MSPDI date (entity-level: the task start is unusable). */
export function malformedDateMspdi(): Uint8Array {
  return mspdiDocument({ tasks: taskXml(1, '<Start>not-a-date</Start>') })
}

/** A malformed MSPDI duration (negative — not representable). */
export function malformedDurationMspdi(): Uint8Array {
  return mspdiDocument({
    tasks: `<Task><UID>1</UID><ID>1</ID><Name>Task 1</Name><OutlineLevel>1</OutlineLevel><Summary>false</Summary><Milestone>false</Milestone><Manual>false</Manual><Duration>-PT8H0M0S</Duration><Priority>500</Priority><PercentComplete>0</PercentComplete><Work>PT0H0M0S</Work><RemainingWork>PT8H0M0S</RemainingWork><ActualWork>PT0H0M0S</ActualWork><Cost>0</Cost><ActualCost>0</ActualCost><RemainingCost>0</RemainingCost><OutlineNumber>1</OutlineNumber></Task>`,
  })
}

/** MSPDI with a baseline slot (capturedAt approximation) and a non-zero
 * PhysicalPercentComplete (drop) — both PROJECT-020 provenance surfaces. */
export function baselineAndPhysicalMspdi(): Uint8Array {
  return mspdiDocument({
    lastSaved: '2026-08-02T08:00:00',
    tasks: taskXml(
      1,
      '<PhysicalPercentComplete>25</PhysicalPercentComplete><Baseline><Start>2026-08-03T09:00:00</Start><Finish>2026-08-03T17:00:00</Finish><Duration>PT8H0M0S</Duration></Baseline>',
    ),
  })
}

/** A malformed `MinutesPerDay` used by three day-format lag dependencies —
 * the producer-contract de-duplication fixture (one declaration diagnostic). */
export function malformedFactorMspdi(): Uint8Array {
  const link = (pred: number): string =>
    `<PredecessorLink><PredecessorUID>${pred}</PredecessorUID><Type>0</Type><LinkLag>10</LinkLag><LinkLagFormat>5</LinkLagFormat></PredecessorLink>`
  const t1 = taskXml(1)
  const t2 = taskXml(2, link(1))
  const t3 = taskXml(3, link(1))
  const t4 = taskXml(4, link(1))
  return mspdiDocument({ tasks: t1 + t2 + t3 + t4, minutesPerDay: 0 })
}

// ---- D08–D16, D19: MPP (synthetic sidecar outcomes) -----------------------

const COUNTS: MppSidecarCounts = {
  tasks: 1,
  resources: 0,
  calendars: 1,
  predecessorLinks: 0,
  assignments: 0,
}

/** Build a successful sidecar conversion outcome over MSPDI bytes. */
export function outcome(mspdiBytes: Uint8Array, format = 'MPP14'): MppConversionOutcome {
  return {
    mspdiBytes,
    frame: {
      version: MPP_SIDECAR_PROTOCOL_VERSION,
      requestId: 'golden-request',
      ok: true,
      format,
      counts: COUNTS,
    },
    sidecarDiagnostics: [],
  }
}

/** Build a failed-conversion outcome carrying one sidecar error. */
export function failedOutcome(code: string, message: string): MppConversionOutcome {
  return {
    mspdiBytes: new Uint8Array(0),
    frame: {
      version: MPP_SIDECAR_PROTOCOL_VERSION,
      requestId: 'golden-request',
      ok: false,
      error: { code, message },
    },
    sidecarDiagnostics: [
      { code, severity: 'error', message, stage: 'sidecar' } satisfies MppDiagnostic,
    ],
  }
}

/** D08 — clean MPP import (clean MSPDI through the outcome boundary). */
export function d08CleanMppOutcome(): MppConversionOutcome {
  return outcome(d04CleanMspdi(), 'MPP14')
}

/** D09 — MPP N1: sentinel CalendarUID -1 on a task. */
export function d09N1Outcome(): MppConversionOutcome {
  return outcome(mspdiDocument({ tasks: taskXml(1, '<CalendarUID>-1</CalendarUID>') }))
}

/** D10 — MPP N2: sentinel BaseCalendarUID -1 on a calendar. */
export function d10N2Outcome(): MppConversionOutcome {
  const calendar = STANDARD_CALENDAR.replace(
    '<IsBaseCalendarDefault>true</IsBaseCalendarDefault>',
    '<IsBaseCalendarDefault>true</IsBaseCalendarDefault><BaseCalendarUID>-1</BaseCalendarUID>',
  )
  return outcome(mspdiDocument({ tasks: taskXml(1), calendars: calendar }))
}

/** D11 — MPP N3: hidden placeholder task (UID 0) + null-name placeholder resource. */
export function d11N3Outcome(): MppConversionOutcome {
  const placeholderTask = `<Task><UID>0</UID><ID>0</ID><Name>Project Summary</Name><OutlineLevel>0</OutlineLevel><WBS>0</WBS><Summary>true</Summary><Milestone>false</Milestone><Manual>false</Manual><Duration>PT0H0M0S</Duration></Task>`
  const placeholderResource = `<Resource><UID>0</UID><ID>0</ID><Name></Name><Type>1</Type></Resource>`
  return outcome(
    mspdiDocument({
      tasks: placeholderTask + taskXml(1),
      resources: placeholderResource,
    }),
  )
}

/** D12 — MPP N4: a working period running "until midnight". */
export function d12N4Outcome(): MppConversionOutcome {
  const calendar = STANDARD_CALENDAR.replaceAll(
    '<ToTime>17:00:00</ToTime>',
    '<ToTime>00:00:00</ToTime>',
  ).replaceAll('<FromTime>09:00:00</FromTime>', '<FromTime>22:00:00</FromTime>')
  return outcome(mspdiDocument({ tasks: taskXml(1), calendars: calendar }))
}

/** D13 — MPP N5: an "unassigned" placeholder assignment. */
export function d13N5Outcome(): MppConversionOutcome {
  const assignment =
    '<Assignment><UID>1</UID><TaskUID>1</TaskUID><ResourceUID>-65535</ResourceUID><Units>1</Units></Assignment>'
  return outcome(mspdiDocument({ tasks: taskXml(1), assignments: assignment }))
}

/** D14 — unsupported MPP version/format (sidecar refusal). */
export function d14UnsupportedVersionOutcome(): MppConversionOutcome {
  return failedOutcome(
    MPP_UNSUPPORTED_FORMAT,
    'MPP input is not a recognized project format (no MPP8/MPP9/MPP12/MPP14 container)',
  )
}

/** D15 — sidecar failure (nonzero exit). */
export function d15SidecarFailureOutcome(): MppConversionOutcome {
  return failedOutcome(MPP_SIDECAR_EXIT, 'sidecar exited with code 1: conversion failed')
}

/** D16 — network isolation unavailable (fail-closed refusal). */
export function d16NetworkIsolationOutcome(): MppConversionOutcome {
  return failedOutcome(
    MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE,
    'network isolation required but the host cannot provide the mechanism; conversion refused (fail closed)',
  )
}

/** The remaining sidecar failure classes (one outcome each, for coverage). */
export function sidecarUnavailableOutcome(): MppConversionOutcome {
  return failedOutcome(MPP_SIDECAR_UNAVAILABLE, 'java executable or MPXJ distribution missing')
}

export function sidecarTimeoutOutcome(): MppConversionOutcome {
  return failedOutcome(MPP_SIDECAR_TIMEOUT, 'conversion exceeded the wall-clock timeout')
}

export function sidecarResponseInvalidOutcome(): MppConversionOutcome {
  return failedOutcome(MPP_SIDECAR_RESPONSE_INVALID, 'sidecar stdout frame is not valid JSON')
}

export function inputUnreadableOutcome(): MppConversionOutcome {
  return failedOutcome(MPP_INPUT_UNREADABLE, 'MPP input file cannot be read: ENOENT')
}

export function inputTooLargeOutcome(): MppConversionOutcome {
  return failedOutcome(MPP_INPUT_TOO_LARGE, 'MPP input is 104857600 bytes (limit 104857600)')
}

/** Malformed MSPDI output from the sidecar (the importer's error, not the
 * normalizer's). */
export function malformedMspdiOutputOutcome(): MppConversionOutcome {
  return outcome(encodeUtf8('this is not xml at all'))
}

/** D19 — the composite: N1 + N3 + N4 + N5 + baseline approximation +
 * PhysicalPercentComplete drop in one MPP conversion. */
export function d19CompositeOutcome(): MppConversionOutcome {
  const placeholderTask = `<Task><UID>0</UID><ID>0</ID><Name>Project Summary</Name><OutlineLevel>0</OutlineLevel><WBS>0</WBS><Summary>true</Summary><Milestone>false</Milestone><Manual>false</Manual><Duration>PT0H0M0S</Duration></Task>`
  const task = taskXml(
    1,
    '<CalendarUID>-1</CalendarUID><PhysicalPercentComplete>40</PhysicalPercentComplete><Baseline><Start>2026-08-03T09:00:00</Start><Finish>2026-08-03T17:00:00</Finish><Duration>PT8H0M0S</Duration></Baseline>',
  )
  const calendar = STANDARD_CALENDAR.replaceAll(
    '<ToTime>17:00:00</ToTime>',
    '<ToTime>00:00:00</ToTime>',
  ).replaceAll('<FromTime>09:00:00</FromTime>', '<FromTime>22:00:00</FromTime>')
  const assignment =
    '<Assignment><UID>1</UID><TaskUID>1</TaskUID><ResourceUID>-65535</ResourceUID><Units>1</Units></Assignment>'
  return outcome(
    mspdiDocument({
      tasks: placeholderTask + task,
      assignments: assignment,
      calendars: calendar,
      lastSaved: '2026-08-02T08:00:00',
    }),
    'MPP9',
  )
}
