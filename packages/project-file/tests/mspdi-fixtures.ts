/**
 * PROJECT-015 — MSPDI golden fixtures (M01–M17).
 *
 * Hand-authored MSPDI XML payloads exercising every required import surface.
 * M01–M13 + M16 are valid (import → zero error diagnostics, pass
 * `validateProjectDocument` + `schedule`). M14 surfaces an unsupported feature
 * (warning, document still valid). M15 / M17 are invalid (error diagnostics,
 * partial/empty recovery). Every fixture is a stable byte payload so
 * `importMspdi` determinism is provable.
 *
 * MSPDI conventions used by these fixtures (documented in
 * `spec/project/requirements.md` PROJECT-015):
 *   - `<SaveVersion>16</SaveVersion>` (Project 2016; in the supported set).
 *   - `<Task><UID>` is the persistent interop id; canonical `task.uid` carries
 *     it verbatim and canonical `task.id = 't'+uid` (identity.ts).
 *   - `<OutlineNumber>` reconstructs `parentTaskId` (canonical identity is NOT
 *     WBS — WBS only rebuilds the hierarchy).
 *   - `<Duration>` / `<Work>` are ISO-8601 (`PT8H0M0S` = 8 working hours =
 *     480 working minutes).
 *   - `<LinkLag>` is in tenths of the unit declared by `<LinkLagFormat>`
 *     (`2400` fmt 1 = 240 minutes; `30` fmt 3 = 180 minutes;
 *     `25` fmt 5 = 2.5 days = 1200 minutes with `MinutesPerDay` 480).
 *   - `<ConstraintType>` 0–7 maps to the eight canonical constraint types.
 *   - `<Type>` on `<Resource>`: 1=work, 2=material, 3=cost.
 *   - `<DayType>` 1=Sunday..7=Saturday (canonical week key = DayType-1).
 */
import { encodeUtf8 } from '../src/utf8.js'

/** Standard Mon–Fri 09:00–17:00 calendar, UID 1, base + default. */
export const STANDARD_CALENDAR_XML = `
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

/** Build a `<Task>` XML element. Fields not provided default to sane values. */
export interface TaskSpec {
  uid: number
  id?: number
  name?: string
  outlineNumber?: string
  outlineLevel?: number
  summary?: boolean
  milestone?: boolean
  manual?: boolean
  type?: number // 0 fixedUnits, 1 fixedDuration, 2 fixedWork
  duration?: string // ISO-8601; default PT8H0M0S
  work?: string
  actualWork?: string
  remainingWork?: string
  cost?: number
  actualCost?: number
  remainingCost?: number
  percentComplete?: number
  priority?: number
  start?: string
  finish?: string
  constraintType?: number
  constraintDate?: string
  deadline?: string
  calendarUid?: number
  notes?: string
  extendedAttrs?: Array<{ fieldId: string; value: string }>
  predecessorLinks?: Array<{ predUid: number; type: number; lag?: number; lagFormat?: number }>
  baselines?: Array<{
    slot: number
    start?: string
    finish?: string
    duration?: string
    work?: string
    cost?: number
  }>
}

function taskXml(s: TaskSpec): string {
  const duration = s.duration ?? 'PT8H0M0S'
  const work = s.work ?? 'PT0H0M0S'
  const actualWork = s.actualWork ?? 'PT0H0M0S'
  const remainingWork = s.remainingWork ?? work
  const lines: string[] = ['      <Task>']
  lines.push(`        <UID>${s.uid}</UID>`)
  lines.push(`        <ID>${s.id ?? s.uid}</ID>`)
  lines.push(`        <Name>${s.name ?? `Task ${s.uid}`}</Name>`)
  if (s.type !== undefined) lines.push(`        <Type>${s.type}</Type>`)
  if (s.outlineNumber !== undefined)
    lines.push(`        <OutlineNumber>${s.outlineNumber}</OutlineNumber>`)
  lines.push(`        <OutlineLevel>${s.outlineLevel ?? 1}</OutlineLevel>`)
  lines.push(`        <Summary>${s.summary === true ? 'true' : 'false'}</Summary>`)
  lines.push(`        <Milestone>${s.milestone === true ? 'true' : 'false'}</Milestone>`)
  lines.push(`        <Manual>${s.manual === true ? 'true' : 'false'}</Manual>`)
  lines.push(`        <Duration>${duration}</Duration>`)
  lines.push(`        <Priority>${s.priority ?? 500}</Priority>`)
  lines.push(`        <PercentComplete>${s.percentComplete ?? 0}</PercentComplete>`)
  lines.push(`        <Work>${work}</Work>`)
  lines.push(`        <RemainingWork>${remainingWork}</RemainingWork>`)
  lines.push(`        <ActualWork>${actualWork}</ActualWork>`)
  lines.push(`        <Cost>${s.cost ?? 0}</Cost>`)
  lines.push(`        <ActualCost>${s.actualCost ?? 0}</ActualCost>`)
  lines.push(`        <RemainingCost>${s.remainingCost ?? 0}</RemainingCost>`)
  if (s.start !== undefined) lines.push(`        <Start>${s.start}</Start>`)
  if (s.finish !== undefined) lines.push(`        <Finish>${s.finish}</Finish>`)
  if (s.constraintType !== undefined)
    lines.push(`        <ConstraintType>${s.constraintType}</ConstraintType>`)
  if (s.constraintDate !== undefined)
    lines.push(`        <ConstraintDate>${s.constraintDate}</ConstraintDate>`)
  if (s.deadline !== undefined) lines.push(`        <Deadline>${s.deadline}</Deadline>`)
  if (s.calendarUid !== undefined) lines.push(`        <CalendarUID>${s.calendarUid}</CalendarUID>`)
  if (s.notes !== undefined) lines.push(`        <Notes>${s.notes}</Notes>`)
  if (s.extendedAttrs !== undefined) {
    for (const ea of s.extendedAttrs) {
      lines.push(
        `        <ExtendedAttribute><FieldID>${ea.fieldId}</FieldID><Value>${ea.value}</Value></ExtendedAttribute>`,
      )
    }
  }
  if (s.predecessorLinks !== undefined) {
    for (const pl of s.predecessorLinks) {
      lines.push(`        <PredecessorLink>`)
      lines.push(`          <PredecessorUID>${pl.predUid}</PredecessorUID>`)
      lines.push(`          <Type>${pl.type}</Type>`)
      lines.push(`          <LinkLag>${pl.lag ?? 0}</LinkLag>`)
      lines.push(`          <LinkLagFormat>${pl.lagFormat ?? 1}</LinkLagFormat>`)
      lines.push(`        </PredecessorLink>`)
    }
  }
  if (s.baselines !== undefined) {
    for (const b of s.baselines) {
      const slotName = b.slot === 0 ? 'Baseline' : `Baseline${b.slot}`
      lines.push(`        <${slotName}>`)
      if (b.start !== undefined) lines.push(`          <Start>${b.start}</Start>`)
      if (b.finish !== undefined) lines.push(`          <Finish>${b.finish}</Finish>`)
      lines.push(`          <Duration>${b.duration ?? 'PT8H0M0S'}</Duration>`)
      lines.push(`          <Work>${b.work ?? 'PT0H0M0S'}</Work>`)
      lines.push(`          <Cost>${b.cost ?? 0}</Cost>`)
      lines.push(`        </${slotName}>`)
    }
  }
  lines.push('      </Task>')
  return lines.join('\n')
}

export interface ResourceSpec {
  uid: number
  name?: string
  type?: number // 1 work, 2 material, 3 cost
  maxUnits?: number
  standardRate?: number
  overtimeRate?: number
  costPerUse?: number
  calendarUid?: number
}

function resourceXml(s: ResourceSpec): string {
  const kind = s.type ?? 1
  const maxUnits = s.maxUnits ?? (kind === 1 ? 1 : 0)
  return [
    '      <Resource>',
    `        <UID>${s.uid}</UID>`,
    `        <Name>${s.name ?? `Resource ${s.uid}`}</Name>`,
    `        <Type>${kind}</Type>`,
    `        <MaxUnits>${maxUnits}</MaxUnits>`,
    `        <StandardRate>${s.standardRate ?? 0}</StandardRate>`,
    `        <OvertimeRate>${s.overtimeRate ?? 0}</OvertimeRate>`,
    `        <CostPerUse>${s.costPerUse ?? 0}</CostPerUse>`,
    s.calendarUid !== undefined ? `        <CalendarUID>${s.calendarUid}</CalendarUID>` : '',
    '      </Resource>',
  ]
    .filter((l) => l.length > 0)
    .join('\n')
}

export interface AssignmentSpec {
  uid: number
  taskUid: number
  resourceUid: number
  units?: number
  work?: string
  actualWork?: string
  remainingWork?: string
  cost?: number
  actualCost?: number
  remainingCost?: number
}

function assignmentXml(s: AssignmentSpec): string {
  return [
    '      <Assignment>',
    `        <UID>${s.uid}</UID>`,
    `        <TaskUID>${s.taskUid}</TaskUID>`,
    `        <ResourceUID>${s.resourceUid}</ResourceUID>`,
    `        <Units>${s.units ?? 1}</Units>`,
    `        <Work>${s.work ?? 'PT0H0M0S'}</Work>`,
    `        <ActualWork>${s.actualWork ?? 'PT0H0M0S'}</ActualWork>`,
    `        <RemainingWork>${s.remainingWork ?? s.work ?? 'PT0H0M0S'}</RemainingWork>`,
    `        <Cost>${s.cost ?? 0}</Cost>`,
    `        <ActualCost>${s.actualCost ?? 0}</ActualCost>`,
    `        <RemainingCost>${s.remainingCost ?? 0}</RemainingCost>`,
    '      </Assignment>',
  ].join('\n')
}

/** Build a full MSPDI `<Project>` document. */
export function projectXml(opts: {
  name?: string
  startDate?: string
  lastSaved?: string
  calendars?: string
  tasks?: string
  resources?: string
  assignments?: string
  extendedAttributes?: string
  saveVersion?: number
  /** Project-level lag conversion settings (root children). */
  minutesPerDay?: number
  minutesPerWeek?: number
  daysPerMonth?: number
}): Uint8Array {
  const sv = opts.saveVersion ?? 16
  const factors =
    (opts.minutesPerDay !== undefined
      ? `  <MinutesPerDay>${opts.minutesPerDay}</MinutesPerDay>\n`
      : '') +
    (opts.minutesPerWeek !== undefined
      ? `  <MinutesPerWeek>${opts.minutesPerWeek}</MinutesPerWeek>\n`
      : '') +
    (opts.daysPerMonth !== undefined ? `  <DaysPerMonth>${opts.daysPerMonth}</DaysPerMonth>\n` : '')
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <SaveVersion>${sv}</SaveVersion>
  <Name>${opts.name ?? 'Test Project'}</Name>
  <ScheduleFromStart>true</ScheduleFromStart>
  <StartDate>${opts.startDate ?? '2026-08-03T09:00:00'}</StartDate>
${factors}  ${opts.lastSaved !== undefined ? `<LastSaved>${opts.lastSaved}</LastSaved>` : ''}
  ${opts.extendedAttributes !== undefined ? `<ExtendedAttributes>${opts.extendedAttributes}</ExtendedAttributes>` : ''}
  <Calendars>${opts.calendars ?? STANDARD_CALENDAR_XML}</Calendars>
  <Tasks>${opts.tasks ?? ''}</Tasks>
  <Resources>${opts.resources ?? ''}</Resources>
  <Assignments>${opts.assignments ?? ''}</Assignments>
</Project>
`
  return encodeUtf8(xml)
}

export { taskXml, resourceXml, assignmentXml }

// ---- the golden fixtures (M01–M18) --------------------------------------

/** M01 — minimal MSPDI project: properties + standard calendar, no tasks. */
export function m01Minimal(): Uint8Array {
  return projectXml({ name: 'M01 Minimal' })
}

/** M02 — WBS hierarchy: summary + two children. */
export function m02Wbs(): Uint8Array {
  const tasks = [
    taskXml({
      uid: 1,
      name: 'Phase A',
      outlineNumber: '1',
      outlineLevel: 1,
      summary: true,
      duration: 'PT16H0M0S',
      work: 'PT16H0M0S',
    }),
    taskXml({
      uid: 2,
      name: 'Design',
      outlineNumber: '1.1',
      outlineLevel: 2,
      duration: 'PT8H0M0S',
    }),
    taskXml({ uid: 3, name: 'Build', outlineNumber: '1.2', outlineLevel: 2, duration: 'PT8H0M0S' }),
  ].join('\n')
  return projectXml({ name: 'M02 WBS', tasks })
}

/** M03 — dependency graph: 3 tasks + FS + SS. */
export function m03Dependencies(): Uint8Array {
  const tasks = [
    taskXml({ uid: 1, name: 'A', outlineNumber: '1' }),
    taskXml({
      uid: 2,
      name: 'B',
      outlineNumber: '2',
      predecessorLinks: [{ predUid: 1, type: 0, lag: 0 }], // FS
    }),
    taskXml({
      uid: 3,
      name: 'C',
      outlineNumber: '3',
      predecessorLinks: [{ predUid: 1, type: 2, lag: 0 }], // SS
    }),
  ].join('\n')
  return projectXml({ name: 'M03 Dependencies', tasks })
}

/** M04 — lag/lead: FS +240min lag and FS -120min lead. */
export function m04LagLead(): Uint8Array {
  const tasks = [
    taskXml({ uid: 1, name: 'First', outlineNumber: '1' }),
    taskXml({
      uid: 2,
      name: 'Second',
      outlineNumber: '2',
      predecessorLinks: [{ predUid: 1, type: 0, lag: 2400, lagFormat: 1 }], // +240 min
    }),
    taskXml({
      uid: 3,
      name: 'Third',
      outlineNumber: '3',
      predecessorLinks: [{ predUid: 2, type: 0, lag: -1200, lagFormat: 1 }], // -120 min (lead)
    }),
  ].join('\n')
  return projectXml({ name: 'M04 Lag Lead', tasks })
}

/** M05 — calendars: base Standard + derived (baseCalendarId link). */
export function m05Calendars(): Uint8Array {
  const derived = `
      <Calendar>
        <UID>2</UID>
        <Name>Early</Name>
        <IsBaseCalendar>false</IsBaseCalendar>
        <BaseCalendarUID>1</BaseCalendarUID>
        <WeekDays>
          <WeekDay><DayType>1</DayType><DayWorking>false</DayWorking></WeekDay>
          <WeekDay><DayType>2</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
          <WeekDay><DayType>3</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
          <WeekDay><DayType>4</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
          <WeekDay><DayType>5</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
          <WeekDay><DayType>6</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
          <WeekDay><DayType>7</DayType><DayWorking>false</DayWorking></WeekDay>
        </WeekDays>
        <Exceptions />
      </Calendar>`
  return projectXml({
    name: 'M05 Calendars',
    calendars: STANDARD_CALENDAR_XML + derived,
  })
}

/** M06 — calendar exceptions: a holiday + a working-Saturday. */
export function m06Exceptions(): Uint8Array {
  const calWithExc = `
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
        <Exceptions>
          <Exception>
            <Name>Holiday</Name>
            <Type>1</Type>
            <Start>2026-08-03T00:00:00</Start>
            <Finish>2026-08-03T23:59:00</Finish>
            <WorkingTimes />
          </Exception>
          <Exception>
            <Name>Half Saturday</Name>
            <Type>1</Type>
            <Start>2026-08-08T00:00:00</Start>
            <Finish>2026-08-08T23:59:00</Finish>
            <WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>13:00:00</ToTime></WorkingTime></WorkingTimes>
          </Exception>
        </Exceptions>
      </Calendar>`
  return projectXml({ name: 'M06 Exceptions', calendars: calWithExc })
}

/** M07 — resources: work, material, cost. */
export function m07Resources(): Uint8Array {
  const resources = [
    resourceXml({ uid: 1, name: 'Alice', type: 1, maxUnits: 1, standardRate: 50, calendarUid: 1 }),
    resourceXml({ uid: 2, name: 'Concrete', type: 2, standardRate: 100 }),
    resourceXml({ uid: 3, name: 'Travel', type: 3 }),
  ].join('\n')
  return projectXml({ name: 'M07 Resources', resources })
}

/** M08 — assignments. */
export function m08Assignments(): Uint8Array {
  const tasks = taskXml({ uid: 1, name: 'Build', outlineNumber: '1', duration: 'PT8H0M0S' })
  const resources = [
    resourceXml({ uid: 1, name: 'Alice', type: 1, maxUnits: 1, standardRate: 50, calendarUid: 1 }),
    resourceXml({ uid: 2, name: 'Concrete', type: 2, standardRate: 100 }),
  ].join('\n')
  const assignments = [
    assignmentXml({ uid: 1, taskUid: 1, resourceUid: 1, units: 1, work: 'PT8H0M0S', cost: 50 }),
    assignmentXml({ uid: 2, taskUid: 1, resourceUid: 2, units: 5, cost: 500 }),
  ].join('\n')
  return projectXml({ name: 'M08 Assignments', tasks, resources, assignments })
}

/** M09 — constraints: SNET + MFO. */
export function m09Constraints(): Uint8Array {
  const tasks = [
    taskXml({
      uid: 1,
      name: 'SNET Task',
      outlineNumber: '1',
      constraintType: 4, // SNET
      constraintDate: '2026-08-03T09:00:00',
    }),
    taskXml({
      uid: 2,
      name: 'MFO Task',
      outlineNumber: '2',
      constraintType: 3, // MFO
      constraintDate: '2026-08-07T17:00:00',
      duration: 'PT8H0M0S',
    }),
  ].join('\n')
  return projectXml({ name: 'M09 Constraints', tasks })
}

/** M10 — deadlines + progress. */
export function m10DeadlinesProgress(): Uint8Array {
  const tasks = [
    taskXml({
      uid: 1,
      name: 'Tracked',
      outlineNumber: '1',
      duration: 'PT8H0M0S',
      deadline: '2026-08-07T17:00:00',
      percentComplete: 75,
    }),
  ].join('\n')
  return projectXml({ name: 'M10 Deadlines Progress', tasks })
}

/** M11 — baseline (single). */
export function m11Baseline(): Uint8Array {
  const tasks = [
    taskXml({
      uid: 1,
      name: 'Baselined',
      outlineNumber: '1',
      duration: 'PT8H0M0S',
      baselines: [
        {
          slot: 0,
          start: '2026-08-03T09:00:00',
          finish: '2026-08-03T17:00:00',
          duration: 'PT8H0M0S',
          work: 'PT8H0M0S',
          cost: 100,
        },
      ],
    }),
  ].join('\n')
  return projectXml({ name: 'M11 Baseline', tasks, lastSaved: '2026-08-02T08:00:00' })
}

/** M12 — multiple baselines. */
export function m12MultipleBaseline(): Uint8Array {
  const tasks = [
    taskXml({
      uid: 1,
      name: 'Multi Baseline',
      outlineNumber: '1',
      duration: 'PT8H0M0S',
      baselines: [
        {
          slot: 0,
          start: '2026-08-03T09:00:00',
          finish: '2026-08-03T17:00:00',
          duration: 'PT8H0M0S',
          work: 'PT8H0M0S',
          cost: 100,
        },
        {
          slot: 1,
          start: '2026-08-04T09:00:00',
          finish: '2026-08-04T17:00:00',
          duration: 'PT8H0M0S',
          work: 'PT8H0M0S',
          cost: 110,
        },
      ],
    }),
  ].join('\n')
  return projectXml({ name: 'M12 Multiple Baseline', tasks, lastSaved: '2026-08-02T08:00:00' })
}

/** M13 — custom fields. */
export function m13CustomFields(): Uint8Array {
  const extAttrs = `
      <ExtendedAttribute>
        <FieldID>188743731</FieldID>
        <FieldName>Text1</FieldName>
        <Alias>Sponsor</Alias>
        <Type>text</Type>
      </ExtendedAttribute>
      <ExtendedAttribute>
        <FieldID>188743734</FieldID>
        <FieldName>Number1</FieldName>
        <Alias>Budget</Alias>
        <Type>number</Type>
      </ExtendedAttribute>`
  const tasks = [
    taskXml({
      uid: 1,
      name: 'With Custom',
      outlineNumber: '1',
      duration: 'PT8H0M0S',
      extendedAttrs: [
        { fieldId: '188743731', value: 'Acme Corp' },
        { fieldId: '188743734', value: '5000' },
      ],
    }),
  ].join('\n')
  return projectXml({ name: 'M13 Custom Fields', tasks, extendedAttributes: extAttrs })
}

/** M14 — unsupported feature: recurring calendar exception (Type=2 yearly). */
export function m14Unsupported(): Uint8Array {
  const calWithRecurring = `
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
        <Exceptions>
          <Exception>
            <Name>Christmas</Name>
            <Type>2</Type>
            <Start>2026-12-25T00:00:00</Start>
            <Finish>2026-12-25T23:59:00</Finish>
            <WorkingTimes />
          </Exception>
        </Exceptions>
      </Calendar>`
  return projectXml({ name: 'M14 Unsupported', calendars: calWithRecurring })
}

/** M15 — malformed references: dependency with a dangling predecessor. */
export function m15MalformedReference(): Uint8Array {
  const tasks = [
    taskXml({ uid: 1, name: 'Real', outlineNumber: '1' }),
    taskXml({
      uid: 2,
      name: 'Ghost Successor',
      outlineNumber: '2',
      predecessorLinks: [{ predUid: 999, type: 0, lag: 0 }], // dangling
    }),
  ].join('\n')
  return projectXml({ name: 'M15 Malformed Refs', tasks })
}

/** M16 — large project: 60 tasks, 50 deps, 10 resources, 60 assignments. */
export function m16Large(): Uint8Array {
  const tasks: string[] = []
  for (let i = 1; i <= 60; i++) {
    const links =
      i >= 2 && i <= 51 ? [{ predUid: i - 1, type: 0, lag: ((i % 3) * 10) as number }] : undefined
    tasks.push(
      taskXml({
        uid: i,
        name: `Task ${i}`,
        outlineNumber: String(i),
        duration: 'PT8H0M0S',
        predecessorLinks: links,
      }),
    )
  }
  const resources: string[] = []
  for (let i = 1; i <= 10; i++) {
    resources.push(
      resourceXml({
        uid: i,
        name: `Resource ${i}`,
        type: 1,
        maxUnits: 1,
        standardRate: 50 + i,
        calendarUid: 1,
      }),
    )
  }
  const assignments: string[] = []
  for (let i = 1; i <= 60; i++) {
    assignments.push(
      assignmentXml({
        uid: i,
        taskUid: i,
        resourceUid: ((i - 1) % 10) + 1,
        units: 1,
        work: 'PT8H0M0S',
      }),
    )
  }
  return projectXml({
    name: 'M16 Large',
    tasks: tasks.join('\n'),
    resources: resources.join('\n'),
    assignments: assignments.join('\n'),
  })
}

/** M17 — adversarial XML: DOCTYPE entity-expansion attack + deep nesting. */
export function m17Adversarial(): Uint8Array {
  const xml = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE bomb [
  <!ENTITY a "0123456789">
  <!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;">
  <!ENTITY c "&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;">
  <!ENTITY d "&c;&c;&c;&c;&c;&c;&c;&c;&c;&c;">
]>
<Project xmlns="http://schemas.microsoft.com/project">
  <SaveVersion>16</SaveVersion>
  <Name>&d;</Name>
</Project>
`
  return encodeUtf8(xml)
}

/** M17b — deeply nested XML (exceeds the depth guard). */
export function m17bDeeplyNested(): Uint8Array {
  let inner = '<Project xmlns="http://schemas.microsoft.com/project"><SaveVersion>16</SaveVersion>'
  for (let i = 0; i < 300; i++) inner += '<Nested>'
  for (let i = 0; i < 300; i++) inner += '</Nested>'
  inner += '</Project>'
  return encodeUtf8('<?xml version="1.0"?>\n' + inner)
}

/** M17c — malformed XML (unclosed root). */
export function m17cMalformed(): Uint8Array {
  return encodeUtf8(
    '<?xml version="1.0"?>\n<Project xmlns="http://schemas.microsoft.com/project"><SaveVersion>16</SaveVersion>',
  )
}

/**
 * M18 — lag units golden (PROJECT-015 correction round 1): one FS dependency
 * per working lag unit, each with an exact expected conversion, plus a
 * negative day-unit lead. Declares the conversion factors explicitly
 * (`MinutesPerDay` 480 / `MinutesPerWeek` 2400 / `DaysPerMonth` 20 — the
 * MSPDI defaults) so the exact values below are unambiguous:
 *
 *   t2  fmt 1  LinkLag  1500 →  150 min  (15 hours of minutes)
 *   t3  fmt 3  LinkLag    30 →  180 min  (3 working hours)
 *   t4  fmt 5  LinkLag    25 → 1200 min  (2.5 working days)
 *   t5  fmt 7  LinkLag    76 → 18240 min (7.6 working weeks)
 *   t6  fmt 9  LinkLag    25 → 24000 min (2.5 working months)
 *   t7  fmt 5  LinkLag    -5 →  -240 min (-0.5 working day lead)
 *
 * The anchor finishes Monday 2026-08-03 13:00 (PT4H); every successor is a
 * 1-hour FS task of the anchor, so each unit's converted working minutes are
 * directly observable in `lagMinutes` and in the scheduled start dates.
 */
export function m18LagUnits(): Uint8Array {
  const tasks = [
    taskXml({ uid: 1, name: 'Anchor', outlineNumber: '1', duration: 'PT4H0M0S' }),
    taskXml({
      uid: 2,
      name: 'Minute Lag',
      outlineNumber: '2',
      duration: 'PT1H0M0S',
      predecessorLinks: [{ predUid: 1, type: 0, lag: 1500, lagFormat: 1 }],
    }),
    taskXml({
      uid: 3,
      name: 'Hour Lag',
      outlineNumber: '3',
      duration: 'PT1H0M0S',
      predecessorLinks: [{ predUid: 1, type: 0, lag: 30, lagFormat: 3 }],
    }),
    taskXml({
      uid: 4,
      name: 'Day Lag',
      outlineNumber: '4',
      duration: 'PT1H0M0S',
      predecessorLinks: [{ predUid: 1, type: 0, lag: 25, lagFormat: 5 }],
    }),
    taskXml({
      uid: 5,
      name: 'Week Lag',
      outlineNumber: '5',
      duration: 'PT1H0M0S',
      predecessorLinks: [{ predUid: 1, type: 0, lag: 76, lagFormat: 7 }],
    }),
    taskXml({
      uid: 6,
      name: 'Month Lag',
      outlineNumber: '6',
      duration: 'PT1H0M0S',
      predecessorLinks: [{ predUid: 1, type: 0, lag: 25, lagFormat: 9 }],
    }),
    taskXml({
      uid: 7,
      name: 'Day Lead',
      outlineNumber: '7',
      duration: 'PT1H0M0S',
      predecessorLinks: [{ predUid: 1, type: 0, lag: -5, lagFormat: 5 }],
    }),
  ].join('\n')
  return projectXml({
    name: 'M18 Lag Units',
    minutesPerDay: 480,
    minutesPerWeek: 2400,
    daysPerMonth: 20,
    tasks,
  })
}
