/**
 * The host-side canonical authority bindings — shared by both shells
 * (established at PROJECT-027; moved to the shared `@genoffice/project-host`
 * layer at PROJECT-028).
 *
 * THE ONLY module in the host binding that imports
 * `@genoffice/project-scheduling`. Every scheduling, calendar, and
 * allocation value the host renders or schedules flows through these three
 * bindings — the renderer-core injection seams the accepted PROJECT-021/025/
 * 026 increments define:
 *
 * - `scheduleRunner`           → the `ScheduleRunner` the renderer session
 *                                injects (PROJECT-021).
 * - `workingTimeQuery`         → the `CalendarWorkingTimeQuery` the Gantt
 *                                calendar surfaces inject (PROJECT-025) —
 *                                the canonical `workingIntervals ∘
 *                                resolveCalendar` composition, exactly as the
 *                                spec documents.
 * - `allocationQuery`          → the `ResourceAllocationQuery` the resource
 *                                utilization surfaces inject (PROJECT-026) —
 *                                the canonical `resourceAllocations`
 *                                authority.
 *
 * The host implements NONE of these semantics (architecture-lock §3/§6): it
 * wires the accepted authorities together. The discipline suite pins this
 * module as the single scheduling import site.
 *
 * Importing the scheduling package also registers the canonical resource
 * leveler with the engine's `LevelResources` dispatch (the package's
 * documented side effect) — the host gains that behavior through the same
 * single import, never by calling the leveler itself.
 */
import type {
  CalendarWorkingTimeQuery,
  ResourceAllocationQuery,
  ScheduleRunner,
} from '@genoffice/project-renderer-core'
import {
  resolveCalendar,
  resourceAllocations,
  schedule,
  workingIntervals,
} from '@genoffice/project-scheduling'

/** The scheduling authority (PROJECT-021 `ScheduleRunner` injection). */
export const scheduleRunner: ScheduleRunner = schedule

/** The working-time authority (PROJECT-025 `CalendarWorkingTimeQuery`). */
export const workingTimeQuery: CalendarWorkingTimeQuery = (calendars, calendarId, start, finish) =>
  workingIntervals(resolveCalendar({ calendars: [...calendars] }, calendarId), start, finish)

/** The allocation authority (PROJECT-026 `ResourceAllocationQuery`). */
export const allocationQuery: ResourceAllocationQuery = resourceAllocations
