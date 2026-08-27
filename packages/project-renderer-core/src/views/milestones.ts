/**
 * PROJECT-022 — milestone geometry.
 *
 * `buildMilestones` maps zero-duration schedule events onto diamond
 * markers in fraction space (hosts multiply by the pixel width). A row
 * carries milestone geometry iff its schedule has a `scheduledStart` AND
 * the row is milestone-like: the canonical `milestone` flag is set, OR the
 * task's duration is zero, OR the schedule window is a zero span
 * (`scheduledStart === scheduledFinish`). The diamond instant is
 * `scheduledStart` — never invented, never defaulted (architecture-lock
 * §11): a row without a schedule has no milestone geometry even when the
 * flag is set.
 *
 * A flagged milestone row that ALSO has a real schedule span
 * (`scheduledFinish > scheduledStart`) keeps BOTH geometries — a bar from
 * `./gantt-bars.js` and a marker here; hosts render both (the documented
 * orthogonal rule).
 */
import type { ISODateTime, TaskId } from '@genoffice/project-contracts'
import type { ProjectViewProjection } from '../projection.js'
import type { TimelineViewport } from '../state.js'
import { parseInstant } from '../state.js'
import { viewportSpanMs } from './gantt-bars.js'
import type { ProjectRowWindow } from './virtualization.js'
import { rowWindowIsEmpty } from './virtualization.js'

/** One milestone diamond's geometry. `fraction` is clamped to [0, 1]; the
 * flags state when the instant falls outside the viewport on either side
 * (hosts may render edge affordances). */
export interface ProjectMilestone {
  readonly taskId: TaskId
  /** Absolute index into the projection's visible rows. */
  readonly rowIndex: number
  /** The diamond instant (the schedule's `scheduledStart`). */
  readonly instant: ISODateTime
  readonly fraction: number
  readonly beforeViewport: boolean
  readonly afterViewport: boolean
}

/** The milestone surface (PROJECT-022): markers of every milestone-like
 * in-window row, in ascending row order. */
export type ProjectMilestones = readonly ProjectMilestone[]

/**
 * Builds the milestone geometry. Pure and deterministic; never mutates its
 * inputs; no marker is ever produced without a schedule.
 */
export function buildMilestones(
  projection: ProjectViewProjection,
  viewport: TimelineViewport,
  rowWindow: ProjectRowWindow,
): ProjectMilestones {
  const span = viewportSpanMs(viewport)
  if (span === undefined || rowWindowIsEmpty(rowWindow)) return []

  const milestones: ProjectMilestone[] = []
  for (
    let index = rowWindow.firstIndex;
    index <= rowWindow.lastIndex && index < projection.rows.length;
    index += 1
  ) {
    const row = projection.rows[index]
    if (row === undefined) continue
    const schedule = row.schedule
    if (schedule === undefined) continue
    const scheduledStart = schedule.scheduledStart
    if (scheduledStart === undefined) continue
    const startMs = parseInstant(scheduledStart)
    if (startMs === undefined) continue
    const finishMs = parseInstant(schedule.scheduledFinish ?? '')
    const zeroSpan = finishMs !== undefined && finishMs === startMs
    if (!row.milestone && row.duration !== 0 && !zeroSpan) continue

    const raw = (startMs - span.start) / span.ms
    milestones.push({
      taskId: row.taskId,
      rowIndex: index,
      instant: scheduledStart,
      fraction: Number.isFinite(raw) ? Math.max(0, Math.min(1, raw)) : 0,
      beforeViewport: raw < 0,
      afterViewport: raw > 1,
    })
  }
  return milestones
}
