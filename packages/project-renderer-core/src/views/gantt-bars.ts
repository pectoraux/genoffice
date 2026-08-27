/**
 * PROJECT-022 — Gantt bar geometry (the right pane's row geometry).
 *
 * `buildGanttBars` maps each VISIBLE, IN-WINDOW projection row's schedule
 * window onto horizontal bar geometry in FRACTION space: positions are
 * fractions in [0, 1] across the timeline viewport (hosts multiply by the
 * pixel width — the core never sees a pixel), rows are addressed by their
 * absolute visible-row index. Bars partially outside the viewport are
 * CLIPPED to [0, 1] with explicit `startsBefore`/`finishesAfter` flags so
 * hosts can render edge affordances without re-deriving anything.
 *
 * Bar dates are exclusively the schedule's `scheduledStart`/`scheduledFinish`
 * (the projection's by-reference join with the scheduling authority). A row
 * without both instants has NO bar — dates are never invented, defaulted,
 * or degraded (architecture-lock §11), and a schedule window with
 * `start >= finish` is not a span and produces no bar either (a
 * zero-duration schedule is MILESTONE geometry, `./milestones.js`).
 *
 * The progress point is a pure linear interpolation of the schedule's
 * percent-complete echo over the bar's raw (unclamped) span — a documented
 * rendering convention, not a scheduling computation: `scheduleStart +
 * (percent/100) × (scheduleFinish − scheduleStart)` with the percent source
 * `schedule.percentComplete ?? task.percentComplete` (the schedule-first
 * precedence; both are canonical echoes) clamped to [0, 100].
 */
import type { TaskId } from '@genoffice/project-contracts'
import type { ProjectViewProjection } from '../projection.js'
import type { TimelineViewport } from '../state.js'
import { parseInstant } from '../state.js'
import type { ProjectRowWindow } from './virtualization.js'
import { rowWindowIsEmpty } from './virtualization.js'

/** One bar's geometry. `startFraction`/`finishFraction` are clamped to
 * [0, 1]; the flags state when the true schedule window extends beyond the
 * viewport on either side. */
export interface ProjectGanttBar {
  readonly taskId: TaskId
  /** Absolute index into the projection's visible rows. */
  readonly rowIndex: number
  /** `leaf` for regular task bars, `summary` for summary rows (the same
   * date source — the schedule's rolled-up window; only the rendering kind
   * differs). */
  readonly kind: 'leaf' | 'summary'
  readonly startFraction: number
  readonly finishFraction: number
  readonly startsBefore: boolean
  readonly finishesAfter: boolean
  /** The progress point as a clamped fraction in [0, 1]: where the
   * percent-complete overlay reaches. Always within the clamped bar span. */
  readonly progressFraction: number
}

/** The Gantt-bar surface (PROJECT-022): the bars of every in-window row
 * that has a real schedule span, in ascending row order. */
export type ProjectGanttBars = readonly ProjectGanttBar[]

/**
 * Builds the bar geometry. Pure and deterministic; never mutates its
 * inputs; no bar is ever produced for a row without a complete schedule
 * span.
 */
export function buildGanttBars(
  projection: ProjectViewProjection,
  viewport: TimelineViewport,
  rowWindow: ProjectRowWindow,
): ProjectGanttBars {
  const span = viewportSpanMs(viewport)
  if (span === undefined || rowWindowIsEmpty(rowWindow)) return []

  const bars: ProjectGanttBar[] = []
  for (
    let index = rowWindow.firstIndex;
    index <= rowWindow.lastIndex && index < projection.rows.length;
    index += 1
  ) {
    const row = projection.rows[index]
    if (row === undefined) continue
    const startMs = parseInstant(row.schedule?.scheduledStart ?? '')
    const finishMs = parseInstant(row.schedule?.scheduledFinish ?? '')
    if (startMs === undefined || finishMs === undefined || finishMs <= startMs) continue

    const rawStart = (startMs - span.start) / span.ms
    const rawFinish = (finishMs - span.start) / span.ms
    // Percent source: the schedule's derived echo when present (rolled up
    // for summaries), else the canonical task field; clamped to [0, 100].
    const percentFraction = clamp01((row.schedule?.percentComplete ?? row.percentComplete) / 100)
    const rawProgress = rawStart + percentFraction * (rawFinish - rawStart)
    const startFraction = clamp01(rawStart)
    const finishFraction = clamp01(rawFinish)
    bars.push({
      taskId: row.taskId,
      rowIndex: index,
      kind: row.summary ? 'summary' : 'leaf',
      startFraction,
      finishFraction,
      startsBefore: rawStart < 0,
      finishesAfter: rawFinish > 1,
      // clamp is monotonic, so the clamped progress point never precedes
      // the clamped bar start:
      progressFraction: clamp01(rawProgress),
    })
  }
  return bars
}

/** The parsed viewport window as `{ start, ms }` where `start` is the epoch
 * start and `ms` the positive span; `undefined` for an unparseable or
 * degenerate window (no geometry is built in that case). */
export function viewportSpanMs(
  viewport: TimelineViewport,
): { start: number; ms: number } | undefined {
  const start = parseInstant(viewport.start)
  const finish = parseInstant(viewport.finish)
  if (start === undefined || finish === undefined || finish <= start) return undefined
  return { start, ms: finish - start }
}

const clamp01 = (value: number): number =>
  Number.isFinite(value) ? Math.max(0, Math.min(1, value)) : 0
