/**
 * PROJECT-022 — dependency link geometry.
 *
 * `buildDependencies` maps the canonical `Dependency` records onto visible
 * link geometry between the Gantt rows, in fraction space (x) and
 * absolute-row-index space (y). This is pure VISIBILITY projection — no
 * dependency semantics are invented, re-derived, or re-validated here: the
 * type/edge mapping is the documented FS/SS/FF/SF rendering convention and
 * nothing else.
 *
 * Endpoint resolution (deterministic, the Microsoft Project behavior):
 *
 * - A dependency endpoint whose task row is VISIBLE and IN THE VIRTUALIZED
 *   WINDOW attaches to that row's own bar edge.
 * - An endpoint hidden by COLLAPSE attaches to its nearest visible
 *   ancestor's row — the collapsed summary's rolled-up bar edge (the
 *   ancestor's own schedule dates; never recomputed).
 * - An endpoint whose row exists but is scrolled OUT of the window
 *   produces NO link (off-screen), and a collapsed endpoint whose nearest
 *   visible ancestor is out of the window produces none either.
 * - When both endpoints resolve to the SAME row (the link is entirely
 *   inside one collapsed subtree) the link is omitted — there is nothing
 *   to draw.
 * - An endpoint whose resolved row has no schedule edge instant (no
 *   schedule, or the needed instant is missing) omits the link — dates are
 *   never invented (architecture-lock §11).
 *
 * The route is a deterministic four-point elbow: exit the source edge
 * horizontally, one vertical segment at the midpoint of the two anchor
 * fractions, enter the target edge horizontally:
 * `[(xFrom, yFrom), (xm, yFrom), (xm, yTo), (xTo, yTo)]` with
 * `xm = (xFrom + xTo) / 2`. Hosts render the polyline (or refine the
 * stroke); hit-testing links needs a pixel tolerance and stays a host
 * concern.
 */
import type {
  DependencyId,
  DependencyType,
  ProjectDocument,
  TaskId,
  TaskSchedule,
} from '@genoffice/project-contracts'
import type { ProjectViewProjection } from '../projection.js'
import type { TimelineViewport } from '../state.js'
import { parseInstant } from '../state.js'
import { viewportSpanMs } from './gantt-bars.js'
import type { ProjectRowWindow } from './virtualization.js'
import { rowWindowIsEmpty } from './virtualization.js'

/** A point on the link surface: x in viewport-fraction space, y in
 * absolute visible-row-index space. Hosts scale both to pixels. */
export interface ProjectLinkPoint {
  readonly fraction: number
  readonly rowIndex: number
}

/** One end of a link: the RESOLVED attach target (the task's own row, or
 * its nearest visible ancestor when collapsed), the bar edge it attaches
 * to, and the edge instant's clamped viewport fraction. */
export interface ProjectLinkEndpoint {
  /** The resolved attach task (may be an ancestor of the dependency's
   * endpoint task). */
  readonly taskId: TaskId
  readonly rowIndex: number
  readonly edge: 'start' | 'finish'
  readonly fraction: number
}

/** One dependency link's geometry. `predecessorTaskId`/`successorTaskId`
 * are the CANONICAL dependency endpoints (never the resolved ancestors) —
 * the resolved rows are on `from`/`to`. */
export interface ProjectDependencyLink {
  readonly dependencyId: DependencyId
  readonly type: DependencyType
  readonly predecessorTaskId: TaskId
  readonly successorTaskId: TaskId
  readonly from: ProjectLinkEndpoint
  readonly to: ProjectLinkEndpoint
  readonly route: readonly ProjectLinkPoint[]
}

/** The dependency-link surface (PROJECT-022): the visible links in
 * canonical document order. */
export type ProjectDependencies = readonly ProjectDependencyLink[]

/** The bar edge each relationship type attaches to (the rendering
 * convention): FS finishes→starts, SS starts→starts, FF finishes→finishes,
 * SF starts→finishes. */
const PRED_EDGE_BY_TYPE: Readonly<Record<DependencyType, 'start' | 'finish'>> = {
  FS: 'finish',
  SS: 'start',
  FF: 'finish',
  SF: 'start',
}
const SUCC_EDGE_BY_TYPE: Readonly<Record<DependencyType, 'start' | 'finish'>> = {
  FS: 'start',
  SS: 'start',
  FF: 'finish',
  SF: 'finish',
}

/**
 * Builds the dependency-link geometry. Pure and deterministic; never
 * mutates its inputs.
 */
export function buildDependencies(
  document: ProjectDocument,
  projection: ProjectViewProjection,
  viewport: TimelineViewport,
  rowWindow: ProjectRowWindow,
): ProjectDependencies {
  const span = viewportSpanMs(viewport)
  if (span === undefined || rowWindowIsEmpty(rowWindow)) return []

  const rowIndexByTask = new Map<TaskId, number>()
  projection.rows.forEach((row, index) => rowIndexByTask.set(row.taskId, index))
  const scheduleByTask = new Map<TaskId, TaskSchedule | undefined>()
  for (const row of projection.rows) scheduleByTask.set(row.taskId, row.schedule)
  const parentByTask = new Map<TaskId, TaskId | undefined>()
  for (const task of document.tasks) {
    if (task.parentTaskId !== undefined) parentByTask.set(task.id, task.parentTaskId)
  }
  const inWindow = (index: number): boolean =>
    index >= rowWindow.firstIndex && index <= rowWindow.lastIndex

  const links: ProjectDependencyLink[] = []
  for (const dependency of document.dependencies) {
    const from = resolveEndpoint(
      dependency.predecessorId,
      PRED_EDGE_BY_TYPE[dependency.type],
      rowIndexByTask,
      scheduleByTask,
      parentByTask,
      inWindow,
      span,
    )
    if (from === undefined) continue
    const to = resolveEndpoint(
      dependency.successorId,
      SUCC_EDGE_BY_TYPE[dependency.type],
      rowIndexByTask,
      scheduleByTask,
      parentByTask,
      inWindow,
      span,
    )
    if (to === undefined || to.rowIndex === from.rowIndex) continue

    const mid = (from.fraction + to.fraction) / 2
    links.push({
      dependencyId: dependency.id,
      type: dependency.type,
      predecessorTaskId: dependency.predecessorId,
      successorTaskId: dependency.successorId,
      from,
      to,
      route: [
        { fraction: from.fraction, rowIndex: from.rowIndex },
        { fraction: mid, rowIndex: from.rowIndex },
        { fraction: mid, rowIndex: to.rowIndex },
        { fraction: to.fraction, rowIndex: to.rowIndex },
      ],
    })
  }
  return links
}

/**
 * Resolves one dependency endpoint to its attach row/edge/fraction:
 * the task's own row when visible and in-window; otherwise the nearest
 * VISIBLE ancestor's row (collapse resolution) when that ancestor is
 * in-window; otherwise `undefined` (no link). The edge instant is the
 * resolved task's schedule `scheduledStart`/`scheduledFinish`; a missing
 * instant resolves to `undefined` — never an invented position.
 */
function resolveEndpoint(
  taskId: TaskId,
  edge: 'start' | 'finish',
  rowIndexByTask: ReadonlyMap<TaskId, number>,
  scheduleByTask: ReadonlyMap<TaskId, TaskSchedule | undefined>,
  parentByTask: ReadonlyMap<TaskId, TaskId | undefined>,
  inWindow: (index: number) => boolean,
  span: { start: number; ms: number },
): ProjectLinkEndpoint | undefined {
  let resolvedTask: TaskId | undefined = taskId
  let guard = 0
  while (resolvedTask !== undefined && guard < 10_000) {
    guard += 1
    const rowIndex = rowIndexByTask.get(resolvedTask)
    if (rowIndex !== undefined) {
      if (!inWindow(rowIndex)) return undefined // visible but scrolled off
      const schedule = scheduleByTask.get(resolvedTask)
      const instant = edge === 'start' ? schedule?.scheduledStart : schedule?.scheduledFinish
      if (instant === undefined) return undefined // never invent a date
      const at = parseInstant(instant)
      if (at === undefined) return undefined
      const raw = (at - span.start) / span.ms
      return {
        taskId: resolvedTask,
        rowIndex,
        edge,
        fraction: Number.isFinite(raw) ? Math.max(0, Math.min(1, raw)) : 0,
      }
    }
    // Hidden by collapse: climb to the nearest visible ancestor.
    resolvedTask = parentByTask.get(resolvedTask)
  }
  return undefined
}
