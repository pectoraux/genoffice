/**
 * PROJECT-022 — the task grid (left pane) view model.
 *
 * `buildTaskGrid` projects the accepted `ProjectViewProjection` rows into
 * the structured, virtualized grid model both hosts render: the resolved
 * column set and one STRUCTURED cell per column per visible row. Cells
 * carry semantic values — durations as working minutes, dates as ISO
 * instants, predecessors as typed link records — never formatted strings:
 * label/number formatting is a host/locale concern (the PROJECT-021
 * locale-free discipline).
 *
 * Column resolution follows the canonical document: the active
 * `ProjectTable`'s `columns` (field-name strings, the `.gproj` convention)
 * are mapped onto the supported field set below; unrecognized field names
 * produce `unsupported` cells (never a crash, never invented data). When no
 * table is active — or the active table declares no columns — the
 * documented DEFAULT column set applies.
 *
 * Scheduling values come exclusively from the projection's by-reference
 * schedule join: start/finish cells are the schedule's scheduled instants
 * and are `empty` when the row has no schedule — dates are never invented
 * (architecture-lock §11). Where both a canonical task field and its
 * derived schedule echo exist (duration, percent complete), the SCHEDULE
 * value wins and the task field is the fallback — the same
 * schedule-first precedence the projection's project window uses.
 */
import type { ISODateTime, ProjectDocument, WorkingMinutes } from '@genoffice/project-contracts'
import type { DependencyType, DependencyId, TaskId } from '@genoffice/project-contracts'
import type { ProjectTaskRow, ProjectViewProjection } from '../projection.js'
import type { ProjectRowWindow } from './virtualization.js'
import { rowWindowIsEmpty } from './virtualization.js'

/** The supported grid column fields. The string keys are the canonical
 * column-name convention (`ProjectTable.columns`, the `.gproj` fixtures). */
export type TaskGridField =
  | 'rowNumber'
  | 'taskName'
  | 'duration'
  | 'start'
  | 'finish'
  | 'percentComplete'
  | 'predecessors'
  | 'resourceNames'
  | 'wbs'
  | 'outlineLevel'
  | 'priority'
  | 'uid'

const FIELD_BY_COLUMN_NAME: ReadonlyMap<string, TaskGridField> = new Map(
  (
    [
      ['rowNumber', 'rowNumber'],
      ['name', 'taskName'],
      ['duration', 'duration'],
      ['start', 'start'],
      ['finish', 'finish'],
      ['percentComplete', 'percentComplete'],
      ['predecessors', 'predecessors'],
      ['resourceNames', 'resourceNames'],
      ['wbs', 'wbs'],
      ['outlineLevel', 'outlineLevel'],
      ['priority', 'priority'],
      ['uid', 'uid'],
    ] as const
  ).map(([name, field]) => [name, field]),
)

/** The documented default column set when no canonical table is active:
 * the Microsoft Project "Entry"-like set (row number, name, duration,
 * scheduled start/finish, predecessors, resource names). */
export const DEFAULT_TASK_GRID_COLUMNS: readonly string[] = [
  'rowNumber',
  'name',
  'duration',
  'start',
  'finish',
  'predecessors',
  'resourceNames',
]

/** One resolved grid column: the mapped field (or `unsupported` for a
 * field name outside the supported set) plus the raw canonical column
 * string it was resolved from. */
export interface ProjectGridColumn {
  readonly field: TaskGridField | 'unsupported'
  /** The raw `ProjectTable.columns` (or default-set) string this column
   * was resolved from. */
  readonly source: string
}

/** One predecessor link inside a `predecessors` cell: the raw canonical
 * dependency record (document order), with the predecessor's uid for
 * MS-Project-style ID display. */
export interface ProjectGridPredecessorLink {
  readonly dependencyId: DependencyId
  readonly predecessorTaskId: TaskId
  readonly predecessorUid: number
  readonly type: DependencyType
  readonly lagMinutes: number
}

/** A structured, unformatted grid cell. Hosts format for display. */
export type ProjectGridCell =
  | { readonly kind: 'rowNumber'; readonly value: number }
  | {
      readonly kind: 'taskName'
      readonly text: string
      readonly outlineLevel: number
      readonly summary: boolean
      readonly milestone: boolean
      readonly collapsed: boolean
    }
  | { readonly kind: 'duration'; readonly minutes: WorkingMinutes }
  | { readonly kind: 'instant'; readonly iso: ISODateTime }
  | { readonly kind: 'empty' }
  | { readonly kind: 'percentComplete'; readonly value: number }
  | { readonly kind: 'predecessors'; readonly links: readonly ProjectGridPredecessorLink[] }
  | { readonly kind: 'resources'; readonly names: readonly string[] }
  | { readonly kind: 'text'; readonly text: string }
  | { readonly kind: 'number'; readonly value: number }
  | { readonly kind: 'unsupported'; readonly source: string }

/** One virtualized grid row: the ABSOLUTE visible-row index, the projection
 * row joined BY REFERENCE, and the cells positionally aligned with
 * `ProjectTaskGrid.columns`. */
export interface ProjectGridRow {
  readonly index: number
  readonly row: ProjectTaskRow
  readonly cells: readonly ProjectGridCell[]
}

/** The task-grid view model (PROJECT-022). */
export interface ProjectTaskGrid {
  readonly columns: readonly ProjectGridColumn[]
  readonly rows: readonly ProjectGridRow[]
  readonly rowWindow: ProjectRowWindow
}

/**
 * Builds the task grid. Pure and deterministic: the same
 * `(document, projection, rowWindow)` triple always produces the same grid;
 * inputs are never mutated; projection rows are joined by reference.
 */
export function buildTaskGrid(
  document: ProjectDocument,
  projection: ProjectViewProjection,
  rowWindow: ProjectRowWindow,
): ProjectTaskGrid {
  const columns = resolveColumns(projection)
  const predecessorsByTask = predecessorsBySuccessor(document)
  const documentIndexByTask = new Map<TaskId, number>()
  document.tasks.forEach((task, index) => documentIndexByTask.set(task.id, index))

  const rows: ProjectGridRow[] = []
  if (!rowWindowIsEmpty(rowWindow)) {
    for (
      let index = rowWindow.firstIndex;
      index <= rowWindow.lastIndex && index < projection.rows.length;
      index += 1
    ) {
      const row = projection.rows[index]
      if (row === undefined) continue
      rows.push({
        index,
        row,
        cells: columns.map((column) =>
          cellFor(column, row, predecessorsByTask, documentIndexByTask),
        ),
      })
    }
  }
  return { columns, rows, rowWindow }
}

/** Column resolution: the active canonical table's field strings mapped
 * onto the supported set, or the documented default set. */
function resolveColumns(projection: ProjectViewProjection): ProjectGridColumn[] {
  const activeColumns = projection.activeTable?.columns
  const sources =
    activeColumns !== undefined && activeColumns.length > 0
      ? activeColumns
      : DEFAULT_TASK_GRID_COLUMNS
  return sources.map((source) => {
    const field = FIELD_BY_COLUMN_NAME.get(source)
    return { field: field ?? 'unsupported', source }
  })
}

/** Dependencies grouped by successor (document order preserved). */
function predecessorsBySuccessor(
  document: ProjectDocument,
): Map<TaskId, ProjectGridPredecessorLink[]> {
  const bySuccessor = new Map<TaskId, ProjectGridPredecessorLink[]>()
  const uidByTask = new Map<TaskId, number>()
  for (const task of document.tasks) uidByTask.set(task.id, task.uid)
  for (const dependency of document.dependencies) {
    const links = bySuccessor.get(dependency.successorId) ?? []
    links.push({
      dependencyId: dependency.id,
      predecessorTaskId: dependency.predecessorId,
      predecessorUid: uidByTask.get(dependency.predecessorId) ?? 0,
      type: dependency.type,
      lagMinutes: dependency.lagMinutes,
    })
    bySuccessor.set(dependency.successorId, links)
  }
  return bySuccessor
}

function cellFor(
  column: ProjectGridColumn,
  row: ProjectTaskRow,
  predecessorsByTask: ReadonlyMap<TaskId, readonly ProjectGridPredecessorLink[]>,
  documentIndexByTask: ReadonlyMap<TaskId, number>,
): ProjectGridCell {
  switch (column.field) {
    case 'rowNumber':
      // Canonical document position (1-based): STABLE under collapse —
      // hidden rows keep their numbers, exactly like the MS Project ID
      // column (never a recomputed visible-sequence number).
      return { kind: 'rowNumber', value: (documentIndexByTask.get(row.taskId) ?? -1) + 1 }
    case 'taskName':
      return {
        kind: 'taskName',
        text: row.name,
        outlineLevel: row.outlineLevel,
        summary: row.summary,
        milestone: row.milestone,
        collapsed: row.collapsed,
      }
    case 'duration':
      // Schedule-first precedence: the derived (rolled-up) duration when a
      // schedule exists, else the canonical task field.
      return { kind: 'duration', minutes: row.schedule?.duration ?? row.duration }
    case 'start':
      return instantCell(row.schedule?.scheduledStart)
    case 'finish':
      return instantCell(row.schedule?.scheduledFinish)
    case 'percentComplete':
      return {
        kind: 'percentComplete',
        value: row.schedule?.percentComplete ?? row.percentComplete,
      }
    case 'predecessors':
      return { kind: 'predecessors', links: predecessorsByTask.get(row.taskId) ?? [] }
    case 'resourceNames':
      return { kind: 'resources', names: row.resourceNames }
    case 'wbs':
      return { kind: 'text', text: row.wbs }
    case 'outlineLevel':
      return { kind: 'number', value: row.outlineLevel }
    case 'priority':
      return { kind: 'number', value: row.priority }
    case 'uid':
      return { kind: 'number', value: row.uid }
    case 'unsupported':
      return { kind: 'unsupported', source: column.source }
    default: {
      // Exhaustiveness guard: every field maps to a cell.
      const exhaustive: never = column.field
      return exhaustive
    }
  }
}

function instantCell(iso: ISODateTime | undefined): ProjectGridCell {
  return iso !== undefined ? { kind: 'instant', iso } : { kind: 'empty' }
}
