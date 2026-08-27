/**
 * PROJECT-021 — the view-state intent reducer.
 *
 * Pure and deterministic: the same `(state, intent, context)` triple always
 * produces the same next state. Hosts dispatch `ProjectViewIntent` values;
 * the reducer validates every entity reference against the live document and
 * ignores (deterministically — the state is returned unchanged) any intent
 * that references an entity that does not exist. Collapse intents carry a
 * stricter rule: `collapsed` contains only SUMMARY task ids (leaves have no
 * subtree to hide), enforced here and in `reconcileViewState`. The
 * reconciled invariants hold after every reduction: selection sets contain
 * only live entity ids, the task-selection `anchorId`/`focusId` are members
 * of the surviving `taskIds` when present (validated against the selection,
 * not mere document existence), the collapsed set contains only live SUMMARY
 * ids, active view references point at live definitions, and the viewport is
 * a well-formed window.
 */
import type {
  DependencyId,
  DerivedSchedule,
  ProjectDocument,
  ResourceId,
  TaskId,
} from '@genoffice/project-contracts'
import type { ProjectViewIntent, SelectMode } from './intents.js'
import {
  type ProjectViewState,
  type TaskSelection,
  parseInstant,
  reconcileViewState,
} from './state.js'
/** Viewport span guards and fit padding are defined once in `./timeline.js`
 * (the module that owns the viewport math) and re-exported here for reducer
 * consumers. */
import {
  MIN_VIEWPORT_SPAN_MS,
  MAX_VIEWPORT_SPAN_MS,
  fitViewport,
  scaleViewport,
} from './timeline.js'

/** The reducer context: the live canonical document and (when a scheduler is
 * wired) the derived schedule of that exact document. */
export interface ViewReducerContext {
  readonly document: ProjectDocument
  readonly schedule?: DerivedSchedule
}

const withTaskSelection = (state: ProjectViewState, tasks: TaskSelection): ProjectViewState => ({
  ...state,
  tasks,
})

const withoutKeys = <T extends object>(value: T): T => {
  const out: Record<string, unknown> = {}
  for (const [key, entry] of Object.entries(value)) {
    if (entry !== undefined) out[key] = entry
  }
  return out as T
}

const selectEntity = <T>(ids: readonly T[], id: T, mode: SelectMode): readonly T[] => {
  const selected = mode === 'toggle' && ids.includes(id)
  if (mode === 'toggle')
    return selected ? ids.filter((candidate) => candidate !== id) : [...ids, id]
  return [id]
}

function selectTask(
  document: ProjectDocument,
  state: ProjectViewState,
  taskId: TaskId,
  mode: SelectMode,
): ProjectViewState {
  const current = state.tasks.taskIds
  if (mode === 'extend') {
    const anchorId = state.tasks.anchorId ?? current[current.length - 1] ?? taskId
    return withTaskSelection(state, {
      taskIds: outlineRange(document, anchorId, taskId),
      anchorId,
      focusId: taskId,
    })
  }
  const next = selectEntity(current, taskId, mode)
  const anchor = mode === 'toggle' && next.includes(taskId) ? taskId : next[next.length - 1]
  return withTaskSelection(
    state,
    withoutKeys({
      taskIds: next,
      anchorId: anchor,
      focusId: next.includes(taskId) ? taskId : next[next.length - 1],
    }),
  )
}

/** The canonical outline-order range [from, to] (inclusive, document order).
 * Unknown endpoints degrade deterministically to exactly the known endpoint(s). */
function outlineRange(document: ProjectDocument, from: TaskId, to: TaskId): readonly TaskId[] {
  const order = document.tasks.map((task) => task.id)
  const fromIndex = order.indexOf(from)
  const toIndex = order.indexOf(to)
  if (fromIndex < 0 && toIndex < 0) return []
  if (fromIndex < 0) return [to]
  if (toIndex < 0) return [from]
  const [lo, hi] = fromIndex <= toIndex ? [fromIndex, toIndex] : [toIndex, fromIndex]
  return order.slice(lo, hi + 1)
}

/**
 * PROJECT-021 — reduce one view intent against a live document context.
 *
 * Unknown entity references make the intent a deterministic no-op (the state
 * object is returned unchanged, reference-equal). The result is always
 * reconciled against the context document.
 */
export function reduceViewState(
  state: ProjectViewState,
  intent: ProjectViewIntent,
  context: ViewReducerContext,
): ProjectViewState {
  const document = context.document
  const exists = (taskId: TaskId): boolean => document.tasks.some((task) => task.id === taskId)
  /** Collapse is a summary-tree operation: the documented state invariant is
   * `collapsed ⊆ summary TaskIds` — a leaf task has no subtree to hide, so
   * collapse intents referencing leaves are deterministic no-ops just like
   * unknown ids. */
  const isSummary = (taskId: TaskId): boolean =>
    document.tasks.some((task) => task.id === taskId && task.summary)

  let next: ProjectViewState
  switch (intent.type) {
    case 'selectTask': {
      if (!exists(intent.taskId)) return state
      next = selectTask(document, state, intent.taskId, intent.mode ?? 'set')
      break
    }
    case 'selectTasks': {
      const valid = [...new Set(intent.taskIds)].filter(exists)
      next = withTaskSelection(
        state,
        valid.length > 0
          ? { taskIds: valid, anchorId: valid[valid.length - 1], focusId: valid[valid.length - 1] }
          : { taskIds: [] },
      )
      break
    }
    case 'clearSelection': {
      next = { ...state, tasks: { taskIds: [] }, dependencies: [], resources: [] }
      break
    }
    case 'selectDependency': {
      if (!document.dependencies.some((dependency) => dependency.id === intent.dependencyId)) {
        return state
      }
      next = {
        ...state,
        dependencies: selectEntity(
          state.dependencies,
          intent.dependencyId as DependencyId,
          intent.mode ?? 'set',
        ),
      }
      break
    }
    case 'selectResource': {
      if (!document.resources.some((resource) => resource.id === intent.resourceId)) return state
      next = {
        ...state,
        resources: selectEntity(
          state.resources,
          intent.resourceId as ResourceId,
          intent.mode ?? 'set',
        ),
      }
      break
    }
    case 'toggleCollapse': {
      // Summary-only: a leaf or unknown id is a deterministic no-op (the
      // state is returned unchanged, reference-equal).
      if (!isSummary(intent.taskId)) return state
      next = {
        ...state,
        collapsed: state.collapsed.includes(intent.taskId)
          ? state.collapsed.filter((id) => id !== intent.taskId)
          : [...state.collapsed, intent.taskId],
      }
      break
    }
    case 'setCollapsed': {
      // Summary-only: leaf ids are ignored (they can never be collapsed;
      // removal requests for them are no-ops by the same invariant).
      const valid = [...new Set(intent.taskIds)].filter(isSummary)
      next = {
        ...state,
        collapsed: intent.collapsed
          ? [...new Set([...state.collapsed, ...valid])]
          : state.collapsed.filter((id) => !valid.includes(id)),
      }
      break
    }
    case 'collapseAll': {
      next = { ...state, collapsed: document.tasks.filter((t) => t.summary).map((t) => t.id) }
      break
    }
    case 'expandAll': {
      next = { ...state, collapsed: [] }
      break
    }
    case 'setViewport': {
      const start = parseInstant(intent.start)
      const finish = parseInstant(intent.finish)
      if (start === undefined || finish === undefined) return state
      const span = finish - start
      if (span < MIN_VIEWPORT_SPAN_MS || span > MAX_VIEWPORT_SPAN_MS) return state
      next = { ...state, viewport: { start: intent.start, finish: intent.finish } }
      break
    }
    case 'scaleViewport': {
      if (!Number.isFinite(intent.factor) || intent.factor <= 0) return state
      next = { ...state, viewport: scaleViewport(state.viewport, intent.factor, intent.focus) }
      break
    }
    case 'fitViewport': {
      next = { ...state, viewport: fitViewport(document, context.schedule) }
      break
    }
    case 'setActiveView': {
      if (
        intent.viewId !== undefined &&
        !document.views.some((view) => view.id === intent.viewId)
      ) {
        return state
      }
      next = { ...withoutKeys({ ...state, activeViewId: intent.viewId }) }
      break
    }
    case 'setActiveTable': {
      if (
        intent.tableId !== undefined &&
        !document.tables.some((table) => table.id === intent.tableId)
      ) {
        return state
      }
      next = { ...withoutKeys({ ...state, activeTableId: intent.tableId }) }
      break
    }
    case 'setActiveFilter': {
      if (
        intent.filterId !== undefined &&
        !document.filters.some((filter) => filter.id === intent.filterId)
      ) {
        return state
      }
      next = { ...withoutKeys({ ...state, activeFilterId: intent.filterId }) }
      break
    }
    case 'setActiveGroup': {
      if (
        intent.groupId !== undefined &&
        !document.groups.some((group) => group.id === intent.groupId)
      ) {
        return state
      }
      next = { ...withoutKeys({ ...state, activeGroupId: intent.groupId }) }
      break
    }
    default: {
      const exhaustive: never = intent
      return exhaustive
    }
  }
  return reconcileViewState(next, document)
}
