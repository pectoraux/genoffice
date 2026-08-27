/**
 * PROJECT-021 — shared renderer view state.
 *
 * The view state is the host-independent INTERACTION state of the Project
 * renderer: what is selected, which summary tasks are collapsed, which part
 * of the timeline is visible, and which canonical view/table/filter/group
 * definitions are active. Desktop (Electron) and web hosts drive the exact
 * same state shape through the same intent reducer (`./reduce.js`), so the
 * two hosts cannot drift apart (R-009). The collapsed set carries the
 * invariant `collapsed ⊆ summary TaskIds` (collapse is a summary-tree
 * operation), enforced by the reducer and by `reconcileViewState`.
 *
 * Architecture-lock §11 (renderer rule): renderer state may cache projections
 * of canonical state but may NOT own authoritative task dates, dependencies,
 * critical path, float, leveling results, or persisted Project semantics.
 * Enforced structurally: every field below is either an entity-ID reference
 * (validated against the live document by `reconcileViewState`), a canonical
 * view-definition reference, or a timeline viewport instant. No scheduling
 * value is ever stored here — schedule values are joined at projection time
 * (`./projection.js`) from the `DerivedSchedule` the scheduling authority
 * produced.
 *
 * The state is plain JSON data by construction (hosts may persist it in their
 * own workspace preferences; it is never written into `.gproj` — the
 * PROJECT-014 rule that the adapter does not persist renderer state).
 */
import type {
  DependencyId,
  ProjectFilterId,
  ProjectGroupId,
  ProjectTableId,
  ProjectViewId,
  ResourceId,
  TaskId,
  ProjectDocument,
  DerivedSchedule,
} from '@genoffice/project-contracts'

/**
 * Task selection. `taskIds` is the ordered selection (first-occurrence order
 * preserved by the intents that build it; never a set-like reordering — the
 * reducer keeps insertion order deterministically). `anchorId` is the
 * shift-extend anchor and `focusId` the most recently focused task (keyboard
 * navigation, PROJECT-023). Both are optional and always members of
 * `taskIds` when present — an invariant the reducer maintains by construction
 * and `reconcileViewState` enforces against the SURVIVING selection (a task
 * that still exists in the document but is not selected can never be the
 * anchor or the focus).
 */
export interface TaskSelection {
  readonly taskIds: readonly TaskId[]
  readonly anchorId?: TaskId
  readonly focusId?: TaskId
}

/**
 * The visible timeline window as ISO-8601 UTC instants. The window is view
 * state (which slice of the schedule the user is looking at), never task
 * state: the instants come from user navigation intents, and the projection
 * maps schedule instants into the window (`./timeline.js`). Invariants kept
 * by the reducer: `start < finish`, both parseable as UTC instants.
 */
export interface TimelineViewport {
  readonly start: string
  readonly finish: string
}

/**
 * The shared Project renderer view state (PROJECT-021). One immutable value
 * per interaction state; the reducer produces new values, hosts never mutate
 * it. Selection sets are typed per entity kind so a dependency selection can
 * never masquerade as a task selection (identity is branded in contracts).
 */
export interface ProjectViewState {
  /** Ordered selected task ids (subset of the document's tasks, always). */
  readonly tasks: TaskSelection
  /** Selected dependency ids (subset of the document's dependencies). */
  readonly dependencies: readonly DependencyId[]
  /** Selected resource ids (subset of the document's resources). */
  readonly resources: readonly ResourceId[]
  /** Collapsed SUMMARY-task ids. The documented invariant is
   * `collapsed ⊆ summary TaskIds`: a leaf has no subtree to hide, so it can
   * never be collapsed. Enforced by the reducer (leaf collapse intents are
   * deterministic no-ops) and by `reconcileViewState` (which also prunes
   * entries whose task became a leaf — e.g. its subtree was deleted and the
   * engine recomputed `summary` — or entries restored from persisted host
   * state), so the invariant holds after every reduction AND after external
   * state restoration. (Subset of the document's tasks, always.) */
  readonly collapsed: readonly TaskId[]
  /** The visible timeline window. */
  readonly viewport: TimelineViewport
  /** Active canonical view definition (document-declared `ProjectView`). */
  readonly activeViewId?: ProjectViewId
  /** Active canonical table definition (document-declared `ProjectTable`). */
  readonly activeTableId?: ProjectTableId
  /** Active canonical filter definition (document-declared `ProjectFilter`). */
  readonly activeFilterId?: ProjectFilterId
  /** Active canonical group definition (document-declared `ProjectGroup`). */
  readonly activeGroupId?: ProjectGroupId
}

/** The default viewport when no scheduling information exists yet: a 30-day
 * window starting at the project's canonical start date. Deterministic. */
const DEFAULT_VIEWPORT_DAYS = 30

const DAY_MS = 24 * 60 * 60 * 1000

/** Parses an ISO-8601 instant to epoch milliseconds; `undefined` when the
 * value is not a parseable UTC instant. */
export function parseInstant(value: string): number | undefined {
  if (typeof value !== 'string' || value.length === 0) return undefined
  const ms = Date.parse(value)
  return Number.isFinite(ms) ? ms : undefined
}

/** Formats epoch milliseconds as a canonical ISO-8601 UTC instant. */
export function formatInstant(ms: number): string {
  return new Date(ms).toISOString()
}

/**
 * The initial view state for a document: nothing selected, nothing collapsed,
 * the viewport fitted to the project window (schedule span when a schedule is
 * available, otherwise the canonical properties start date ± the default
 * window). Pure and deterministic: the same document + schedule always
 * produce the same initial state.
 */
export function createViewState(
  document: ProjectDocument,
  schedule?: DerivedSchedule,
): ProjectViewState {
  return {
    tasks: { taskIds: [] },
    dependencies: [],
    resources: [],
    collapsed: [],
    viewport: initialViewport(document, schedule),
  }
}

function initialViewport(document: ProjectDocument, schedule?: DerivedSchedule): TimelineViewport {
  const startMs =
    (schedule?.projectStart !== undefined
      ? parseInstant(schedule.projectStart)
      : parseInstant(document.properties.startDate)) ?? 0
  const finishMsRaw =
    schedule?.projectFinish !== undefined
      ? parseInstant(schedule.projectFinish)
      : document.properties.finishDate !== undefined
        ? parseInstant(document.properties.finishDate)
        : undefined
  const finishMs =
    finishMsRaw !== undefined && finishMsRaw > startMs
      ? finishMsRaw
      : startMs + DEFAULT_VIEWPORT_DAYS * DAY_MS
  return {
    start: formatInstant(startMs),
    finish: formatInstant(Math.max(finishMs, startMs + DAY_MS)),
  }
}

/**
 * Reconciles a view state against a document: every entity reference that no
 * longer exists is dropped (deterministically, preserving the surviving
 * order); collapse entries whose task is no longer a SUMMARY (deleted
 * subtree — the engine recomputes `summary` — or a leaf id from restored
 * host state) are pruned the same way, keeping the invariant
 * `collapsed ⊆ summary TaskIds`; the task-selection `anchorId`/`focusId` are
 * validated against the SURVIVING `tasks.taskIds` (mere existence in the
 * document is not enough — a live-but-unselected anchor/focus, e.g. from a
 * malformed externally restored state, is dropped, keeping the documented
 * `TaskSelection` invariant that both are selection members when present);
 * active view-definition references that vanished are cleared. The viewport
 * is time, not an entity reference, and is left untouched.
 *
 * Hosts call this after ANY document replacement (session command, undo,
 * redo, file load) so the cached projections never reference dead entities —
 * the `reduceViewState` reducer applies it automatically after every intent.
 * Hosts restoring a persisted view state MUST run it once before first use
 * (it is the security net for states produced outside the reducer).
 */
export function reconcileViewState(
  state: ProjectViewState,
  document: ProjectDocument,
): ProjectViewState {
  const taskIds = new Set(document.tasks.map((task) => task.id))
  const summaryIds = new Set(document.tasks.filter((task) => task.summary).map((task) => task.id))
  const dependencyIds = new Set(document.dependencies.map((dependency) => dependency.id))
  const resourceIds = new Set(document.resources.map((resource) => resource.id))

  const taskSelection: TaskSelection = {
    taskIds: state.tasks.taskIds.filter((id) => taskIds.has(id)),
  }
  // Anchor/focus must be members of the SURVIVING selection, not merely
  // live tasks: a restored/malformed state can carry an anchor that exists
  // in the document while not being selected.
  const selectedTaskIds = new Set(taskSelection.taskIds)
  const anchorId =
    state.tasks.anchorId !== undefined && selectedTaskIds.has(state.tasks.anchorId)
      ? state.tasks.anchorId
      : undefined
  const focusId =
    state.tasks.focusId !== undefined && selectedTaskIds.has(state.tasks.focusId)
      ? state.tasks.focusId
      : undefined

  return {
    tasks: {
      taskIds: taskSelection.taskIds,
      ...(anchorId !== undefined ? { anchorId } : {}),
      ...(focusId !== undefined ? { focusId } : {}),
    },
    dependencies: state.dependencies.filter((id) => dependencyIds.has(id)),
    resources: state.resources.filter((id) => resourceIds.has(id)),
    collapsed: state.collapsed.filter((id) => summaryIds.has(id)),
    viewport: state.viewport,
    ...(state.activeViewId !== undefined &&
    document.views.some((view) => view.id === state.activeViewId)
      ? { activeViewId: state.activeViewId }
      : {}),
    ...(state.activeTableId !== undefined &&
    document.tables.some((table) => table.id === state.activeTableId)
      ? { activeTableId: state.activeTableId }
      : {}),
    ...(state.activeFilterId !== undefined &&
    document.filters.some((filter) => filter.id === state.activeFilterId)
      ? { activeFilterId: state.activeFilterId }
      : {}),
    ...(state.activeGroupId !== undefined &&
    document.groups.some((group) => group.id === state.activeGroupId)
      ? { activeGroupId: state.activeGroupId }
      : {}),
  }
}
