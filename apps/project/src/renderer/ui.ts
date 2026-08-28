/**
 * PROJECT-027 — the desktop host DOM rendering layer.
 *
 * Maps the shared renderer-core view models (the pure `ProjectGanttView`:
 * task grid, time axis, bars, milestones, dependency links, calendar
 * surfaces, critical-path floats, resource utilization) onto real DOM.
 * This module computes NO Project semantics: geometry arrives in fraction
 * space and is multiplied by host pixel constants; instants are formatted
 * by string slicing; the canonical `hitTestGantt` inverse resolves pointer
 * hits. Pixel layout (row height, column widths, pane width, scroll) is
 * HOST state by the accepted PROJECT-021/022 rules — never view state.
 */
import type { DependencyId, ProjectDocument, TaskId } from '@genoffice/project-contracts'
import { plainMinutes } from '@genoffice/project-contracts'
import type {
  EditableDependencyField,
  EditableTaskField,
  ProjectGanttView,
  ProjectGridCell,
  ProjectGridRow,
  ProjectViewProjection,
  ProjectViewState,
  SelectMode,
} from '@genoffice/project-renderer-core'
import { hitTestGantt, viewportFraction } from '@genoffice/project-renderer-core'
import type { HostFileFormat } from './document.js'
import type { StatusMessage } from './app.js'

/** Host pixel constants (presentation only — never canonical). */
export const ROW_HEIGHT = 28
export const RESOURCE_ROW_HEIGHT = 26
export const RESOURCE_LABEL_WIDTH = 130
const GRID_HEADER_HEIGHT = 30
const FALLBACK_ROW_CAPACITY = 24

/** Fixed host column widths per grid field (the 022 no-canonical-widths
 * rule: widths are host layout, not document state). */
const COLUMN_WIDTHS: Record<string, number> = {
  rowNumber: 48,
  taskName: 260,
  duration: 96,
  start: 150,
  finish: 150,
  percentComplete: 96,
  predecessors: 130,
  resourceNames: 150,
  wbs: 90,
  outlineLevel: 70,
  priority: 70,
  uid: 60,
}

const COLUMN_LABELS: Record<string, string> = {
  rowNumber: 'ID',
  taskName: 'Task Name',
  duration: 'Duration (min)',
  start: 'Start',
  finish: 'Finish',
  percentComplete: '% Complete',
  predecessors: 'Predecessors',
  resourceNames: 'Resource Names',
  wbs: 'WBS',
  outlineLevel: 'Level',
  priority: 'Priority',
  uid: 'UID',
}

/** The editable-field column mapping (cell activation → edit intents). */
const EDITABLE_FIELD_BY_COLUMN: Record<string, EditableTaskField> = {
  taskName: 'taskName',
  duration: 'duration',
  start: 'start',
  finish: 'finish',
}

export interface UICallbacks {
  onRowClick(taskId: TaskId, mode: SelectMode): void
  onLinkClick(dependencyId: DependencyId, mode: SelectMode): void
  onCellActivate(taskId: TaskId, field: EditableTaskField): void
  onLinkActivate(dependencyId: DependencyId, field: EditableDependencyField): void
  onDraftChange(draft: string): void
  onScroll(firstRow: number): void
  onWidthChange(width: number): void
}

export interface UIUpdateInputs {
  readonly document: ProjectDocument
  readonly projection: ProjectViewProjection
  readonly view: ProjectGanttView
  readonly viewState: ProjectViewState
  readonly filePath: string | null
  readonly format: HostFileFormat
  readonly dirty: boolean
  readonly status: StatusMessage
  readonly diagnostics: readonly { severity: string }[]
  readonly canUndo: boolean
  readonly canRedo: boolean
  readonly editing: boolean
}

export interface UI {
  readonly visibleRowCapacity: number
  update(inputs: UIUpdateInputs): void
}

const el = (tag: string, className?: string): HTMLElement => {
  const node = document.createElement(tag)
  if (className !== undefined) node.className = className
  return node
}

/** Presentation-only instant label: `YYYY-MM-DD HH:mm` (UTC, by slicing). */
const formatInstantLabel = (iso: string): string => `${iso.slice(0, 10)} ${iso.slice(11, 16)}`

const basename = (path: string): string =>
  path.slice(Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\')) + 1)

export function createUI(root: HTMLElement, callbacks: UICallbacks): UI {
  // ---- skeleton (built once) ----------------------------------------------
  root.innerHTML = ''
  const app = el('div', 'gp-app')
  app.dataset.testid = 'project-app'

  const workspace = el('div', 'gp-workspace')

  const gridPane = el('div', 'gp-grid')
  gridPane.dataset.testid = 'task-grid'
  const gridHeader = el('div', 'gp-grid-header')
  const gridScroll = el('div', 'gp-grid-scroll')
  gridScroll.tabIndex = 0
  gridScroll.setAttribute('role', 'grid')
  gridScroll.setAttribute('aria-label', 'Task grid')
  const gridSpacer = el('div', 'gp-grid-spacer')
  gridScroll.appendChild(gridSpacer)
  gridPane.append(gridHeader, gridScroll)

  const timelinePane = el('div', 'gp-timeline')
  timelinePane.dataset.testid = 'timeline'
  const axis = el('div', 'gp-axis')
  axis.dataset.testid = 'time-axis'
  const timelineScroll = el('div', 'gp-timeline-scroll')
  const timelineSpacer = el('div', 'gp-timeline-spacer')
  timelineScroll.appendChild(timelineSpacer)
  timelinePane.append(axis, timelineScroll)

  workspace.append(gridPane, timelinePane)

  const statusbar = el('div', 'gp-statusbar')
  statusbar.dataset.testid = 'statusbar'

  app.append(workspace, statusbar)
  root.appendChild(app)

  // The persistent editor input (moved + refocused per active edit; the
  // draft flows through the view state, never through local input state).
  const editor = document.createElement('input')
  editor.className = 'gp-editor'
  editor.dataset.testid = 'cell-editor'
  editor.setAttribute('aria-label', 'Cell editor')
  editor.spellcheck = false
  editor.addEventListener('input', () => callbacks.onDraftChange(editor.value))

  // ---- state captured across updates ---------------------------------------
  let lastView: ProjectGanttView | null = null
  let timelineWidth = Math.max(320, timelinePane.clientWidth || 900)
  let columnLefts: number[] = []
  let syncingScroll = false

  // ---- pointer translation (mode from the modifier keys) ------------------
  const modeFromEvent = (event: MouseEvent): SelectMode =>
    event.shiftKey ? 'extend' : event.ctrlKey || event.metaKey ? 'toggle' : 'set'

  // ---- scroll synchronization (one window, two panes) ---------------------
  gridScroll.addEventListener('scroll', () => {
    if (syncingScroll) return
    syncingScroll = true
    timelineScroll.scrollTop = gridScroll.scrollTop
    syncingScroll = false
    callbacks.onScroll(Math.max(0, Math.floor(gridScroll.scrollTop / ROW_HEIGHT)))
  })
  timelineScroll.addEventListener('scroll', () => {
    if (syncingScroll) return
    syncingScroll = true
    gridScroll.scrollTop = timelineScroll.scrollTop
    syncingScroll = false
    callbacks.onScroll(Math.max(0, Math.floor(timelineScroll.scrollTop / ROW_HEIGHT)))
  })

  // ---- timeline pointer interaction (the canonical hit-test inverse) ------
  // Links: the double-press is detected across POINTERDOWN events (id-keyed,
  // not element-keyed — a selection re-render replaces the polyline between
  // presses, which would swallow the browser's synthetic dblclick).
  let lastLinkPress: { dependencyId: DependencyId; at: number } | null = null
  timelineScroll.addEventListener('pointerdown', (event) => {
    if (!(event.target instanceof Element) || lastView === null) return
    const link = event.target.closest('[data-dependency-id]')
    if (link !== null) {
      const dependencyId = link.getAttribute('data-dependency-id') as DependencyId
      const now = event.timeStamp
      if (
        lastLinkPress !== null &&
        lastLinkPress.dependencyId === dependencyId &&
        now - lastLinkPress.at < 500
      ) {
        lastLinkPress = null
        callbacks.onLinkActivate(dependencyId, 'lag')
        return
      }
      lastLinkPress = { dependencyId, at: now }
      callbacks.onLinkClick(dependencyId, modeFromEvent(event))
      return
    }
    lastLinkPress = null
    const bounds = timelineScroll.getBoundingClientRect()
    const x = event.clientX - bounds.left + timelineScroll.scrollLeft
    const y = event.clientY - bounds.top + timelineScroll.scrollTop
    const fraction = Math.min(1, Math.max(0, x / timelineWidth))
    const rowIndex = Math.floor(y / ROW_HEIGHT)
    const hit = hitTestGantt(lastView.timeline, { rowIndex, fraction })
    if (hit !== undefined) callbacks.onRowClick(hit.taskId, modeFromEvent(event))
  })

  // ---- grid body interaction ----------------------------------------------
  gridScroll.addEventListener('pointerdown', (event) => {
    if (!(event.target instanceof Element)) return
    const cell = event.target.closest('[data-task-id]')
    if (cell === null) return
    callbacks.onRowClick(cell.getAttribute('data-task-id') as TaskId, modeFromEvent(event))
  })
  gridScroll.addEventListener('dblclick', (event) => {
    if (!(event.target instanceof Element)) return
    const cell = event.target.closest('[data-column]')
    if (cell === null) return
    const taskId = cell.getAttribute('data-task-id')
    const column = cell.getAttribute('data-column')
    if (taskId === null || column === null) return
    const field = EDITABLE_FIELD_BY_COLUMN[column]
    if (field === undefined) return
    callbacks.onCellActivate(taskId as TaskId, field)
  })

  // ---- width observation (host layout) -------------------------------------
  const resizeObserver =
    typeof ResizeObserver !== 'undefined'
      ? new ResizeObserver((entries) => {
          const width = entries[0]?.contentRect.width
          if (width !== undefined && Math.abs(width - timelineWidth) >= 1) {
            timelineWidth = Math.max(320, Math.floor(width))
            callbacks.onWidthChange(timelineWidth)
          }
        })
      : null
  if (resizeObserver !== null) resizeObserver.observe(timelinePane)

  // ---- cell rendering -------------------------------------------------------

  const cellText = (cell: ProjectGridCell): string => {
    switch (cell.kind) {
      case 'rowNumber':
        return String(cell.value)
      case 'taskName':
        return cell.text
      case 'duration':
        return String(plainMinutes(cell.minutes))
      case 'instant':
        return cell.iso === '' ? '' : formatInstantLabel(cell.iso)
      case 'percentComplete':
        return `${cell.value}%`
      case 'predecessors':
        return cell.links
          .map(
            (link) =>
              `${link.predecessorUid}${link.type}${
                link.lagMinutes === 0
                  ? ''
                  : link.lagMinutes > 0
                    ? `+${link.lagMinutes}`
                    : String(link.lagMinutes)
              }`,
          )
          .join(', ')
      case 'resources':
        return cell.names.join(', ')
      case 'text':
        return cell.text
      case 'number':
        return String(cell.value)
      default:
        return ''
    }
  }

  /**
   * Grid rows are KEYED and updated IN PLACE (`taskId` → element): a
   * selection re-render must not replace the element under an in-flight
   * double click — replacing it between pointerdown and pointerup swallows
   * the browser's synthetic click/dblclick, making cell activation
   * impossible. Rows leaving the virtual window are detached; rows coming
   * back are re-attached.
   */
  const rowEls = new Map<TaskId, HTMLElement>()
  const emptyState = el('div', 'gp-empty')
  emptyState.dataset.testid = 'empty-state'
  emptyState.textContent = 'No tasks — press Insert or Task ▸ New Task'

  function createGridRow(
    gridRow: ProjectGridRow,
    columns: readonly { field: string; source: string }[],
  ): HTMLElement {
    const row = gridRow.row
    const rowEl = el('div', 'gp-grid-row')
    rowEl.dataset.testid = 'task-row'
    rowEl.dataset.taskId = row.taskId
    rowEl.setAttribute('role', 'row')
    gridRow.cells.forEach((cell, columnIndex) => {
      const column = columns[columnIndex]
      const field = column?.field ?? 'unsupported'
      const width = COLUMN_WIDTHS[field] ?? 140
      const cellEl = el('div', 'gp-grid-cell')
      cellEl.dataset.column = field === 'unsupported' ? (column?.source ?? 'unsupported') : field
      cellEl.dataset.taskId = row.taskId
      cellEl.setAttribute('role', 'gridcell')
      if (cell.kind === 'taskName') {
        cellEl.classList.add('gp-cell-name')
        cellEl.appendChild(el('span', 'gp-name-marker'))
        cellEl.appendChild(el('span', 'gp-name-text'))
      } else if (cell.kind === 'duration' || cell.kind === 'instant') {
        cellEl.classList.add('gp-cell-schedule')
      }
      cellEl.style.left = `${columnLefts[columnIndex] ?? 0}px`
      cellEl.style.width = `${width}px`
      rowEl.appendChild(cellEl)
    })
    updateGridRow(rowEl, gridRow)
    return rowEl
  }

  function updateGridRow(rowEl: HTMLElement, gridRow: ProjectGridRow): void {
    const row = gridRow.row
    rowEl.style.top = `${gridRow.index * ROW_HEIGHT}px`
    rowEl.style.height = `${ROW_HEIGHT}px`
    rowEl.dataset.rowIndex = String(gridRow.index)
    rowEl.dataset.selected = String(row.selected)
    rowEl.dataset.focused = String(row.focused)
    rowEl.dataset.summary = String(row.summary)
    rowEl.setAttribute('aria-selected', String(row.selected))
    const cells = [...rowEl.children]
    gridRow.cells.forEach((cell, columnIndex) => {
      const cellEl = cells[columnIndex]
      if (!(cellEl instanceof HTMLElement)) return
      if (cell.kind === 'taskName') {
        cellEl.style.paddingLeft = `${(cell.outlineLevel - 1) * 14 + 8}px`
        cellEl.classList.toggle('gp-cell-summary', cell.summary)
        cellEl.classList.toggle('gp-cell-milestone', cell.milestone)
        cellEl.classList.toggle('gp-cell-collapsed', cell.collapsed)
        const marker = cellEl.querySelector('.gp-name-marker')
        if (marker !== null) {
          marker.textContent = cell.summary
            ? cell.collapsed
              ? '▸'
              : '▾'
            : cell.milestone
              ? '◆'
              : '·'
        }
        const text = cellEl.querySelector('.gp-name-text')
        if (text !== null) text.textContent = cellText(cell)
        return
      }
      cellEl.classList.toggle(
        'gp-cell-schedule',
        cell.kind === 'duration' || cell.kind === 'instant',
      )
      cellEl.textContent = cellText(cell)
    })
  }

  // ---- the persistent editor overlay ----------------------------------------
  function positionEditor(
    inputs: UIUpdateInputs,
    projection: ProjectViewProjection,
    columns: readonly { field: string; source: string }[],
  ): void {
    const editing = inputs.viewState.editing
    const dependencyEditing = inputs.viewState.dependencyEditing
    if (editing !== undefined) {
      const rowIndex = projection.rows.findIndex((row) => row.taskId === editing.taskId)
      const columnIndex = columns.findIndex((column) => column.field === editing.field)
      if (rowIndex >= 0 && columnIndex >= 0) {
        const width = COLUMN_WIDTHS[editing.field] ?? 140
        editor.style.top = `${rowIndex * ROW_HEIGHT}px`
        editor.style.left = `${columnLefts[columnIndex] ?? 0}px`
        editor.style.width = `${width}px`
        editor.style.height = `${ROW_HEIGHT - 2}px`
        editor.dataset.taskId = editing.taskId
        editor.dataset.field = editing.field
        delete editor.dataset.dependencyId
        if (editor.parentElement !== gridSpacer) gridSpacer.appendChild(editor)
        if (editor.value !== editing.draft) editor.value = editing.draft
        if (document.activeElement !== editor) {
          editor.focus()
          editor.setSelectionRange(editor.value.length, editor.value.length)
        }
        return
      }
    }
    if (dependencyEditing !== undefined && lastView !== null) {
      const link = lastView.timeline.links.find(
        (candidate) => candidate.dependencyId === dependencyEditing.dependencyId,
      )
      if (link !== undefined) {
        const point = link.route[0]
        editor.style.top = `${(point?.rowIndex ?? 0) * ROW_HEIGHT - 2}px`
        editor.style.left = `${(point?.fraction ?? 0) * timelineWidth + 8}px`
        editor.style.width = '80px'
        editor.style.height = '22px'
        delete editor.dataset.taskId
        editor.dataset.dependencyId = dependencyEditing.dependencyId
        editor.dataset.field = dependencyEditing.field
        if (editor.parentElement !== timelineSpacer) timelineSpacer.appendChild(editor)
        if (editor.value !== dependencyEditing.draft) editor.value = dependencyEditing.draft
        if (document.activeElement !== editor) {
          editor.focus()
          editor.setSelectionRange(editor.value.length, editor.value.length)
        }
        return
      }
    }
    // No active edit: retire the overlay and restore grid focus.
    editor.remove()
    if (document.activeElement === document.body) gridScroll.focus()
  }

  // ---- update ----------------------------------------------------------------

  function update(inputs: UIUpdateInputs): void {
    const view = inputs.view
    const rowCount = inputs.projection.rows.length

    // ---- grid header ----
    gridHeader.style.height = `${GRID_HEADER_HEIGHT}px`
    gridHeader.innerHTML = ''
    columnLefts = []
    let headerX = 0
    for (const column of view.taskGrid.columns) {
      const width = COLUMN_WIDTHS[column.field] ?? 140
      const headerCell = el('div', 'gp-grid-header-cell')
      headerCell.textContent = COLUMN_LABELS[column.field] ?? column.source
      headerCell.style.left = `${headerX}px`
      headerCell.style.width = `${width}px`
      headerCell.dataset.column = column.field === 'unsupported' ? column.source : column.field
      gridHeader.appendChild(headerCell)
      columnLefts.push(headerX)
      headerX += width
    }
    gridPane.style.width = `${headerX}px`

    // ---- grid rows (virtualized window, keyed in place) ----
    gridSpacer.style.height = `${rowCount * ROW_HEIGHT}px`
    const wantedIds = new Set(view.taskGrid.rows.map((gridRow) => gridRow.row.taskId))
    for (const [taskId, rowEl] of rowEls) {
      if (!wantedIds.has(taskId)) {
        rowEl.remove()
        rowEls.delete(taskId)
      }
    }
    if (rowCount === 0) {
      if (emptyState.parentElement !== gridSpacer) gridSpacer.appendChild(emptyState)
    } else {
      emptyState.remove()
    }
    for (const gridRow of view.taskGrid.rows) {
      const taskId = gridRow.row.taskId
      let rowEl = rowEls.get(taskId)
      if (rowEl === undefined) {
        rowEl = createGridRow(gridRow, view.taskGrid.columns)
        rowEls.set(taskId, rowEl)
        gridSpacer.appendChild(rowEl)
      } else {
        if (rowEl.parentElement !== gridSpacer) gridSpacer.appendChild(rowEl)
        updateGridRow(rowEl, gridRow)
      }
    }

    // ---- time axis ----
    axis.style.height = `${GRID_HEADER_HEIGHT}px`
    axis.dataset.axisLevel = view.timeline.axisLevel
    axis.innerHTML = ''
    for (const band of view.timeline.bands) {
      const left = viewportFraction(view.timeline.viewport, band.start) * timelineWidth
      const right = viewportFraction(view.timeline.viewport, band.finish) * timelineWidth
      const bandEl = el('div', 'gp-axis-band')
      bandEl.style.left = `${left}px`
      bandEl.style.width = `${Math.max(1, right - left)}px`
      bandEl.dataset.bandStart = band.start
      bandEl.dataset.bandFinish = band.finish
      bandEl.textContent = band.level === 'month' ? band.start.slice(0, 7) : band.start.slice(0, 10)
      axis.appendChild(bandEl)
    }

    // ---- timeline body ----
    const resources = view.timeline.resourceUtilization?.resources ?? []
    const resourceStripHeight =
      resources.length > 0 ? resources.length * RESOURCE_ROW_HEIGHT + 4 : 0
    const bodyHeight = rowCount * ROW_HEIGHT + resourceStripHeight
    timelineSpacer.style.height = `${bodyHeight}px`
    timelineSpacer.style.width = `${timelineWidth}px`
    timelineSpacer.innerHTML = ''
    lastView = view

    // calendar background surface (PROJECT-025)
    if (view.timeline.calendar?.bands !== undefined) {
      for (const band of view.timeline.calendar.bands) {
        const left = viewportFraction(view.timeline.viewport, band.start) * timelineWidth
        const right = viewportFraction(view.timeline.viewport, band.finish) * timelineWidth
        const bandEl = el('div', 'gp-calendar-band')
        bandEl.style.left = `${left}px`
        bandEl.style.width = `${Math.max(0, right - left)}px`
        bandEl.style.height = `${bodyHeight}px`
        bandEl.dataset.testid = 'calendar-band'
        bandEl.dataset.working = String(band.working)
        timelineSpacer.appendChild(bandEl)
      }
    }

    // per-row calendar surfaces (PROJECT-025) — non-working shading only
    for (const rowCalendar of view.timeline.rowCalendars ?? []) {
      if (rowCalendar.surface.bands === undefined) continue
      for (const band of rowCalendar.surface.bands) {
        if (band.working) continue
        const left = viewportFraction(view.timeline.viewport, band.start) * timelineWidth
        const right = viewportFraction(view.timeline.viewport, band.finish) * timelineWidth
        const bandEl = el('div', 'gp-row-calendar-band')
        bandEl.style.left = `${left}px`
        bandEl.style.width = `${Math.max(0, right - left)}px`
        bandEl.style.top = `${rowCalendar.rowIndex * ROW_HEIGHT}px`
        bandEl.style.height = `${ROW_HEIGHT}px`
        bandEl.dataset.testid = 'row-calendar-band'
        bandEl.dataset.taskId = rowCalendar.taskId
        timelineSpacer.appendChild(bandEl)
      }
    }

    // critical floats lookup (PROJECT-026 verbatim echo)
    const floatByTask = new Map(
      (view.timeline.criticalPath?.floats ?? []).map((float) => [float.taskId, float]),
    )
    const criticalLinks = new Set(view.timeline.criticalPath?.criticalDependencyIds ?? [])

    // dependency links (PROJECT-022/024) — SVG over the rows
    if (view.timeline.links.length > 0) {
      const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg')
      svg.setAttribute('class', 'gp-links')
      svg.setAttribute('width', String(timelineWidth))
      svg.setAttribute('height', String(rowCount * ROW_HEIGHT))
      for (const link of view.timeline.links) {
        const points = link.route
          .map(
            (point) =>
              `${point.fraction * timelineWidth},${point.rowIndex * ROW_HEIGHT + ROW_HEIGHT / 2}`,
          )
          .join(' ')
        const hit = document.createElementNS('http://www.w3.org/2000/svg', 'polyline')
        hit.setAttribute('points', points)
        hit.setAttribute('class', 'gp-link-hit')
        hit.dataset.testid = 'dependency-link'
        hit.dataset.dependencyId = link.dependencyId
        hit.dataset.linkType = link.type
        hit.dataset.selected = String(link.selected)
        hit.dataset.critical = String(criticalLinks.has(link.dependencyId))
        const visible = document.createElementNS('http://www.w3.org/2000/svg', 'polyline')
        visible.setAttribute('points', points)
        visible.setAttribute('class', link.selected ? 'gp-link gp-link-selected' : 'gp-link')
        if (criticalLinks.has(link.dependencyId)) visible.classList.add('gp-link-critical')
        visible.dataset.dependencyId = link.dependencyId
        svg.append(hit, visible)
      }
      timelineSpacer.appendChild(svg)
    }

    // slack bars (PROJECT-026 — the authority's positive-float geometry)
    for (const float of view.timeline.criticalPath?.floats ?? []) {
      if (float.slack === undefined) continue
      const left = float.slack.startFraction * timelineWidth
      const right = float.slack.finishFraction * timelineWidth
      const slackEl = el('div', 'gp-slack-bar')
      slackEl.dataset.testid = 'slack-bar'
      slackEl.dataset.taskId = float.taskId
      slackEl.dataset.totalSlack = String(float.totalSlack)
      slackEl.style.left = `${left}px`
      slackEl.style.width = `${Math.max(0, right - left)}px`
      slackEl.style.top = `${float.rowIndex * ROW_HEIGHT + ROW_HEIGHT / 2 - 2}px`
      timelineSpacer.appendChild(slackEl)
    }

    // bars (PROJECT-022)
    for (const bar of view.timeline.bars) {
      const float = floatByTask.get(bar.taskId)
      const left = bar.startFraction * timelineWidth
      const right = bar.finishFraction * timelineWidth
      const barEl = el('div', bar.kind === 'summary' ? 'gp-bar gp-bar-summary' : 'gp-bar')
      barEl.dataset.testid = 'gantt-bar'
      barEl.dataset.taskId = bar.taskId
      barEl.dataset.kind = bar.kind
      barEl.dataset.critical = String(float?.critical ?? false)
      if (bar.startsBefore) barEl.dataset.startsBefore = 'true'
      if (bar.finishesAfter) barEl.dataset.finishesAfter = 'true'
      barEl.style.left = `${left}px`
      barEl.style.width = `${Math.max(2, right - left)}px`
      barEl.style.top = `${bar.rowIndex * ROW_HEIGHT + 6}px`
      barEl.style.height = `${ROW_HEIGHT - 12}px`
      if (bar.progressFraction > bar.startFraction) {
        const progress = el('div', 'gp-bar-progress')
        progress.style.width = `${Math.min(1, bar.progressFraction - bar.startFraction) * 100}%`
        barEl.appendChild(progress)
      }
      timelineSpacer.appendChild(barEl)
    }

    // milestones (PROJECT-022)
    for (const milestone of view.timeline.milestones) {
      const milestoneEl = el('div', 'gp-milestone')
      milestoneEl.dataset.testid = 'gantt-milestone'
      milestoneEl.dataset.taskId = milestone.taskId
      milestoneEl.dataset.critical = String(floatByTask.get(milestone.taskId)?.critical ?? false)
      milestoneEl.style.left = `${milestone.fraction * timelineWidth - 6}px`
      milestoneEl.style.top = `${milestone.rowIndex * ROW_HEIGHT + (ROW_HEIGHT - 12) / 2}px`
      timelineSpacer.appendChild(milestoneEl)
    }

    // resource utilization strip (PROJECT-026)
    let resourceTop = rowCount * ROW_HEIGHT
    for (const resource of resources) {
      const label = el('div', 'gp-resource-label')
      label.dataset.testid = 'resource-row'
      label.dataset.resourceId = resource.resourceId
      label.textContent = resource.name ?? resource.resourceId
      label.style.top = `${resourceTop}px`
      label.style.width = `${RESOURCE_LABEL_WIDTH}px`
      timelineSpacer.appendChild(label)
      for (const band of resource.bands) {
        const left = viewportFraction(view.timeline.viewport, band.start) * timelineWidth
        const right = viewportFraction(view.timeline.viewport, band.finish) * timelineWidth
        const bandEl = el('div', 'gp-resource-band')
        bandEl.dataset.testid = 'resource-band'
        bandEl.dataset.resourceId = band.resourceId
        bandEl.dataset.overallocated = String(band.overallocated)
        bandEl.dataset.demandUnits = String(band.demandUnits)
        bandEl.dataset.capacityUnits = String(band.capacityUnits)
        bandEl.style.left = `${left + RESOURCE_LABEL_WIDTH}px`
        bandEl.style.width = `${Math.max(1, right - left)}px`
        bandEl.style.top = `${resourceTop}px`
        bandEl.title = `${band.demandUnits}/${band.capacityUnits}`
        timelineSpacer.appendChild(bandEl)
      }
      resourceTop += RESOURCE_ROW_HEIGHT
    }

    // ---- editor overlay ----
    positionEditor(inputs, inputs.projection, view.taskGrid.columns)

    // ---- statusbar ----
    statusbar.innerHTML = ''
    const file = el('span', 'gp-status-file')
    file.dataset.testid = 'file-label'
    file.textContent = inputs.filePath === null ? 'Untitled' : basename(inputs.filePath)
    const dirtyEl = el('span', 'gp-status-dirty')
    dirtyEl.dataset.testid = 'dirty-indicator'
    dirtyEl.dataset.dirty = String(inputs.dirty)
    dirtyEl.textContent = inputs.dirty ? '●' : ''
    const formatEl = el('span', 'gp-status-format')
    formatEl.dataset.testid = 'format-label'
    formatEl.textContent = inputs.format.toUpperCase()
    const statusEl = el(
      'span',
      inputs.status.kind === 'error' ? 'gp-status-text gp-status-error' : 'gp-status-text',
    )
    statusEl.dataset.testid = 'status-text'
    statusEl.textContent = inputs.status.text
    const diagEl = el('span', 'gp-status-diag')
    diagEl.dataset.testid = 'diagnostics-label'
    diagEl.textContent = inputs.diagnostics.length > 0 ? `⚠ ${inputs.diagnostics.length}` : ''
    const historyEl = el('span', 'gp-status-history')
    historyEl.dataset.testid = 'history-label'
    historyEl.dataset.canUndo = String(inputs.canUndo)
    historyEl.dataset.canRedo = String(inputs.canRedo)
    historyEl.textContent = `${inputs.canUndo ? 'undo' : ''}${
      inputs.canUndo && inputs.canRedo ? '/' : ''
    }${inputs.canRedo ? 'redo' : ''}`
    const countEl = el('span', 'gp-status-count')
    countEl.dataset.testid = 'row-count'
    countEl.textContent = `${rowCount} ${rowCount === 1 ? 'task' : 'tasks'}`
    statusbar.append(file, dirtyEl, formatEl, statusEl, diagEl, historyEl, countEl)
  }

  return {
    get visibleRowCapacity(): number {
      const height = gridScroll.clientHeight
      return Math.max(
        1,
        Math.floor((height > 0 ? height : FALLBACK_ROW_CAPACITY * ROW_HEIGHT) / ROW_HEIGHT),
      )
    },
    update,
  }
}
