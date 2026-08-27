import { describe, expect, it } from 'vitest'
import pkgRaw from '../package.json?raw'
import workflow from '../../.github/workflows/project-foundation.yml?raw'
import requirements from '../../../spec/project/requirements.md?raw'
import workItems from '../../../spec/project/work-items.md?raw'
import dependencyGraph from '../../../spec/project/dependency-graph.md?raw'
import verificationMatrix from '../../../spec/project/verification-matrix.md?raw'
import architectureLock from '../../../spec/project/architecture-lock.md?raw'
import acr from '../../../spec/project/architecture-changes/ACR-001-project-file-adapter-boundary.md?raw'

/**
 * PROJECT-021/022 — architecture discipline guards.
 *
 * Static source-level checks mirroring (and exceeding) the CI foundation
 * boundary grep: the renderer core imports ONLY the contracts and engine
 * packages, never React/Electron/Node/browser/HTTP (architecture-lock §13),
 * never the scheduling package (scheduling authority — lock §3/§6; the
 * scheduler is injected as a structural runner), and never the file/host
 * packages (host-specific transport stays outside the renderer — R-009).
 * Determinism guards: no wall clock, no randomness, no locale comparisons.
 * This suite itself uses ONLY vitest + `?raw` module sources — no `node:`
 * imports, satisfying the same CI boundary grep as every foundation package.
 *
 * Roadmap reconciliation increment — extended with the clarified §13 rule
 * (ACR-001): the renderer MUST NOT import project-file, MSPDI parser
 * internals, or MPP parser internals — the renderer must not acquire
 * file-format knowledge (file-format implementations remain behind the
 * project-file adapter boundary).
 */
const srcModules = import.meta.glob('../src/**/*.ts', {
  query: '?raw',
  import: 'default',
  eager: true,
}) as Record<string, string>

const srcFiles = Object.entries(srcModules)
const pkg = JSON.parse(pkgRaw) as {
  dependencies?: Record<string, string>
  devDependencies?: Record<string, string>
}

const FORBIDDEN_IMPORT_PATTERNS: ReadonlyArray<{ pattern: RegExp; label: string }> = [
  { pattern: /from ['"]react(?:-dom)?['"]/, label: 'React' },
  { pattern: /from ['"]electron['"]/, label: 'Electron' },
  { pattern: /from ['"]node:/, label: 'Node built-ins' },
  { pattern: /from ['"]https?['"]/, label: 'HTTP(S) client' },
  { pattern: /from ['"]@genoffice\/project-scheduling['"]/, label: 'the scheduling package' },
  {
    pattern: /from ['"]@genoffice\/project-file(?:\/[^'"]*)?['"]/,
    label: 'the file adapter package (public surface or internals)',
  },
  { pattern: /from ['"]@genoffice\/project-mpp-host['"]/, label: 'the MPP host package' },
  {
    pattern: /from ['"][^'"]*(?:mspdi|gproj)[^'"]*['"]/i,
    label: 'MSPDI/.gproj file-format parser internals',
  },
  {
    pattern: /from ['"][^'"]*\bmpp\b[^'"]*['"]/i,
    label: 'MPP file-format parser internals',
  },
  {
    pattern: /from ['"]fs['"]|from ['"]path['"]|from ['"]child_process['"]/,
    label: 'Node core modules',
  },
]

describe('PROJECT-021 architecture — package boundaries', () => {
  it('ships exactly the contracts + engine runtime dependencies', () => {
    expect(Object.keys(pkg.dependencies ?? {}).sort()).toEqual([
      '@genoffice/project-contracts',
      '@genoffice/project-engine',
    ])
    expect(Object.keys(pkg.devDependencies ?? {}).sort()).toEqual(['typescript', 'vitest'])
  })

  it('never imports React/Electron/Node/HTTP/scheduling/file/host packages or file-format parser internals in src', () => {
    expect(srcFiles.length).toBeGreaterThan(0)
    for (const [file, source] of srcFiles) {
      for (const { pattern, label } of FORBIDDEN_IMPORT_PATTERNS) {
        if (pattern.test(source)) {
          throw new Error(`${file} imports ${label}`)
        }
      }
    }
  })

  it('contains no wall clock, randomness, or locale comparisons in src', () => {
    for (const [file, source] of srcFiles) {
      expect(source, `${file} uses Date.now`).not.toContain('Date.now(')
      expect(source, `${file} uses Math.random`).not.toContain('Math.random(')
      expect(source, `${file} uses localeCompare`).not.toContain('localeCompare')
    }
  })

  it('keeps the view state free of scheduling-derived values (lock §11)', () => {
    const stateSource = srcModules['../src/state.ts'] ?? ''
    const intentsSource = srcModules['../src/intents.ts'] ?? ''
    const editingSource = srcModules['../src/editing.ts'] ?? ''
    const editFlowSource = srcModules['../src/edit-flow.ts'] ?? ''
    const dependencyEditingSource = srcModules['../src/dependency-editing.ts'] ?? ''
    for (const source of [
      stateSource,
      intentsSource,
      editingSource,
      editFlowSource,
      dependencyEditingSource,
    ]) {
      // No scheduling-derived types or field declarations anywhere in the
      // interaction-state modules (comment prose quoting the lock is fine;
      // declarations are not).
      expect(source).not.toContain('TaskSchedule')
      expect(source).not.toMatch(/\btotalSlack\b/)
      expect(source).not.toMatch(/\bfreeSlack\b/)
      expect(source).not.toMatch(/critical\??:/)
      expect(source).not.toMatch(/scheduled(Start|Finish)\??:/)
    }
    // The view state's only time values are the viewport window instants.
    expect(stateSource).toContain('readonly start: string')
    expect(stateSource).toContain('readonly finish: string')
    // The editing state carries an entity reference + field + draft text
    // only — the draft is user input en route to a command, never a cached
    // scheduling value.
    expect(editingSource).toContain('readonly taskId: TaskId')
    expect(editingSource).toContain('readonly field: EditableTaskField')
    expect(editingSource).toContain('readonly draft: string')
  })

  it('exposes the public surface from the index (projection, control, state, timeline)', () => {
    const index = srcModules['../src/index.ts'] ?? ''
    for (const symbol of [
      'ProjectViewState',
      'createViewState',
      'reconcileViewState',
      'ProjectViewIntent',
      'reduceViewState',
      'projectDocumentView',
      'ProjectTaskRow',
      'ProjectViewProjection',
      'buildCreateTaskCommand',
      'buildCreateTaskInSiblingGroupCommand',
      'buildIndentCommand',
      'buildOutdentCommand',
      'buildDeleteSelectionCommands',
      'nextTaskIdentity',
      'ProjectRendererSession',
      'createRendererSession',
      'applyRendererCommand',
      'undoRendererCommand',
      'redoRendererCommand',
      'ScheduleRunner',
      'buildTimeAxis',
      'scaleViewport',
      'fitViewport',
    ]) {
      expect(index).toContain(symbol)
    }
  })

  it('exposes the PROJECT-022 view-model surface from the index (grid, timeline, bars, dependencies, milestones, virtualization, hit testing)', () => {
    const index = srcModules['../src/index.ts'] ?? ''
    for (const symbol of [
      'ProjectTaskGrid',
      'buildTaskGrid',
      'ProjectTimeline',
      'buildTimeline',
      'ProjectGanttBars',
      'buildGanttBars',
      'ProjectDependencies',
      'buildDependencies',
      'ProjectMilestones',
      'buildMilestones',
      'ProjectRowWindow',
      'buildRowWindow',
      'ProjectGanttView',
      'buildGanttView',
      'hitTestGantt',
    ]) {
      expect(index).toContain(symbol)
    }
    // The five architect-named surfaces exist as exported types:
    for (const symbol of [
      'type ProjectTaskGrid',
      'type ProjectTimeline',
      'type ProjectGanttBars',
      'type ProjectDependencies',
      'type ProjectMilestones',
    ]) {
      expect(index).toContain(symbol)
    }
  })

  it('exposes the PROJECT-023 selection/editing surface from the index (editing model, commit flow, navigation)', () => {
    const index = srcModules['../src/index.ts'] ?? ''
    for (const symbol of [
      'type EditableTaskField',
      'type TaskEditing',
      'type TaskEditCommit',
      'type TaskEditInvalidReason',
      'type TaskEditFlowOutcome',
      'type MoveFocusDirection',
      'EDITABLE_TASK_FIELDS',
      'editableTaskFields',
      'isTaskFieldEditable',
      'initialTaskEditDraft',
      'commitTaskEdit',
      'commitTaskEditThroughSession',
    ]) {
      expect(index).toContain(symbol)
    }
    // The editing intents and the keyboard-navigation intent exist on the
    // frozen intent union (the selection intents were PROJECT-021).
    const intents = srcModules['../src/intents.ts'] ?? ''
    for (const intent of [
      "type: 'beginTaskEdit'",
      "type: 'updateTaskEditDraft'",
      "type: 'endTaskEdit'",
      "type: 'moveTaskFocus'",
    ]) {
      expect(intents).toContain(intent)
    }
    // The editing state slice is an additive optional field of the view
    // state, reconciled when its task dies.
    const state = srcModules['../src/state.ts'] ?? ''
    expect(state).toContain('readonly editing?: TaskEditing')
  })

  it('exposes the PROJECT-024 dependency-editing surface from the index (builders, editing model, commit flow)', () => {
    const index = srcModules['../src/index.ts'] ?? ''
    for (const symbol of [
      'type AddDependencyOptions',
      'type EditableDependencyField',
      'type DependencyEditing',
      'type DependencyEditCommit',
      'type DependencyEditInvalidReason',
      'type DependencyEditFlowOutcome',
      'DEFAULT_DEPENDENCY_TYPE',
      'DEFAULT_DEPENDENCY_LAG_MINUTES',
      'DEPENDENCY_TYPE_CODES',
      'EDITABLE_DEPENDENCY_FIELDS',
      'editableDependencyFields',
      'initialDependencyEditDraft',
      'commitDependencyEdit',
      'commitDependencyEditThroughSession',
      'buildAddDependencyCommand',
      'buildRemoveDependencySelectionCommands',
      'nextDependencyIdentity',
    ]) {
      expect(index).toContain(symbol)
    }
    // The dependency-editing intents exist on the frozen intent union.
    const intents = srcModules['../src/intents.ts'] ?? ''
    for (const intent of [
      "type: 'beginDependencyEdit'",
      "type: 'updateDependencyEditDraft'",
      "type: 'endDependencyEdit'",
    ]) {
      expect(intents).toContain(intent)
    }
    // The dependency-editing state slice is an additive optional field of
    // the view state, reconciled when its link dies.
    const state = srcModules['../src/state.ts'] ?? ''
    expect(state).toContain('readonly dependencyEditing?: DependencyEditing')
    // The link view model carries the interaction-state reflection.
    const dependencies = srcModules['../src/views/dependencies.ts'] ?? ''
    expect(dependencies).toContain('readonly selected: boolean')
    expect(dependencies).toContain('readonly editingField?: EditableDependencyField')
  })

  it('exposes the PROJECT-025 calendar-visualization surface from the index (query, catalog, bands, surfaces)', () => {
    const index = srcModules['../src/index.ts'] ?? ''
    for (const symbol of [
      'type CalendarWorkingInterval',
      'type CalendarWorkingTimeQuery',
      'type CalendarSurfaceStatus',
      'type CalendarViewInput',
      'type ProjectCalendarBand',
      'type ProjectCalendarCatalog',
      'type ProjectCalendarCatalogEntry',
      'type ProjectCalendarSurface',
      'type ProjectRowCalendar',
      'CALENDAR_EVALUATION_FAILED',
      'buildCalendarCatalog',
      'buildCalendarSurface',
      'classifyCalendarBands',
    ]) {
      expect(index).toContain(symbol)
    }
    // The timeline carries the additive calendar surfaces (optional —
    // present iff a working-time query was threaded, never invented).
    const timeline = srcModules['../src/views/timeline.ts'] ?? ''
    expect(timeline).toContain('readonly calendar?: ProjectCalendarSurface')
    expect(timeline).toContain('readonly rowCalendars?: readonly ProjectRowCalendar[]')
    // The gantt view threads the calendar input.
    const ganttView = srcModules['../src/views/gantt-view.ts'] ?? ''
    expect(ganttView).toContain('calendar?: CalendarViewInput')
  })

  it('keeps the calendar projection free of calendar-evaluation semantics (no second calendar engine)', () => {
    // The calendar module may echo definitions and do interval algebra, but
    // it must never DECOMPOSE a date (weekday/date-part extraction is the
    // primitive every calendar evaluation needs) nor key exceptions by
    // date — without these, no working-time derivation is possible.
    const calendar = srcModules['../src/calendar.ts'] ?? ''
    expect(calendar).not.toMatch(
      /getUTCDay|getUTCFullYear|getUTCMonth|getUTCDate|getUTCHours|getUTCMinutes/,
    )
    expect(calendar).not.toMatch(/new Date\(Date\.UTC\(|dateKey/)
    expect(calendar).not.toContain('toISOString().slice')
    // The only evaluation entry is the INJECTED query (a type declaration,
    // never an implementation) — the ScheduleRunner precedent.
    expect(calendar).toContain('export type CalendarWorkingTimeQuery =')
    expect(calendar).toContain('workingTime: CalendarWorkingTimeQuery')
    // The scheduling package stays un-imported (also covered by the global
    // import scan; asserted here for the calendar surface explicitly).
    expect(calendar).not.toContain("from '@genoffice/project-scheduling'")
    // The timeline composes surfaces through the injected query only.
    const timeline = srcModules['../src/views/timeline.ts'] ?? ''
    expect(timeline).toContain('buildCalendarSurface')
    expect(timeline).not.toContain("from '@genoffice/project-scheduling'")
  })

  it('keeps the view models free of pixel/DOM APIs (fraction space only)', () => {
    const viewFiles = Object.entries(srcModules).filter(([file]) =>
      file.startsWith('../src/views/'),
    )
    expect(viewFiles.length).toBeGreaterThanOrEqual(7)
    for (const [file, source] of viewFiles) {
      expect(source, `${file} touches a pixel-layout API`).not.toMatch(
        /clientWidth|offsetWidth|innerWidth|getBoundingClientRect|devicePixelRatio|scrollHeight/,
      )
      // DOM document APIs — the canonical `ProjectDocument` parameter (e.g.
      // `document.dependencies`) is NOT a DOM reference:
      expect(source, `${file} touches the DOM document`).not.toMatch(
        /document\.(getElementById|createElement|querySelector|addEventListener|body)/,
      )
      expect(source, `${file} uses localeCompare`).not.toContain('localeCompare')
      expect(source, `${file} uses Date.now`).not.toContain('Date.now(')
      expect(source, `${file} uses Math.random`).not.toContain('Math.random(')
    }
  })

  it('documents the scheduling-authority injection on the session module', () => {
    const session = srcModules['../src/session.ts'] ?? ''
    expect(session).toContain('INJECTED')
    expect(session).toContain('ScheduleRunner')
    expect(session).not.toContain("from '@genoffice/project-scheduling'")
  })
})

describe('PROJECT-021 architecture — CI and spec lockstep', () => {
  it('the foundation CI gate covers the renderer core (boundary grep + typecheck + test)', () => {
    expect(workflow).toContain('packages/project-renderer-core')
    expect(workflow).toContain('Typecheck project-renderer-core')
    expect(workflow).toContain('Test project-renderer-core')
  })

  it('the spec set carries the PROJECT-021 sections in lockstep', () => {
    expect(requirements).toContain('PROJECT-021 — Shared renderer core')
    expect(workItems).toMatch(/\|\s*PROJECT-021\s*\|/)
    expect(workItems).toContain('ProjectRendererSession')
    expect(dependencyGraph).toContain('Package dependency edges (PROJECT-021)')
    expect(dependencyGraph).toContain('project-renderer-core')
    expect(verificationMatrix).toContain('PROJECT-021 evidence requirements')
  })

  it('the spec set carries the PROJECT-022 sections in lockstep', () => {
    expect(requirements).toContain('PROJECT-022 — Gantt / task grid / timeline views')
    expect(workItems).toMatch(/\|\s*PROJECT-022\s*\|/)
    expect(workItems).toContain('ProjectGanttView')
    expect(dependencyGraph).toContain('Package dependency edges (PROJECT-022)')
    expect(verificationMatrix).toContain('PROJECT-022 evidence requirements')
  })

  it('the spec set carries the PROJECT-023 sections in lockstep', () => {
    expect(requirements).toContain('PROJECT-023 — Selection / editing')
    expect(workItems).toMatch(/\|\s*PROJECT-023\s*\|/)
    expect(workItems).toContain('commitTaskEditThroughSession')
    expect(dependencyGraph).toContain('Package dependency edges (PROJECT-023)')
    expect(verificationMatrix).toContain('PROJECT-023 evidence requirements')
  })

  it('the spec set carries the PROJECT-025 sections in lockstep', () => {
    expect(requirements).toContain('PROJECT-025 — Calendar visualization')
    expect(workItems).toMatch(/\|\s*PROJECT-025\s*\|/)
    expect(workItems).toContain('CalendarWorkingTimeQuery')
    expect(dependencyGraph).toContain('Package dependency edges (PROJECT-025)')
    expect(verificationMatrix).toContain('PROJECT-025 evidence requirements')
  })

  it('leaves the frozen architecture lock untouched (renderer boundary already sanctioned)', () => {
    expect(architectureLock).toContain(
      '`packages/project-renderer-core`: shared renderer boundary until PROJECT-021+; no scheduling authority.',
    )
    expect(architectureLock).toContain('Status: FROZEN')
  })

  it('documents the clarified §13 rule (ACR-001): the renderer acquires no file-format knowledge', () => {
    // The clarified §13 (roadmap reconciliation increment, ACR-001): the
    // renderer is a foundation semantic/runtime package — it must not import
    // external MSPDI/MPP parser implementations nor any format-specific
    // parser internals; file-format implementations stay behind the
    // project-file adapter boundary.
    expect(architectureLock).toContain(
      'Foundation semantic/runtime packages (`project-contracts`, `project-engine`, `project-scheduling`, `project-renderer-core`) must not import external MSPDI/MPP parser implementations.',
    )
    expect(architectureLock).toContain(
      'No `project-engine`, `project-scheduling`, `project-renderer-core`, or host package may directly import format-specific parser internals.',
    )
    expect(architectureLock).toContain(
      'File-format implementations remain behind the `project-file` adapter boundary.',
    )
    expect(architectureLock).toContain('ACR-001')
    expect(acr).toContain('## 9. Principal Architect approval reference')
    // The renderer's runtime dependency set stays exactly contracts + engine
    // (no file adapter package, no format internals).
    expect(Object.keys(pkg.dependencies ?? {}).sort()).toEqual([
      '@genoffice/project-contracts',
      '@genoffice/project-engine',
    ])
  })
})
