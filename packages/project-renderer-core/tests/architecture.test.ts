import { describe, expect, it } from 'vitest'
import pkgRaw from '../package.json?raw'
import workflow from '../../.github/workflows/project-foundation.yml?raw'
import requirements from '../../../spec/project/requirements.md?raw'
import workItems from '../../../spec/project/work-items.md?raw'
import dependencyGraph from '../../../spec/project/dependency-graph.md?raw'
import verificationMatrix from '../../../spec/project/verification-matrix.md?raw'
import architectureLock from '../../../spec/project/architecture-lock.md?raw'

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
  { pattern: /from ['"]@genoffice\/project-file['"]/, label: 'the file package' },
  { pattern: /from ['"]@genoffice\/project-mpp-host['"]/, label: 'the MPP host package' },
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

  it('never imports React/Electron/Node/HTTP/scheduling/file/host packages in src', () => {
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
    for (const source of [stateSource, intentsSource, editingSource, editFlowSource]) {
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

  it('leaves the frozen architecture lock untouched (renderer boundary already sanctioned)', () => {
    expect(architectureLock).toContain(
      '`packages/project-renderer-core`: shared renderer boundary until PROJECT-021+; no scheduling authority.',
    )
    expect(architectureLock).toContain('Status: FROZEN')
  })
})
