import { describe, expect, it } from 'vitest'
import { buildDependencyGraph, DependencyGraphError } from '../src/index.js'
import { makeDependency, makeDocument, makeTask, taskId, wm } from './fixtures.js'
import type { ProjectDocument } from '@genoffice/project-contracts'

const baseDocument = (overrides: Partial<ProjectDocument> = {}): ProjectDocument => ({
  ...makeDocument({
    tasks: [makeTask({ id: 'a' }), makeTask({ id: 'b' })],
  }),
  ...overrides,
})

describe('dependency graph', () => {
  it('accepts FS, SS, FF, and SF links with positive and negative lag', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'a' }),
        makeTask({ id: 'b' }),
        makeTask({ id: 'c' }),
        makeTask({ id: 'd' }),
      ],
      dependencies: [
        makeDependency('d1', 'a', 'b', 'FS', 60),
        makeDependency('d2', 'a', 'c', 'SS', -120),
        makeDependency('d3', 'a', 'd', 'FF', 30),
        makeDependency('d4', 'b', 'd', 'SF', 0),
      ],
    })
    const graph = buildDependencyGraph(document)
    expect(graph.topologicalOrder).toEqual([taskId('a'), taskId('b'), taskId('c'), taskId('d')])
    expect(graph.predecessors.get(taskId('b'))).toHaveLength(1)
    expect(graph.successors.get(taskId('a'))).toHaveLength(3)
  })

  it('rejects self dependencies, cycles, and missing references', () => {
    expect(() =>
      buildDependencyGraph(
        baseDocument({
          dependencies: [makeDependency('d', 'a', 'a', 'FS')],
        }),
      ),
    ).toThrowError(/self-referential/)
    expect(() =>
      buildDependencyGraph(
        baseDocument({
          dependencies: [
            makeDependency('d1', 'a', 'b', 'FS'),
            makeDependency('d2', 'b', 'a', 'FS'),
          ],
        }),
      ),
    ).toThrowError(/cycle/i)
    expect(() =>
      buildDependencyGraph(
        baseDocument({
          dependencies: [makeDependency('d', 'a', 'ghost', 'FS')],
        }),
      ),
    ).toThrowError(/unknown task/i)
    expect(() =>
      buildDependencyGraph(
        baseDocument({
          dependencies: [makeDependency('d', 'ghost', 'a', 'FS')],
        }),
      ),
    ).toThrowError(/unknown task/i)
  })

  it('breaks ties deterministically by task id regardless of input order', () => {
    const ordered = makeDocument({
      tasks: [
        makeTask({ id: 'm' }),
        makeTask({ id: 'z' }),
        makeTask({ id: 'a' }),
        makeTask({ id: 'q' }),
      ],
      dependencies: [makeDependency('d1', 'z', 'm', 'FS'), makeDependency('d2', 'a', 'm', 'FS')],
    })
    const shuffled = makeDocument({
      tasks: [
        makeTask({ id: 'q' }),
        makeTask({ id: 'a' }),
        makeTask({ id: 'm' }),
        makeTask({ id: 'z' }),
      ],
      dependencies: [makeDependency('d2', 'a', 'm', 'FS'), makeDependency('d1', 'z', 'm', 'FS')],
    })
    const first = buildDependencyGraph(ordered)
    const second = buildDependencyGraph(shuffled)
    expect(first.topologicalOrder).toEqual([taskId('a'), taskId('q'), taskId('z'), taskId('m')])
    expect(second.topologicalOrder).toEqual(first.topologicalOrder)
  })

  it('orders children before their summary parent', () => {
    const document = makeDocument({
      tasks: [
        makeTask({ id: 'summary', summary: true, duration: wm(0) }),
        makeTask({ id: 'child-a', parentTaskId: taskId('summary'), outlineLevel: 2 }),
        makeTask({ id: 'child-b', parentTaskId: taskId('summary'), outlineLevel: 2 }),
      ],
    })
    const graph = buildDependencyGraph(document)
    expect(graph.topologicalOrder.indexOf(taskId('child-a'))).toBeLessThan(
      graph.topologicalOrder.indexOf(taskId('summary')),
    )
    expect(graph.topologicalOrder.indexOf(taskId('child-b'))).toBeLessThan(
      graph.topologicalOrder.indexOf(taskId('summary')),
    )
  })

  it('rejects hierarchy cycles, self parents, and missing parents', () => {
    expect(() =>
      buildDependencyGraph(
        makeDocument({
          tasks: [
            makeTask({ id: 'a', parentTaskId: taskId('b') }),
            makeTask({ id: 'b', parentTaskId: taskId('a') }),
          ],
        }),
      ),
    ).toThrowError(/cycle/i)
    expect(() =>
      buildDependencyGraph(
        makeDocument({
          tasks: [makeTask({ id: 'a', parentTaskId: taskId('a') })],
        }),
      ),
    ).toThrowError(/cannot parent itself/i)
    expect(() =>
      buildDependencyGraph(
        makeDocument({
          tasks: [makeTask({ id: 'a', parentTaskId: taskId('ghost') })],
        }),
      ),
    ).toThrowError(/missing parent/i)
  })

  it('rejects a dependency that would contradict the hierarchy order', () => {
    // summary -> child dependency conflicts with child-before-parent roll-up.
    expect(() =>
      buildDependencyGraph(
        makeDocument({
          tasks: [
            makeTask({ id: 'summary', summary: true }),
            makeTask({ id: 'child', parentTaskId: taskId('summary'), outlineLevel: 2 }),
          ],
          dependencies: [makeDependency('d', 'summary', 'child', 'FS')],
        }),
      ),
    ).toThrowError(DependencyGraphError)
  })

  it('returns a canonical order on repeated builds', () => {
    const document = makeDocument({
      tasks: [makeTask({ id: 'x' }), makeTask({ id: 'y' }), makeTask({ id: 'w' })],
      dependencies: [
        makeDependency('d1', 'x', 'y', 'FS'),
        makeDependency('d2', 'w', 'x', 'SS', 30),
      ],
    })
    expect(buildDependencyGraph(document).topologicalOrder).toEqual(
      buildDependencyGraph(document).topologicalOrder,
    )
  })
})
