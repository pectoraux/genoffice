import { describe, expect, it } from 'vitest'
import workItemsRaw from '../../../spec/project/work-items.md?raw'
import dependencyGraphRaw from '../../../spec/project/dependency-graph.md?raw'
import architectureLockRaw from '../../../spec/project/architecture-lock.md?raw'
import acrRaw from '../../../spec/project/architecture-changes/ACR-001-project-file-adapter-boundary.md?raw'
import verificationMatrixRaw from '../../../spec/project/verification-matrix.md?raw'

/**
 * Roadmap reconciliation increment — deterministic specification-consistency
 * guard.
 *
 * Verifies, by SOURCE-LEVEL TEXT PARSING ONLY (no runtime package, no
 * filesystem access beyond vitest `?raw` module sources — the established
 * project-file discipline-suite precedent), that `work-items.md` and
 * `dependency-graph.md` express the SAME direct dependency information:
 *
 *   - every numbered work item exists in both documents;
 *   - every direct dependency declared by work-items.md is represented in
 *     the dependency graph;
 *   - every dependency-graph edge has a corresponding work-item dependency;
 *   - there are no contradictory predecessor declarations (per-item set
 *     equality, no self-dependencies, dependencies point to strictly
 *     lower-numbered items only, global acyclicity);
 *   - no unknown work-item IDs appear in either document's dependency data;
 *   - PROJECT-019A remains correctly represented as the accepted rescope
 *     decision (never as a separate graph node);
 *   - the next authorized product increment is correctly recorded
 *     (PROJECT-027 after the lockstep advance upon PROJECT-026's
 *     acceptance), with all of its direct dependencies inside the accepted
 *     set.
 *
 * Scope: the CANONICAL structures are the work-items table, the dependency
 * graph's canonical direct-dependency block, and the accepted-frontier
 * record. The per-project "Package dependency edges" sections are narrative
 * history (static package edges), not work-item dependency declarations,
 * and are intentionally not parsed as edges.
 */

// ---------------------------------------------------------------------------
// work-items.md — parse the table (ID | Title | Depends on | … | …).
// ---------------------------------------------------------------------------

const WORK_ITEM_ROW_RE = /^\|\s*(PROJECT-\d{3})\s*\|\s*([^|]+?)\s*\|\s*([^|]*?)\s*\|/

function parseDependeeList(raw: string, context: string): string[] {
  const trimmed = raw.trim()
  if (trimmed === 'none' || trimmed === '') return []
  const parts = trimmed.split(',').map((p) => p.trim())
  for (const part of parts) {
    if (!/^\d{3}$/.test(part)) {
      throw new Error(`${context}: malformed dependency token ${JSON.stringify(part)}`)
    }
  }
  return parts.map((p) => `PROJECT-${p}`)
}

const workItems = new Map<string, { title: string; deps: string[] }>()
const workItemOrder: string[] = []
for (const line of workItemsRaw.split('\n')) {
  const match = WORK_ITEM_ROW_RE.exec(line)
  if (!match) continue
  const [, id, title, depsRaw] = match
  if (workItems.has(id)) {
    throw new Error(`work-items.md: duplicate row ${id}`)
  }
  workItems.set(id, {
    title: title.trim(),
    deps: parseDependeeList(depsRaw, `work-items.md ${id}`),
  })
  workItemOrder.push(id)
}

// ---------------------------------------------------------------------------
// dependency-graph.md — parse the canonical direct-dependency block and the
// accepted-frontier record (both strict line formats; any deviation fails).
// ---------------------------------------------------------------------------

function extractFencedBlock(document: string, sectionHeading: string): string[] {
  const headingIndex = document.indexOf(sectionHeading)
  if (headingIndex < 0) {
    throw new Error(`dependency-graph.md: missing section ${JSON.stringify(sectionHeading)}`)
  }
  const fenceStart = document.indexOf('```text', headingIndex)
  if (fenceStart < 0) {
    throw new Error(`dependency-graph.md: missing text fence under ${sectionHeading}`)
  }
  const fenceEnd = document.indexOf('```', fenceStart + '```text'.length)
  if (fenceEnd < 0) {
    throw new Error(`dependency-graph.md: unterminated fence under ${sectionHeading}`)
  }
  return document
    .slice(fenceStart + '```text'.length, fenceEnd)
    .split('\n')
    .filter((line) => line.trim() !== '')
}

const DAG_LINE_RE = /^PROJECT-(\d{3}) ← (none|\d{3}(?:, \d{3})*)$/

const graphDeps = new Map<string, string[]>()
const graphOrder: string[] = []
for (const line of extractFencedBlock(dependencyGraphRaw, '## Canonical direct-dependency graph')) {
  const match = DAG_LINE_RE.exec(line)
  if (!match) {
    throw new Error(
      `dependency-graph.md canonical block: line does not match the strict per-item format: ${JSON.stringify(line)}`,
    )
  }
  const id = `PROJECT-${match[1]}`
  if (graphDeps.has(id)) {
    throw new Error(`dependency-graph.md canonical block: duplicate node ${id}`)
  }
  graphDeps.set(id, parseDependeeList(match[2], `dependency-graph.md ${id}`))
  graphOrder.push(id)
}

const FRONTIER_ACCEPTED_RE =
  /^Accepted: PROJECT-(\d{3})\.\.PROJECT-(\d{3}) \(PROJECT-019 completed as the 019A rescope decision record\)$/
const FRONTIER_NEXT_RE = /^Next authorized: PROJECT-(\d{3})$/

const frontierLines = extractFencedBlock(dependencyGraphRaw, '## Accepted frontier')
if (frontierLines.length !== 2) {
  throw new Error(
    `dependency-graph.md accepted frontier: expected exactly 2 lines, got ${frontierLines.length}`,
  )
}
const acceptedMatch = FRONTIER_ACCEPTED_RE.exec(frontierLines[0] ?? '')
const nextMatch = FRONTIER_NEXT_RE.exec(frontierLines[1] ?? '')
if (!acceptedMatch || !nextMatch) {
  throw new Error(
    `dependency-graph.md accepted frontier: malformed record lines ${JSON.stringify(frontierLines)}`,
  )
}
const frontierAcceptedFirst = Number(acceptedMatch[1] as string)
const frontierAcceptedLast = Number(acceptedMatch[2] as string)
const frontierNext = `PROJECT-${nextMatch[1] as string}`
const acceptedSet = new Set<string>()
for (let n = frontierAcceptedFirst; n <= frontierAcceptedLast; n += 1) {
  acceptedSet.add(`PROJECT-${String(n).padStart(3, '0')}`)
}

const numeric = (id: string): number => Number(id.slice('PROJECT-'.length))

// ---------------------------------------------------------------------------
// Guard battery.
// ---------------------------------------------------------------------------

describe('specification consistency — work-items ↔ dependency-graph', () => {
  it('every numbered work item exists in both documents (49 items, same set)', () => {
    expect(workItems.size).toBe(49)
    expect(graphDeps.size).toBe(49)
    expect(new Set(workItemOrder)).toEqual(new Set(graphOrder))
  })

  it('both documents list items in the authoritative PROJECT-number order', () => {
    const numbersOf = (ids: string[]): number[] => ids.map(numeric)
    expect(numbersOf(workItemOrder)).toEqual([...numbersOf(workItemOrder)].sort((a, b) => a - b))
    expect(numbersOf(graphOrder)).toEqual([...numbersOf(graphOrder)].sort((a, b) => a - b))
  })

  it('every direct dependency declared by work-items.md is represented in the dependency graph', () => {
    for (const [id, { deps }] of workItems) {
      const graphEntry = graphDeps.get(id)
      expect(graphEntry, `${id} missing from the dependency graph`).toBeDefined()
      for (const dep of deps) {
        expect(graphEntry, `${id} dependency ${dep} missing from the graph edges`).toContain(dep)
      }
    }
  })

  it('every dependency-graph edge has a corresponding work-item dependency', () => {
    for (const [id, deps] of graphDeps) {
      const workItem = workItems.get(id)
      expect(workItem, `${id} missing from work-items.md`).toBeDefined()
      for (const dep of deps) {
        expect(workItem?.deps, `${id} graph edge ${dep} has no work-item dependency`).toContain(dep)
      }
    }
  })

  it('no contradictory predecessor declarations (set equality, no self-deps, backwards-only, acyclic)', () => {
    for (const id of workItemOrder) {
      const fromTable = [...(workItems.get(id)?.deps ?? [])].sort()
      const fromGraph = [...(graphDeps.get(id) ?? [])].sort()
      // Per-item set equality is the exact "no contradictions between the two
      // predecessor declarations" contract.
      expect(fromGraph, `${id}: graph and work-items disagree on direct dependencies`).toEqual(
        fromTable,
      )
      expect(fromTable, `${id}: self-dependency declared`).not.toContain(id)
      for (const dep of fromTable) {
        // The PROJECT-number order IS the authoritative roadmap order, so a
        // dependency on a later-numbered item would reorder the roadmap.
        expect(
          numeric(dep),
          `${id}: dependency ${dep} points forward in the roadmap order`,
        ).toBeLessThan(numeric(id))
      }
    }
    // Global acyclicity (implied by backwards-only, asserted independently).
    const visited = new Set<string>()
    const stack = new Set<string>()
    const visit = (id: string): void => {
      if (visited.has(id)) return
      if (stack.has(id)) throw new Error(`dependency cycle through ${id}`)
      stack.add(id)
      for (const dep of graphDeps.get(id) ?? []) visit(dep)
      stack.delete(id)
      visited.add(id)
    }
    for (const id of graphOrder) visit(id)
    expect(visited.size).toBe(49)
  })

  it('no unknown work-item IDs appear in either document’s dependency data', () => {
    const known = new Set(workItemOrder)
    for (const [id, { deps }] of workItems) {
      for (const dep of deps) {
        expect(known, `work-items.md ${id} references unknown ${dep}`).toContain(dep)
      }
    }
    for (const [id, deps] of graphDeps) {
      for (const dep of deps) {
        expect(known, `dependency-graph.md ${id} references unknown ${dep}`).toContain(dep)
      }
    }
  })

  it('PROJECT-019A remains correctly represented as the accepted rescope decision', () => {
    // The rescope is recorded in the 019 work-item row…
    expect(workItems.get('PROJECT-019')?.title).toBe('MPP export strategy / rescope (019A)')
    // …as a decision record in the dependency graph's PROJECT-019 section…
    expect(dependencyGraphRaw).toContain('## Package dependency edges (PROJECT-019)')
    expect(dependencyGraphRaw).toContain('DECISION RECORD')
    expect(dependencyGraphRaw).toContain('019A')
    // …in the accepted-frontier record…
    expect(frontierLines[0]).toContain('019A rescope decision record')
    // …and NEVER as a separate graph node or work-item row.
    expect(graphDeps.has('PROJECT-019A')).toBe(false)
    expect(workItems.has('PROJECT-019A')).toBe(false)
    expect(graphOrder).toHaveLength(49)
  })

  it('PROJECT-027 remains the next authorized product increment', () => {
    expect(frontierNext).toBe('PROJECT-027')
    // The frontier must be well-formed: the accepted set is contiguous from
    // PROJECT-001, fully populated in the roadmap, and the next authorized
    // item sits beyond it.
    expect(frontierAcceptedFirst).toBe(1)
    expect(frontierAcceptedLast).toBe(26)
    for (let n = 1; n <= 26; n += 1) {
      const id = `PROJECT-${String(n).padStart(3, '0')}`
      expect(workItems.has(id), `${id} missing although inside the accepted frontier`).toBe(true)
      expect(graphDeps.has(id), `${id} missing from the graph although accepted`).toBe(true)
    }
    expect(numeric(frontierNext)).toBeGreaterThan(frontierAcceptedLast)
    // The next authorized item exists in both documents with matching deps,
    // and EVERY direct dependency is inside the accepted set (the
    // authorization gate: "a work item cannot be authorized until all direct
    // dependencies are objectively accepted").
    expect(workItems.get('PROJECT-027')?.title).toBe('Electron shell')
    const depsFromTable = [...(workItems.get('PROJECT-027')?.deps ?? [])].sort()
    const depsFromGraph = [...(graphDeps.get('PROJECT-027') ?? [])].sort()
    expect(depsFromGraph).toEqual(depsFromTable)
    expect(depsFromTable).toEqual(['PROJECT-021', 'PROJECT-022'])
    for (const dep of depsFromTable) {
      expect(acceptedSet, `PROJECT-027 dependency ${dep} is not yet accepted`).toContain(dep)
    }
  })

  it('the canonical dependency-graph block uses ONLY the strict per-item format', () => {
    // The parse itself rejects deviations; this test documents the contract
    // and pins the superseded notations out of the document.
    expect(dependencyGraphRaw).toContain('## Canonical direct-dependency graph')
    expect(dependencyGraphRaw).not.toContain('## Full roadmap dependency intent')
    expect(dependencyGraphRaw).not.toContain('## Foundation DAG')
    for (const line of extractFencedBlock(
      dependencyGraphRaw,
      '## Canonical direct-dependency graph',
    )) {
      expect(DAG_LINE_RE.test(line), `non-canonical graph line: ${line}`).toBe(true)
    }
  })
})

describe('specification consistency — reconciliation artifacts in lockstep', () => {
  it('architecture-lock §13 carries the clarified rule and the ACR-001 reference (not silent)', () => {
    expect(architectureLockRaw).toContain(
      'Foundation semantic/runtime packages (`project-contracts`, `project-engine`, `project-scheduling`, `project-renderer-core`) must not import external MSPDI/MPP parser implementations.',
    )
    expect(architectureLockRaw).toContain(
      '`packages/project-file` is the sanctioned file-adapter boundary and may contain format-specific parser/serializer implementations.',
    )
    expect(architectureLockRaw).toContain(
      'No `project-engine`, `project-scheduling`, `project-renderer-core`, or host package may directly import format-specific parser internals.',
    )
    expect(architectureLockRaw).toContain(
      'File-format implementations remain behind the `project-file` adapter boundary.',
    )
    expect(architectureLockRaw).toContain('ACR-001')
    expect(architectureLockRaw).toContain('Status: FROZEN')
  })

  it('ACR-001 carries the mandated record sections and the approval reference', () => {
    for (const section of [
      '## 1. Motivation',
      '## 2. Affected invariant',
      '## 3. Old interpretation',
      '## 4. New interpretation',
      '## 5. Alternatives considered',
      '## 6. Compatibility impact',
      '## 7. Migration impact',
      '## 8. Verification impact',
      '## 9. Principal Architect approval reference',
    ]) {
      expect(acrRaw, `ACR-001 missing ${section}`).toContain(section)
    }
    expect(acrRaw).toContain('Principal Architect')
  })

  it('the verification matrix carries the reconciliation evidence requirements', () => {
    expect(verificationMatrixRaw).toContain(
      '## Architecture reconciliation evidence requirements (roadmap-lockstep increment)',
    )
    expect(verificationMatrixRaw).toContain('spec-consistency.test.ts')
  })
})
