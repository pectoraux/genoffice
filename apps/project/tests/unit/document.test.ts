/**
 * PROJECT-027 — the host document battery (real adapters, real engine).
 *
 * Proves the new-document template is a VALID, schedulable canonical
 * document (the engine validates it; the scheduler derives from it), the
 * open/save composition runs through the canonical adapters with honest
 * error outcomes, and the format/extension mapping is exact.
 */
import { describe, expect, it } from 'vitest'
import { asTaskId } from '@genoffice/project-contracts'
import { applyProjectCommand, validateProjectDocument } from '@genoffice/project-engine'
import { schedule } from '@genoffice/project-scheduling'
import { buildCreateTaskCommand } from '@genoffice/project-renderer-core'
import {
  adapterForFormat,
  defaultFileNameFor,
  exportDocumentBytes,
  extensionForFormat,
  formatForPath,
  importDocumentBytes,
  newProjectDocument,
} from '../../src/renderer/document.js'

describe('the new-document template', () => {
  it('is a valid canonical document (the engine accepts it)', () => {
    const validation = validateProjectDocument(newProjectDocument())
    expect(validation.accepted).toBe(true)
    expect(validation.diagnostics).toEqual([])
  })

  it('schedules through the real canonical scheduler', () => {
    const derived = schedule(newProjectDocument())
    expect(derived.diagnostics).toEqual([])
    expect(derived.projectStart).toBeDefined()
  })

  it('schedules a created task through the canonical pipeline', () => {
    // Grow the template the way the host does: the renderer-core builder →
    // the engine command application → the real scheduler.
    const document = newProjectDocument('Fixture')
    const execution = applyProjectCommand(document, buildCreateTaskCommand(document), 'c1')
    expect(execution.result.accepted).toBe(true)
    const derived = schedule(execution.document)
    expect(derived.diagnostics).toEqual([])
    expect(derived.taskSchedules[asTaskId('t1')]).toBeDefined()
  })

  it('is deterministic (the same call produces the same document)', () => {
    expect(newProjectDocument()).toEqual(newProjectDocument('Project1'))
    expect(newProjectDocument('Alpha').properties.name).toBe('Alpha')
  })
})

describe('format mapping', () => {
  it('maps extensions to formats and back', () => {
    expect(formatForPath('/tmp/a.gproj')).toBe('gproj')
    expect(formatForPath('/tmp/a.GPROJ')).toBe('gproj')
    expect(formatForPath('/tmp/a.xml')).toBe('mspdi')
    expect(formatForPath('/tmp/a.mpp')).toBeNull()
    expect(formatForPath('/tmp/a.docx')).toBeNull()
    expect(extensionForFormat('gproj')).toBe('gproj')
    expect(extensionForFormat('mspdi')).toBe('xml')
  })

  it('selects the canonical adapters', () => {
    expect(adapterForFormat('gproj').format).toBe('gproj')
    expect(adapterForFormat('mspdi').format).toBe('mspdi')
  })

  it('derives the save-dialog default name from the document name', () => {
    expect(defaultFileNameFor(newProjectDocument('Site Works'), 'gproj')).toBe('Site Works.gproj')
    expect(defaultFileNameFor(newProjectDocument('  '), 'gproj')).toBe('Untitled.gproj')
  })
})

describe('open through the canonical adapters', () => {
  it('round-trips a .gproj export → import byte-identically (document equality)', () => {
    const document = newProjectDocument('Round Trip')
    const exported = exportDocumentBytes(document, 'gproj')
    expect(exported.diagnostics).toEqual([])
    const outcome = importDocumentBytes('/tmp/round-trip.gproj', exported.bytes)
    expect(outcome.kind).toBe('imported')
    if (outcome.kind === 'imported') {
      expect(outcome.imported.document).toEqual(document)
      expect(outcome.imported.format).toBe('gproj')
      // The adapter's own info-level read marker travels verbatim; no
      // warnings, no errors.
      expect(outcome.imported.diagnostics.every((d) => d.severity === 'info')).toBe(true)
    }
  })

  it('imports an MSPDI export through the same composition', () => {
    const document = newProjectDocument('Interchange')
    const exported = exportDocumentBytes(document, 'mspdi')
    const outcome = importDocumentBytes('/tmp/interchange.xml', exported.bytes)
    expect(outcome.kind).toBe('imported')
    if (outcome.kind === 'imported') {
      expect(outcome.imported.format).toBe('mspdi')
      // The MSPDI interchange is the accepted 016 surface: the document
      // round-trips schedulable.
      const derived = schedule(outcome.imported.document)
      expect(derived.diagnostics).toEqual([])
    }
  })

  it('surfaces adapter error diagnostics as a failed open (never a load)', () => {
    const garbage = new TextEncoder().encode('this is not a project file')
    const outcome = importDocumentBytes('/tmp/broken.gproj', garbage)
    expect(outcome.kind).toBe('error')
    if (outcome.kind === 'error') {
      expect(outcome.message.length).toBeGreaterThan(0)
    }
  })

  it('rejects unsupported extensions', () => {
    const outcome = importDocumentBytes('/tmp/plan.mpp', new Uint8Array())
    expect(outcome.kind).toBe('error')
    if (outcome.kind === 'error') {
      expect(outcome.message).toContain('Unsupported project file')
    }
  })
})
