/**
 * PROJECT-018 — full MPP→canonical pipeline (host side).
 *
 *   MPP bytes / file
 *     → input-size validation                (MPP_INPUT_TOO_LARGE)
 *     → isolated temp workspace              (unique dir per import,
 *                                            deterministically removed)
 *     → MPXJ sidecar conversion              (launcher; stage 'sidecar')
 *     → importMppFromMspdi                   (foundation: N1–N5 → accepted
 *                                            PROJECT-015 importer → engine
 *                                            validation; stages
 *                                            'normalization'/'mspdi'/
 *                                            'canonical')
 *     → schedule()                           (stage 'scheduling')
 *
 * Error atomicity: a fatal CONVERSION failure (sidecar error, malformed
 * MSPDI, canonical rejection) yields `emptyProjectDocument()` — never a
 * partially authoritative document — with every stage's diagnostics
 * preserved. A scheduling failure keeps the (valid) canonical document and
 * returns the schedule in its rejected shape (`{ taskSchedules: {},
 * diagnostics }`); scheduling is derived state, not conversion.
 *
 * The input file is NEVER mutated; on failure no repository/project state
 * changes (the caller receives values only — this library owns no state).
 */
import { mkdtempSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type {
  DerivedSchedule,
  ProjectDocument,
  ProjectFileMetadata,
} from '@genoffice/project-contracts'
import {
  MPP_INPUT_TOO_LARGE,
  MPP_MAX_INPUT_BYTES,
  emptyProjectDocument,
  importMppFromMspdi,
  type MppConversionOutcome,
  type MppDiagnostic,
} from '@genoffice/project-file'
import { schedule } from '@genoffice/project-scheduling'
import { MppSidecarLauncher } from './launcher.js'

/** The full staged result of importing an MPP file. */
export interface MppFullImportResult {
  /** The canonical document (empty on fatal conversion failure). */
  readonly document: ProjectDocument
  /** The derived schedule. Present only when the conversion succeeded;
   * a scheduling failure keeps the rejected shape (`{ taskSchedules: {},
   * diagnostics: [] }` becomes the scheduling-stage diagnostics). */
  readonly schedule?: DerivedSchedule
  /** Staged diagnostics, provenance preserved (sidecar → normalization →
   * mspdi → canonical → scheduling). */
  readonly diagnostics: readonly MppDiagnostic[]
}

export interface MppImportOptions {
  readonly launcher: MppSidecarLauncher
  /** Optional metadata forwarded to the importer (like every adapter). */
  readonly metadata?: ProjectFileMetadata
  /** Base directory for temp workspaces (default: the OS temp dir). */
  readonly tempBase?: string
}

/** Import an MPP file from disk. The file is read by the sidecar only — it
 * is never mutated, and nothing outside the per-import temp workspace is
 * written. */
export async function importMppFromFile(
  inputPath: string,
  options: MppImportOptions,
): Promise<MppFullImportResult> {
  const diagnostics: MppDiagnostic[] = []
  let inputSize: number
  try {
    inputSize = statSync(inputPath).size
  } catch {
    // An unreadable input is an input-side failure, reported at the sidecar
    // stage (the launcher never sees it — atomic, no partial state):
    diagnostics.push({
      code: MPP_INPUT_TOO_LARGE,
      severity: 'error',
      message: `MPP input file cannot be read: ${inputPath}`,
      stage: 'sidecar',
    })
    return { document: emptyProjectDocument(), diagnostics }
  }
  if (inputSize > MPP_MAX_INPUT_BYTES) {
    diagnostics.push({
      code: MPP_INPUT_TOO_LARGE,
      severity: 'error',
      message: `MPP input is ${inputSize} bytes (limit ${MPP_MAX_INPUT_BYTES})`,
      stage: 'sidecar',
    })
    return { document: emptyProjectDocument(), diagnostics }
  }

  const workspace = mkdtempSync(join(options.tempBase ?? tmpdir(), 'genoffice-mpp-'))
  try {
    const outputPath = join(workspace, 'converted.mspdi')
    const conversion = await options.launcher.convert(inputPath, outputPath)
    if (!conversion.ok) {
      return {
        document: emptyProjectDocument(),
        diagnostics: [...diagnostics, ...conversion.diagnostics],
      }
    }
    return finishImport(
      {
        mspdiBytes: conversion.mspdiBytes,
        frame: conversion.frame,
        sidecarDiagnostics: diagnostics,
      },
      options,
    )
  } finally {
    // Deterministic cleanup — the workspace is ALWAYS removed, success or
    // failure, converted payload included (the result holds the bytes it
    // needs in memory).
    rmSync(workspace, { recursive: true, force: true })
  }
}

/** Import MPP bytes. The bytes are written to an isolated temp workspace
 * (never adjacent to user files) and processed exactly like a file. */
export async function importMppFromBytes(
  bytes: Uint8Array,
  options: MppImportOptions,
): Promise<MppFullImportResult> {
  const diagnostics: MppDiagnostic[] = []
  if (bytes.byteLength > MPP_MAX_INPUT_BYTES) {
    diagnostics.push({
      code: MPP_INPUT_TOO_LARGE,
      severity: 'error',
      message: `MPP input is ${bytes.byteLength} bytes (limit ${MPP_MAX_INPUT_BYTES})`,
      stage: 'sidecar',
    })
    return { document: emptyProjectDocument(), diagnostics }
  }
  const workspace = mkdtempSync(join(options.tempBase ?? tmpdir(), 'genoffice-mpp-'))
  try {
    const inputPath = join(workspace, 'input.mpp')
    const outputPath = join(workspace, 'converted.mspdi')
    writeFileSync(inputPath, bytes)
    const conversion = await options.launcher.convert(inputPath, outputPath)
    if (!conversion.ok) {
      return {
        document: emptyProjectDocument(),
        diagnostics: [...diagnostics, ...conversion.diagnostics],
      }
    }
    return finishImport(
      {
        mspdiBytes: conversion.mspdiBytes,
        frame: conversion.frame,
        sidecarDiagnostics: diagnostics,
      },
      options,
    )
  } finally {
    rmSync(workspace, { recursive: true, force: true })
  }
}

/** Shared tail of the pipeline: foundation import + scheduling. */
function finishImport(
  outcome: MppConversionOutcome,
  options: MppImportOptions,
): MppFullImportResult {
  const imported = importMppFromMspdi(outcome, options.metadata)
  const diagnostics: MppDiagnostic[] = [...imported.diagnostics]
  const derived = schedule(imported.document)
  diagnostics.push(...stageSchedulingDiagnostics(derived.diagnostics))
  return { document: imported.document, schedule: derived, diagnostics }
}

/**
 * Map scheduling-engine diagnostics to their staged form (pure — exported
 * for the scheduling-failure-path test: the mapping itself must preserve
 * code/severity/message and stamp `stage: 'scheduling'`).
 */
export function stageSchedulingDiagnostics(
  diagnostics: ReadonlyArray<{
    code: string
    severity: 'info' | 'warning' | 'error'
    message: string
  }>,
): MppDiagnostic[] {
  return diagnostics.map((d) => ({ ...d, stage: 'scheduling' }))
}
