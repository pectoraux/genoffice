/**
 * PROJECT-018 — full MPP→canonical pipeline (host side).
 *
 *   MPP bytes / file
 *     → input validation                    (MPP_INPUT_UNREADABLE for a
 *                                          missing/unreadable path —
 *                                          deliberately distinct from
 *                                          MPP_INPUT_TOO_LARGE for size;
 *                                          enforced before any process
 *                                          is ever started)
 *     → isolated temp workspace              (unique dir per import,
 *                                            deterministically removed)
 *     → MPXJ sidecar conversion              (launcher, inside the
 *                                            OS-enforced network-isolated
 *                                            context; stage 'sidecar')
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
import { closeSync, mkdtempSync, openSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type {
  DerivedSchedule,
  ProjectDocument,
  ProjectFileMetadata,
} from '@genoffice/project-contracts'
import {
  MPP_INPUT_TOO_LARGE,
  MPP_INPUT_UNREADABLE,
  MPP_MAX_INPUT_BYTES,
  buildCompatibilityReport,
  emptyProjectDocument,
  importMppFromMspdi,
  type CompatibilityReport,
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
  /** The source format the sidecar detected (e.g. `"MPP14"`), when a
   * frame was produced — the honest compatibility `sourceVersion`. */
  readonly sourceFormat?: string
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
    // Readability proof: stat alone does NOT require read permission on the
    // file itself (only search permission on the parent directory), so a
    // permission-denied input would otherwise surface later as a
    // misleading sidecar-stage "unsupported format" refusal. Opening the
    // file for reading (no data consumed) turns every input-side
    // readability failure — missing path (ENOENT), permission (EACCES),
    // any other I/O error — into the precise MPP_INPUT_UNREADABLE
    // diagnostic, before any process is ever started:
    closeSync(openSync(inputPath, 'r'))
  } catch (error) {
    diagnostics.push({
      code: MPP_INPUT_UNREADABLE,
      severity: 'error',
      message: `MPP input file cannot be read: ${inputPath} (${errorMessage(error)})`,
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
  return {
    document: imported.document,
    schedule: derived,
    ...(outcome.frame.format !== undefined ? { sourceFormat: outcome.frame.format } : {}),
    diagnostics,
  }
}

/** The full-pipeline result paired with its PROJECT-020 compatibility
 * report (same document/schedule as `importMppFromFile`, plus the
 * canonical compatibility summary over the SAME staged diagnostics). */
export interface MppCompatibilityImportResult {
  readonly document: ProjectDocument
  readonly schedule?: DerivedSchedule
  readonly report: CompatibilityReport
}

/**
 * Import an MPP file with the canonical compatibility report (PROJECT-020).
 *
 * Runs the production pipeline unchanged — sidecar → N1–N5 normalization →
 * accepted MSPDI importer → canonical validation → `schedule()` — then
 * aggregates the staged diagnostics into the `CompatibilityReport` via the
 * foundation `buildCompatibilityReport` (host → foundation direction
 * only; the report itself is host-neutral). The scheduling channel is the
 * already-staged scheduling diagnostics of the full pipeline: scheduling
 * status is 'success'/'failure' only over an authoritative document — the
 * trivial scheduling of the atomic empty document after a fatal conversion
 * is NOT an attempt, and the report says 'not-attempted'.
 */
export async function importMppFromFileWithCompatibility(
  inputPath: string,
  options: MppImportOptions,
): Promise<MppCompatibilityImportResult> {
  const result = await importMppFromFile(inputPath, options)
  // Split the staged diagnostics by channel: the scheduling-stage entries
  // enter via the scheduling channel (their presence — even empty — signals
  // the scheduler ran), everything else via the import channel. (The full
  // pipeline ALWAYS schedules when it reaches the tail, so presence here is
  // honest; a fatal conversion short-circuits earlier and the aggregator's
  // own rule marks scheduling 'not-attempted'.)
  const schedulingDiagnostics = result.diagnostics.filter((d) => d.stage === 'scheduling')
  const importDiagnostics = result.diagnostics.filter((d) => d.stage !== 'scheduling')
  const report = buildCompatibilityReport({
    format: 'mpp',
    ...(result.sourceFormat !== undefined ? { sourceVersion: result.sourceFormat } : {}),
    diagnostics: importDiagnostics,
    schedulingDiagnostics,
  })
  return {
    document: result.document,
    ...(result.schedule !== undefined ? { schedule: result.schedule } : {}),
    report,
  }
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

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
