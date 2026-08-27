/**
 * PROJECT-020 — import compatibility diagnostics test suite.
 *
 * The 40 required compatibility tests (brief §REQUIRED TEST SUITE, numbered
 * in each test's leading comment) plus the discipline guards (classification
 * lockstep, architecture boundaries, security, aggregation determinism at
 * scale).
 *
 * Conventions:
 *   - Every golden asserts the COMPLETE diagnostic object (code, severity,
 *     message, format, stage, loss, recoverability, entityType, entityId) —
 *     never just a code.
 *   - The real canonical `schedule` is INJECTED at the test layer (the
 *     package itself stays scheduling-free — the accepted dependency edge).
 *   - Determinism is proven by repeat runs (3× on valid fixtures): report
 *     JSON, canonical document bytes, and DerivedSchedule JSON must all be
 *     byte-identical.
 */
import { describe, expect, it } from 'vitest'
import { schedule } from '@genoffice/project-scheduling'
import modelSource from '../src/compatibility/model.ts?raw'
import classificationSource from '../src/compatibility/classification.ts?raw'
import aggregateSource from '../src/compatibility/aggregate.ts?raw'
import pipelineSource from '../src/compatibility/pipeline.ts?raw'
import compatibilityIndexSource from '../src/compatibility/index.ts?raw'
import type { ImportDiagnostic, ProjectDocument } from '@genoffice/project-contracts'
import type { MppConversionOutcome } from '../src/mpp/types.js'
import {
  GPROJ_DIAGNOSTIC_CODES,
  MSPDI_DIAGNOSTIC_CODES,
  MPP_DIAGNOSTIC_CODES,
  emptyProjectDocument,
  serializeGproj,
  deserializeGproj,
  buildCompatibilityReport,
  classifyImportDiagnostic,
  COMPATIBILITY_CODE_CLASSIFICATIONS,
  COMPATIBILITY_STAGE_PIPELINE,
  deriveCompatibilityEntityType,
  importGprojWithCompatibility,
  importMspdiWithCompatibility,
  importMppWithCompatibility,
  isFatalImport,
  MPP_DROPPED_UNASSIGNED_ASSIGNMENT,
  MPP_NORMALIZED_BASE_CALENDAR_SENTINEL,
  MPP_NORMALIZED_MIDNIGHT_PERIOD,
  MPP_NORMALIZED_PLACEHOLDER_RECORD,
  MPP_NORMALIZED_SENTINEL_REFERENCE,
  MPP_SIDECAR_EXIT,
  MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE,
  MPP_SIDECAR_RESPONSE_INVALID,
  MPP_SIDECAR_TIMEOUT,
  MPP_SIDECAR_UNAVAILABLE,
  MPP_INPUT_TOO_LARGE,
  MPP_INPUT_UNREADABLE,
  MPP_UNSUPPORTED_FORMAT,
  MSPDI_BASELINE_CAPTURED_AT_APPROXIMATED,
  MSPDI_PHYSICAL_PERCENT_COMPLETE_DROPPED,
  MSPDI_READ,
  INVALID_GPROJ,
  INVALID_MSPDI,
  INVALID_MSPDI_DATE,
  INVALID_MSPDI_DURATION,
  INVALID_MSPDI_REFERENCE,
  UNSUPPORTED_GPROJ_VERSION,
  UNSUPPORTED_MSPDI_FEATURE,
} from '../src/index.js'
import {
  baselineAndPhysicalMspdi,
  d01CleanGproj,
  d02MalformedGproj,
  d03UnsupportedGprojVersion,
  d04CleanMspdi,
  d05UnsupportedFeatureMspdi,
  d06InvalidReferenceMspdi,
  d07CalendarDegradationMspdi,
  d08CleanMppOutcome,
  d09N1Outcome,
  d10N2Outcome,
  d11N3Outcome,
  d12N4Outcome,
  d13N5Outcome,
  d14UnsupportedVersionOutcome,
  d15SidecarFailureOutcome,
  d16NetworkIsolationOutcome,
  d17CanonicalFailureMspdi,
  d19CompositeOutcome,
  inputTooLargeOutcome,
  inputUnreadableOutcome,
  malformedDateMspdi,
  malformedDurationMspdi,
  malformedFactorMspdi,
  malformedMspdiOutputOutcome,
  mspdiDocument,
  sidecarResponseInvalidOutcome,
  sidecarTimeoutOutcome,
  sidecarUnavailableOutcome,
  taskXml,
} from './compatibility-goldens.js'

/** JSON canonical form for byte-identity assertions. */
const json = (value: unknown): string => JSON.stringify(value)

const SCHEDULED = { schedule: (document: ProjectDocument) => schedule(document) }

// ─────────────────────────────────────────────────────────────────────────────
// GPROJ compatibility (required tests 1–3)
// ─────────────────────────────────────────────────────────────────────────────

describe('PROJECT-020 — .gproj compatibility', () => {
  it('1. clean GPROJ import: success on every dimension, one preserved read sentinel', () => {
    const { document, report } = importGprojWithCompatibility(d01CleanGproj(), undefined, SCHEDULED)
    expect(report.format).toBe('gproj')
    expect(report.sourceVersion).toBe('1')
    expect(report.status).toEqual({
      import: 'success',
      validation: 'success',
      scheduling: 'success',
    })
    expect(report.authoritative).toBe(true)
    expect(report.saveEligibility).toBe('allowed')
    expect(report.diagnostics).toEqual([
      {
        code: 'GPROJ_READ',
        severity: 'info',
        message: 'Read .gproj formatVersion 1',
        format: 'gproj',
        stage: 'gproj',
        loss: 'none',
        recoverability: 'preserved',
      },
    ])
    expect(report.preservedCount).toBe(1)
    expect(report.normalizedCount).toBe(0)
    expect(report.approximatedCount).toBe(0)
    expect(report.droppedCount).toBe(0)
    expect(report.unsupportedCount).toBe(0)
    expect(report.errorCount).toBe(0)
    expect(report.warningCount).toBe(0)
    expect(document.tasks).toHaveLength(3)
  })

  it('2. GPROJ malformed envelope: fatal, no authoritative document, no fabricated sourceVersion', () => {
    const { document, report } = importGprojWithCompatibility(
      d02MalformedGproj(),
      undefined,
      SCHEDULED,
    )
    expect(report.status).toEqual({
      import: 'failure',
      validation: 'not-attempted',
      scheduling: 'not-attempted',
    })
    expect(report.authoritative).toBe(false)
    expect(report.saveEligibility).toBe('prohibited')
    expect(report.sourceVersion).toBeUndefined()
    expect('sourceVersion' in report).toBe(false)
    expect(report.diagnostics).toEqual([
      {
        code: INVALID_GPROJ,
        severity: 'error',
        message: 'Input is not valid JSON',
        format: 'gproj',
        stage: 'gproj',
        loss: 'invalid',
        recoverability: 'fatal',
      },
    ])
    expect(report.errorCount).toBe(1)
    expect(json(document)).toBe(json(emptyProjectDocument()))
  })

  it('3. GPROJ unsupported version: fatal refusal with the honest sourceVersion', () => {
    const { document, report } = importGprojWithCompatibility(
      d03UnsupportedGprojVersion(),
      undefined,
      SCHEDULED,
    )
    expect(report.sourceVersion).toBe('99')
    expect(report.status.import).toBe('failure')
    expect(report.status.validation).toBe('not-attempted')
    expect(report.diagnostics).toEqual([
      {
        code: UNSUPPORTED_GPROJ_VERSION,
        severity: 'error',
        message: 'Unsupported .gproj format version: 99',
        format: 'gproj',
        stage: 'gproj',
        loss: 'unsupported',
        recoverability: 'fatal',
      },
    ])
    expect(report.unsupportedCount).toBe(1)
    expect(json(document)).toBe(json(emptyProjectDocument()))
  })
})

// ─────────────────────────────────────────────────────────────────────────────
// MSPDI compatibility (required tests 4–10)
// ─────────────────────────────────────────────────────────────────────────────

describe('PROJECT-020 — MSPDI compatibility', () => {
  it('4. clean MSPDI import: success on every dimension', () => {
    const { document, report } = importMspdiWithCompatibility(d04CleanMspdi(), undefined, SCHEDULED)
    expect(report.format).toBe('mspdi')
    expect(report.sourceVersion).toBe('16')
    expect(report.status).toEqual({
      import: 'success',
      validation: 'success',
      scheduling: 'success',
    })
    expect(report.authoritative).toBe(true)
    expect(report.saveEligibility).toBe('allowed')
    expect(report.diagnostics).toEqual([
      {
        code: MSPDI_READ,
        severity: 'info',
        message: 'Imported MSPDI (SaveVersion 16) as canonical ProjectDocument schemaVersion 1',
        format: 'mspdi',
        stage: 'mspdi',
        loss: 'none',
        recoverability: 'preserved',
      },
    ])
    expect(report.preservedCount).toBe(1)
    expect(document.tasks).toHaveLength(1)
  })

  it('5. MSPDI unsupported feature (percentage lag): warning, document still authoritative', () => {
    const { report } = importMspdiWithCompatibility(
      d05UnsupportedFeatureMspdi(),
      undefined,
      SCHEDULED,
    )
    expect(report.status).toEqual({
      import: 'success-with-warnings',
      validation: 'success',
      scheduling: 'success',
    })
    expect(report.authoritative).toBe(true)
    expect(report.warningCount).toBe(1)
    expect(report.unsupportedCount).toBe(1)
    expect(report.diagnostics).toEqual([
      {
        code: UNSUPPORTED_MSPDI_FEATURE,
        severity: 'warning',
        message:
          '<PredecessorLink> from 1 has percentage lag (format 35); not faithfully representable as integer working-minutes, defaulted to 0',
        format: 'mspdi',
        stage: 'mspdi',
        loss: 'unsupported',
        recoverability: 'canonical',
      },
      {
        code: MSPDI_READ,
        severity: 'info',
        message: 'Imported MSPDI (SaveVersion 16) as canonical ProjectDocument schemaVersion 1',
        format: 'mspdi',
        stage: 'mspdi',
        loss: 'none',
        recoverability: 'preserved',
      },
    ])
  })

  it('6. MSPDI malformed date: entity-level error, partial recovery, still authoritative', () => {
    const { document, report } = importMspdiWithCompatibility(
      malformedDateMspdi(),
      undefined,
      SCHEDULED,
    )
    expect(report.status).toEqual({
      import: 'success-with-errors',
      validation: 'success',
      scheduling: 'success',
    })
    expect(report.authoritative).toBe(true)
    expect(report.diagnostics[0]).toEqual({
      code: INVALID_MSPDI_DATE,
      severity: 'error',
      message: '<Start> is not a valid ISO-8601 date: "not-a-date"',
      entityId: 't1',
      entityType: 'task',
      format: 'mspdi',
      stage: 'mspdi',
      loss: 'invalid',
      recoverability: 'partial',
    })
    // The task survives without the malformed start:
    expect(document.tasks).toHaveLength(1)
    expect(document.tasks[0]!.start).toBeUndefined()
  })

  it('7. MSPDI malformed duration: entity-level error, value recovered to zero minutes', () => {
    const { document, report } = importMspdiWithCompatibility(
      malformedDurationMspdi(),
      undefined,
      SCHEDULED,
    )
    expect(report.status.import).toBe('success-with-errors')
    expect(report.diagnostics[0]).toEqual({
      code: INVALID_MSPDI_DURATION,
      severity: 'error',
      message: '<Duration> cannot be converted to integer working minutes: "-PT8H0M0S" (invalid)',
      entityId: 't1',
      entityType: 'task',
      format: 'mspdi',
      stage: 'mspdi',
      loss: 'invalid',
      recoverability: 'partial',
    })
    expect(document.tasks[0]!.duration).toBe(0)
  })

  it('8. MSPDI invalid reference: dependency dropped with an explicit invalid-class diagnostic', () => {
    const { document, report } = importMspdiWithCompatibility(
      d06InvalidReferenceMspdi(),
      undefined,
      SCHEDULED,
    )
    expect(report.status.import).toBe('success-with-errors')
    expect(report.authoritative).toBe(true)
    expect(report.diagnostics[0]).toEqual({
      code: INVALID_MSPDI_REFERENCE,
      severity: 'error',
      message:
        '<PredecessorLink> references <PredecessorUID> 99 which is not a declared task; dependency dropped',
      format: 'mspdi',
      stage: 'mspdi',
      loss: 'invalid',
      recoverability: 'partial',
    })
    expect(document.dependencies).toHaveLength(0)
  })

  it('9. MSPDI calendar degradation (recurring exception): warning with entity provenance', () => {
    const { document, report } = importMspdiWithCompatibility(
      d07CalendarDegradationMspdi(),
      undefined,
      SCHEDULED,
    )
    expect(report.status).toEqual({
      import: 'success-with-warnings',
      validation: 'success',
      scheduling: 'success',
    })
    expect(report.diagnostics[0]).toEqual({
      code: UNSUPPORTED_MSPDI_FEATURE,
      severity: 'warning',
      message:
        '<Calendar uid=1> has a recurring exception on 2026-08-10; the canonical model only supports single-date exceptions — mapped to a single 2026-08-10 exception',
      entityId: 'c1',
      entityType: 'calendar',
      format: 'mspdi',
      stage: 'mspdi',
      loss: 'unsupported',
      recoverability: 'canonical',
    })
    expect(document.calendars[0]!.exceptions).toEqual([
      { date: '2026-08-10', periods: [{ startMinute: 540, endMinute: 780 }] },
    ])
  })

  it('10. MSPDI baseline warning: capturedAt approximated + PhysicalPercentComplete dropped (previously silent)', () => {
    const { document, report } = importMspdiWithCompatibility(
      baselineAndPhysicalMspdi(),
      undefined,
      SCHEDULED,
    )
    expect(report.status.import).toBe('success-with-warnings')
    expect(report.approximatedCount).toBe(1)
    expect(report.droppedCount).toBe(1)
    expect(report.diagnostics[0]).toEqual({
      code: MSPDI_BASELINE_CAPTURED_AT_APPROXIMATED,
      severity: 'warning',
      message:
        'MSPDI carries no per-baseline captured date; baseline slot 0 capturedAt approximated from <LastSaved> (2026-08-02T08:00:00.000Z)',
      entityId: 'b0',
      entityType: 'baseline',
      format: 'mspdi',
      stage: 'mspdi',
      loss: 'approximated',
      recoverability: 'canonical',
    })
    expect(report.diagnostics[1]).toEqual({
      code: MSPDI_PHYSICAL_PERCENT_COMPLETE_DROPPED,
      severity: 'warning',
      message:
        '<Task uid=1> carries <PhysicalPercentComplete> 25 which the canonical import does not reconstruct (documented round-trip limitation); value dropped',
      entityId: 't1',
      entityType: 'task',
      format: 'mspdi',
      stage: 'mspdi',
      loss: 'dropped',
      recoverability: 'canonical',
    })
    expect(document.baselines[0]!.capturedAt).toBe('2026-08-02T08:00:00.000Z')
    expect(document.tasks[0]!.physicalPercentComplete).toBeUndefined()
  })
})

// ─────────────────────────────────────────────────────────────────────────────
// MPP compatibility (required tests 11–24)
// ─────────────────────────────────────────────────────────────────────────────

describe('PROJECT-020 — MPP compatibility', () => {
  it('11. clean MPP import: staged provenance, success on every dimension', () => {
    const { document, report } = importMppWithCompatibility(
      d08CleanMppOutcome(),
      undefined,
      SCHEDULED,
    )
    expect(report.format).toBe('mpp')
    expect(report.sourceVersion).toBe('MPP14')
    expect(report.status).toEqual({
      import: 'success',
      validation: 'success',
      scheduling: 'success',
    })
    expect(report.authoritative).toBe(true)
    expect(report.diagnostics).toEqual([
      {
        code: MSPDI_READ,
        severity: 'info',
        message: 'Imported MSPDI (SaveVersion 16) as canonical ProjectDocument schemaVersion 1',
        format: 'mpp',
        stage: 'mspdi',
        loss: 'none',
        recoverability: 'preserved',
      },
    ])
    expect(document.tasks).toHaveLength(1)
  })

  it('12. MPP N1 normalization: sentinel CalendarUID stripped, semantics preserved', () => {
    const { document, report } = importMppWithCompatibility(d09N1Outcome(), undefined, SCHEDULED)
    expect(report.status.import).toBe('success')
    expect(report.normalizedCount).toBe(1)
    expect(report.diagnostics[0]).toEqual({
      code: MPP_NORMALIZED_SENTINEL_REFERENCE,
      severity: 'info',
      message:
        'MPP sentinel CalendarUID -1 ("inherit the default calendar") was stripped from a task before import',
      entityId: 't1',
      entityType: 'task',
      format: 'mpp',
      stage: 'normalization',
      loss: 'normalized',
      recoverability: 'preserved',
    })
    expect(document.tasks[0]!.calendarId).toBeUndefined()
  })

  it('13. MPP N2 normalization: sentinel BaseCalendarUID stripped', () => {
    const { report } = importMppWithCompatibility(d10N2Outcome(), undefined, SCHEDULED)
    expect(report.normalizedCount).toBe(1)
    expect(report.diagnostics[0]).toEqual({
      code: MPP_NORMALIZED_BASE_CALENDAR_SENTINEL,
      severity: 'info',
      message:
        'MPP sentinel BaseCalendarUID -1 ("no base calendar") was stripped from a calendar before import',
      entityId: 'c1',
      entityType: 'calendar',
      format: 'mpp',
      stage: 'normalization',
      loss: 'normalized',
      recoverability: 'preserved',
    })
  })

  it('14. MPP N3 normalization: placeholder task and resource filtered', () => {
    const { document, report } = importMppWithCompatibility(d11N3Outcome(), undefined, SCHEDULED)
    expect(report.normalizedCount).toBe(2)
    const placeholder = report.diagnostics.filter(
      (d) => d.code === MPP_NORMALIZED_PLACEHOLDER_RECORD,
    )
    // Canonical sort: same stage/severity/code → entityType 'resource'
    // sorts before 'task':
    expect(placeholder.map((d) => d.entityId)).toEqual(['r0', 't0'])
    expect(placeholder.every((d) => d.stage === 'normalization' && d.loss === 'normalized')).toBe(
      true,
    )
    expect(document.tasks.map((t) => String(t.id))).toEqual(['t1'])
    expect(document.resources).toHaveLength(0)
  })

  it('15. MPP N4 normalization: midnight working periods rewritten to the day-end', () => {
    const { document, report } = importMppWithCompatibility(d12N4Outcome(), undefined, SCHEDULED)
    // Five weekday periods rewritten — five DISTINCT diagnostics (countable
    // occurrences, never collapsed by de-duplication):
    expect(
      report.diagnostics.filter((d) => d.code === MPP_NORMALIZED_MIDNIGHT_PERIOD),
    ).toHaveLength(5)
    expect(report.normalizedCount).toBe(5)
    expect(
      report.diagnostics
        .filter((d) => d.code === MPP_NORMALIZED_MIDNIGHT_PERIOD)
        .every(
          (d) => d.loss === 'normalized' && d.severity === 'info' && d.stage === 'normalization',
        ),
    ).toBe(true)
    expect(document.calendars[0]!.workingWeek[2]).toEqual([{ startMinute: 1320, endMinute: 1440 }])
  })

  it('16. MPP N5 normalization: unassigned placeholder assignment dropped as diagnosed loss', () => {
    const { document, report } = importMppWithCompatibility(d13N5Outcome(), undefined, SCHEDULED)
    expect(report.status.import).toBe('success-with-warnings')
    expect(report.droppedCount).toBe(1)
    expect(report.warningCount).toBe(1)
    expect(report.diagnostics[0]).toEqual({
      code: MPP_DROPPED_UNASSIGNED_ASSIGNMENT,
      severity: 'warning',
      message:
        'MPP "unassigned" placeholder assignment (ResourceUID -65535) was dropped — the canonical model has no unassigned assignment (expected loss)',
      entityId: 't1',
      entityType: 'task',
      format: 'mpp',
      stage: 'normalization',
      loss: 'dropped',
      recoverability: 'canonical',
    })
    expect(document.assignments).toHaveLength(0)
  })

  it('17. unsupported MPP version: sidecar refusal is fatal and classified unsupported', () => {
    const { document, report } = importMppWithCompatibility(
      d14UnsupportedVersionOutcome(),
      undefined,
      SCHEDULED,
    )
    expect(report.status).toEqual({
      import: 'failure',
      validation: 'not-attempted',
      scheduling: 'not-attempted',
    })
    expect(report.authoritative).toBe(false)
    expect(report.saveEligibility).toBe('prohibited')
    expect(report.unsupportedCount).toBe(1)
    expect(report.diagnostics).toEqual([
      {
        code: MPP_UNSUPPORTED_FORMAT,
        severity: 'error',
        message:
          'MPP input is not a recognized project format (no MPP8/MPP9/MPP12/MPP14 container)',
        format: 'mpp',
        stage: 'sidecar',
        loss: 'unsupported',
        recoverability: 'fatal',
      },
    ])
    expect(json(document)).toBe(json(emptyProjectDocument()))
  })

  it('18. MPP unreadable input: readability failure, distinct from size failure', () => {
    const { report } = importMppWithCompatibility(inputUnreadableOutcome(), undefined, SCHEDULED)
    expect(report.diagnostics[0]!.code).toBe(MPP_INPUT_UNREADABLE)
    expect(report.diagnostics[0]!.stage).toBe('sidecar')
    expect(report.diagnostics[0]!.loss).toBe('invalid')
    expect(report.diagnostics[0]!.recoverability).toBe('fatal')
    expect(report.status.import).toBe('failure')
  })

  it('19. MPP oversized input: size failure, distinct from readability failure', () => {
    const { report } = importMppWithCompatibility(inputTooLargeOutcome(), undefined, SCHEDULED)
    expect(report.diagnostics[0]!.code).toBe(MPP_INPUT_TOO_LARGE)
    expect(report.diagnostics[0]!.loss).toBe('invalid')
    expect(report.diagnostics[0]!.recoverability).toBe('fatal')
    expect(report.status.import).toBe('failure')
  })

  it('20. sidecar unavailable: fatal sidecar-stage failure', () => {
    const { report } = importMppWithCompatibility(sidecarUnavailableOutcome(), undefined, SCHEDULED)
    expect(report.diagnostics[0]!.code).toBe(MPP_SIDECAR_UNAVAILABLE)
    expect(report.diagnostics[0]!.stage).toBe('sidecar')
    expect(report.diagnostics[0]!.recoverability).toBe('fatal')
    expect(report.status).toEqual({
      import: 'failure',
      validation: 'not-attempted',
      scheduling: 'not-attempted',
    })
  })

  it('21. sidecar timeout: fatal sidecar-stage failure', () => {
    const { report } = importMppWithCompatibility(sidecarTimeoutOutcome(), undefined, SCHEDULED)
    expect(report.diagnostics[0]!.code).toBe(MPP_SIDECAR_TIMEOUT)
    expect(report.diagnostics[0]!.recoverability).toBe('fatal')
    expect(report.status.import).toBe('failure')
  })

  it('22. network isolation unavailable: the fail-closed refusal is fatal', () => {
    const { report } = importMppWithCompatibility(
      d16NetworkIsolationOutcome(),
      undefined,
      SCHEDULED,
    )
    expect(report.diagnostics).toEqual([
      {
        code: MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE,
        severity: 'error',
        message:
          'network isolation required but the host cannot provide the mechanism; conversion refused (fail closed)',
        format: 'mpp',
        stage: 'sidecar',
        loss: 'invalid',
        recoverability: 'fatal',
      },
    ])
    expect(report.authoritative).toBe(false)
  })

  it('23. malformed sidecar response: protocol failure is fatal', () => {
    const { report } = importMppWithCompatibility(
      sidecarResponseInvalidOutcome(),
      undefined,
      SCHEDULED,
    )
    expect(report.diagnostics[0]!.code).toBe(MPP_SIDECAR_RESPONSE_INVALID)
    expect(report.diagnostics[0]!.stage).toBe('sidecar')
    expect(report.diagnostics[0]!.recoverability).toBe('fatal')
    expect(report.status.import).toBe('failure')
  })

  it('24. malformed MSPDI output: the importer error surfaces through the MPP pipeline (fatal)', () => {
    const { document, report } = importMppWithCompatibility(
      malformedMspdiOutputOutcome(),
      undefined,
      SCHEDULED,
    )
    // The normalizer passes malformed XML through untouched; the accepted
    // importer reports it — staged 'mspdi', fatal for the MPP pipeline. The
    // accepted pipeline then validates the ATOMIC EMPTY document, whose
    // canonical-stage error (MISSING_CALENDAR — the empty document has no
    // calendars) is an atomicity artifact the report records honestly:
    expect(report.status).toEqual({
      import: 'failure',
      validation: 'failure',
      scheduling: 'not-attempted',
    })
    expect(report.diagnostics[0]!.code).toBe(INVALID_MSPDI)
    expect(report.diagnostics[0]!.stage).toBe('mspdi')
    expect(report.diagnostics[0]!.loss).toBe('invalid')
    expect(report.diagnostics[0]!.recoverability).toBe('fatal')
    expect(report.diagnostics[1]!.code).toBe('MISSING_CALENDAR')
    expect(report.diagnostics[1]!.stage).toBe('canonical')
    expect(report.diagnostics[1]!.recoverability).toBe('fatal')
    expect(json(document)).toBe(json(emptyProjectDocument()))
    // The sidecar-failure code family (D15) is equally fatal (and its
    // refusal short-circuits BEFORE any canonical validation):
    const exit = importMppWithCompatibility(d15SidecarFailureOutcome(), undefined, SCHEDULED)
    expect(exit.report.diagnostics[0]!.code).toBe(MPP_SIDECAR_EXIT)
    expect(exit.report.diagnostics[0]!.stage).toBe('sidecar')
    expect(exit.report.status).toEqual({
      import: 'failure',
      validation: 'not-attempted',
      scheduling: 'not-attempted',
    })
  })
})

// ─────────────────────────────────────────────────────────────────────────────
// Status dimensions, atomicity, save eligibility (required tests 25–26, 34–35, 40)
// ─────────────────────────────────────────────────────────────────────────────

describe('PROJECT-020 — status dimensions, atomicity, save eligibility', () => {
  it('25. canonical validation failure: import succeeds structurally, validation fails, scheduling not attempted', () => {
    const { document, report } = importMspdiWithCompatibility(
      d17CanonicalFailureMspdi(),
      undefined,
      SCHEDULED,
    )
    // The brief's canonical example shape:
    expect(report.status).toEqual({
      import: 'success',
      validation: 'failure',
      scheduling: 'not-attempted',
    })
    expect(report.authoritative).toBe(false)
    expect(report.saveEligibility).toBe('prohibited')
    expect(report.errorCount).toBe(1)
    // Stage order: the mspdi read sentinel (stage 'mspdi') sorts before the
    // canonical-stage rejection:
    expect(report.diagnostics[0]!.code).toBe(MSPDI_READ)
    expect(report.diagnostics[1]).toEqual({
      code: 'DEPENDENCY_CYCLE',
      severity: 'error',
      message: 'Dependency graph contains a cycle',
      format: 'mspdi',
      stage: 'canonical',
      loss: 'invalid',
      recoverability: 'partial',
    })
    // The constructed (invalid) document is still returned by the accepted
    // adapter — the report says it is NOT authoritative:
    expect(document.tasks).toHaveLength(2)
  })

  it('25b. MPP canonical rejection is ATOMIC: import failure, empty document, full provenance', () => {
    const { document, report } = importMppWithCompatibility(
      { ...d08CleanMppOutcome(), mspdiBytes: d17CanonicalFailureMspdi() },
      undefined,
      SCHEDULED,
    )
    // The MPP pipeline returns the atomic empty document on canonical
    // rejection (accepted PROJECT-018 semantics) — the report states it:
    expect(report.status).toEqual({
      import: 'failure',
      validation: 'failure',
      scheduling: 'not-attempted',
    })
    expect(report.authoritative).toBe(false)
    expect(json(document)).toBe(json(emptyProjectDocument()))
    expect(
      report.diagnostics.some((d) => d.code === 'DEPENDENCY_CYCLE' && d.stage === 'canonical'),
    ).toBe(true)
  })

  it('26. scheduling failure: derived-state failure, document stays authoritative and save-eligible', () => {
    const failingRunner = (): { diagnostics: ImportDiagnostic[] } => ({
      diagnostics: [
        {
          code: 'CALENDAR_SEARCH_EXHAUSTED',
          severity: 'error',
          message: 'calendar working-time search exhausted the horizon',
        },
      ],
    })
    const { document, report } = importMspdiWithCompatibility(d04CleanMspdi(), undefined, {
      schedule: failingRunner,
    })
    expect(report.status).toEqual({
      import: 'success',
      validation: 'success',
      scheduling: 'failure',
    })
    expect(report.authoritative).toBe(true)
    expect(report.saveEligibility).toBe('allowed')
    // Stage order: the mspdi read sentinel sorts before the scheduling error:
    expect(report.diagnostics[0]!.code).toBe(MSPDI_READ)
    expect(report.diagnostics[1]).toEqual({
      code: 'CALENDAR_SEARCH_EXHAUSTED',
      severity: 'error',
      message: 'calendar working-time search exhausted the horizon',
      format: 'mspdi',
      stage: 'scheduling',
      loss: 'invalid',
      recoverability: 'canonical',
    })
    expect(document.tasks).toHaveLength(1)
  })

  it('34. warning import can save .gproj: save allowed and the round-trip preserves semantics', () => {
    const { document, report } = importMspdiWithCompatibility(
      d05UnsupportedFeatureMspdi(),
      undefined,
      SCHEDULED,
    )
    expect(report.saveEligibility).toBe('allowed')
    expect(report.status.validation).toBe('success')
    const bytes = serializeGproj(document)
    const reopened = deserializeGproj(bytes)
    expect(reopened.diagnostics.filter((d) => d.severity === 'error')).toEqual([])
    expect(serializeGproj(reopened.document)).toEqual(bytes)
  })

  it('35. invalid import cannot save: canonical validation errors prohibit save', () => {
    const { report } = importMspdiWithCompatibility(
      d17CanonicalFailureMspdi(),
      undefined,
      SCHEDULED,
    )
    expect(report.status.validation).toBe('failure')
    expect(report.saveEligibility).toBe('prohibited')
    expect(report.authoritative).toBe(false)
    // Every fatal fixture is equally prohibited:
    for (const fatal of [d02MalformedGproj(), d03UnsupportedGprojVersion()] as const) {
      expect(importGprojWithCompatibility(fatal, undefined, SCHEDULED).report.saveEligibility).toBe(
        'prohibited',
      )
    }
  })

  it('40. fatal errors never yield an authoritative document', () => {
    const fatalMpp: Array<[string, MppConversionOutcome]> = [
      ['MPP_UNSUPPORTED_FORMAT', d14UnsupportedVersionOutcome()],
      ['MPP_SIDECAR_EXIT', d15SidecarFailureOutcome()],
      ['MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE', d16NetworkIsolationOutcome()],
      ['MPP_SIDECAR_UNAVAILABLE', sidecarUnavailableOutcome()],
      ['MPP_SIDECAR_TIMEOUT', sidecarTimeoutOutcome()],
      ['MPP_SIDECAR_RESPONSE_INVALID', sidecarResponseInvalidOutcome()],
      ['MPP_INPUT_UNREADABLE', inputUnreadableOutcome()],
      ['MPP_INPUT_TOO_LARGE', inputTooLargeOutcome()],
      ['INVALID_MSPDI (malformed sidecar output)', malformedMspdiOutputOutcome()],
    ]
    for (const [label, outcome] of fatalMpp) {
      const { document, report } = importMppWithCompatibility(outcome, undefined, SCHEDULED)
      expect(report.authoritative, label).toBe(false)
      expect(report.status.import, label).toBe('failure')
      expect(json(document), label).toBe(json(emptyProjectDocument()))
      expect(
        report.diagnostics.every((d) => d.recoverability === 'fatal'),
        label,
      ).toBe(true)
      expect(report.saveEligibility, label).toBe('prohibited')
    }
  })
})

// ─────────────────────────────────────────────────────────────────────────────
// Classification, ordering, dedup, determinism (required tests 27–33, 36–39)
// ─────────────────────────────────────────────────────────────────────────────

describe('PROJECT-020 — classification, ordering, de-duplication, determinism', () => {
  it('27. diagnostic provenance: every stage is represented with its format and entity association', () => {
    const composite = importMppWithCompatibility(d19CompositeOutcome(), undefined, SCHEDULED)
    const stages = composite.report.diagnostics.map((d) => d.stage)
    expect(stages).toContain('normalization')
    expect(stages).toContain('mspdi')
    expect(new Set(composite.report.diagnostics.map((d) => d.format))).toEqual(new Set(['mpp']))
    // canonical + scheduling provenance come from the other paths:
    const canonical = importMspdiWithCompatibility(d17CanonicalFailureMspdi(), undefined, SCHEDULED)
    expect(canonical.report.diagnostics.some((d) => d.stage === 'canonical')).toBe(true)
    const scheduling = importMspdiWithCompatibility(d04CleanMspdi(), undefined, {
      schedule: () => ({ diagnostics: [{ code: 'CYCLE', severity: 'error', message: 'cycle' }] }),
    })
    expect(scheduling.report.diagnostics.some((d) => d.stage === 'scheduling')).toBe(true)
    const sidecar = importMppWithCompatibility(d15SidecarFailureOutcome(), undefined, SCHEDULED)
    expect(sidecar.report.diagnostics.every((d) => d.stage === 'sidecar')).toBe(true)
    // And the documented pipeline composition per format:
    expect(COMPATIBILITY_STAGE_PIPELINE.mpp).toEqual([
      'sidecar',
      'normalization',
      'mspdi',
      'canonical',
      'scheduling',
    ])
    expect(COMPATIBILITY_STAGE_PIPELINE.mspdi).toEqual(['mspdi', 'canonical', 'scheduling'])
    expect(COMPATIBILITY_STAGE_PIPELINE.gproj).toEqual(['gproj', 'canonical', 'scheduling'])
  })

  it('28. severity classification: info preserves, warning degrades, error rejects — recoverability follows', () => {
    const { report } = importMppWithCompatibility(d19CompositeOutcome(), undefined, SCHEDULED)
    const byCode = new Map(report.diagnostics.map((d) => [d.code, d]))
    // info = mechanical normalization that preserves semantics (N1):
    const n1 = byCode.get(MPP_NORMALIZED_SENTINEL_REFERENCE)!
    expect(n1.severity).toBe('info')
    expect(n1.recoverability).toBe('preserved')
    expect(n1.loss).toBe('normalized')
    // info bookkeeping is lossless (the read sentinel):
    expect(byCode.get(MSPDI_READ)!.loss).toBe('none')
    // warning = semantic loss, document still canonical (N5):
    const n5 = byCode.get(MPP_DROPPED_UNASSIGNED_ASSIGNMENT)!
    expect(n5.severity).toBe('warning')
    expect(n5.recoverability).toBe('canonical')
    // error = rejection; non-fatal errors are partial, fatal ones fatal:
    const nonFatalError = importMspdiWithCompatibility(malformedDateMspdi(), undefined, SCHEDULED)
    expect(nonFatalError.report.diagnostics[0]!.severity).toBe('error')
    expect(nonFatalError.report.diagnostics[0]!.recoverability).toBe('partial')
    const fatalError = importGprojWithCompatibility(d02MalformedGproj(), undefined, SCHEDULED)
    expect(fatalError.report.diagnostics[0]!.recoverability).toBe('fatal')
  })

  it('29. data-loss classification: one real instance of every class', () => {
    const clean = importGprojWithCompatibility(d01CleanGproj(), undefined, SCHEDULED).report
    const n1 = importMppWithCompatibility(d09N1Outcome(), undefined, SCHEDULED).report
    const baseline = importMspdiWithCompatibility(
      baselineAndPhysicalMspdi(),
      undefined,
      SCHEDULED,
    ).report
    const n5 = importMppWithCompatibility(d13N5Outcome(), undefined, SCHEDULED).report
    const unsupported = importMspdiWithCompatibility(
      d05UnsupportedFeatureMspdi(),
      undefined,
      SCHEDULED,
    ).report
    const invalid = importMspdiWithCompatibility(
      d06InvalidReferenceMspdi(),
      undefined,
      SCHEDULED,
    ).report
    expect(clean.diagnostics[0]!.loss).toBe('none')
    expect(n1.diagnostics[0]!.loss).toBe('normalized')
    expect(baseline.diagnostics[0]!.loss).toBe('approximated')
    expect(n5.diagnostics[0]!.loss).toBe('dropped')
    expect(unsupported.diagnostics[0]!.loss).toBe('unsupported')
    expect(invalid.diagnostics[0]!.loss).toBe('invalid')
  })

  it('30. deterministic diagnostic ordering: the canonical sort key (stage → severity → code → entityType → entityId → message)', () => {
    const { report } = importMppWithCompatibility(d19CompositeOutcome(), undefined, SCHEDULED)
    expect(
      report.diagnostics.map((d) => `${d.stage}:${d.severity}:${d.code}:${d.entityId ?? ''}`),
    ).toEqual([
      'normalization:warning:MPP_DROPPED_UNASSIGNED_ASSIGNMENT:t1',
      'normalization:info:MPP_NORMALIZED_MIDNIGHT_PERIOD:',
      'normalization:info:MPP_NORMALIZED_MIDNIGHT_PERIOD:',
      'normalization:info:MPP_NORMALIZED_MIDNIGHT_PERIOD:',
      'normalization:info:MPP_NORMALIZED_MIDNIGHT_PERIOD:',
      'normalization:info:MPP_NORMALIZED_MIDNIGHT_PERIOD:',
      'normalization:info:MPP_NORMALIZED_PLACEHOLDER_RECORD:t0',
      'normalization:info:MPP_NORMALIZED_SENTINEL_REFERENCE:t1',
      'mspdi:warning:MSPDI_BASELINE_CAPTURED_AT_APPROXIMATED:b0',
      'mspdi:warning:MSPDI_PHYSICAL_PERCENT_COMPLETE_DROPPED:t1',
      'mspdi:info:MSPDI_READ:',
    ])
  })

  it('31. deterministic de-duplication policy: declaration-level uniqueness by the producer, occurrences never collapsed', () => {
    // (a) One malformed <MinutesPerDay> declaration used by THREE
    // dependencies → exactly ONE declaration diagnostic (producer contract):
    const { report } = importMspdiWithCompatibility(malformedFactorMspdi(), undefined, SCHEDULED)
    const factorDiags = report.diagnostics.filter((d) => d.code === INVALID_MSPDI)
    expect(factorDiags).toHaveLength(1)
    expect(factorDiags[0]!.message).toContain('MinutesPerDay')
    expect(report.diagnostics.filter((d) => d.code === INVALID_MSPDI_REFERENCE)).toHaveLength(0)
    // (b) Five distinct midnight-period occurrences → five diagnostics
    // (identical objects preserved — their multiplicity is information):
    const n4 = importMppWithCompatibility(d12N4Outcome(), undefined, SCHEDULED).report
    expect(n4.diagnostics.filter((d) => d.code === MPP_NORMALIZED_MIDNIGHT_PERIOD)).toHaveLength(5)
  })

  it('32. repeated import gives identical diagnostics, document, and schedule (3× repeat)', () => {
    for (const input of [d01CleanGproj(), d04CleanMspdi(), baselineAndPhysicalMspdi()] as const) {
      const first = importMspdiOrGproj(input, SCHEDULED)
      for (let run = 2; run <= 3; run++) {
        const again = importMspdiOrGproj(input, SCHEDULED)
        expect(json(again.report)).toBe(json(first.report))
        expect(serializeGproj(again.document)).toEqual(serializeGproj(first.document))
        expect(json(schedule(again.document))).toBe(json(schedule(first.document)))
      }
    }
    const mppFirst = importMppWithCompatibility(d19CompositeOutcome(), undefined, SCHEDULED)
    for (let run = 2; run <= 3; run++) {
      const again = importMppWithCompatibility(d19CompositeOutcome(), undefined, SCHEDULED)
      expect(json(again.report)).toBe(json(mppFirst.report))
      expect(serializeGproj(again.document)).toEqual(serializeGproj(mppFirst.document))
      expect(json(schedule(again.document))).toBe(json(schedule(mppFirst.document)))
    }
  })

  it('33. reordered source elements give canonical diagnostic stability', () => {
    const taskA = taskXml(2, '<Start>not-a-date</Start>')
    const taskB = taskXml(3, '<Finish>also-not-a-date</Finish>')
    const first = importMspdiWithCompatibility(
      mspdiDocument({ tasks: taskXml(1) + taskA + taskB }),
      undefined,
      SCHEDULED,
    )
    const reordered = importMspdiWithCompatibility(
      mspdiDocument({ tasks: taskB + taskA + taskXml(1) }),
      undefined,
      SCHEDULED,
    )
    // Entity-keyed diagnostics sort identically regardless of source order
    // (canonical identity is not array position — architecture-lock §4):
    expect(json(reordered.report)).toBe(json(first.report))
    expect(reordered.report.diagnostics.map((d) => d.entityId)).toEqual(['t2', 't3', undefined])
    expect(reordered.document.tasks.map((t) => String(t.id)).sort()).toEqual(
      first.document.tasks.map((t) => String(t.id)).sort(),
    )
    // Same-entity child-element reordering keeps the document byte-identical
    // (the accepted PROJECT-015 invariance):
    const childA = taskXml(4, '<Deadline>2026-09-01T09:00:00</Deadline>')
    const childB = `<Task><Deadline>2026-09-01T09:00:00</Deadline><OutlineNumber>4</OutlineNumber><RemainingCost>0</RemainingCost><ActualCost>0</ActualCost><Cost>0</Cost><ActualWork>PT0H0M0S</ActualWork><RemainingWork>PT8H0M0S</RemainingWork><Work>PT0H0M0S</Work><PercentComplete>0</PercentComplete><Priority>500</Priority><Manual>false</Manual><Milestone>false</Milestone><Summary>false</Summary><OutlineLevel>1</OutlineLevel><Duration>PT8H0M0S</Duration><Name>Task 4</Name><UID>4</UID><ID>4</ID></Task>`
    const docA = importMspdiWithCompatibility(
      mspdiDocument({ tasks: childA }),
      undefined,
      SCHEDULED,
    )
    const docB = importMspdiWithCompatibility(
      mspdiDocument({ tasks: childB }),
      undefined,
      SCHEDULED,
    )
    expect(serializeGproj(docB.document)).toEqual(serializeGproj(docA.document))
    expect(json(docB.report)).toBe(json(docA.report))
  })

  it('36. compatibility summary counts: the complete D19 composite report', () => {
    const { report } = importMppWithCompatibility(d19CompositeOutcome(), undefined, SCHEDULED)
    expect(report.sourceVersion).toBe('MPP9')
    expect(report.status).toEqual({
      import: 'success-with-warnings',
      validation: 'success',
      scheduling: 'success',
    })
    expect(report.authoritative).toBe(true)
    expect(report.saveEligibility).toBe('allowed')
    expect(report.diagnostics).toHaveLength(11)
    expect(report.preservedCount).toBe(1)
    expect(report.normalizedCount).toBe(7)
    expect(report.approximatedCount).toBe(1)
    expect(report.droppedCount).toBe(2)
    expect(report.unsupportedCount).toBe(0)
    expect(report.errorCount).toBe(0)
    expect(report.warningCount).toBe(3)
  })

  it('37. stage separation: the three dimensions are independent, never a single boolean', () => {
    const clean = importMspdiWithCompatibility(d04CleanMspdi(), undefined, SCHEDULED).report
    expect(clean.status).toEqual({
      import: 'success',
      validation: 'success',
      scheduling: 'success',
    })
    const validationFailure = importMspdiWithCompatibility(
      d17CanonicalFailureMspdi(),
      undefined,
      SCHEDULED,
    ).report
    expect(validationFailure.status).toEqual({
      import: 'success',
      validation: 'failure',
      scheduling: 'not-attempted',
    })
    const schedulingFailure = importMspdiWithCompatibility(d04CleanMspdi(), undefined, {
      schedule: () => ({ diagnostics: [{ code: 'CYCLE', severity: 'error', message: 'cycle' }] }),
    }).report
    expect(schedulingFailure.status).toEqual({
      import: 'success',
      validation: 'success',
      scheduling: 'failure',
    })
    const importErrors = importMspdiWithCompatibility(
      malformedDateMspdi(),
      undefined,
      SCHEDULED,
    ).report
    expect(importErrors.status).toEqual({
      import: 'success-with-errors',
      validation: 'success',
      scheduling: 'success',
    })
  })

  it('38. entity/field provenance: entityType derivation from the deterministic identity prefixes', () => {
    expect(deriveCompatibilityEntityType('t17')).toBe('task')
    expect(deriveCompatibilityEntityType('r2')).toBe('resource')
    expect(deriveCompatibilityEntityType('a3')).toBe('assignment')
    expect(deriveCompatibilityEntityType('c1')).toBe('calendar')
    expect(deriveCompatibilityEntityType('b0')).toBe('baseline')
    expect(deriveCompatibilityEntityType('d-t2-t1-FS')).toBe('dependency')
    expect(deriveCompatibilityEntityType('188743731')).toBe('custom-field')
    expect(deriveCompatibilityEntityType(undefined)).toBeUndefined()
    expect(deriveCompatibilityEntityType('hand-authored-id')).toBeUndefined()
    // And in reports: the referenced entity carries both id and type:
    const { report } = importMspdiWithCompatibility(
      d07CalendarDegradationMspdi(),
      undefined,
      SCHEDULED,
    )
    expect(report.diagnostics[0]!.entityId).toBe('c1')
    expect(report.diagnostics[0]!.entityType).toBe('calendar')
  })

  it('39. accepted normalization is never classified as silent loss: every N1–N5 application is diagnosed and classified', () => {
    const { report } = importMppWithCompatibility(d19CompositeOutcome(), undefined, SCHEDULED)
    const nCodes = [
      MPP_NORMALIZED_SENTINEL_REFERENCE,
      MPP_NORMALIZED_BASE_CALENDAR_SENTINEL,
      MPP_NORMALIZED_PLACEHOLDER_RECORD,
      MPP_NORMALIZED_MIDNIGHT_PERIOD,
      MPP_DROPPED_UNASSIGNED_ASSIGNMENT,
    ]
    for (const code of nCodes) {
      const instances = report.diagnostics.filter((d) => d.code === code)
      // N2 is the only one the composite does not trigger:
      if (code === MPP_NORMALIZED_BASE_CALENDAR_SENTINEL) {
        const n2 = importMppWithCompatibility(d10N2Outcome(), undefined, SCHEDULED).report
        expect(n2.diagnostics.some((d) => d.code === code && d.loss !== 'none')).toBe(true)
        continue
      }
      expect(instances.length, code).toBeGreaterThan(0)
      expect(
        instances.every((d) => d.loss !== 'none'),
        code,
      ).toBe(true)
    }
    // The classification table agrees for the whole N family:
    for (const code of MPP_DIAGNOSTIC_CODES) {
      if (nCodes.includes(code as (typeof nCodes)[number])) {
        expect(COMPATIBILITY_CODE_CLASSIFICATIONS[code]!.loss).not.toBe('none')
      }
    }
  })
})

// ─────────────────────────────────────────────────────────────────────────────
// Discipline: classification lockstep, architecture, security, aggregation
// ─────────────────────────────────────────────────────────────────────────────

describe('PROJECT-020 — discipline guards', () => {
  it('41a. classification table covers every exported adapter code family', () => {
    for (const code of [...GPROJ_DIAGNOSTIC_CODES, ...MPP_DIAGNOSTIC_CODES]) {
      expect(COMPATIBILITY_CODE_CLASSIFICATIONS[code], code).toBeDefined()
    }
    for (const code of MSPDI_DIAGNOSTIC_CODES) {
      expect(COMPATIBILITY_CODE_CLASSIFICATIONS[code], code).toBeDefined()
    }
  })

  it('41b. classification table stays in lockstep with the engine + scheduling sources', async () => {
    const engineSource = (await import('../../project-engine/src/document.ts?raw'))
      .default as string
    const scheduleSource = (await import('../../project-scheduling/src/schedule.ts?raw'))
      .default as string
    const graphSource = (await import('../../project-scheduling/src/graph.ts?raw'))
      .default as string
    const calendarSource = (await import('../../project-scheduling/src/calendar.ts?raw'))
      .default as string
    for (const source of [engineSource, scheduleSource, graphSource, calendarSource]) {
      const codes = new Set(source.match(/'([A-Z][A-Z_]{3,})'/g)?.map((m) => m.slice(1, -1)) ?? [])
      for (const code of codes) {
        // Non-diagnostic uppercase strings (regex constants, format roots,
        // lag-format constants…) are excluded by shape; every remaining
        // uppercase token must have a classification entry:
        if (/_RE$|_ELEMENT$|_NS$|^LAG_|^PERCENT_|^ELAPSED_|^WORKING_|^DEFAULT_/.test(code)) continue
        expect(
          COMPATIBILITY_CODE_CLASSIFICATIONS[code],
          `engine/scheduling code ${code}`,
        ).toBeDefined()
      }
    }
  })

  it('41c. classifyImportDiagnostic is deterministic for unknown codes (forward compatibility)', () => {
    const unknownError: ImportDiagnostic = {
      code: 'SOME_FUTURE_CODE',
      severity: 'error',
      message: 'x',
    }
    const a = classifyImportDiagnostic('mspdi', unknownError, false)
    const b = classifyImportDiagnostic('mspdi', unknownError, false)
    expect(a).toEqual(b)
    expect(a.stage).toBe('mspdi')
    expect(a.loss).toBe('invalid')
    expect(a.recoverability).toBe('partial')
    const unknownWarning: ImportDiagnostic = {
      code: 'ANOTHER_FUTURE_CODE',
      severity: 'warning',
      message: 'y',
    }
    expect(classifyImportDiagnostic('gproj', unknownWarning, false).loss).toBe('unsupported')
    expect(classifyImportDiagnostic('gproj', unknownWarning, false).stage).toBe('gproj')
  })

  it('42. architecture: the compatibility layer stays inside the foundation boundary', () => {
    const sources = [
      modelSource,
      classificationSource,
      aggregateSource,
      pipelineSource,
      compatibilityIndexSource,
    ]
    // Forbidden import literals are assembled (not written verbatim) so this
    // test file itself passes the CI foundation boundary grep unchanged.
    const forbiddenImports = ['node:', 'react', 'electron', 'http', 'https'].map(
      (m) => `from '${m}`,
    )
    for (const source of sources) {
      expect(source).not.toContain('project-scheduling')
      for (const pattern of forbiddenImports) {
        expect(source, pattern).not.toContain(pattern)
      }
      expect(source).not.toContain('.localeCompare(')
      expect(source).not.toContain('Date.now(')
      expect(source).not.toContain('Math.random(')
    }
  })

  it('43. security: reports expose no environment, credentials, host paths, or process internals', () => {
    const reports = [
      importGprojWithCompatibility(d01CleanGproj(), undefined, SCHEDULED).report,
      importMspdiWithCompatibility(baselineAndPhysicalMspdi(), undefined, SCHEDULED).report,
      importMppWithCompatibility(d15SidecarFailureOutcome(), undefined, SCHEDULED).report,
      importMppWithCompatibility(inputUnreadableOutcome(), undefined, SCHEDULED).report,
    ]
    for (const report of reports) {
      const text = json(report)
      expect(text).not.toContain('process.env')
      expect(text).not.toContain('HOME')
      expect(text).not.toContain('token')
      expect(text).not.toContain('secret')
      expect(text).not.toContain('credential')
      expect(text).not.toContain('/home/')
      expect(text).not.toContain('/tmp/')
      expect(text).not.toContain('node_modules')
    }
  })

  it('44. aggregation is a pure, repeatable function at scale (no quadratic de-duplication scan)', () => {
    const diagnostics: ImportDiagnostic[] = []
    for (let i = 0; i < 4000; i++) {
      diagnostics.push(
        i % 2 === 0
          ? {
              code: 'INVALID_MSPDI_DURATION',
              severity: 'error',
              message: `synthetic ${i}`,
              entityId: `t${i}`,
            }
          : { code: 'MSPDI_READ', severity: 'info', message: `synthetic ${i}` },
      )
    }
    const first = buildCompatibilityReport({ format: 'mspdi', diagnostics })
    const second = buildCompatibilityReport({ format: 'mspdi', diagnostics })
    expect(json(second)).toBe(json(first))
    expect(first.diagnostics).toHaveLength(4000)
    expect(first.errorCount).toBe(2000)
    expect(first.preservedCount).toBe(2000)
    expect(isFatalImport('mspdi', diagnostics)).toBe(false)
  })
})

// ─────────────────────────────────────────────────────────────────────────────
// helpers
// ─────────────────────────────────────────────────────────────────────────────

/** Route a byte payload through the right compatibility pipeline (used by
 * the determinism test to cover `.gproj` + MSPDI uniformly). */
function importMspdiOrGproj(
  input: Uint8Array,
  options: {
    schedule: (document: ProjectDocument) => { diagnostics: readonly ImportDiagnostic[] }
  },
): {
  document: ProjectDocument
  report: ReturnType<typeof importGprojWithCompatibility>['report']
} {
  const decoded = new TextDecoder().decode(input)
  if (decoded.includes('"format":"gproj"')) {
    return importGprojWithCompatibility(input, undefined, options)
  }
  return importMspdiWithCompatibility(input, undefined, options)
}
