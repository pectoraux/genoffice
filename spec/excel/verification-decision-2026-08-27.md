# Excel Parity — Principal Architect Verification Decision

Date: 2026-08-27

This document is the durable Principal Architect decision record for EXCEL-018 through EXCEL-023. It is authoritative for the workflow state and prevents verification status from depending on chat context.

## Decision

| Work item | Status | Decision basis |
|---|---|---|
| EXCEL-018 Remove Duplicates | VERIFIED | Structural `remove-rows` implementation independently audited; mandatory formula-preservation cases, mixed absolute/relative references, package-level byte proof, browser E2E, deployed E2E, CI web gate all green. |
| EXCEL-020 Protection | VERIFIED | Sheet/workbook protection read/write path independently audited; Lock and Unlock both exercised through the real ribbon; password-bearing paths fail closed; browser and deployed E2E green; CI web gate green. |
| EXCEL-021 Tables | VERIFIED | Canonical table read/write path, relationship resolution, create/delete/structural behavior, byte-preservation evidence, browser/deployed E2E, and CI web gate all green. |
| EXCEL-022 Images / Drawings | VERIFIED | Canonical image read/write/delete cascade independently audited; public Univer facade only; absolute-anchor fail-closed behavior verified; browser/deployed E2E and CI web gate green. |
| EXCEL-023 Charts | VERIFIED | Canonical chart read/write path independently audited; browser purity/private-internals constraints satisfied; chart dirty-state and create/edit/delete/reopen coverage green; deployed E2E and CI web gate green. |

## Evidence anchor

The verification campaign is recorded in the merged evidence PR #25:

- PR: https://github.com/pectoraux/genoffice/pull/25
- Merged commit: `65c7ba45d721e45a7a4e92d34ce6f8f17ef65b59`
- Evidence artifact: `packages/xlsx-gateway/tests/xlsx-dedupe-byte-proof.test.ts`
- Evidence artifact: `apps/web/tests/e2e/ribbon-protection.spec.ts` test 6
- Cross-feature browser regression: 112/112 local
- Production browser regression: 57/57
- Principal verification prerequisites: satisfied for all five items

## Architect findings

No implementation defect remains open for EXCEL-018, EXCEL-020, EXCEL-021, EXCEL-022, or EXCEL-023.

The repository-wide `test`/Electron failures do not block these VERIFIED decisions because they are the documented pre-existing frozen-surface baseline and the canonical `web` gate is green.

The foreign `tay-nurs-projects` Vercel integration is not the production authority for this project and does not affect the verification decision. Production evidence was established against `genoffice.vercel.app`.

## Workflow consequence

EXCEL-018 through EXCEL-023 are CLOSED as VERIFIED work items. Their implementation work is not to be reopened unless a future regression, architecture change, or explicit Architecture Change Request invalidates the decision.

**EXCEL-024 Conditional Formatting is now UNBLOCKED for its already-completed Phase A forensic audit and may proceed to implementation under the established workflow.**
