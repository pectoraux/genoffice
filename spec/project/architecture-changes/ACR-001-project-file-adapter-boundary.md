# ACR-001 — `project-file` is the sanctioned file-adapter boundary (architecture-lock §13 clarification)

- Status: APPROVED by explicit Principal Architect directive (the reconciliation increment authorization); implemented together with this record; final acceptance via Independent Architect Review of the reconciliation pull request.
- Type: Clarification of a frozen invariant (architecture-lock §13). No prohibition is weakened, no runtime behavior changes, no package edge changes.
- Recorded under: architecture-lock §12 — a change to a frozen invariant requires a recorded proposal containing motivation, affected invariants, alternatives, compatibility impact, migration plan, verification impact, and explicit authority approval.
- Baseline at record time: PROJECT-001 through PROJECT-024 objectively accepted; `main` = merge commit of PR #24 (`b3368664d6d81b02b224c516d654cc9643b80544`).

## 1. Motivation

Architecture-lock §13 as originally worded read:

> Foundation packages must not import React/React DOM, Electron, Node filesystem/process APIs, browser globals, HTTP clients/server route modules, Excel renderer packages, or `.mpp`/MSPDI parser implementation code.

The final clause, read literally against the §3 package list, appears to prohibit MSPDI/MPP parser implementation code in **every** foundation package — including `packages/project-file`, which §3 defines as "file adapter boundary" and §10 defines as the home of "file adapters [that] translate to/from the canonical model". The accepted corpus has treated `project-file` as exactly that boundary since PROJECT-014: the pure-TypeScript `.gproj` serializer (PROJECT-014), the pure-TypeScript MSPDI XML tokenizer and writer (PROJECT-015/016), and the host-neutral MPP adapter contract plus N1–N5 MSPDI normalization (PROJECT-018) all live in `packages/project-file` and were independently reviewed and accepted. `spec/project/architecture.md` ("Owns file-adapter interfaces and, in later increments, `.gproj`, MSPDI, and MPP translators") and `spec/project/dependency-graph.md` (the PROJECT-014–020 package-edge sections) record the same sanctioned role.

The Principal Architect's roadmap audit (this reconciliation increment's authorization) found this specification-consistency drift — the lock's letter contradicts the sanctioned adapter role — with NO fundamental architecture failure. This record resolves the contradiction by clarifying the invariant, exactly as directed.

## 2. Affected invariant

- `spec/project/architecture-lock.md` §13 (Forbidden dependencies) — clarified only. This is the ONLY change to the lock; §1–§12 and §14 are untouched, and the document remains `Status: FROZEN`.
- No other frozen invariant is affected: §3 (canonical package boundaries), §10 (files), and §2 (layer order) already describe `project-file` as the adapter boundary and are consistent with the clarified reading.

## 3. Old interpretation

- Letter: "Foundation packages … must not import … `.mpp`/MSPDI parser implementation code" — on a literal reading this includes `packages/project-file`, which would make the accepted in-package `.gproj`/MSPDI implementations (PROJECT-014–018) a lock violation.
- As practiced and accepted: the clause banned **external** parser implementations and process code from all foundation packages; pure-TypeScript format implementations inside the sanctioned `project-file` adapter boundary were accepted (e.g. requirements.md PROJECT-015: "a safe pure-TypeScript XML parser (no `DOMParser`, no Node `fs`, no external library — architecture-lock §13)"), and the externalized MPXJ sidecar was placed in the HOST package `@genoffice/project-mpp-host` (PROJECT-018) precisely because MPXJ is an external Java parser implementation.

## 4. New interpretation (the clarified rule — normative)

> Foundation semantic/runtime packages must not import external MSPDI/MPP parser implementations.
>
> `packages/project-file` is the sanctioned file-adapter boundary and may contain format-specific parser/serializer implementations.
>
> No `project-engine`, `project-scheduling`, `project-renderer-core`, or host package may directly import format-specific parser internals.
>
> File-format implementations remain behind the `project-file` adapter boundary.

Concretely, within the clarified §13 text in the lock:

- "Foundation semantic/runtime packages" are `project-contracts`, `project-engine`, `project-scheduling`, and `project-renderer-core` (the `project-file` adapter package is addressed separately, below).
- "External MSPDI/MPP parser implementations" are third-party parser libraries (MPXJ, Aspose.Tasks, `tsmpp`, and similar) — these remain forbidden in **every** foundation package, `project-file` included; their only sanctioned location stays the host-managed sidecar behind `@genoffice/project-mpp-host` (the accepted PROJECT-018 architecture, unchanged).
- "Format-specific parser/serializer implementations" are the in-package pure-TypeScript implementations (`src/gproj/**`, `src/mspdi/**` incl. `xml-parser.ts`/`xml-writer.ts`, `src/mpp/**` adapter + normalization code) — sanctioned inside `project-file` only.
- "Format-specific parser internals" are the deep modules of those implementations. Consumers (including the host package) may import the `@genoffice/project-file` PUBLIC adapter surface (`ProjectFileAdapter`/`gprojFileAdapter`, `MspdiFileAdapter`, `importMppFromMspdi`, compatibility pipeline) but never internal parser modules via deep paths.

All other §13 prohibitions (React/React DOM, Electron, Node filesystem/process APIs, browser globals, HTTP clients/server route modules, Excel renderer packages) continue to apply to every foundation package, `project-file` included, unchanged.

## 5. Alternatives considered

1. **Leave §13 as-is and rely on prose interpretation.** Rejected: the letter/behavior contradiction persists and grows with every file-format increment; future reviewers cannot distinguish a sanctioned adapter implementation from a lock violation by reading the lock.
2. **Move format implementations out of `project-file` into host packages.** Rejected: contradicts lock §3/§10 and `architecture.md` (file adapters are a foundation boundary, not host code), would force format knowledge into hosts, and would rewrite the accepted PROJECT-014–020 package edges — a genuine architecture change with large migration cost, for zero behavioral benefit.
3. **Rewrite §13 from scratch.** Rejected: minimal-change discipline. The clarification keeps every previously enforced prohibition verbatim and adds the scoping sentences; nothing is reorganized beyond the minimum needed to remove the contradiction.
4. **Do nothing and record the interpretation in requirements.md only.** Rejected: §12 requires frozen-invariant changes to be recorded as architecture-change proposals with the lock itself updated — a prose-only note would leave the controlling document self-contradictory.

## 6. Compatibility impact

- None at runtime. No source file, package edge, public surface, or behavior changes as a consequence of the clarification.
- Every previously enforced prohibition remains enforced: the React/Electron/Node/browser/HTTP/Excel bans are unchanged for all foundation packages; external MSPDI/MPP parser libraries remain banned from all foundation packages (the MPXJ sidecar stays confined to `@genoffice/project-mpp-host`, exactly as accepted in PROJECT-018); the MPP import architecture and the PROJECT-019A MPP-export deferral decision are untouched.
- The accepted in-package `.gproj`/MSPDI implementations in `project-file` move from "sanctioned by interpretation" to "sanctioned by the lock's letter" — the only semantic delta.

## 7. Migration impact

- None. No code moves, no dependency changes, no CI workflow changes. The CI foundation boundary grep is unchanged (it never scanned for format internals; it scans for React/Electron/Node/HTTP imports, which remain forbidden).
- The previously merged verbatim §13 guard assertions (`packages/project-file/tests/mpp-feasibility.test.ts`, `packages/project-file/tests/mpp-export-strategy.test.ts`) are updated in the same change set to assert the clarified §13 sentences verbatim — strengthened, not weakened (they additionally assert this record exists).
- New per-package architecture guards added by the same reconciliation increment enforce the clarified rule mechanically (import-specifier allowlists; cross-package scans; sanctioned-surface presence assertions).

## 8. Verification impact

- The clarified rule is enforced by the strengthened architecture-boundary test suites added in this increment:
  - `packages/project-file/tests/architecture.test.ts` — `project-file` MAY contain the `.gproj`/MSPDI/MPP adapter implementations; `project-engine`, `project-scheduling`, and `project-renderer-core` MUST NOT import `project-file`, MSPDI parser internals, or MPP parser internals; the host package may import only the `project-file` PUBLIC surface (never deep parser internals).
  - `packages/project-engine/tests/architecture.test.ts` and `packages/project-scheduling/tests/architecture.test.ts` — self-scans with exact dependency sets and import-specifier allowlists (the domain engine and the scheduling engine acquire no file-format knowledge).
  - `packages/project-renderer-core/tests/architecture.test.ts` — extended with the file-format-internals import prohibition (the renderer acquires no file-format knowledge).
- The specification-consistency guard (`packages/project-file/tests/spec-consistency.test.ts`) asserts the clarified §13 text and this record's presence, so the lock and the record stay in lockstep.
- The PROJECT-001–024 regression suite must remain green with zero skipped tests; no Excel surfaces, renderer behavior, scheduling behavior, or file-format semantics may change.

## 9. Principal Architect approval reference

- The Principal Architect's reconciliation directive issued on the accepted state "PROJECT-001 through PROJECT-024 are accepted; PROJECT-024 is already merged into main", authorizing exactly this clarification: "1. ARCHITECTURE-LOCK CLARIFICATION — The current architecture-lock §13 wording is too broad because packages/project-file is explicitly the sanctioned file-adapter boundary… Prepare an explicit architecture-change/clarification record containing: motivation, affected invariant, old interpretation, new interpretation, alternatives considered, compatibility impact, migration impact, verification impact, explicit Principal Architect approval reference. The intended clarified rule is: [the four rules quoted verbatim in §4 above]. Do NOT silently modify architecture-lock.md without the architecture-change record."
- The clarified rule text in architecture-lock §13 is the architect's verbatim intended rule from that directive.
- Per §12 ("Until accepted, the proposal is not implemented"), the directive's explicit pre-approval authorizes shipping this record together with the §13 clarification as ONE review unit; final acceptance is the Independent Architect Review of the reconciliation pull request that carries both.

## 10. Explicitly out of scope

- No scheduling-semantics change, no renderer-semantics change, no file-format-semantics change (`.gproj`, MSPDI import/export behavior, MPP import behavior, and the 019A MPP-export deferral decision are all untouched).
- No new package, no package-edge change, no CI workflow change.
- No implementation of PROJECT-025 (Calendar visualization) — it remains the next authorized product increment, under separate review (PR #26).
