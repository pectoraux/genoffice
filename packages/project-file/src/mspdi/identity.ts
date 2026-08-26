/**
 * PROJECT-015 — Deterministic MSPDI → canonical identity mapping.
 *
 * MSPDI uses its own integer UIDs (Task UID, Resource UID, Assignment UID,
 * Calendar UID). The canonical GenOffice model uses branded-string identities
 * (`TaskId`, `ResourceId`, `AssignmentId`, `CalendarId`, `DependencyId`,
 * `BaselineId`) plus a separate numeric `uid` field on `Task` and `Resource`
 * for interop. Per the architecture-lock §4: "UIDs are persistent
 * source/interoperability identifiers; local IDs are canonical application
 * identities." So the canonical `task.uid`/`resource.uid` carry the MSPDI UID
 * (the persistent interop id), and the canonical branded `id` is the
 * deterministic application identity derived from it.
 *
 * The mapping is:
 *
 *   - deterministic — `f(uid)` is a pure function of the integer UID.
 *   - stable within an imported document — the same UID maps to the same id
 *     everywhere (tasks, predecessor links, assignment TaskUID/ResourceUID,
 *     calendar references) within one import.
 *   - collision-safe — every entity family uses a distinct alphabetic prefix
 *     (`t`, `r`, `a`, `c`, `b`, `d`) so the branded-id spaces are disjoint
 *     (they are already disjoint by brand, but the prefixes make diagnostics
 *     unambiguous and prevent accidental cross-family confusion in logs).
 *   - reproducible — re-importing byte-identical MSPDI reproduces
 *     byte-identical ids (and therefore byte-identical `ProjectDocument`
 *     bytes via `serializeGproj`). No random IDs, no `Date.now()`, no
 *     host-supplied seed.
 *
 * Identity is NOT WBS. WBS / outline number reconstructs hierarchy only
 * (parentTaskId + outlineLevel + wbs); the canonical identity is the MSPDI UID
 * mapped through this module.
 *
 * Documented mapping (per PROJECT-015 brief: "Document the mapping"):
 *
 *   | MSPDI source                  | Canonical field     | Mapping                 |
 *   | ----------------------------- | -------------------- | ----------------------- |
 *   | `<Task><UID>` int            | `task.uid`           | verbatim (number)       |
 *   | `<Task><UID>` int            | `task.id`            | `asTaskId('t'+uid)`     |
 *   | `<Resource><UID>` int        | `resource.uid`       | verbatim (number)       |
 *   | `<Resource><UID>` int        | `resource.id`        | `asResourceId('r'+uid)` |
 *   | `<Assignment><UID>` int      | `assignment.id`      | `asAssignmentId('a'+uid)`|
 *   | `<Calendar><UID>` int        | `calendar.id`        | `asCalendarId('c'+uid)` |
 *   | `<Calendar><BaseCalendarUID>`| `calendar.baseCalendarId` | `asCalendarId('c'+uid)` |
 *   | `<Task><CalendarUID>`         | `task.calendarId`     | `asCalendarId('c'+uid)` |
 *   | `<Resource><CalendarUID>`     | `resource.calendarId`| `asCalendarId('c'+uid)` |
 *   | `<Project><DefaultCalendarUID>`? | `properties.defaultCalendarId` | `asCalendarId('c'+uid)` |
 *   | Predecessor link (succ, pred, type) | `dependency.id` | `asDependencyId('d-'+succ+'-'+pred+'-'+type)` |
 *   | Baseline slot index (0..10)   | `baseline.id`        | `asBaselineId('b'+index)` |
 */
import {
  asAssignmentId,
  asBaselineId,
  asCalendarId,
  asDependencyId,
  asResourceId,
  asTaskId,
} from '@genoffice/project-contracts'
import type {
  AssignmentId,
  BaselineId,
  CalendarId,
  DependencyId,
  DependencyType,
  ResourceId,
  TaskId,
} from '@genoffice/project-contracts'

/** MSPDI Task UID → canonical `TaskId` (`'t' + uid`). */
export function uidToTaskId(uid: number | string): TaskId {
  return asTaskId(`t${uid}`)
}

/** MSPDI Resource UID → canonical `ResourceId` (`'r' + uid`). */
export function uidToResourceId(uid: number | string): ResourceId {
  return asResourceId(`r${uid}`)
}

/** MSPDI Assignment UID → canonical `AssignmentId` (`'a' + uid`). */
export function uidToAssignmentId(uid: number | string): AssignmentId {
  return asAssignmentId(`a${uid}`)
}

/** MSPDI Calendar UID → canonical `CalendarId` (`'c' + uid`). */
export function uidToCalendarId(uid: number | string): CalendarId {
  return asCalendarId(`c${uid}`)
}

/** MSPDI baseline slot index (0=Baseline, 1=Baseline1, …, 10=Baseline10)
 * → canonical `BaselineId` (`'b' + index`). */
export function baselineIndexToId(index: number): BaselineId {
  return asBaselineId(`b${index}`)
}

/** Deterministic `DependencyId` from the successor id, predecessor id, and
 * link type. Two links between the same successor/predecessor pair with the
 * same type collide (→ the engine emits `DUPLICATE_DEPENDENCY_ID`); two links
 * between the same pair with different types are distinct. */
export function dependencyId(
  successorId: string,
  predecessorId: string,
  type: DependencyType,
): DependencyId {
  return asDependencyId(`d-${successorId}-${predecessorId}-${type}`)
}
