/**
 * PROJECT-018 — the five PROJECT-017-approved mechanical normalizations
 * (N1–N5) applied to MPXJ-generated MSPDI XML before it enters the accepted
 * PROJECT-015 importer (feasibility report §10 — the definitions are
 * implemented verbatim, nothing improvised, nothing added):
 *
 *   N1  strip `-1` sentinel `CalendarUID` references on tasks/resources
 *       (they mean "inherit the default calendar", not a dangling ref)
 *   N2  strip `-1` sentinel `BaseCalendarUID` on calendars (root-cause twin
 *       of N1, same pass, separately diagnosed for traceability)
 *   N3  filter the hidden placeholder task (`UID 0` / `OutlineLevel 0` /
 *       `WBS "0"` — the MPP summary artifact) and the analogous null-name
 *       placeholder resource
 *   N4  rewrite `WorkingTime` `ToTime 00:00:00` ("until midnight", the
 *       Microsoft/MPXJ convention) to the ISO-8601 day-end `24:00:00`
 *       (canonical `endMinute: 1440`)
 *   N5  drop "unassigned" assignments (`ResourceUID -65535`) with a
 *       dedicated expected-loss diagnostic
 *
 * Implementation: parse with the accepted pure-TypeScript tokenizer
 * (`parseXml`), transform a mutable copy of the element tree, re-serialize
 * deterministically (compact form — no whitespace injection, document-order
 * attributes, explicit escaping via the PROJECT-016 writer primitives). The
 * layer NEVER invents semantic values: it only strips sentinels, filters
 * placeholder records, and re-expresses the day-end boundary.
 *
 * If the input is not well-formed XML the bytes are returned UNCHANGED with
 * zero diagnostics — malformed input is the accepted importer's error to
 * report (stage `'mspdi'`), not the normalizer's.
 */
import { encodeUtf8 } from '../utf8.js'
import {
  MPP_DROPPED_UNASSIGNED_ASSIGNMENT,
  MPP_NORMALIZED_BASE_CALENDAR_SENTINEL,
  MPP_NORMALIZED_MIDNIGHT_PERIOD,
  MPP_NORMALIZED_PLACEHOLDER_RECORD,
  MPP_NORMALIZED_SENTINEL_REFERENCE,
} from './diagnostics.js'
import type { MppDiagnostic } from './types.js'
import { XmlParseError, parseXml, type XmlNode } from '../mspdi/xml-parser.js'
import { escapeXmlAttribute, escapeXmlText, isValidXmlName } from '../mspdi/xml-writer.js'

/** The result of normalizing an MPXJ-generated MSPDI document. */
export interface MppNormalizationResult {
  /** The normalized MSPDI XML bytes (UTF-8). Identical to the input when no
   * normalization applied or when the input is not well-formed XML. */
  readonly bytes: Uint8Array
  /** Diagnostics for every applied normalization (stage `'normalization'`). */
  readonly diagnostics: readonly MppDiagnostic[]
}

/** A mutable deep copy of a parsed element tree (internal working type —
 * structurally assignable to `XmlNode`). */
interface MutableNode {
  name: string
  attributes: Record<string, string>
  children: MutableNode[]
  text: string
}

/** Deep-clone a parsed element tree (pure structural copy). */
function cloneNode(node: XmlNode): MutableNode {
  return {
    name: node.name,
    attributes: { ...node.attributes },
    children: node.children.map(cloneNode),
    text: node.text,
  }
}

/** First direct child element with the given local name (mutable view). */
function mFirstChild(node: MutableNode, name: string): MutableNode | undefined {
  return node.children.find((child) => child.name === name)
}

/** Direct-child text, with the accepted `childText` convention (empty text
 * → undefined). */
function mChildText(node: MutableNode, name: string): string | undefined {
  const child = mFirstChild(node, name)
  return child !== undefined && child.text !== '' ? child.text : undefined
}

/** Replace a node's child-element list in place (surgical filter helper). */
function retainChildren(node: MutableNode, keep: (child: MutableNode) => boolean): void {
  node.children = node.children.filter(keep)
}

/**
 * Deterministically serialize an element tree back to compact XML text.
 * Compact form guarantees no whitespace text nodes are invented; attributes
 * serialize in document order (parser insertion order); text and attribute
 * values are escaped with the PROJECT-016 writer primitives.
 */
export function serializeXmlNode(node: MutableNode): string {
  if (!isValidXmlName(node.name)) {
    // Unreachable for trees produced by parseXml (the tokenizer validates
    // names); guarded anyway so malformed markup can never be emitted.
    throw new XmlParseError(`invalid element name: ${node.name}`)
  }
  let attrs = ''
  for (const [name, value] of Object.entries(node.attributes)) {
    if (!isValidXmlName(name)) {
      throw new XmlParseError(`invalid attribute name: ${name}`)
    }
    attrs += ` ${name}="${escapeXmlAttribute(value)}"`
  }
  const open = `<${node.name}${attrs}`
  if (node.children.length === 0 && node.text === '') return `${open}/>`
  let inner = escapeXmlText(node.text)
  for (const child of node.children) inner += serializeXmlNode(child)
  return `${open}>${inner}</${node.name}>`
}

/** Strip a `-1` sentinel `CalendarUID` child from a task/resource (N1).
 * Returns true when a sentinel was stripped. */
function stripSentinelCalendarUid(entity: MutableNode): boolean {
  const calendarUid = mFirstChild(entity, 'CalendarUID')
  if (calendarUid === undefined || calendarUid.text.trim() !== '-1') return false
  retainChildren(entity, (child) => child !== calendarUid)
  return true
}

/**
 * Apply the N1–N5 normalizations to an MPXJ-generated MSPDI document.
 *
 * The five normalizations are exactly the PROJECT-017-approved set — each
 * rewrite/drop emits a stage-`'normalization'` diagnostic naming the
 * affected entity; nothing is ever dropped or rewritten silently.
 */
export function normalizeMspdiForCanonicalImport(input: Uint8Array): MppNormalizationResult {
  const diagnostics: MppDiagnostic[] = []
  let root: XmlNode
  try {
    root = parseXml(input)
  } catch {
    // Not well-formed XML — let the accepted importer report it (stage
    // 'mspdi'). The normalizer never masks importer-level errors.
    return { bytes: input, diagnostics }
  }

  // Work on a copy; the input bytes are never mutated.
  const tree = cloneNode(root)

  // ── N3 (tasks) + N1 (task calendars) ──────────────────────────────────
  const tasks = mFirstChild(tree, 'Tasks')
  if (tasks !== undefined) {
    retainChildren(tasks, (task) => {
      if (mChildText(task, 'UID')?.trim() === '0') {
        diagnostics.push({
          code: MPP_NORMALIZED_PLACEHOLDER_RECORD,
          severity: 'info',
          message:
            'MPP hidden placeholder task (UID 0 / OutlineLevel 0 / WBS "0" — the MPP summary artifact) was filtered before import',
          entityId: 't0',
          stage: 'normalization',
        })
        return false
      }
      if (stripSentinelCalendarUid(task)) {
        diagnostics.push({
          code: MPP_NORMALIZED_SENTINEL_REFERENCE,
          severity: 'info',
          message:
            'MPP sentinel CalendarUID -1 ("inherit the default calendar") was stripped from a task before import',
          entityId: `t${mChildText(task, 'UID')?.trim() ?? ''}`,
          stage: 'normalization',
        })
      }
      return true
    })
  }

  // ── N3 (resources) + N1 (resource calendars) ──────────────────────────
  const resources = mFirstChild(tree, 'Resources')
  if (resources !== undefined) {
    retainChildren(resources, (resource) => {
      const name = mChildText(resource, 'Name')
      if (name === undefined || name.trim() === '') {
        diagnostics.push({
          code: MPP_NORMALIZED_PLACEHOLDER_RECORD,
          severity: 'info',
          message:
            'MPP null-name placeholder resource was filtered before import (documented MPXJ artifact)',
          entityId: `r${mChildText(resource, 'UID')?.trim() ?? ''}`,
          stage: 'normalization',
        })
        return false
      }
      if (stripSentinelCalendarUid(resource)) {
        diagnostics.push({
          code: MPP_NORMALIZED_SENTINEL_REFERENCE,
          severity: 'info',
          message:
            'MPP sentinel CalendarUID -1 ("inherit the default calendar") was stripped from a resource before import',
          entityId: `r${mChildText(resource, 'UID')?.trim() ?? ''}`,
          stage: 'normalization',
        })
      }
      return true
    })
  }

  // ── N2: strip `-1` sentinel BaseCalendarUID on calendars. ──────────────
  const calendars = mFirstChild(tree, 'Calendars')
  if (calendars !== undefined) {
    for (const calendar of calendars.children) {
      if (calendar.name !== 'Calendar') continue
      const baseUid = mFirstChild(calendar, 'BaseCalendarUID')
      if (baseUid === undefined || baseUid.text.trim() !== '-1') continue
      retainChildren(calendar, (child) => child !== baseUid)
      diagnostics.push({
        code: MPP_NORMALIZED_BASE_CALENDAR_SENTINEL,
        severity: 'info',
        message:
          'MPP sentinel BaseCalendarUID -1 ("no base calendar") was stripped from a calendar before import',
        entityId: `c${mChildText(calendar, 'UID')?.trim() ?? ''}`,
        stage: 'normalization',
      })
    }
  }

  // ── N4: rewrite "until midnight" ToTime 00:00:00 → 24:00:00. ──────────
  rewriteMidnightWorkingTimes(tree, diagnostics)

  // ── N5: drop "unassigned" placeholder assignments (ResourceUID -65535). ─
  const assignments = mFirstChild(tree, 'Assignments')
  if (assignments !== undefined) {
    retainChildren(assignments, (assignment) => {
      if (mChildText(assignment, 'ResourceUID')?.trim() !== '-65535') return true
      const taskUid = mChildText(assignment, 'TaskUID')?.trim() ?? ''
      diagnostics.push({
        code: MPP_DROPPED_UNASSIGNED_ASSIGNMENT,
        severity: 'warning',
        message:
          'MPP "unassigned" placeholder assignment (ResourceUID -65535) was dropped — the canonical model has no unassigned assignment (expected loss)',
        entityId: `t${taskUid}`,
        stage: 'normalization',
      })
      return false
    })
  }

  if (diagnostics.length === 0) return { bytes: input, diagnostics }
  const xml = `<?xml version="1.0" encoding="UTF-8"?>\n${serializeXmlNode(tree)}`
  return { bytes: encodeUtf8(xml), diagnostics }
}

/** Recursively rewrite every `WorkingTime` whose `ToTime` is exactly
 * `00:00:00` to the ISO-8601 day-end `24:00:00` (N4). */
function rewriteMidnightWorkingTimes(node: MutableNode, diagnostics: MppDiagnostic[]): void {
  if (node.name === 'WorkingTime') {
    const toTime = mFirstChild(node, 'ToTime')
    if (toTime !== undefined && toTime.text.trim() === '00:00:00') {
      toTime.text = '24:00:00'
      diagnostics.push({
        code: MPP_NORMALIZED_MIDNIGHT_PERIOD,
        severity: 'info',
        message:
          'MPP working period running "until midnight" (ToTime 00:00:00) was rewritten to the ISO-8601 day-end expression 24:00:00 (canonical endMinute 1440)',
        stage: 'normalization',
      })
    }
  }
  for (const child of node.children) rewriteMidnightWorkingTimes(child, diagnostics)
}
