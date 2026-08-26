/**
 * PROJECT-014 — Canonical (deterministic) JSON serialization.
 *
 * Two `ProjectDocument` values with identical semantic content MUST produce
 * byte-identical `.gproj` files. `JSON.stringify` alone does NOT guarantee this
 * because it preserves object-key insertion order, and object-key insertion
 * order is NOT part of the canonical Project model (the model uses arrays for
 * ordered collections and `Record<...>` maps for keyed lookups whose key order
 * is semantically irrelevant).
 *
 * `canonicalJson` therefore:
 *   - Preserves array order verbatim (task arrays, dependency arrays, notes,
 *     availability windows, calendar exceptions — order IS semantically
 *     meaningful and is part of the canonical document form).
 *   - Sorts object keys by Unicode code point (NOT `localeCompare`, which the
 *     PROJECT-014 brief forbids) so two semantically-equivalent objects with
 *     differently-ordered keys serialize to identical bytes.
 *   - Uses a fixed 2-space indent and a single trailing newline so the bytes
 *     are stable across runtimes.
 *
 * The serializer builds the canonical envelope object in code (so its key
 * order is already canonical), and `canonicalJson` guarantees that any
 * `Record<...>` maps nested inside (e.g. `task.customFields`,
 * `baseline.taskSnapshots`, `calendar.workingWeek`) are also emitted in a
 * stable order regardless of how the in-memory object was assembled.
 */

/** Comparator that sorts strings by Unicode code point (not locale). */
function compareCodePoint(a: string, b: string): number {
  const len = Math.min(a.length, b.length)
  for (let i = 0; i < len; i++) {
    const ca = a.codePointAt(i) as number
    const cb = b.codePointAt(i) as number
    if (ca !== cb) return ca - cb
  }
  return a.length - b.length
}

/**
 * Produce a "canonical" (key-sorted, array-order-preserved) deep clone of a
 * JSON-safe value. Used as the pre-stringify pass so `JSON.stringify` output
 * is deterministic.
 */
export function canonicalValue(value: unknown, depth = 0): unknown {
  if (depth > 4096) throw new Error('canonicalValue: structure too deep')
  if (value === null || typeof value !== 'object') return value
  if (Array.isArray(value)) {
    return value.map((item) => canonicalValue(item, depth + 1))
  }
  const obj = value as Record<string, unknown>
  const keys = Object.keys(obj).sort(compareCodePoint)
  const out: Record<string, unknown> = {}
  for (const key of keys) out[key] = canonicalValue(obj[key], depth + 1)
  return out
}

/** Serialize a JSON-safe value to deterministic UTF-8 bytes. */
export function canonicalJson(value: unknown): Uint8Array {
  const stable = canonicalValue(value)
  const text = JSON.stringify(stable, null, 2) + '\n'
  // UTF-8 encode in pure TS (host-neutral — see utf8.ts).
  return utf8Encode(text)
}

// Local import kept inline to avoid a circular re-export at the module level.
// `encodeUtf8` is a pure function with no side effects.
import { encodeUtf8 as utf8Encode } from './utf8.js'
