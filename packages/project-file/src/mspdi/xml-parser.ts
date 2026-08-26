/**
 * PROJECT-015 — Pure-TypeScript XML parser for MSPDI import.
 *
 * The `@genoffice/project-file` package is a foundation package and MUST remain
 * host-neutral (`spec/project/architecture-lock.md` §13 forbids Node `fs`/
 * `process`, browser globals like `DOMParser`, and host-provided APIs). To
 * avoid pulling in a Node-only or browser-only XML library (which the
 * architecture lock forbids without an explicit proposal), the MSPDI importer
 * ships a tiny, safe, hand-written XML parser in pure TypeScript — the XML
 * analog of `../utf8.ts`.
 *
 * Safety properties (per PROJECT-015 XML-parsing requirements):
 *
 *   - No `eval`, no `Function`, no arbitrary constructor invocation, no
 *     prototype-key deserialization. The output is a plain tree of `{ name,
 *     attributes, children, text }` records — never user-controlled class
 *     instances.
 *   - `<!DOCTYPE>` is REJECTED (throws `XmlParseError`). MSPDI files never
 *     carry a DOCTYPE; rejecting it closes the entity-expansion / external-
 *     entity attack surface (no DTD subset processing at all).
 *   - External / parameter entities are never resolved (there is no DTD
 *     processing path that could resolve them).
 *   - Numeric (`&#NN;` / `&#xNN;`) and the five built-in named entities
 *     (`&lt; &gt; &amp; &quot; &apos;`) are resolved; any other named entity
 *     is rejected (throws).
 *   - A decoded-text-size cap guards against the billion-laughs / quadratic-
 *     expansion family of attacks: the total decoded output length is bounded
 *     by `MSPDI_MAX_DECODED_BYTES` regardless of input encoding tricks.
 *   - A depth cap (`MSPDI_MAX_PARSE_DEPTH`) guards against pathologically
 *     nested elements (and is also the stack-overflow guard, since the
 *     parser is recursive-descent bounded by this constant).
 *   - A byte-size cap (`MSPDI_MAX_INPUT_BYTES`) is enforced before any decode.
 *   - Malformed XML (unclosed tags, bad entity, stray `<`, trailing content,
 *     mismatched close) throws `XmlParseError`; the importer converts that to
 *     a single `INVALID_MSPDI` diagnostic.
 *
 * The parser handles the MSPDI surface: an optional XML declaration, comments,
 * CDATA sections, processing instructions, elements with attributes, self-
 * closing tags, and text content. Namespace prefixes (none appear in MSPDI,
 * but a hostile file might include them) are stripped to the local name so
 * element lookup is by local name only.
 *
 * This module is pure and host-independent.
 */
import { decodeUtf8 } from '../utf8.js'

/** Maximum accepted input size (bytes), shared with the .gproj envelope. */
export const MSPDI_MAX_INPUT_BYTES = 100 * 1024 * 1024 // 100 MiB

/** Maximum accepted element-nesting depth. MSPDI is shallow (< 12); this caps
 * hostile nesting and bounds the recursive-descent stack. */
export const MSPDI_MAX_PARSE_DEPTH = 256

/** Maximum decoded-text length across the whole parse (entity-expansion guard).
 * Generous enough for any real MSPDI, small enough to make billion-laughs
 * infeasible. */
export const MSPDI_MAX_DECODED_BYTES = 200 * 1024 * 1024 // 200 MiB

/** Thrown on any malformed, oversized, or hostile XML input. The importer
 * catches this and converts it to an `INVALID_MSPDI` error diagnostic. */
export class XmlParseError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'XmlParseError'
  }
}

/** A parsed XML attribute (name already local-name-normalized). */
export interface XmlAttribute {
  readonly name: string
  readonly value: string
}

/** A parsed XML element node. */
export interface XmlNode {
  /** Local element name (namespace prefix stripped). */
  readonly name: string
  /** Attributes keyed by local attribute name. */
  readonly attributes: Readonly<Record<string, string>>
  /** Child elements, in document order. */
  readonly children: XmlNode[]
  /** Concatenated decoded text content of this element's direct text/CDATA
   * nodes (NOT including children's text). */
  readonly text: string
}

const NAMED_ENTITIES: Readonly<Record<string, string>> = {
  lt: '<',
  gt: '>',
  amp: '&',
  quot: '"',
  apos: "'",
}

/**
 * Decode XML entity references in a text/attribute string. Supports the five
 * built-in named entities and numeric (`&#NN;` / `&#xNN;`) entities. Any other
 * named entity throws `XmlParseError` (MSPDI declares no entities of its own,
 * and resolving arbitrary named entities would require DTD processing, which
 * is explicitly out of scope and unsafe).
 *
 * `decodedSoFarRef` is the running total decoded length across the whole parse;
 * the function throws when the cap is exceeded (billion-laughs guard).
 */
export function decodeEntities(input: string, decodedSoFarRef: { value: number }): string {
  if (input.indexOf('&') === -1) {
    decodedSoFarRef.value += input.length
    if (decodedSoFarRef.value > MSPDI_MAX_DECODED_BYTES) {
      throw new XmlParseError('decoded XML text exceeds size limit')
    }
    return input
  }
  let out = ''
  let i = 0
  while (i < input.length) {
    const ch = input.charCodeAt(i)
    if (ch !== 0x26 /* & */) {
      out += input[i]
      i += 1
      continue
    }
    // Find the closing ';'.
    const semi = input.indexOf(';', i)
    if (semi === -1) {
      throw new XmlParseError("unterminated entity reference (missing ';')")
    }
    const body = input.slice(i + 1, semi)
    if (body.length === 0) {
      throw new XmlParseError("empty entity reference '&;'")
    }
    let replacement: string
    if (body[0] === '#') {
      // Numeric entity.
      let code: number
      if (body[1] === 'x' || body[1] === 'X') {
        const hex = body.slice(2)
        if (!/^[0-9A-Fa-f]+$/.test(hex) || hex.length === 0) {
          throw new XmlParseError(`malformed hex entity '&#x${hex};'`)
        }
        code = parseInt(hex, 16)
      } else {
        const dec = body.slice(1)
        if (!/^[0-9]+$/.test(dec)) {
          throw new XmlParseError(`malformed numeric entity '&#${dec};'`)
        }
        code = parseInt(dec, 10)
      }
      if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) {
        throw new XmlParseError(`numeric entity out of range: &#${body.slice(1)};`)
      }
      // BMP direct; supplementary plane via surrogate pair (matches JSON/JS).
      if (code <= 0xffff) {
        replacement = String.fromCharCode(code)
      } else {
        const adjusted = code - 0x10000
        replacement = String.fromCharCode(0xd800 | (adjusted >> 10), 0xdc00 | (adjusted & 0x3ff))
      }
    } else {
      if (!/^[A-Za-z][A-Za-z0-9._-]*$/.test(body)) {
        throw new XmlParseError(`malformed named entity '&${body};'`)
      }
      const named = NAMED_ENTITIES[body]
      if (named === undefined) {
        throw new XmlParseError(
          `unsupported named entity '&${body};' (only the five built-in entities are resolved; DTD entities are not processed)`,
        )
      }
      replacement = named
    }
    out += replacement
    decodedSoFarRef.value += replacement.length
    if (decodedSoFarRef.value > MSPDI_MAX_DECODED_BYTES) {
      throw new XmlParseError('decoded XML text exceeds size limit (entity expansion guard)')
    }
    i = semi + 1
  }
  return out
}

/** Strip a namespace prefix (`a:b` → `b`). If no prefix, returns the name. */
function localName(name: string): string {
  const colon = name.indexOf(':')
  return colon === -1 ? name : name.slice(colon + 1)
}

/** Read a qualified name starting at `s[i]` (letter/_ then letter/digit/_/-/.).
 * Returns the name and advances `i`. */
function readName(s: string, i: number): { name: string; next: number } {
  const start = i
  if (i >= s.length) throw new XmlParseError('expected element name')
  const c0 = s.charCodeAt(i)
  const isNameStart =
    (c0 >= 0x41 && c0 <= 0x5a) || // A-Z
    (c0 >= 0x61 && c0 <= 0x7a) || // a-z
    c0 === 0x5f // _
  if (!isNameStart) throw new XmlParseError(`invalid name start at offset ${i}`)
  i += 1
  while (i < s.length) {
    const c = s.charCodeAt(i)
    if (
      (c >= 0x41 && c <= 0x5a) ||
      (c >= 0x61 && c <= 0x7a) ||
      (c >= 0x30 && c <= 0x39) ||
      c === 0x5f ||
      c === 0x2d || // -
      c === 0x2e // .
    ) {
      i += 1
    } else {
      break
    }
  }
  return { name: s.slice(start, i), next: i }
}

function skipWhitespace(s: string, i: number): number {
  while (i < s.length) {
    const c = s.charCodeAt(i)
    if (c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d) {
      i += 1
    } else {
      break
    }
  }
  return i
}

/**
 * Parse XML text into a single root `XmlNode`. Throws `XmlParseError` on any
 * malformed, oversized, or hostile input.
 *
 * Accepts either a `Uint8Array` (decoded as UTF-8 via `../utf8.ts`, keeping
 * the package host-neutral) or a `string` (already decoded).
 */
export function parseXml(
  input: Uint8Array | string,
  opts?: { maxBytes?: number; maxDepth?: number },
): XmlNode {
  const maxBytes = opts?.maxBytes ?? MSPDI_MAX_INPUT_BYTES
  const maxDepth = opts?.maxDepth ?? MSPDI_MAX_PARSE_DEPTH
  const text: string =
    typeof input === 'string'
      ? input
      : (() => {
          if (input.byteLength > maxBytes) {
            throw new XmlParseError(`input exceeds byte limit (${input.byteLength} > ${maxBytes})`)
          }
          return decodeUtf8(input)
        })()
  if (text.length > maxBytes) {
    throw new XmlParseError(`decoded input exceeds byte limit (${text.length} > ${maxBytes})`)
  }

  const decodedTotal = { value: 0 }
  let i = 0
  const len = text.length
  let root: XmlNode | null = null
  // Stack of in-progress elements whose children/text are still being built.
  const stack: {
    name: string
    attributes: Record<string, string>
    children: XmlNode[]
    text: string
  }[] = []

  while (i < len) {
    // Skip whitespace between tokens at the document level.
    if (stack.length === 0) {
      const skipped = skipWhitespace(text, i)
      if (skipped !== i) {
        i = skipped
        continue
      }
    }
    if (text.charCodeAt(i) !== 0x3c /* < */) {
      // Text content inside the current element.
      const start = i
      while (i < len && text.charCodeAt(i) !== 0x3c) i += 1
      const raw = text.slice(start, i)
      const decoded = decodeEntities(raw, decodedTotal)
      const top = stack[stack.length - 1]
      if (top !== undefined) {
        top.text += decoded
      } else if (decoded.trim().length !== 0) {
        throw new XmlParseError(
          `stray text outside root element: ${JSON.stringify(decoded.slice(0, 40))}`,
        )
      }
      continue
    }
    // We are at '<'.
    const c1 = text.charCodeAt(i + 1)
    if (c1 === 0x3f /* ? */) {
      // Processing instruction `<?...?>`
      const end = text.indexOf('?>', i + 2)
      if (end === -1) throw new XmlParseError('unterminated processing instruction')
      i = end + 2
      continue
    }
    if (c1 === 0x21 /* ! */) {
      // `<!--` comment, `<![CDATA[`, or `<!DOCTYPE`.
      if (text.startsWith('<!--', i)) {
        const end = text.indexOf('-->', i + 4)
        if (end === -1) throw new XmlParseError('unterminated comment')
        i = end + 3
        continue
      }
      if (text.startsWith('<![CDATA[', i)) {
        const end = text.indexOf(']]>', i + 9)
        if (end === -1) throw new XmlParseError('unterminated CDATA section')
        const cdata = text.slice(i + 9, end)
        decodedTotal.value += cdata.length
        if (decodedTotal.value > MSPDI_MAX_DECODED_BYTES) {
          throw new XmlParseError('decoded XML text exceeds size limit (CDATA)')
        }
        const top = stack[stack.length - 1]
        if (top !== undefined) {
          top.text += cdata
        } else {
          throw new XmlParseError('CDATA outside root element')
        }
        i = end + 3
        continue
      }
      if (text.startsWith('<!DOCTYPE', i) || text.startsWith('<!doctype', i)) {
        // Reject DOCTYPE entirely — no DTD subset processing, no external
        // entity resolution, no internal subset. MSPDI never carries one.
        throw new XmlParseError(
          '<!DOCTYPE> is not supported (entity/external-entity expansion is disabled for safety)',
        )
      }
      throw new XmlParseError(`unknown markup declaration at offset ${i}`)
    }
    if (c1 === 0x2f /* / */) {
      // Closing tag `</name>`
      const afterSlash = i + 2
      const { name: rawName, next: afterName } = readName(text, afterSlash)
      const close = skipWhitespace(text, afterName)
      if (text.charCodeAt(close) !== 0x3e /* > */) {
        throw new XmlParseError(`expected '>' in closing tag </${rawName}>`)
      }
      const local = localName(rawName)
      const top = stack.pop()
      if (top === undefined) {
        throw new XmlParseError(`unmatched closing tag </${rawName}>`)
      }
      if (top.name !== local) {
        throw new XmlParseError(`mismatched closing tag: </${rawName}> closes <${top.name}>`)
      }
      // Finalize this node.
      const node: XmlNode = {
        name: top.name,
        attributes: top.attributes,
        children: top.children,
        text: top.text,
      }
      if (stack.length === 0) {
        // This was the root element.
        if (root !== null) {
          throw new XmlParseError('multiple root elements')
        }
        root = node
      } else {
        stack[stack.length - 1].children.push(node)
      }
      i = close + 1
      continue
    }
    // Opening tag `<name ...>` or `<name .../>`.
    const { name: rawName, next: afterName } = readName(text, i + 1)
    const local = localName(rawName)
    const attributes: Record<string, string> = {}
    let j = skipWhitespace(text, afterName)
    let selfClosing = false
    while (true) {
      // End of tag?
      if (text.charCodeAt(j) === 0x2f /* / */) {
        if (text.charCodeAt(j + 1) !== 0x3e /* > */) {
          throw new XmlParseError(`expected '/>' at offset ${j}`)
        }
        selfClosing = true
        j += 2
        break
      }
      if (text.charCodeAt(j) === 0x3e /* > */) {
        j += 1
        break
      }
      // Attribute: name = "value" or name = 'value'
      const { name: aName, next: afterAttrName } = readName(text, j)
      const k = skipWhitespace(text, afterAttrName)
      if (text.charCodeAt(k) !== 0x3d /* = */) {
        throw new XmlParseError(`attribute '${aName}' has no value (missing '=')`)
      }
      const m = skipWhitespace(text, k + 1)
      const quote = text.charCodeAt(m)
      if (quote !== 0x22 /* " */ && quote !== 0x27 /* ' */) {
        throw new XmlParseError(`attribute '${aName}' value must be quoted`)
      }
      const close = text.indexOf(String.fromCharCode(quote), m + 1)
      if (close === -1) {
        throw new XmlParseError(`attribute '${aName}' value is unterminated`)
      }
      const rawValue = text.slice(m + 1, close)
      const value = decodeEntities(rawValue, decodedTotal)
      attributes[localName(aName)] = value
      j = skipWhitespace(text, close + 1)
    }
    // Depth guard (and stack-overflow guard).
    if (stack.length >= maxDepth) {
      throw new XmlParseError(`element nesting exceeds depth limit (${maxDepth})`)
    }
    const building = {
      name: local,
      attributes,
      children: [] as XmlNode[],
      text: '',
    }
    if (selfClosing) {
      const node: XmlNode = {
        name: local,
        attributes,
        children: [],
        text: '',
      }
      if (stack.length === 0) {
        if (root !== null) throw new XmlParseError('multiple root elements')
        root = node
      } else {
        stack[stack.length - 1].children.push(node)
      }
    } else {
      stack.push(building)
    }
    i = j
  }

  if (stack.length !== 0) {
    throw new XmlParseError(`unclosed element <${stack[stack.length - 1].name}>`)
  }
  if (root === null) {
    throw new XmlParseError('no root element')
  }
  return root
}

// ---- traversal helpers (used by the MSPDI importer) ---------------------

/** All direct child elements with the given local name, in document order. */
export function childrenNamed(node: XmlNode, name: string): XmlNode[] {
  return node.children.filter((c) => c.name === name)
}

/** The first direct child element with the given name, or `undefined`. */
export function firstChild(node: XmlNode, name: string): XmlNode | undefined {
  return node.children.find((c) => c.name === name)
}

/** The decoded text of the first direct child with the given name, trimmed,
 * or `undefined` if absent/empty. */
export function childText(node: XmlNode, name: string): string | undefined {
  const child = firstChild(node, name)
  if (child === undefined) return undefined
  const t = child.text.trim()
  return t.length === 0 ? undefined : t
}
