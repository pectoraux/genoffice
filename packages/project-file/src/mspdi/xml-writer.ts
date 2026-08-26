/**
 * PROJECT-016 — Deterministic pure-TypeScript XML writer for MSPDI export.
 *
 * The XML analog of the accepted PROJECT-015 pure-TS tokenizer
 * (`./xml-parser.ts`) and of the `.gproj` UTF-8 codec (`../utf8.ts`): the
 * `@genoffice/project-file` package is a foundation package and MUST remain
 * host-neutral (`spec/project/architecture-lock.md` §13 forbids Node
 * `fs`/`process`, browser globals like `XMLSerializer`/`DOMParser`, and
 * external XML libraries). The writer is a tiny hand-written serializer —
 * no `eval`, no `Function`, no arbitrary constructors, no prototype-key
 * deserialization, no filesystem, no network.
 *
 * Determinism properties (PROJECT-016 brief — "Deterministic XML"):
 *
 *   - The writer emits a fixed physical layout: an XML declaration, LF line
 *     endings, two-space indentation per nesting level, one element per line,
 *     and inline text leaves (`<Name>value</Name>`). The same logical tree
 *     always produces byte-identical output.
 *   - Element order is decided by the CALLER (the exporter defines the
 *     canonical element order explicitly; the writer never sorts).
 *   - Attribute order is the caller-provided key order (the exporter only
 *     ever emits the single MSPDI default-namespace declaration, so there is
 *     no attribute-order ambiguity in practice).
 *   - The writer never reads the clock, never uses randomness, and never
 *     consults the host locale.
 *
 * Escaping (PROJECT-016 brief — "XML REPRESENTATION"):
 *
 *   - Text content escapes `&` → `&amp;`, `<` → `&lt;`, `>` → `&gt;`, and
 *     the carriage return `\r` → `&#xD;` (a literal CR in text content is
 *     normalized to LF by conforming XML parsers; the escaped form survives
 *     both the accepted PROJECT-015 parser and external ones).
 *   - Attribute values additionally escape `"` → `&quot;`, `'` → `&apos;`,
 *     and the other whitespace characters subject to XML attribute-value
 *     normalization (`\t` → `&#x9;`, `\n` → `&#xA;`, `\r` → `&#xD;`).
 *   - The five built-in named entities are exactly the set the accepted
 *     PROJECT-015 parser resolves, so exported text round-trips through
 *     `parseXml` + `decodeEntities` losslessly.
 *
 * The writer cannot produce malformed XML by construction: element names are
 * validated against the XML name production used by the parser, and all text
 * passes through the escapers above.
 */

/** Characters that must be escaped in XML element text content. */
export function escapeXmlText(input: string): string {
  if (
    input.indexOf('&') === -1 &&
    input.indexOf('<') === -1 &&
    input.indexOf('>') === -1 &&
    input.indexOf('\r') === -1
  ) {
    return input
  }
  return input
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\r/g, '&#xD;')
}

/** Characters that must be escaped in a double-quoted XML attribute value. */
export function escapeXmlAttribute(input: string): string {
  return input
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
    .replace(/\t/g, '&#x9;')
    .replace(/\n/g, '&#xA;')
    .replace(/\r/g, '&#xD;')
}

/**
 * Validate an element/attribute name against the same restricted name
 * production the PROJECT-015 parser accepts (`letter` or `_`, then letters,
 * digits, `_`, `-`, `.`). Namespace prefixes (`a:b`) are allowed and their
 * colon is preserved verbatim. Throws on anything else so the exporter can
 * never emit malformed markup.
 */
export function isValidXmlName(name: string): boolean {
  return /^[A-Za-z_][A-Za-z0-9_.-]*(:[A-Za-z_][A-Za-z0-9_.-]*)?$/.test(name)
}

const INDENT_UNIT = '  '

/**
 * A minimal deterministic XML document builder. The builder accumulates
 * string parts; `toString()` returns the document text. Indentation is
 * tracked by depth; every open/close/leaf emits exactly one line.
 */
export class XmlWriter {
  private readonly parts: string[] = []
  private depth = 0

  private indent(): string {
    return INDENT_UNIT.repeat(this.depth)
  }

  /** Emit the XML declaration. Must be the first call. */
  declaration(version = '1.0', encoding = 'utf-8'): this {
    this.parts.push(`<?xml version="${version}" encoding="${encoding}"?>`)
    return this
  }

  /** Open `<name attrs…>` on its own line, increasing the nesting depth. */
  open(name: string, attributes?: Record<string, string>): this {
    if (!isValidXmlName(name)) {
      throw new Error(`invalid XML element name: ${JSON.stringify(name)}`)
    }
    let attrs = ''
    if (attributes !== undefined) {
      for (const key of Object.keys(attributes)) {
        if (!isValidXmlName(key)) {
          throw new Error(`invalid XML attribute name: ${JSON.stringify(key)}`)
        }
        attrs += ` ${key}="${escapeXmlAttribute(attributes[key])}"`
      }
    }
    this.parts.push(`${this.indent()}<${name}${attrs}>`)
    this.depth += 1
    return this
  }

  /** Close `</name>` on its own line, decreasing the nesting depth. */
  close(name: string): this {
    if (!isValidXmlName(name)) {
      throw new Error(`invalid XML element name: ${JSON.stringify(name)}`)
    }
    this.depth -= 1
    this.parts.push(`${this.indent()}</${name}>`)
    return this
  }

  /** Emit a self-closing `<name />` element (no children, no text). */
  selfClosing(name: string, attributes?: Record<string, string>): this {
    if (!isValidXmlName(name)) {
      throw new Error(`invalid XML element name: ${JSON.stringify(name)}`)
    }
    let attrs = ''
    if (attributes !== undefined) {
      for (const key of Object.keys(attributes)) {
        if (!isValidXmlName(key)) {
          throw new Error(`invalid XML attribute name: ${JSON.stringify(key)}`)
        }
        attrs += ` ${key}="${escapeXmlAttribute(attributes[key])}"`
      }
    }
    this.parts.push(`${this.indent()}<${name}${attrs} />`)
    return this
  }

  /** Emit an inline text leaf `<name>escaped-text</name>` on one line. */
  leaf(name: string, text: string): this {
    if (!isValidXmlName(name)) {
      throw new Error(`invalid XML element name: ${JSON.stringify(name)}`)
    }
    this.parts.push(`${this.indent()}<${name}>${escapeXmlText(text)}</${name}>`)
    return this
  }

  /**
   * Emit an inline empty leaf `<name></name>` on one line. Distinct from
   * `selfClosing` so the writer can express both deterministic forms when a
   * present-but-empty element is semantically meaningful.
   */
  emptyLeaf(name: string): this {
    if (!isValidXmlName(name)) {
      throw new Error(`invalid XML element name: ${JSON.stringify(name)}`)
    }
    this.parts.push(`${this.indent()}<${name}></${name}>`)
    return this
  }

  /** The accumulated document text (LF line endings, trailing newline). */
  toString(): string {
    if (this.depth !== 0) {
      throw new Error('XmlWriter is not balanced (unclosed elements remain)')
    }
    return this.parts.join('\n') + '\n'
  }
}
