import type { WorkbookStyleEdit } from '../types.js'
import type {
  CellBorderStyleName,
  CellBorderState,
  CellFormatState,
} from '../domain/workbook.types.js'

/**
 * Built-in (ECMA-376 §18.8.30) numFmtId → formatCode map. numFmtIds 0..163
 * are reserved; 164+ are user-defined and live in the file's <numFmts>.
 * The StylesheetReader resolves a cellXfs numFmtId to its pattern through
 * this map (when the id is built-in) or through the file's <numFmts>
 * (when the id is user-defined). Only the most common built-ins are listed
 * — anything else falls through to "no numberFormat" (the file's own XML
 * keeps the format regardless).
 */
const BUILTIN_NUMFMTS: ReadonlyMap<number, string> = new Map<number, string>([
  [1, '0'],
  [2, '0.00'],
  [3, '#,##0'],
  [4, '#,##0.00'],
  [5, '$#,##0_);($#,##0)'],
  [6, '$#,##0_);[Red]($#,##0)'],
  [7, '$#,##0.00_);($#,##0.00)'],
  [8, '$#,##0.00_);[Red]($#,##0.00)'],
  [9, '0%'],
  [10, '0.00%'],
  [11, '0.00E+00'],
  [12, '# ?/?'],
  [13, '# ??/??'],
  [14, 'mm-dd-yy'],
  [15, 'd-mmm-yy'],
  [16, 'd-mmm'],
  [17, 'mmm-yy'],
  [18, 'h:mm AM/PM'],
  [19, 'h:mm:ss AM/PM'],
  [20, 'h:mm'],
  [21, 'h:mm:ss'],
  [22, 'm/d/yy h:mm'],
  [37, '#,##0_);(#,##0)'],
  [38, '#,##0_);[Red](#,##0)'],
  [39, '#,##0.00_);(#,##0.00)'],
  [40, '#,##0.00_);[Red](#,##0.00)'],
  [41, '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'],
  [42, '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'],
  [44, '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'],
  [45, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'],
  [46, '[$-404]e-m-d'],
  [47, 'mm:ss'],
  [48, '[h]:mm:ss'],
  [49, 'mmss.0'],
])

/// Copy-on-write editor for xl/styles.xml. Existing entries are never
/// modified — every changed cell gets a new cellXfs entry (deduped) derived
/// from its current one, so untouched cells keep their exact formatting.
export class StylesheetEditor {
  private readonly source: string
  private numFmts: string[]
  private readonly fonts: string[]
  private readonly fills: string[]
  private readonly borders: string[]
  private readonly hadBordersSection: boolean
  private readonly cellXfs: string[]
  private readonly dxfs: string[]
  private readonly hadDxfsSection: boolean
  private readonly originalCounts: {
    numFmts: number
    fonts: number
    fills: number
    borders: number
    cellXfs: number
    dxfs: number
  }
  private readonly cache = new Map<string, number>()
  private nextNumFmtId: number

  constructor(stylesXml: string) {
    this.source = stylesXml
    this.numFmts = extractElements(sectionInner(stylesXml, 'numFmts') ?? '', 'numFmt')
    const fontsInner = sectionInner(stylesXml, 'fonts')
    const fillsInner = sectionInner(stylesXml, 'fills')
    const bordersInner = sectionInner(stylesXml, 'borders')
    const cellXfsInner = sectionInner(stylesXml, 'cellXfs')
    if (fontsInner === null || fillsInner === null || cellXfsInner === null) {
      throw new Error(
        'The workbook stylesheet is missing fonts, fills, or cellXfs — style edits cannot be saved.',
      )
    }
    this.fonts = extractElements(fontsInner, 'font')
    this.fills = extractElements(fillsInner, 'fill')
    this.hadBordersSection = bordersInner !== null
    this.borders = bordersInner === null ? ['<border/>'] : extractElements(bordersInner, 'border')
    if (this.borders.length === 0) this.borders.push('<border/>')
    this.cellXfs = extractElements(cellXfsInner, 'xf')
    const dxfsInner = sectionInner(stylesXml, 'dxfs')
    this.hadDxfsSection = dxfsInner !== null
    this.dxfs = dxfsInner === null ? [] : extractElements(dxfsInner, 'dxf')
    if (this.fonts.length === 0 || this.cellXfs.length === 0) {
      throw new Error(
        'The workbook stylesheet has no base font or cell format — style edits cannot be saved.',
      )
    }
    this.originalCounts = {
      numFmts: this.numFmts.length,
      fonts: this.fonts.length,
      fills: this.fills.length,
      borders: this.borders.length,
      cellXfs: this.cellXfs.length,
      dxfs: this.dxfs.length,
    }
    this.nextNumFmtId =
      this.numFmts.reduce(
        (maximum, entry) => Math.max(maximum, Number(readAttribute(entry, 'numFmtId') ?? 0)),
        163,
      ) + 1
  }

  get changed(): boolean {
    return (
      this.numFmts.length !== this.originalCounts.numFmts ||
      this.fonts.length !== this.originalCounts.fonts ||
      this.fills.length !== this.originalCounts.fills ||
      this.borders.length !== this.originalCounts.borders ||
      this.cellXfs.length !== this.originalCounts.cellXfs ||
      this.dxfs.length !== this.originalCounts.dxfs
    )
  }

  /// Conditional-formatting highlight styles; deduped like every other list.
  internDxf(dxfXml: string): number {
    return internElement(this.dxfs, dxfXml)
  }

  /// Returns the cellXfs index of a format equal to the base format with the
  /// delta applied, appending new numFmt/font/fill/xf entries as needed.
  resolveStyle(baseXfIndex: number, delta: WorkbookStyleEdit): number {
    const cacheKey = `${baseXfIndex}|${JSON.stringify(delta, Object.keys(delta).sort())}`
    const cached = this.cache.get(cacheKey)
    if (cached !== undefined) return cached

    const baseXf = this.cellXfs[baseXfIndex] ?? this.cellXfs[0] ?? '<xf/>'
    let fontId = Number(readAttribute(baseXf, 'fontId') ?? 0)
    let fillId = Number(readAttribute(baseXf, 'fillId') ?? 0)
    let numFmtId = Number(readAttribute(baseXf, 'numFmtId') ?? 0)
    let borderId = Number(readAttribute(baseXf, 'borderId') ?? 0)
    const xfId = readAttribute(baseXf, 'xfId')

    if (hasFontDelta(delta)) {
      fontId = this.internFont(buildFont(this.fonts[fontId] ?? '<font/>', delta))
    }
    if (delta.fillColor !== undefined) {
      // Fill index 0 is the stylesheet's mandatory "none" pattern.
      fillId = delta.fillColor === null ? 0 : this.internFill(buildSolidFill(delta.fillColor))
    }
    if (delta.numberFormat !== undefined) {
      numFmtId = this.internNumberFormat(delta.numberFormat)
    }
    if (hasBorderDelta(delta)) {
      borderId = internElement(
        this.borders,
        buildBorder(this.borders[borderId] ?? '<border/>', delta),
      )
    }
    const alignment = buildAlignment(baseXf, delta)
    const protection = buildProtection(baseXf, delta)

    const attributes = [
      `numFmtId="${numFmtId}"`,
      `fontId="${fontId}"`,
      `fillId="${fillId}"`,
      `borderId="${borderId}"`,
      ...(xfId === undefined ? [] : [`xfId="${xfId}"`]),
      ...(numFmtId !== 0 ? ['applyNumberFormat="1"'] : []),
      ...(fontId !== 0 ? ['applyFont="1"'] : []),
      ...(fillId !== 0 ? ['applyFill="1"'] : []),
      ...(borderId !== 0 ? ['applyBorder="1"'] : []),
      ...(alignment !== '' ? ['applyAlignment="1"'] : []),
      ...(protection !== '' ? ['applyProtection="1"'] : []),
    ].join(' ')
    const children = `${alignment}${protection}`
    const xf = children === '' ? `<xf ${attributes}/>` : `<xf ${attributes}>${children}</xf>`

    const index = internElement(this.cellXfs, xf)
    this.cache.set(cacheKey, index)
    return index
  }

  serialize(): string {
    let result = this.source
    result = replaceSection(result, 'fonts', this.fonts)
    result = replaceSection(result, 'fills', this.fills)
    if (this.hadBordersSection) {
      result = replaceSection(result, 'borders', this.borders)
    } else if (this.borders.length > this.originalCounts.borders) {
      // Schema order: borders comes immediately after fills.
      const section = `<borders count="${this.borders.length}">${this.borders.join('')}</borders>`
      result = result.replace(/<\/fills>|<fills\b[^>]*\/>/, (match) => `${match}${section}`)
    }
    result = replaceSection(result, 'cellXfs', this.cellXfs)
    if (this.hadDxfsSection) {
      result = replaceSection(result, 'dxfs', this.dxfs)
    } else if (this.dxfs.length > 0) {
      const section = `<dxfs count="${this.dxfs.length}">${this.dxfs.join('')}</dxfs>`
      // Schema order: dxfs follows cellStyles (when present) or cellXfs.
      const anchor =
        /<\/cellStyles>|<cellStyles\b[^>]*\/>/.exec(result) ??
        /<\/cellXfs>|<cellXfs\b[^>]*\/>/.exec(result)
      if (anchor) {
        const at = anchor.index + anchor[0].length
        result = result.slice(0, at) + section + result.slice(at)
      }
    }
    if (this.numFmts.length > 0) {
      const section = `<numFmts count="${this.numFmts.length}">${this.numFmts.join('')}</numFmts>`
      if (sectionInner(result, 'numFmts') !== null) {
        result = result.replace(
          /<numFmts\b[^>]*>[\s\S]*?<\/numFmts>|<numFmts\b[^>]*\/>/,
          () => section,
        )
      } else {
        // Schema order: numFmts comes immediately before fonts.
        result = result.replace(/<fonts\b/, () => `${section}<fonts`)
      }
    }
    return result
  }

  private internFont(fontXml: string): number {
    return internElement(this.fonts, fontXml)
  }

  private internFill(fillXml: string): number {
    return internElement(this.fills, fillXml)
  }

  private internNumberFormat(pattern: string): number {
    const builtin = BUILTIN_NUMBER_FORMATS.get(pattern)
    if (builtin !== undefined) return builtin
    for (const entry of this.numFmts) {
      if (readAttribute(entry, 'formatCode') === escapeXmlAttribute(pattern)) {
        return Number(readAttribute(entry, 'numFmtId') ?? 0)
      }
    }
    const id = this.nextNumFmtId
    this.nextNumFmtId += 1
    this.numFmts.push(`<numFmt numFmtId="${id}" formatCode="${escapeXmlAttribute(pattern)}"/>`)
    return id
  }
}

const BUILTIN_NUMBER_FORMATS = new Map<string, number>([
  ['General', 0],
  ['0', 1],
  ['0.00', 2],
  ['#,##0', 3],
  ['#,##0.00', 4],
  ['0%', 9],
  ['0.00%', 10],
  ['0.00E+00', 11],
  ['@', 49],
])

function hasFontDelta(delta: WorkbookStyleEdit): boolean {
  return (
    delta.bold !== undefined ||
    delta.italic !== undefined ||
    delta.underline !== undefined ||
    delta.underlineStyle !== undefined ||
    delta.strikethrough !== undefined ||
    delta.fontFamily !== undefined ||
    delta.fontSize !== undefined ||
    delta.fontColor !== undefined
  )
}

/// Applies the delta to a copy of the base font XML. Only overridden child
/// elements are replaced; everything else (family, scheme, charset) is kept.
function buildFont(baseFontXml: string, delta: WorkbookStyleEdit): string {
  let inner = /<font\b[^>]*>([\s\S]*?)<\/font>/.exec(baseFontXml)?.[1] ?? ''
  const added: string[] = []
  const override = (pattern: RegExp, replacement: string): void => {
    inner = inner.replace(pattern, '')
    if (replacement !== '') added.push(replacement)
  }
  if (delta.bold !== undefined) override(/<b\b[^>]*\/?>/g, delta.bold ? '<b/>' : '')
  if (delta.italic !== undefined) override(/<i\b[^>]*\/?>/g, delta.italic ? '<i/>' : '')
  if (delta.underline !== undefined || delta.underlineStyle !== undefined) {
    const on = delta.underline ?? true
    override(
      /<u\b[^>]*\/?>/g,
      on ? (delta.underlineStyle === 'double' ? '<u val="double"/>' : '<u/>') : '',
    )
  }
  if (delta.strikethrough !== undefined) {
    override(/<strike\b[^>]*\/?>/g, delta.strikethrough ? '<strike/>' : '')
  }
  if (delta.fontSize !== undefined) {
    override(/<sz\b[^>]*\/?>/g, `<sz val="${delta.fontSize}"/>`)
  }
  if (delta.fontColor !== undefined) {
    override(
      /<color\b[^>]*\/?>/g,
      delta.fontColor === null ? '' : `<color rgb="${toArgb(delta.fontColor)}"/>`,
    )
  }
  if (delta.fontFamily !== undefined) {
    override(/<name\b[^>]*\/?>/g, `<name val="${escapeXmlAttribute(delta.fontFamily)}"/>`)
  }
  const content = `${added.join('')}${inner}`
  return content === '' ? '<font/>' : `<font>${content}</font>`
}

function buildSolidFill(fillColor: string): string {
  return `<fill><patternFill patternType="solid"><fgColor rgb="${toArgb(fillColor)}"/><bgColor indexed="64"/></patternFill></fill>`
}

const BORDER_EDGE_TAGS = ['left', 'right', 'top', 'bottom'] as const
const BORDER_DELTA_KEYS = {
  left: 'borderLeft',
  right: 'borderRight',
  top: 'borderTop',
  bottom: 'borderBottom',
} as const

function hasBorderDelta(delta: WorkbookStyleEdit): boolean {
  return (
    delta.borderTop !== undefined ||
    delta.borderBottom !== undefined ||
    delta.borderLeft !== undefined ||
    delta.borderRight !== undefined
  )
}

/// Applies edge deltas to a copy of the base border XML. Children rebuild in
/// schema order (left, right, top, bottom, diagonal); untouched edges and the
/// border element's own attributes (diagonalUp/Down) are kept verbatim.
function buildBorder(baseBorderXml: string, delta: WorkbookStyleEdit): string {
  const attributes = /<border\b([^>]*?)\/?>/.exec(baseBorderXml)?.[1] ?? ''
  const inner = /<border\b[^>]*>([\s\S]*?)<\/border>/.exec(baseBorderXml)?.[1] ?? ''
  const childOf = (tag: string): string =>
    new RegExp(`<${tag}\\b[^>]*/>|<${tag}\\b[^>]*>[\\s\\S]*?</${tag}>`).exec(inner)?.[0] ?? ''
  const children = BORDER_EDGE_TAGS.map((tag) => {
    const edge = delta[BORDER_DELTA_KEYS[tag]]
    if (edge === undefined) return childOf(tag)
    if (edge === null) return `<${tag}/>`
    const color = edge.color === undefined ? '' : `<color rgb="${toArgb(edge.color)}"/>`
    return color === ''
      ? `<${tag} style="${edge.style}"/>`
      : `<${tag} style="${edge.style}">${color}</${tag}>`
  })
  children.push(childOf('diagonal'))
  const content = children.join('')
  return content === '' && attributes.trim() === ''
    ? '<border/>'
    : `<border${attributes}>${content}</border>`
}

/** CT_CellAlignment attributes with no model field; readingOrder is the cell's RTL flag */
const ALIGNMENT_CARRIED = ['relativeIndent', 'justifyLastLine', 'shrinkToFit', 'readingOrder']

function buildAlignment(baseXf: string, delta: WorkbookStyleEdit): string {
  const baseAlignment = /<alignment\b[^>]*\/?>/.exec(baseXf)?.[0] ?? ''
  const horizontal = delta.horizontalAlignment ?? readAttribute(baseAlignment, 'horizontal')
  const vertical =
    delta.verticalAlignment !== undefined
      ? XLSX_VERTICAL[delta.verticalAlignment]
      : readAttribute(baseAlignment, 'vertical')
  const wrap =
    delta.wrapText !== undefined ? delta.wrapText : readAttribute(baseAlignment, 'wrapText') === '1'
  // 0 clears the rotation (the attribute's absence is "no rotation").
  const rotation =
    delta.textRotation !== undefined
      ? delta.textRotation === 0
        ? undefined
        : String(delta.textRotation)
      : readAttribute(baseAlignment, 'textRotation')
  const indent =
    delta.indent !== undefined
      ? delta.indent === 0
        ? undefined
        : String(delta.indent)
      : readAttribute(baseAlignment, 'indent')
  const modeled = [
    ...(horizontal ? [`horizontal="${horizontal}"`] : []),
    ...(vertical ? [`vertical="${vertical}"`] : []),
    ...(wrap ? ['wrapText="1"'] : []),
    ...(rotation ? [`textRotation="${rotation}"`] : []),
    ...(indent ? [`indent="${indent}"`] : []),
  ]
  // carrying must not create an <alignment> that would not otherwise exist, which would
  // start applying an alignment the cell was inheriting; once the element is there anyway
  // the cell already applies it, so the rest of its attributes belong with it
  // applyAlignment is xsd:boolean, which spells true both ways
  const applyAlignment = readCoreAttribute(baseXf, 'applyAlignment')
  const applies = modeled.length > 0 || applyAlignment === '1' || applyAlignment === 'true'
  const attributes =
    applies && ownsItsAlignment(baseXf)
      ? [...modeled, ...carriedAttributes(baseAlignment, ALIGNMENT_CARRIED)]
      : modeled
  return attributes.length === 0 ? '' : `<alignment ${attributes.join(' ')}/>`
}

/**
 * Whether the regex reads above can be trusted to have found the cell's own alignment.
 * CT_Xf allows only alignment, protection and extLst, so an xf holding nothing but the
 * first two cannot hide a second <alignment> or an applyAlignment belonging to someone
 * else. Anything further, an extension or an mc:AlternateContent branch, could supply
 * either, and a vendor's reading order is not the cell's; carry nothing and leave the
 * result as it was.
 */
function ownsItsAlignment(xf: string): boolean {
  for (const element of xf.matchAll(/<\/?([\w.-]+(?::[\w.-]+)?)/g)) {
    if (!/^(xf|alignment|protection)$/.test(element[1] ?? '')) return false
  }
  return true
}

/// Merges protection flag deltas over the base xf's <protection>. Attributes
/// at their OOXML defaults (locked=1, hidden=0) are omitted; an empty result
/// drops the element.
function buildProtection(baseXf: string, delta: WorkbookStyleEdit): string {
  const baseProtection = /<protection\b[^>]*\/?>/.exec(baseXf)?.[0] ?? ''
  const locked =
    delta.protectionLocked ?? (readAttribute(baseProtection, 'locked') === '0' ? false : undefined)
  const hidden =
    delta.protectionHidden ?? (readAttribute(baseProtection, 'hidden') === '1' ? true : undefined)
  const attributes = [
    ...(locked === false ? ['locked="0"'] : []),
    ...(hidden === true ? ['hidden="1"'] : []),
  ]
  return attributes.length === 0 ? '' : `<protection ${attributes.join(' ')}/>`
}

const XLSX_VERTICAL: Record<string, string> = {
  top: 'top',
  center: 'center',
  bottom: 'bottom',
}

function internElement(list: string[], element: string): number {
  const existing = list.indexOf(element)
  if (existing !== -1) return existing
  list.push(element)
  return list.length - 1
}

function sectionInner(xml: string, tag: string): string | null {
  const selfClosing = new RegExp(`<${tag}\\b[^>]*/>`)
  if (selfClosing.test(xml)) return ''
  const match = new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`).exec(xml)
  return match?.[1] ?? null
}

function replaceSection(xml: string, tag: string, elements: readonly string[]): string {
  const section = `<${tag} count="${elements.length}">${elements.join('')}</${tag}>`
  return xml.replace(
    new RegExp(`<${tag}\\b[^>]*>[\\s\\S]*?</${tag}>|<${tag}\\b[^>]*/>`),
    () => section,
  )
}

function extractElements(inner: string, tag: string): string[] {
  return [
    ...inner.matchAll(new RegExp(`<${tag}\\b[^>]*/>|<${tag}\\b[^>]*>[\\s\\S]*?</${tag}>`, 'g')),
  ].map((match) => match[0])
}

function readAttribute(element: string, name: string): string | undefined {
  return new RegExp(`\\b${name}="([^"]*)"`).exec(element)?.[1]
}

/** like readAttribute, but never matches a namespace-prefixed name */
function readCoreAttribute(element: string, name: string): string | undefined {
  return new RegExp(`(?<![\\w:.-])${name}="([^"]*)"`).exec(element)?.[1]
}

/** attributes the style model does not represent, kept verbatim off the base element */
function carriedAttributes(element: string, names: readonly string[]): string[] {
  return names.flatMap((name) => {
    const value = readCoreAttribute(element, name)
    return value === undefined ? [] : [`${name}="${value}"`]
  })
}

function toArgb(hexColor: string): string {
  return `FF${hexColor.slice(1).toUpperCase()}`
}

function escapeXmlAttribute(input: string): string {
  return input
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;')
}

// ── StylesheetReader (read path: cellXfs index → CellFormatState) ───────────

/** OOXML ARGB ("AARRGGBB" / "RRGGBB") → 6-digit RGB without '#'. */
function argbToRgb(argb: string): string | undefined {
  const hex = argb.trim()
  if (!/^[0-9A-Fa-f]{6}$/.test(hex) && !/^[0-9A-Fa-f]{8}$/.test(hex)) return undefined
  return (hex.length === 8 ? hex.slice(2) : hex).toUpperCase()
}

/**
 * Read-only resolver for the browser-facing presentation snapshot: maps a
 * cell's cellXfs index to the editable subset of its resolved format
 * (CellFormatState). Unmodeled properties (diagonal borders, theme colors,
 * exotic border styles) stay in the file's own XML — the byte-preserving
 * save path keeps them for cells that are not re-emitted.
 *
 * Mirrors the delta vocabulary of StylesheetEditor: only properties the
 * editor can round-trip are resolved, so a rendered style can always be
 * edited back.
 */
export class StylesheetReader {
  private readonly fonts: readonly string[]
  private readonly fills: readonly string[]
  private readonly cellXfs: readonly string[]
  /** Raw `<border>` entries from the styles.xml borders section, in file order. */
  private readonly borderEntries: readonly string[]
  /**
   * numFmtId → formatCode map, parsed from the styles.xml <numFmts> section.
   * Built once at construction time. Used by resolve() to attach a
   * numberFormat pattern to CellFormatState when the cellXfs entry references
   * a custom numFmt (numFmtId ≥ 164, the threshold above which user-defined
   * formats live). Built-in numFmtIds (0..163) are mapped via BUILTIN_NUMFMTS
   * so common formats (General, 0, 0.00, #,##0, $#,##0.00, …) survive
   * round-trip without needing a <numFmt> entry.
   */
  private readonly numFmtByCode: ReadonlyMap<number, string>
  /** Raw `<dxf>` entries from the styles.xml dxfs section, in file order. */
  private readonly dxfEntries: readonly string[]
  private readonly cache = new Map<number, CellFormatState | undefined>()

  constructor(stylesXml: string) {
    const fontsInner = sectionInner(stylesXml, 'fonts')
    const fillsInner = sectionInner(stylesXml, 'fills')
    const cellXfsInner = sectionInner(stylesXml, 'cellXfs')
    this.fonts = fontsInner === null ? [] : extractElements(fontsInner, 'font')
    this.fills = fillsInner === null ? [] : extractElements(fillsInner, 'fill')
    this.cellXfs = cellXfsInner === null ? [] : extractElements(cellXfsInner, 'xf')
    const bordersInner = sectionInner(stylesXml, 'borders')
    this.borderEntries = bordersInner === null ? [] : extractElements(bordersInner, 'border')
    const dxfsInner = sectionInner(stylesXml, 'dxfs')
    this.dxfEntries = dxfsInner === null ? [] : extractElements(dxfsInner, 'dxf')
    const numFmtsInner = sectionInner(stylesXml, 'numFmts')
    const numFmtEntries = numFmtsInner === null ? [] : extractElements(numFmtsInner, 'numFmt')
    const byCode = new Map<number, string>()
    for (const entry of numFmtEntries) {
      const id = Number(readAttribute(entry, 'numFmtId') ?? 'NaN')
      const code = readAttribute(entry, 'formatCode')
      if (Number.isInteger(id) && code) byCode.set(id, decodeXmlText(code))
    }
    this.numFmtByCode = byCode
  }

  /**
   * Raw `<dxf>` XML at a differential-format index; undefined when the index
   * is out of range. EXCEL-024: the conditional-formatting reader resolves
   * rule styling through this accessor so the browser never sees style XML.
   */
  dxfAt(dxfIndex: number): string | undefined {
    return this.dxfEntries[dxfIndex]
  }

  /**
   * Resolved editable format of cellXfs[index]; undefined when the index is
   * out of range, or the resolved format carries no property the editor
   * models (absent = "no explicit format", per WorksheetState.styles).
   */
  formatAt(xfIndex: number): CellFormatState | undefined {
    if (xfIndex < 0 || xfIndex >= this.cellXfs.length) return undefined
    const cached = this.cache.get(xfIndex)
    if (cached !== undefined || this.cache.has(xfIndex)) return cached
    const resolved = this.resolve(xfIndex)
    this.cache.set(xfIndex, resolved)
    return resolved
  }

  private resolve(xfIndex: number): CellFormatState | undefined {
    const xf = this.cellXfs[xfIndex] ?? ''
    const format: {
      bold?: boolean
      italic?: boolean
      underline?: boolean
      strikethrough?: boolean
      fontSize?: number
      fontFamily?: string
      fontColor?: string
      fillColor?: string
      horizontalAlign?: 'left' | 'center' | 'right'
      verticalAlign?: 'top' | 'center' | 'bottom'
      wrapText?: boolean
      numberFormat?: string
      textRotation?: number | 'vertical'
      indent?: number
      border?: CellBorderState
    } = {}
    // Number format — read the cellXfs numFmtId and resolve to a pattern.
    // Custom numFmtIds (≥164) live in <numFmts>; built-in ids (0..163) are
    // mapped via BUILTIN_NUMFMTS. numFmtId 0 ("General") is skipped so a
    // default cell doesn't claim an explicit numberFormat.
    const numFmtIdAttr = readAttribute(xf, 'numFmtId')
    if (numFmtIdAttr !== undefined) {
      const id = Number(numFmtIdAttr)
      if (Number.isInteger(id) && id > 0) {
        const pattern = this.numFmtByCode.get(id) ?? BUILTIN_NUMFMTS.get(id)
        if (pattern) format.numberFormat = pattern
      }
    }
    // Font-derived marks
    const fontId = Number(readAttribute(xf, 'fontId') ?? 0)
    const font = this.fonts[fontId] ?? ''
    if (/<b\b[^>]*\/?>/.test(font)) format.bold = true
    if (/<i\b[^>]*\/?>/.test(font)) format.italic = true
    const uVal = readAttribute(/<u\b[^>]*\/?>/.exec(font)?.[0] ?? '', 'val')
    if (/<u\b[^>]*\/?>/.test(font) && uVal !== 'none') format.underline = true
    if (/<strike\b[^>]*\/?>/.test(font)) format.strikethrough = true
    const sz = Number(readAttribute(/<sz\b[^>]*\/?>/.exec(font)?.[0] ?? '', 'val'))
    if (Number.isFinite(sz) && sz > 0) format.fontSize = sz
    const name = readAttribute(/<name\b[^>]*\/?>/.exec(font)?.[0] ?? '', 'val')
    if (name) format.fontFamily = decodeXmlText(name)
    const fontColor = argbToRgb(
      readAttribute(/<color\b[^>]*\/?>/.exec(font)?.[0] ?? '', 'rgb') ?? '',
    )
    if (fontColor) format.fontColor = fontColor
    // Solid fill
    const fillId = Number(readAttribute(xf, 'fillId') ?? 0)
    const fill = this.fills[fillId] ?? ''
    const pattern = /<patternFill\b([^>]*)>/.exec(fill)?.[1] ?? ''
    if (readAttribute(pattern, 'patternType') === 'solid') {
      const fg = argbToRgb(readAttribute(/<fgColor\b[^>]*\/?>/.exec(fill)?.[0] ?? '', 'rgb') ?? '')
      if (fg) format.fillColor = fg
    }
    // Alignment (child of the xf)
    const alignment = /<alignment\b[^>]*\/?>/.exec(xf)?.[0] ?? ''
    if (alignment) {
      const horizontal = readAttribute(alignment, 'horizontal')
      if (horizontal === 'left' || horizontal === 'center' || horizontal === 'right') {
        format.horizontalAlign = horizontal
      }
      const vertical = readAttribute(alignment, 'vertical')
      if (vertical === 'top' || vertical === 'center' || vertical === 'bottom') {
        format.verticalAlign = vertical
      }
      if (readAttribute(alignment, 'wrapText') === '1') format.wrapText = true
      // Text rotation (EXCEL-027): OOXML stores 0..180 (91..180 = clockwise
      // down as 90 + degrees) plus the special 255 = vertically stacked.
      // 0 is the schema default and never modeled; values outside the OOXML
      // domain (181..254, negatives, non-integers) are ignored for modeling —
      // the raw attribute stays byte-preserved on no-op saves because such a
      // cell never carries a journaled rotation delta.
      const rotationRaw = readAttribute(alignment, 'textRotation')
      if (rotationRaw !== undefined) {
        const rotation = Number(rotationRaw)
        if (Number.isInteger(rotation)) {
          if (rotation === 255) format.textRotation = 'vertical'
          else if (rotation >= 1 && rotation <= 180) format.textRotation = rotation
        }
      }
      // Indent (EXCEL-027): OOXML unsignedInt steps; 0 is the default and
      // never modeled, malformed values are ignored for modeling.
      const indentRaw = readAttribute(alignment, 'indent')
      if (indentRaw !== undefined) {
        const indent = Number(indentRaw)
        if (Number.isInteger(indent) && indent >= 1) format.indent = indent
      }
    }
    // Borders (EXCEL-027): the xf's borderId resolves through the borders
    // section; each of the four modeled sides appears only when the file
    // carries a border there with one of the 13 canonical ST_BorderStyle
    // names. Diagonal edges and unmodelable styles are skipped — the writer
    // preserves them verbatim for cells that are not re-emitted.
    const borderId = Number(readAttribute(xf, 'borderId') ?? 0)
    const borderXml = this.borderEntries[borderId]
    if (borderXml !== undefined) {
      const border = readBorderState(borderXml)
      if (border !== undefined) format.border = border
    }
    return Object.keys(format).length > 0 ? format : undefined
  }
}

/** The 13 ST_BorderStyle names the canonical model round-trips. */
const MODELABLE_BORDER_STYLES: ReadonlySet<string> = new Set([
  'thin',
  'medium',
  'thick',
  'dashed',
  'dotted',
  'double',
  'hair',
  'dashDot',
  'dashDotDot',
  'mediumDashed',
  'mediumDashDot',
  'mediumDashDotDot',
  'slantDashDot',
])

/**
 * Parses the four modeled sides off one `<border>` element. A side appears
 * only when its element carries a modelable style; the color attaches only
 * for a readable rgb value (theme/indexed/auto colors leave color unset).
 * Returns undefined when no side carries a modelable border.
 */
function readBorderState(borderXml: string): CellBorderState | undefined {
  const sideOf = (
    tag: string,
  ): { readonly style: CellBorderStyleName; readonly color?: string | undefined } | undefined => {
    const edge =
      new RegExp(`<${tag}\\b[^>]*/>|<${tag}\\b[^>]*>[\\s\\S]*?</${tag}>`).exec(borderXml)?.[0] ?? ''
    if (edge === '') return undefined
    const style = readAttribute(edge, 'style')
    if (style === undefined || !MODELABLE_BORDER_STYLES.has(style)) return undefined
    const color = argbToRgb(readAttribute(/<color\b[^>]*\/?>/.exec(edge)?.[0] ?? '', 'rgb') ?? '')
    return {
      style: style as CellBorderStyleName,
      ...(color === undefined ? {} : { color }),
    }
  }
  const state: {
    top?: { readonly style: CellBorderStyleName; readonly color?: string | undefined }
    bottom?: { readonly style: CellBorderStyleName; readonly color?: string | undefined }
    left?: { readonly style: CellBorderStyleName; readonly color?: string | undefined }
    right?: { readonly style: CellBorderStyleName; readonly color?: string | undefined }
  } = {}
  const left = sideOf('left')
  if (left !== undefined) state.left = left
  const right = sideOf('right')
  if (right !== undefined) state.right = right
  const topEdge = sideOf('top')
  if (topEdge !== undefined) state.top = topEdge
  const bottom = sideOf('bottom')
  if (bottom !== undefined) state.bottom = bottom
  return Object.keys(state).length > 0 ? state : undefined
}

/** XML entity decoding shared with the gateway's text decoding. */
function decodeXmlText(input: string): string {
  return input
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_m, code: string) => String.fromCodePoint(Number(code)))
    .replace(/&amp;/g, '&')
}
