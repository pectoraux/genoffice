import type { TableArea } from './xlsx-table-add'

/// Table (ListObject) READ side (EXCEL-021): resolves each worksheet's
/// <tableParts> through its rels and parses the table parts into the
/// canonical SheetTableInfo — metadata plus the RESOLVED banding colors the
/// browser paints (header/stripe/totals/first-last-column fills, frame
/// color). Color resolution ports the desktop sidecar's semantics exactly:
/// theme accents with Excel's HSL tint transform for the built-in
/// Light/Medium/Dark families, and the file's own <tableStyle> dxfs for
/// custom styles. Tables the model cannot represent are skipped per table
/// (desktop read_sheet_tables parity); a worksheet whose table wiring is
/// unreadable fails closed PER SHEET (TableReadError) — the workbook still
/// opens and a no-op save preserves the file's XML byte-for-byte.

export class TableReadError extends Error {}

export interface SheetTableInfo {
  /// 0-based range including the header and totals rows.
  readonly area: TableArea
  readonly headerRowCount: number
  /// Rows at the bottom styled as the totals band. Absent = 0.
  readonly totalsRowCount?: number
  readonly showRowStripes: boolean
  readonly showColumnStripes: boolean
  /// displayName ?? name — the token structured references use. Absent when
  /// the part carries neither attribute (still renderable banding).
  readonly name?: string
  readonly columns: readonly string[]
  readonly styleName?: string
  // ── Resolved colors (#RRGGBB). Absent = paint nothing for that band. ──
  readonly headerFill?: string
  readonly headerFontColor?: string
  readonly stripeFill?: string
  readonly secondRowStripeFill?: string
  readonly columnStripeFill?: string
  readonly secondColumnStripeFill?: string
  readonly wholeTableFill?: string
  /// Emphasis fills — already gated by showFirstColumn/showLastColumn
  /// (absent when the flag is off, mirroring the desktop's TableInfo).
  readonly firstColumnFill?: string
  readonly lastColumnFill?: string
  readonly totalRowFill?: string
  readonly totalRowFontColor?: string
  readonly firstHeaderCellFontColor?: string
  /// Style frame color (outline + header rule) for border-drawn families.
  readonly borderColor?: string
}

interface CustomTablePalette {
  wholeTableFill?: string
  headerFill?: string
  headerFontColor?: string
  totalRowFill?: string
  totalRowFontColor?: string
  firstColumnFill?: string
  lastColumnFill?: string
  stripeFill?: string
  secondRowStripeFill?: string
  columnStripeFill?: string
  secondColumnStripeFill?: string
  firstHeaderCellFontColor?: string
}

/// Theme palette + custom <tableStyle> dxfs, parsed once per workbook.
export interface TableStyleContext {
  /// [lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink] — the `theme`
  /// attribute index order (light/dark swapped vs clrScheme document
  /// order). null when theme1.xml is absent or any slot is unresolvable
  /// (ColorContext::default() parity — callers fall back to Office
  /// defaults).
  readonly theme: readonly (readonly [number, number, number])[] | null
  readonly customStyles: ReadonlyMap<string, CustomTablePalette>
}

type Rgb = readonly [number, number, number]

const TABLE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'

/// Office default accents — the fallback when the theme palette is absent.
const DEFAULT_ACCENTS: readonly Rgb[] = [
  [0x44, 0x72, 0xc4],
  [0xed, 0x7d, 0x31],
  [0xa5, 0xa5, 0xa5],
  [0xff, 0xc0, 0x00],
  [0x5b, 0x9b, 0xd5],
  [0x70, 0xad, 0x47],
]

/// Legacy indexed palette, ECMA-376 §18.8.27. Indexes 64/65 are the system
/// window text/background colors.
const INDEXED_COLORS: readonly string[] = [
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF',
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF',
  '800000',
  '008000',
  '000080',
  '808000',
  '800080',
  '008080',
  'C0C0C0',
  '808080',
  '9999FF',
  '993366',
  'FFFFCC',
  'CCFFFF',
  '660066',
  'FF8080',
  '0066CC',
  'CCCCFF',
  '000080',
  'FF00FF',
  'FFFF00',
  '00FFFF',
  '800080',
  '800000',
  '008080',
  '0000FF',
  '00CCFF',
  'CCFFFF',
  'CCFFCC',
  'FFFF99',
  '99CCFF',
  'FF99CC',
  'CC99FF',
  'FFCC99',
  '3366FF',
  '33CCCC',
  '99CC00',
  'FFCC00',
  'FF9900',
  'FF6600',
  '666699',
  '969696',
  '003366',
  '339966',
  '003300',
  '333300',
  '993300',
  '993366',
  '333399',
  '333333',
  '000000',
  'FFFFFF',
]

/// Parses xl/theme/theme1.xml's <a:clrScheme> into the theme-attribute
/// index order. null when the part is missing or any slot lacks a usable
/// srgbClr/sysClr value (desktop read_theme_palette parity).
export async function readTableThemePalette(
  reader: Pick<EntryReader, 'readText' | 'has'>,
): Promise<readonly Rgb[] | null> {
  if (!(await reader.has('xl/theme/theme1.xml'))) return null
  const xml = await reader.readText('xl/theme/theme1.xml')
  const scheme = /<a:clrScheme\b[^>]*>([\s\S]*?)<\/a:clrScheme>/.exec(xml)?.[1]
  if (scheme === undefined) return null
  const order = [
    'lt1',
    'dk1',
    'lt2',
    'dk2',
    'accent1',
    'accent2',
    'accent3',
    'accent4',
    'accent5',
    'accent6',
    'hlink',
    'folHlink',
  ]
  const theme: Rgb[] = []
  for (const slot of order) {
    const element = new RegExp(`<a:${slot}\\b[^>]*>([\\s\\S]*?)</a:${slot}>`).exec(scheme)?.[1]
    // Self-closing slots (<a:lt1 val="..."/>) carry no color of their own.
    if (element === undefined) return null
    const hex =
      /<a:srgbClr\b[^>]*\bval="([^"]+)"/.exec(element)?.[1] ??
      /<a:sysClr\b[^>]*\blastClr="([^"]+)"/.exec(element)?.[1]
    const rgb = hex === undefined ? undefined : parseHexRgb(hex)
    if (rgb === undefined) return null
    theme.push(rgb)
  }
  return theme
}

function parseHexRgb(hex: string): Rgb | undefined {
  const value = hex.length === 8 ? hex.slice(2) : hex
  if (!/^[0-9A-Fa-f]{6}$/.test(value)) return undefined
  return [
    Number.parseInt(value.slice(0, 2), 16),
    Number.parseInt(value.slice(2, 4), 16),
    Number.parseInt(value.slice(4, 6), 16),
  ]
}

/// Reads xl/styles.xml's <tableStyles> into per-style band palettes: each
/// <tableStyleElement type=… dxfId=…> maps to the dxfs entry's fill/font
/// color. Missing parts or sections yield an empty map (no custom styles).
export async function readCustomTableStyles(
  reader: Pick<EntryReader, 'readText' | 'has'>,
  theme: readonly Rgb[] | null,
): Promise<ReadonlyMap<string, CustomTablePalette>> {
  const palettes = new Map<string, CustomTablePalette>()
  if (!(await reader.has('xl/styles.xml'))) return palettes
  const xml = await reader.readText('xl/styles.xml')
  const dxfs = parseDxfSection(xml, theme)
  for (const match of xml.matchAll(
    /<tableStyle\b[^>]*\bname="([^"]+)"[^>]*>([\s\S]*?)<\/tableStyle>/g,
  )) {
    const name = decodeXmlText(match[1] ?? '')
    const body = match[2] ?? ''
    if (name === '') continue
    const palette: CustomTablePalette = {}
    for (const elementMatch of body.matchAll(/<tableStyleElement\b([^>]*)\/?>/g)) {
      const attributes = elementMatch[1] ?? ''
      const kind = /(?:^|\s)type="([^"]*)"/.exec(attributes)?.[1]
      const dxfId = /(?:^|\s)dxfId="([^"]*)"/.exec(attributes)?.[1]
      if (kind === undefined || dxfId === undefined) continue
      const dxf = dxfs[Number(dxfId)]
      if (dxf === undefined) continue
      switch (kind) {
        case 'wholeTable':
          if (dxf.fillColor !== undefined) palette.wholeTableFill = dxf.fillColor
          break
        case 'headerRow':
          if (dxf.fillColor !== undefined) palette.headerFill = dxf.fillColor
          if (dxf.fontColor !== undefined) palette.headerFontColor = dxf.fontColor
          break
        case 'totalRow':
          if (dxf.fillColor !== undefined) palette.totalRowFill = dxf.fillColor
          if (dxf.fontColor !== undefined) palette.totalRowFontColor = dxf.fontColor
          break
        case 'firstColumn':
          if (dxf.fillColor !== undefined) palette.firstColumnFill = dxf.fillColor
          break
        case 'lastColumn':
          if (dxf.fillColor !== undefined) palette.lastColumnFill = dxf.fillColor
          break
        case 'firstRowStripe':
          if (dxf.fillColor !== undefined) palette.stripeFill = dxf.fillColor
          break
        case 'secondRowStripe':
          if (dxf.fillColor !== undefined) palette.secondRowStripeFill = dxf.fillColor
          break
        case 'firstColumnStripe':
          if (dxf.fillColor !== undefined) palette.columnStripeFill = dxf.fillColor
          break
        case 'secondColumnStripe':
          if (dxf.fillColor !== undefined) palette.secondColumnStripeFill = dxf.fillColor
          break
        case 'firstHeaderCell':
          if (dxf.fontColor !== undefined) palette.firstHeaderCellFontColor = dxf.fontColor
          break
        default:
          break
      }
    }
    palettes.set(name, palette)
  }
  return palettes
}

interface DxfColors {
  fillColor?: string
  fontColor?: string
}

/// The <dxfs> section. Solid dxf fills carry the color in bgColor, unlike
/// cell fills (fgColor-first) — the OOXML differential-format quirk.
function parseDxfSection(stylesXml: string, theme: readonly Rgb[] | null): DxfColors[] {
  const section = /<dxfs\b[^>]*>([\s\S]*?)<\/dxfs>/.exec(stylesXml)?.[1]
  if (section === undefined) return []
  const out: DxfColors[] = []
  for (const match of section.matchAll(/<dxf\b[^>]*>([\s\S]*?)<\/dxf>|<dxf\b[^>]*\/>/g)) {
    const body = match[1] ?? ''
    const fontColor = resolveDxfColor(/<font\b[\s\S]*?<color\b([^>]*)\/?>/.exec(body)?.[1], theme)
    const pattern = /<patternFill\b([^>]*)>([\s\S]*?)<\/patternFill>/.exec(body)
    let fillColor: string | undefined
    if (pattern !== null) {
      const patternAttributes = pattern[1] ?? ''
      const patternBody = pattern[2] ?? ''
      const patternType = /(?:^|\s)patternType="([^"]*)"/.exec(patternAttributes)?.[1]
      if (patternType !== 'none') {
        const bgAttributes = /<bgColor\b([^>]*)\/?>/.exec(patternBody)?.[1]
        const fgAttributes = /<fgColor\b([^>]*)\/?>/.exec(patternBody)?.[1]
        fillColor = resolveDxfColor(bgAttributes ?? fgAttributes, theme)
      }
    }
    out.push({
      ...(fillColor !== undefined ? { fillColor } : {}),
      ...(fontColor !== undefined ? { fontColor } : {}),
    })
  }
  return out
}

/// One color element's attribute string → #RRGGBB, resolving rgb / indexed /
/// theme+tint in that order (desktop resolve_color parity).
function resolveDxfColor(
  attributes: string | undefined,
  theme: readonly Rgb[] | null,
): string | undefined {
  if (attributes === undefined) return undefined
  const rgb = /(?:^|\s)rgb="([^"]*)"/.exec(attributes)?.[1]
  if (rgb !== undefined) {
    const value = rgb.length === 8 ? rgb.slice(2) : rgb
    return `#${value}`
  }
  const indexed = /(?:^|\s)indexed="([^"]*)"/.exec(attributes)?.[1]
  if (indexed !== undefined && /^\d+$/.test(indexed)) {
    const value = INDEXED_COLORS[Number(indexed)]
    return value === undefined ? undefined : `#${value}`
  }
  const themeIndex = /(?:^|\s)theme="([^"]*)"/.exec(attributes)?.[1]
  if (themeIndex === undefined || !/^\d+$/.test(themeIndex)) return undefined
  const base = theme?.[Number(themeIndex)]
  if (base === undefined) return undefined
  const tintText = /(?:^|\s)tint="([^"]*)"/.exec(attributes)?.[1]
  const tint =
    tintText !== undefined && /^[+-]?\d*\.?\d+(?:[eE][+-]?\d+)?$/.test(tintText)
      ? Number(tintText)
      : 0
  return tintToHex(base, tint)
}

/// Built-in table-style palette (header fill, header font, row stripe) for
/// the Light/Medium/Dark families — a faithful port of the desktop
/// sidecar's table_style_palette. A nameless tableStyleInfo is Excel's
/// style "None": nothing gets painted. Unknown names resolve as Medium2.
export function resolveTableStylePalette(
  styleName: string | undefined,
  theme: readonly Rgb[] | null,
): { headerFill?: string; headerFontColor?: string; stripeFill?: string } {
  if (styleName === undefined) return {}
  let family: 'light' | 'medium' | 'dark'
  let number: number
  if (styleName.startsWith('TableStyleLight')) {
    family = 'light'
    number = parseStyleNumber(styleName.slice('TableStyleLight'.length), 1)
  } else if (styleName.startsWith('TableStyleMedium')) {
    family = 'medium'
    number = parseStyleNumber(styleName.slice('TableStyleMedium'.length), 2)
  } else if (styleName.startsWith('TableStyleDark')) {
    family = 'dark'
    number = parseStyleNumber(styleName.slice('TableStyleDark'.length), 1)
  } else {
    family = 'medium'
    number = 2
  }
  const shifted = number === 0 ? 0 : number - 1
  const accentIndex = shifted % 7
  const variant = Math.floor(shifted / 7)
  // Column 1 of each family is the theme's neutral text color (dk1), not a
  // literal gray.
  const base: Rgb =
    accentIndex === 0
      ? (theme?.[1] ?? [0x00, 0x00, 0x00])
      : (theme?.[3 + accentIndex] ?? DEFAULT_ACCENTS[accentIndex - 1]!)
  if (family === 'light') {
    const stripe = tintToHex(base, 0.85)
    if (variant === 1) {
      return { headerFill: rgbHex(base), headerFontColor: '#FFFFFF', stripeFill: stripe }
    }
    // Light 1-7 / 15-21: unfilled bold header in the accent color.
    return { headerFontColor: rgbHex(base), stripeFill: stripe }
  }
  if (family === 'dark') {
    return {
      headerFill: tintToHex(base, -0.25),
      headerFontColor: '#FFFFFF',
      stripeFill: tintToHex(base, 0.4),
    }
  }
  const header = variant === 2 ? tintToHex(base, -0.25) : rgbHex(base)
  // Medium 8-14 are Excel's "full color" block: darker stripes (accent tint
  // 0.6, e.g. Medium9 #B8CCE4) than the white-bodied blocks (tint 0.8, e.g.
  // Medium2 #DCE6F1).
  const stripeTint = variant === 1 ? 0.6 : 0.8
  return { headerFill: header, headerFontColor: '#FFFFFF', stripeFill: tintToHex(base, stripeTint) }
}

/// Table style column 1 of the Light block (Light1-7) draws its frame in
/// the base color; the filled families' borders stay unmodelled.
export function resolveTableStyleBorder(
  styleName: string | undefined,
  theme: readonly Rgb[] | null,
): string | undefined {
  if (styleName === undefined || !styleName.startsWith('TableStyleLight')) return undefined
  const number = parseStyleNumber(styleName.slice('TableStyleLight'.length), 1)
  if (Math.floor((number === 0 ? 0 : number - 1) / 7) !== 0) return undefined
  const accentIndex = (number === 0 ? 0 : number - 1) % 7
  const base: Rgb =
    accentIndex === 0
      ? (theme?.[1] ?? [0x00, 0x00, 0x00])
      : (theme?.[3 + accentIndex] ?? DEFAULT_ACCENTS[accentIndex - 1]!)
  return rgbHex(base)
}

function parseStyleNumber(rest: string, fallback: number): number {
  return /^\d+$/.test(rest) ? Number(rest) : fallback
}

/// Excel's tint transform: scale HSL luminance toward black (tint < 0) or
/// white (tint > 0).
function tintToHex(base: Rgb, tint: number): string {
  const [red, green, blue] = applyTint(base, tint)
  return `#${hex2(red)}${hex2(green)}${hex2(blue)}`
}

function rgbHex(base: Rgb): string {
  return `#${hex2(base[0])}${hex2(base[1])}${hex2(base[2])}`
}

function hex2(value: number): string {
  return Math.round(value).toString(16).padStart(2, '0').toUpperCase()
}

function applyTint(rgb: Rgb, tint: number): Rgb {
  if (tint === 0) return rgb
  const [hue, saturation, luminance] = rgbToHsl(rgb)
  const next = tint < 0 ? luminance * (1 + tint) : luminance * (1 - tint) + tint
  return hslToRgb(hue, saturation, Math.min(1, Math.max(0, next)))
}

function rgbToHsl([red8, green8, blue8]: Rgb): [number, number, number] {
  const red = red8 / 255
  const green = green8 / 255
  const blue = blue8 / 255
  const maximum = Math.max(red, green, blue)
  const minimum = Math.min(red, green, blue)
  const luminance = (maximum + minimum) / 2
  if (maximum === minimum) return [0, 0, luminance]
  const delta = maximum - minimum
  const saturation = luminance > 0.5 ? delta / (2 - maximum - minimum) : delta / (maximum + minimum)
  let hue: number
  if (maximum === red) {
    hue = (green - blue) / delta + (green < blue ? 6 : 0)
  } else if (maximum === green) {
    hue = (blue - red) / delta + 2
  } else {
    hue = (red - green) / delta + 4
  }
  return [hue / 6, saturation, luminance]
}

function hslToRgb(hue: number, saturation: number, luminance: number): Rgb {
  if (saturation === 0) {
    const value = Math.round(luminance * 255)
    return [value, value, value]
  }
  const q =
    luminance < 0.5 ? luminance * (1 + saturation) : luminance + saturation - luminance * saturation
  const p = 2 * luminance - q
  const channel = (t0: number): number => {
    let t = t0
    if (t < 0) t += 1
    if (t > 1) t -= 1
    let value: number
    if (t < 1 / 6) {
      value = p + (q - p) * 6 * t
    } else if (t < 1 / 2) {
      value = q
    } else if (t < 2 / 3) {
      value = p + (q - p) * (2 / 3 - t) * 6
    } else {
      value = p
    }
    return Math.round(value * 255)
  }
  return [channel(hue + 1 / 3), channel(hue), channel(hue - 1 / 3)]
}

interface EntryReader {
  readText(path: string): Promise<string>
  has(path: string): Promise<boolean>
}

/// Resolves a worksheet's table parts through its rels (COMMENTS_REL_TYPE
/// two-step lookup pattern — attribute order varies by producer, so never
/// assume Id precedes Target). A worksheet whose <tablePart> references a
/// missing relationship fails closed with TableReadError.
export async function parseSheetTables(
  reader: EntryReader,
  worksheetPath: string,
  worksheetXml: string,
  context: TableStyleContext,
): Promise<readonly SheetTableInfo[]> {
  const tablePartIds = [...worksheetXml.matchAll(/<tablePart\b[^>]*\br:id="([^"]+)"/g)].map(
    (match) => match[1]!,
  )
  if (tablePartIds.length === 0) return []
  const relsPath = worksheetPath.replace(/^(xl\/worksheets\/)([^/]+)$/, '$1_rels/$2.rels')
  if (!(await reader.has(relsPath))) {
    throw new TableReadError(`${worksheetPath} references tables but ${relsPath} is missing.`)
  }
  const relsXml = await reader.readText(relsPath)
  const tables: SheetTableInfo[] = []
  for (const relId of tablePartIds) {
    const relationshipXml = new RegExp(
      `<Relationship\\b[^>]*\\bId="${escapeRegExp(relId)}"[^>]*/?>`,
    ).exec(relsXml)?.[0]
    const target =
      relationshipXml === undefined ? undefined : /\bTarget="([^"]+)"/.exec(relationshipXml)?.[1]
    if (target === undefined) {
      throw new TableReadError(
        `${worksheetPath} references table ${relId} but its relationship is missing.`,
      )
    }
    const tablePath = resolveRelTarget(worksheetPath, target)
    if (!(await reader.has(tablePath))) continue
    const table = parseTablePart(await reader.readText(tablePath), context)
    if (table !== null) tables.push(table)
  }
  return tables
}

/// One table part → SheetTableInfo. null when the part carries no readable
/// ref (the desktop skips such tables — banding cannot address cells).
function parseTablePart(tableXml: string, context: TableStyleContext): SheetTableInfo | null {
  const open = /<table\b[^>]*>/.exec(tableXml)?.[0]
  if (open === undefined) return null
  const ref = /\bref="([^"]+)"/.exec(open)?.[1]
  const area = ref === undefined ? undefined : parseTableRef(ref)
  if (area === undefined) return null
  const headerRowCount = /\bheaderRowCount="([0-9]+)"/.exec(open)?.[1]
  const totalsRowCount = /\btotalsRowCount="([0-9]+)"/.exec(open)?.[1]
  // displayName is the token structured references use; name is a fallback.
  const name = /\bdisplayName="([^"]*)"/.exec(open)?.[1] ?? /\bname="([^"]*)"/.exec(open)?.[1]
  const columns = [...tableXml.matchAll(/<tableColumn\b[^>]*\bname="([^"]*)"/g)].map((match) =>
    decodeXmlText(match[1] ?? ''),
  )
  const styleInfo = /<tableStyleInfo\b([^>]*)\/?>/.exec(tableXml)?.[1] ?? ''
  const styleName = /(?:^|\s)name="([^"]*)"/.exec(styleInfo)?.[1]
  const showRowStripes = isOn(/(?:^|\s)showRowStripes="([^"]*)"/.exec(styleInfo)?.[1])
  const showColumnStripes = isOn(/(?:^|\s)showColumnStripes="([^"]*)"/.exec(styleInfo)?.[1])
  const showFirstColumn = isOn(/(?:^|\s)showFirstColumn="([^"]*)"/.exec(styleInfo)?.[1])
  const showLastColumn = isOn(/(?:^|\s)showLastColumn="([^"]*)"/.exec(styleInfo)?.[1])
  const custom =
    styleName !== undefined ? context.customStyles.get(decodeXmlText(styleName)) : undefined
  const palette =
    custom !== undefined
      ? {
          headerFill: custom.headerFill,
          headerFontColor: custom.headerFontColor,
          stripeFill: custom.stripeFill,
        }
      : resolveTableStylePalette(
          styleName !== undefined ? decodeXmlText(styleName) : undefined,
          context.theme,
        )
  const borderColor =
    custom !== undefined
      ? undefined
      : resolveTableStyleBorder(
          styleName !== undefined ? decodeXmlText(styleName) : undefined,
          context.theme,
        )
  return {
    area,
    headerRowCount: headerRowCount === undefined ? 1 : Number(headerRowCount),
    ...(totalsRowCount !== undefined && totalsRowCount !== '0'
      ? { totalsRowCount: Number(totalsRowCount) }
      : {}),
    showRowStripes,
    showColumnStripes,
    ...(name !== undefined && name !== '' ? { name: decodeXmlText(name) } : {}),
    columns,
    ...(styleName !== undefined ? { styleName: decodeXmlText(styleName) } : {}),
    ...(palette.headerFill !== undefined ? { headerFill: palette.headerFill } : {}),
    ...(palette.headerFontColor !== undefined ? { headerFontColor: palette.headerFontColor } : {}),
    ...(palette.stripeFill !== undefined ? { stripeFill: palette.stripeFill } : {}),
    ...(custom?.secondRowStripeFill !== undefined
      ? { secondRowStripeFill: custom.secondRowStripeFill }
      : {}),
    ...(custom?.columnStripeFill !== undefined
      ? { columnStripeFill: custom.columnStripeFill }
      : {}),
    ...(custom?.secondColumnStripeFill !== undefined
      ? { secondColumnStripeFill: custom.secondColumnStripeFill }
      : {}),
    ...(custom?.wholeTableFill !== undefined ? { wholeTableFill: custom.wholeTableFill } : {}),
    // First/last column emphasis (and the first-header corner cell) only
    // paint when tableStyleInfo turns them on.
    ...(custom !== undefined && showFirstColumn && custom.firstColumnFill !== undefined
      ? { firstColumnFill: custom.firstColumnFill }
      : {}),
    ...(custom !== undefined && showLastColumn && custom.lastColumnFill !== undefined
      ? { lastColumnFill: custom.lastColumnFill }
      : {}),
    ...(custom?.totalRowFill !== undefined ? { totalRowFill: custom.totalRowFill } : {}),
    ...(custom?.totalRowFontColor !== undefined
      ? { totalRowFontColor: custom.totalRowFontColor }
      : {}),
    ...(custom !== undefined && showFirstColumn && custom.firstHeaderCellFontColor !== undefined
      ? { firstHeaderCellFontColor: custom.firstHeaderCellFontColor }
      : {}),
    ...(borderColor !== undefined ? { borderColor } : {}),
  }
}

function isOn(value: string | undefined): boolean {
  return value === '1' || value === 'true'
}

function parseTableRef(ref: string): TableArea | undefined {
  const match = /^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/.exec(ref.replace(/\$/g, ''))
  if (match === null || match[1] === undefined || match[2] === undefined) return undefined
  const startColumn = labelToIndex(match[1].toUpperCase())
  const startRow = Number(match[2]) - 1
  const endColumn = match[3] !== undefined ? labelToIndex(match[3].toUpperCase()) : startColumn
  const endRow = match[4] !== undefined ? Number(match[4]) - 1 : startRow
  if (startRow < 0 || startColumn < 0 || endRow < startRow || endColumn < startColumn) {
    return undefined
  }
  return { startRow, startColumn, endRow, endColumn }
}

function labelToIndex(label: string): number {
  let index = 0
  for (const char of label) index = index * 26 + (char.charCodeAt(0) - 64)
  return index - 1
}

function resolveRelTarget(fromPart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const base = fromPart.split('/').slice(0, -1)
  for (const segment of target.split('/')) {
    if (segment === '..') base.pop()
    else if (segment !== '.' && segment !== '') base.push(segment)
  }
  return base.join('/')
}

function escapeRegExp(input: string): string {
  return input.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function decodeXmlText(input: string): string {
  return input
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, code: string) => String.fromCodePoint(Number(code)))
    .replace(/&amp;/g, '&')
}
