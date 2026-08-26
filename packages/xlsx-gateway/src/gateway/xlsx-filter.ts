/// Writes a declarative filter snapshot into worksheet XML: the
/// `<autoFilter>` element (with per-column value / custom criteria) plus row
/// visibility inside the filter's row span. Unsupported criteria fail closed.

export class FilterEditError extends Error {}

/// Reading `<autoFilter>` back failed closed: the element carries criteria
/// the canonical model cannot represent (top10, dynamicFilter, iconFilter,
/// dateGroup, colorFilters, …). The filter is NOT rendered by the browser and
/// a no-op save preserves the file's XML byte-for-byte.
export class FilterReadError extends Error {}

interface CellArea {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

export interface FilterColumnState {
  readonly colId: number
  readonly values?: readonly string[] | undefined
  readonly blank?: boolean | undefined
  readonly customs?:
    | {
        readonly and?: boolean | undefined
        readonly filters: readonly {
          readonly val: string | number
          readonly operator?: string | undefined
        }[]
      }
    | undefined
}

export interface SheetFilterState {
  readonly sheetName: string
  readonly filter: {
    readonly range: CellArea
    readonly columns: readonly FilterColumnState[]
  } | null
  readonly hiddenRows: readonly number[]
  readonly visibilityRange: CellArea
}

const CUSTOM_OPERATORS = new Set([
  'equal',
  'notEqual',
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
])

/// Child elements of <filterColumn> the canonical model cannot represent.
/// Encountering any of them fails the read closed (no filterState surfaced —
/// the browser never renders an unfaithful filter).
const UNSUPPORTED_FILTER_COLUMN_CHILDREN = [
  'top10',
  'dynamicFilter',
  'iconFilter',
  'colorFilters',
  'dateGroup',
  'dateFilters',
  'customFilters10',
]

export function applyFilterState(worksheetXml: string, state: SheetFilterState): string {
  const element = state.filter === null ? '' : serializeAutoFilter(state.filter)
  const existing = /<autoFilter\b[^>]*\/>|<autoFilter\b[^>]*>[\s\S]*?<\/autoFilter>/.exec(
    worksheetXml,
  )
  let xml = worksheetXml
  if (existing) {
    xml =
      worksheetXml.slice(0, existing.index) +
      element +
      worksheetXml.slice(existing.index + existing[0].length)
  } else if (element !== '') {
    xml = insertAfterSheetData(worksheetXml, element)
  }
  return applyRowVisibility(xml, state.visibilityRange, new Set(state.hiddenRows))
}

/// Parse a worksheet's <autoFilter> element into the canonical
/// SheetFilterState read model. The state carries:
///   - filter: the range + per-column criteria (values / blank / customs)
///   - hiddenRows: 0-based rows inside the filter's data span carrying
///     hidden="1" — the file's stored row visibility
///   - visibilityRange: the filter range (header + data rows)
///
/// Fail-closed: any filterColumn child the canonical model cannot serialize
/// (top10, dynamicFilter, iconFilter, colorFilters, dateGroup, …) or any
/// customFilter operator outside the supported set throws FilterReadError —
/// the caller surfaces NO filterState, so the browser never renders a filter
/// it cannot save faithfully, and a no-op save leaves the XML untouched.
export function parseAutoFilter(worksheetXml: string, sheetName: string): SheetFilterState | null {
  const element = /<autoFilter\b([^>]*)\/>|<autoFilter\b([^>]*)>[\s\S]*?<\/autoFilter>/.exec(
    worksheetXml,
  )
  if (!element) {
    // No autoFilter element: no filter state. The caller omits the field
    // entirely — "filter: null" would mean "cleared", which implies
    // unhiding rows and removing XML that does not exist.
    return null
  }
  const attributes = element[1] ?? element[2] ?? ''
  const ref = /(?:^|\s)ref="([^"]+)"/.exec(attributes)?.[1]
  const range = ref === undefined ? null : parseRefRange(ref)
  if (range === null) {
    throw new FilterReadError(`autoFilter has no readable ref (${ref ?? 'none'}).`)
  }
  const inner =
    element[2] === undefined
      ? ''
      : element[0].slice(element[0].indexOf('>') + 1, element[0].length - '</autoFilter>'.length)
  const columns: FilterColumnState[] = []
  for (const columnMatch of inner.matchAll(/<filterColumn\b([^>]*)>([\s\S]*?)<\/filterColumn>/g)) {
    columns.push(parseFilterColumn(columnMatch[1] ?? '', columnMatch[2] ?? ''))
  }
  // A self-closing <filterColumn/> or one with only unsupported children
  // must not silently vanish — its criteria would be dropped on the next
  // filter save. Fail closed instead.
  if (/<filterColumn\b/g.test(inner) && columns.length !== countFilterColumns(inner)) {
    throw new FilterReadError('autoFilter carries a filterColumn without readable criteria.')
  }
  const hiddenRows = parseHiddenRows(worksheetXml, range)
  return {
    sheetName,
    filter: { range, columns },
    hiddenRows,
    visibilityRange: range,
  }
}

function countFilterColumns(inner: string): number {
  const matches = [...inner.matchAll(/<filterColumn\b[^>]*(?:\/>|>)/g)]
  // Self-closing <filterColumn/> entries count as columns too.
  return matches.length
}

function parseFilterColumn(attributes: string, inner: string): FilterColumnState {
  const colIdText = /(?:^|\s)colId="([0-9]+)"/.exec(attributes)?.[1]
  if (colIdText === undefined) {
    throw new FilterReadError('filterColumn has no colId attribute.')
  }
  const colId = Number(colIdText)
  if (inner.trim() === '') throw new FilterReadError('filterColumn carries no criteria.')
  for (const unsupported of UNSUPPORTED_FILTER_COLUMN_CHILDREN) {
    if (inner.includes(`<${unsupported}`)) {
      throw new FilterReadError(
        `Filter criteria "${unsupported}" cannot be represented in the canonical filter model.`,
      )
    }
  }
  const valuesMatch = /<filters\b([^>]*)>([\s\S]*?)<\/filters>|<filters\b([^>]*)\/>/.exec(inner)
  const customMatch =
    /<customFilters\b([^>]*)>([\s\S]*?)<\/customFilters>|<customFilters\b([^>]*)\/>/.exec(inner)
  if (!valuesMatch && !customMatch) {
    throw new FilterReadError('filterColumn carries no readable criteria.')
  }
  const column: {
    colId: number
    values?: readonly string[]
    blank?: boolean
    customs?: {
      and?: boolean
      filters: { val: string | number; operator?: string }[]
    }
  } = { colId }
  if (valuesMatch) {
    const filterAttrs = valuesMatch[1] ?? valuesMatch[3] ?? ''
    const valuesInner = valuesMatch[2] ?? ''
    const values = [...valuesInner.matchAll(/<filter\b[^>]*?\/>/g)].map(
      (m) => /(?:^|\s)val="([^"]*)"/.exec(m[0])?.[1] ?? '',
    )
    const blank = /(?:^|\s)blank="([^"]*)"/.exec(filterAttrs)?.[1] === '1'
    if (values.length > 0) column.values = values
    if (blank) column.blank = true
    if (values.length === 0 && !blank) {
      throw new FilterReadError('filters element carries no values and no blank flag.')
    }
  }
  if (customMatch) {
    const customAttrs = customMatch[1] ?? customMatch[3] ?? ''
    const customsInner = customMatch[2] ?? ''
    const filters: { val: string | number; operator?: string }[] = []
    for (const custom of customsInner.matchAll(
      /<customFilter\b([^>]*?)(?:\/|><\/customFilter)>/g,
    )) {
      const attrs = custom[1] ?? ''
      const valText = /(?:^|\s)val="([^"]*)"/.exec(attrs)?.[1]
      if (valText === undefined) {
        throw new FilterReadError('customFilter has no val attribute.')
      }
      const operator = /(?:^|\s)operator="([^"]+)"/.exec(attrs)?.[1]
      if (operator !== undefined && !CUSTOM_OPERATORS.has(operator)) {
        throw new FilterReadError(`Filter condition "${operator}" cannot be saved as XLSX yet.`)
      }
      // Numeric when the text parses as a finite number (matching how Univer
      // round-trips customFilter vals); otherwise keep the string.
      const asNumber = Number(valText)
      const val: string | number =
        valText.trim() !== '' && Number.isFinite(asNumber) ? asNumber : decodeXmlText(valText)
      filters.push(operator === undefined ? { val } : { val, operator })
    }
    if (filters.length === 0) {
      throw new FilterReadError('customFilters element carries no customFilter entries.')
    }
    const and = /(?:^|\s)and="([^"]*)"/.exec(customAttrs)?.[1] === '1'
    column.customs = and ? { and: true, filters } : { filters }
  }
  return column
}

/// Rows carrying hidden="1" inside the filter's data span (0-based, header
/// excluded). Rows hidden outside the span belong to the user, not the
/// filter — they are NOT part of the filter state.
function parseHiddenRows(worksheetXml: string, range: CellArea): number[] {
  const hidden: number[] = []
  const rowPattern = /<row\b([^>]*)\/?>/g
  let match: RegExpExecArray | null
  while ((match = rowPattern.exec(worksheetXml)) !== null) {
    const attrs = match[1] ?? ''
    if (/(?:^|\s)hidden="(?:1|true)"/.test(attrs) === false) continue
    const rowNumber = /(?:^|\s)r="([0-9]+)"/.exec(attrs)?.[1]
    if (rowNumber === undefined) continue
    const rowIndex = Number(rowNumber) - 1
    if (rowIndex > range.startRow && rowIndex <= range.endRow) hidden.push(rowIndex)
  }
  return hidden
}

function parseRefRange(ref: string): CellArea | null {
  const [startRef, endRef] = ref.split(':')
  const start = parseA1(startRef ?? '')
  const end = parseA1(endRef ?? startRef ?? '')
  if (!start || !end) return null
  return {
    startRow: start.row,
    endRow: end.row,
    startColumn: start.column,
    endColumn: end.column,
  }
}

function parseA1(address: string): { row: number; column: number } | null {
  const match = /^([A-Z]{1,3})([0-9]+)$/.exec(address)
  if (!match) return null
  return { column: lettersToColumn(match[1]!), row: Number(match[2]!) - 1 }
}

function decodeXmlText(input: string): string {
  return input
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&amp;', '&')
}

function serializeAutoFilter(filter: NonNullable<SheetFilterState['filter']>): string {
  const ref = toRef(filter.range)
  const columns = [...filter.columns]
    .sort((left, right) => left.colId - right.colId)
    .map(serializeFilterColumn)
    .join('')
  return columns === ''
    ? `<autoFilter ref="${ref}"/>`
    : `<autoFilter ref="${ref}">${columns}</autoFilter>`
}

function serializeFilterColumn(column: FilterColumnState): string {
  const parts: string[] = []
  if (column.values !== undefined || column.blank) {
    const blank = column.blank ? ' blank="1"' : ''
    const values = (column.values ?? [])
      .map((value) => `<filter val="${escapeXmlAttribute(value)}"/>`)
      .join('')
    parts.push(`<filters${blank}>${values}</filters>`)
  }
  if (column.customs) {
    for (const custom of column.customs.filters) {
      if (custom.operator !== undefined && !CUSTOM_OPERATORS.has(custom.operator)) {
        throw new FilterEditError(
          `Filter condition "${custom.operator}" cannot be saved as XLSX yet.`,
        )
      }
    }
    const and = column.customs.and ? ' and="1"' : ''
    const filters = column.customs.filters
      .map((custom) => {
        const operator =
          custom.operator === undefined || custom.operator === 'equal'
            ? ''
            : ` operator="${custom.operator}"`
        return `<customFilter${operator} val="${escapeXmlAttribute(String(custom.val))}"/>`
      })
      .join('')
    parts.push(`<customFilters${and}>${filters}</customFilters>`)
  }
  if (parts.length === 0) return ''
  return `<filterColumn colId="${column.colId}">${parts.join('')}</filterColumn>`
}

/// Schema order places autoFilter after sheetData (and after the protection
/// block when present).
function insertAfterSheetData(worksheetXml: string, element: string): string {
  let insertAt = -1
  for (const pattern of [
    /<sheetData\s*\/>|<\/sheetData>/,
    /<sheetProtection\b[^>]*\/>/,
    /<protectedRanges\b[^>]*\/>|<\/protectedRanges>/,
    /<\/scenarios>/,
  ]) {
    const match = pattern.exec(worksheetXml)
    if (match) insertAt = Math.max(insertAt, match.index + match[0].length)
  }
  if (insertAt < 0) throw new FilterEditError('Worksheet has no sheetData element.')
  return worksheetXml.slice(0, insertAt) + element + worksheetXml.slice(insertAt)
}

/// Declarative row visibility inside the filter's data rows: listed rows are
/// hidden, every other row in the span is unhidden (matching how Excel
/// re-evaluates a filter). The header row is never touched.
function applyRowVisibility(
  worksheetXml: string,
  range: CellArea,
  hiddenRows: ReadonlySet<number>,
): string {
  const firstDataRow = range.startRow + 1
  const seen = new Set<number>()
  let xml = worksheetXml.replace(
    /<row\b([^>]*?)(\/>|>)/g,
    (full, attributes: string, close: string) => {
      const rowNumber = /(?:^|\s)r="([0-9]+)"/.exec(attributes)?.[1]
      if (rowNumber === undefined) return full
      const rowIndex = Number(rowNumber) - 1
      if (rowIndex < firstDataRow || rowIndex > range.endRow) return full
      seen.add(rowIndex)
      const withoutHidden = attributes.replace(/\s*hidden="[^"]*"/, '')
      const hidden = hiddenRows.has(rowIndex) ? ' hidden="1"' : ''
      return `<row${withoutHidden}${hidden}${close}`
    },
  )
  const missing = [...hiddenRows]
    .filter(
      (rowIndex) => !seen.has(rowIndex) && rowIndex >= firstDataRow && rowIndex <= range.endRow,
    )
    .sort((left, right) => left - right)
  for (const rowIndex of missing) {
    xml = insertEmptyHiddenRow(xml, rowIndex + 1)
  }
  return xml
}

function insertEmptyHiddenRow(worksheetXml: string, rowNumber: number): string {
  const newRow = `<row r="${rowNumber}" hidden="1"/>`
  const rowStartPattern = /<row\b[^>]*?\br="([1-9][0-9]*)"/g
  let match: RegExpExecArray | null
  while ((match = rowStartPattern.exec(worksheetXml)) !== null) {
    if (Number(match[1]) > rowNumber) {
      return worksheetXml.slice(0, match.index) + newRow + worksheetXml.slice(match.index)
    }
  }
  if (worksheetXml.includes('</sheetData>')) {
    return worksheetXml.replace('</sheetData>', () => `${newRow}</sheetData>`)
  }
  const emptySheetData = /<sheetData\s*\/>/
  if (emptySheetData.test(worksheetXml)) {
    return worksheetXml.replace(emptySheetData, () => `<sheetData>${newRow}</sheetData>`)
  }
  throw new FilterEditError('Worksheet has no sheetData element.')
}

function toRef(range: CellArea): string {
  return (
    `${columnToLetters(range.startColumn)}${range.startRow + 1}` +
    `:${columnToLetters(range.endColumn)}${range.endRow + 1}`
  )
}

function columnToLetters(column: number): string {
  let letters = ''
  let remaining = column + 1
  while (remaining > 0) {
    remaining -= 1
    letters = String.fromCharCode(65 + (remaining % 26)) + letters
    remaining = Math.floor(remaining / 26)
  }
  return letters
}

function lettersToColumn(letters: string): number {
  let column = 0
  for (const character of letters) {
    column = column * 26 + character.charCodeAt(0) - 64
  }
  return column - 1
}

function escapeXmlAttribute(input: string): string {
  return input
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;')
}
