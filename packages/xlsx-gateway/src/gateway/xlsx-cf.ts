/// Declarative conditional-formatting save: the renderer snapshots the full
/// Univer rule set of a dirty sheet and this module rewrites every
/// `<conditionalFormatting>` section from it (mirroring the filter recipe).
/// Highlight styles intern as new dxf entries in the stylesheet.

export class CfEditError extends Error {}

export interface CfCellArea {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

/// One rule in the Univer conditional-formatting model shape (validated
/// structurally here — unknown shapes fail the save rather than guess).
export interface CfWireRule {
  readonly ranges: readonly CfCellArea[]
  readonly stopIfTrue: boolean
  readonly rule: Record<string, unknown>
}

export interface DxfSink {
  internDxf(dxfXml: string): number
}

/// OOXML-representable icon sets; Univer's extras (3Triangles, 3Stars,
/// 5Boxes, …) are x14-only and fail closed.
export const OOXML_ICON_SETS = new Set([
  '3Arrows',
  '3ArrowsGray',
  '3Flags',
  '3TrafficLights1',
  '3TrafficLights2',
  '3Signs',
  '3Symbols',
  '3Symbols2',
  '4Arrows',
  '4ArrowsGray',
  '4RedToBlack',
  '4Rating',
  '4TrafficLights',
  '5Arrows',
  '5ArrowsGray',
  '5Quarters',
  '5Rating',
])

/// Univer's iconMap lists most sets best-icon-first, but the rating sets run
/// worst-first, so their iconId sequence flips relative to the file order.
export const WORST_FIRST_ICON_SETS = new Set(['4Rating', '5Rating'])

/// Whether a Univer icon-set config round-trips to base OOXML: one
/// whitelisted set with icons in natural or fully reversed order.
export function iconSetSaveable(config: unknown): boolean {
  if (!Array.isArray(config) || config.length < 2) return false
  const entries = config as { iconType?: unknown; iconId?: unknown }[]
  const iconTypes = new Set(entries.map((entry) => String(entry?.iconType)))
  if (iconTypes.size !== 1 || !OOXML_ICON_SETS.has([...iconTypes][0] ?? '')) return false
  const ids = entries.map((entry) => String(entry?.iconId))
  return (
    ids.every((id, index) => id === String(index)) ||
    ids.every((id, index) => id === String(entries.length - 1 - index))
  )
}

/// Dry-runs the rule serializer so the UI can reject a rule the save would
/// fail closed on (same code path — zero drift). Returns the save-side error
/// message, or null when the rule is saveable.
export function cfRuleUnsaveableReason(rule: Record<string, unknown>): string | null {
  try {
    serializeCfRule(rule, 1, false, 'A1', { internDxf: () => 0 })
    return null
  } catch (error) {
    return error instanceof Error ? error.message : String(error)
  }
}

interface PreservedBlock {
  readonly sqref: string
  readonly text: string
  matched: boolean
}

export function applyCfRules(
  worksheetXml: string,
  rules: readonly CfWireRule[],
  dxfs: DxfSink,
): string {
  // Blocks whose cfRule carries an extLst are the base half of an x14
  // extension (linked via x14:id) — kept verbatim and guarded below. The
  // x14 part in the worksheet extLst is never rewritten.
  const preserved: PreservedBlock[] = []
  const xml = worksheetXml.replace(
    /<conditionalFormatting\b[^>]*>[\s\S]*?<\/conditionalFormatting>|<conditionalFormatting\b[^>]*\/>/g,
    (block) => {
      if (!/<extLst\b/.test(block)) return ''
      const sqref = /\bsqref="([^"]*)"/.exec(block)?.[1]
      if (sqref === undefined) throw new CfEditError(linkedMessage(block))
      preserved.push({ sqref, text: block, matched: false })
      return block
    },
  )

  // x14 rules and preserved blocks keep their priorities; new rules take the
  // free values.
  const used = new Set<number>()
  for (const match of xml.matchAll(/<(?:\w+:)?cfRule\b[^>]*?\spriority="(\d+)"/g)) {
    used.add(Number(match[1]))
  }
  let priority = 0
  const nextPriority = (): number => {
    do priority += 1
    while (used.has(priority))
    return priority
  }

  const sections: string[] = []
  for (const rule of rules) {
    const sqref = rule.ranges.map(toRef).join(' ')
    const linked = preserved.find((block) => !block.matched && block.sqref === sqref)
    if (linked) {
      // Only a byte-identical round trip proves the rule is unchanged; any
      // difference means the extension's base half was edited.
      const original = /<cfRule\b[^>]*?\spriority="(\d+)"/.exec(linked.text)
      const bare = linked.text.replace(/<extLst\b[\s\S]*?<\/extLst>/, '')
      let probe: string | null
      try {
        probe = original === null ? null : serializeRule(rule, Number(original[1]), dxfs)
      } catch {
        probe = null
      }
      if (probe !== bare) throw new CfEditError(linkedMessage(linked.text))
      linked.matched = true
      continue
    }
    sections.push(serializeRule(rule, nextPriority(), dxfs))
  }
  const removed = preserved.find((block) => !block.matched)
  if (removed) throw new CfEditError(linkedMessage(removed.text))

  if (sections.length === 0) return xml
  const body = sections.join('')
  const last = preserved[preserved.length - 1]
  if (last) {
    const end = xml.lastIndexOf(last.text) + last.text.length
    return xml.slice(0, end) + body + xml.slice(end)
  }
  const anchor =
    /<dataValidations\b|<hyperlinks\b|<printOptions\b|<pageMargins\b|<pageSetup\b|<headerFooter\b|<rowBreaks\b|<colBreaks\b|<drawing\b|<legacyDrawing\b|<picture\b|<oleObjects\b|<tableParts\b|<extLst\b/.exec(
      xml,
    )
  if (anchor) {
    return xml.slice(0, anchor.index) + body + xml.slice(anchor.index)
  }
  const end = xml.lastIndexOf('</worksheet>')
  if (end === -1) throw new CfEditError('Worksheet has no closing element.')
  return xml.slice(0, end) + body + xml.slice(end)
}

function linkedMessage(block: string): string {
  return block.includes('type="dataBar"')
    ? 'This range has a data-bar extension format (x14) that cannot be modified yet'
    : 'This range has extended conditional formatting (x14) that cannot be modified yet'
}

function serializeRule(wireRule: CfWireRule, priority: number, dxfs: DxfSink): string {
  if (wireRule.ranges.length === 0) {
    throw new CfEditError('A conditional-formatting rule has no ranges.')
  }
  const sqref = wireRule.ranges.map(toRef).join(' ')
  const anchor = `${columnToLetters(wireRule.ranges[0]?.startColumn ?? 0)}${(wireRule.ranges[0]?.startRow ?? 0) + 1}`
  const cfRule = serializeCfRule(wireRule.rule, priority, wireRule.stopIfTrue, anchor, dxfs)
  return `<conditionalFormatting sqref="${sqref}">${cfRule}</conditionalFormatting>`
}

function serializeCfRule(
  rule: Record<string, unknown>,
  priority: number,
  stopIfTrue: boolean,
  anchor: string,
  dxfs: DxfSink,
): string {
  const type = rule.type
  if (type === 'colorScale') return colorScaleRule(rule, priority, stopIfTrue)
  if (type === 'dataBar') return dataBarRule(rule, priority, stopIfTrue)
  if (type === 'iconSet') return iconSetRule(rule, priority, stopIfTrue)
  if (type === 'highlightCell') return highlightRule(rule, priority, stopIfTrue, anchor, dxfs)
  throw new CfEditError(`Unsupported conditional-formatting rule type "${String(type)}".`)
}

function attributes(base: string, priority: number, stopIfTrue: boolean, dxfId?: number): string {
  return (
    `type="${base}"${dxfId === undefined ? '' : ` dxfId="${dxfId}"`} priority="${priority}"` +
    (stopIfTrue ? ' stopIfTrue="1"' : '')
  )
}

function highlightRule(
  rule: Record<string, unknown>,
  priority: number,
  stopIfTrue: boolean,
  anchor: string,
  dxfs: DxfSink,
): string {
  const dxfId = dxfs.internDxf(buildDxfXml(rule.style))
  const subType = rule.subType
  const value = rule.value
  const operator = typeof rule.operator === 'string' ? rule.operator : undefined
  const element = (attrs: string, formulas: readonly string[] = []): string => {
    const body = formulas.map((formula) => `<formula>${escapeXmlText(formula)}</formula>`).join('')
    return body === '' ? `<cfRule ${attrs}/>` : `<cfRule ${attrs}>${body}</cfRule>`
  }
  switch (subType) {
    case 'number': {
      if (operator === undefined) throw new CfEditError('A number rule needs an operator.')
      const values = Array.isArray(value) ? value : [value]
      const formulas = values.map((entry) => {
        if (typeof entry !== 'number' || !Number.isFinite(entry)) {
          throw new CfEditError('A number rule needs finite values.')
        }
        return String(entry)
      })
      return element(
        `${attributes('cellIs', priority, stopIfTrue, dxfId)} operator="${operator}"`,
        formulas,
      )
    }
    case 'text': {
      const text = typeof value === 'string' ? value : ''
      const quoted = `"${text.replaceAll('"', '""')}"`
      switch (operator) {
        case 'containsText':
          return element(
            `${attributes('containsText', priority, stopIfTrue, dxfId)} operator="containsText" text="${escapeXmlAttribute(text)}"`,
            [`NOT(ISERROR(SEARCH(${quoted},${anchor})))`],
          )
        case 'notContainsText':
          return element(
            `${attributes('notContainsText', priority, stopIfTrue, dxfId)} operator="notContains" text="${escapeXmlAttribute(text)}"`,
            [`ISERROR(SEARCH(${quoted},${anchor}))`],
          )
        case 'beginsWith':
          return element(
            `${attributes('beginsWith', priority, stopIfTrue, dxfId)} operator="beginsWith" text="${escapeXmlAttribute(text)}"`,
            [`LEFT(${anchor},LEN(${quoted}))=${quoted}`],
          )
        case 'endsWith':
          return element(
            `${attributes('endsWith', priority, stopIfTrue, dxfId)} operator="endsWith" text="${escapeXmlAttribute(text)}"`,
            [`RIGHT(${anchor},LEN(${quoted}))=${quoted}`],
          )
        case 'equal':
          return element(`${attributes('cellIs', priority, stopIfTrue, dxfId)} operator="equal"`, [
            quoted,
          ])
        case 'notEqual':
          return element(
            `${attributes('cellIs', priority, stopIfTrue, dxfId)} operator="notEqual"`,
            [quoted],
          )
        case 'containsBlanks':
          return element(attributes('containsBlanks', priority, stopIfTrue, dxfId), [
            `LEN(TRIM(${anchor}))=0`,
          ])
        case 'notContainsBlanks':
          return element(attributes('notContainsBlanks', priority, stopIfTrue, dxfId), [
            `LEN(TRIM(${anchor}))>0`,
          ])
        case 'containsErrors':
          return element(attributes('containsErrors', priority, stopIfTrue, dxfId), [
            `ISERROR(${anchor})`,
          ])
        case 'notContainsErrors':
          return element(attributes('notContainsErrors', priority, stopIfTrue, dxfId), [
            `NOT(ISERROR(${anchor}))`,
          ])
        default:
          throw new CfEditError(`Unsupported text operator "${String(operator)}".`)
      }
    }
    case 'duplicateValues':
      return element(attributes('duplicateValues', priority, stopIfTrue, dxfId))
    case 'uniqueValues':
      return element(attributes('uniqueValues', priority, stopIfTrue, dxfId))
    case 'rank': {
      const rank = typeof value === 'number' && Number.isFinite(value) ? value : undefined
      if (rank === undefined) throw new CfEditError('A top/bottom rule needs a rank value.')
      return element(
        `${attributes('top10', priority, stopIfTrue, dxfId)}` +
          `${rule.isPercent === true ? ' percent="1"' : ''}` +
          `${rule.isBottom === true ? ' bottom="1"' : ''} rank="${rank}"`,
      )
    }
    case 'average': {
      // OOXML models above/below (with optional equality); exact equal /
      // notEqual to the average has no representation.
      const map: Record<string, string> = {
        greaterThan: '',
        greaterThanOrEqual: ' equalAverage="1"',
        lessThan: ' aboveAverage="0"',
        lessThanOrEqual: ' aboveAverage="0" equalAverage="1"',
      }
      const extra = operator === undefined ? undefined : map[operator]
      if (extra === undefined) {
        throw new CfEditError(`Average rule operator "${String(operator)}" cannot be saved.`)
      }
      return element(`${attributes('aboveAverage', priority, stopIfTrue, dxfId)}${extra}`)
    }
    case 'formula': {
      if (typeof value !== 'string' || value.length === 0) {
        throw new CfEditError('A formula rule needs a formula.')
      }
      return element(attributes('expression', priority, stopIfTrue, dxfId), [
        value.startsWith('=') ? value.slice(1) : value,
      ])
    }
    case 'timePeriod':
      throw new CfEditError('Date-occurring rules cannot be saved yet.')
    default:
      throw new CfEditError(`Unsupported highlight rule "${String(subType)}".`)
  }
}

function colorScaleRule(
  rule: Record<string, unknown>,
  priority: number,
  stopIfTrue: boolean,
): string {
  const config = rule.config
  if (!Array.isArray(config) || config.length < 2) {
    throw new CfEditError('A color scale needs at least two stops.')
  }
  const stops = [...config].sort(
    (left, right) => Number(left?.index ?? 0) - Number(right?.index ?? 0),
  )
  const cfvos = stops.map((stop) => serializeCfvo(stop?.value)).join('')
  const colors = stops
    .map(
      (stop) => `<color rgb="${toArgb(requireColor(stop?.color ?? '#FFFFFF', 'color scale'))}"/>`,
    )
    .join('')
  return (
    `<cfRule ${attributes('colorScale', priority, stopIfTrue)}>` +
    `<colorScale>${cfvos}${colors}</colorScale></cfRule>`
  )
}

function dataBarRule(rule: Record<string, unknown>, priority: number, stopIfTrue: boolean): string {
  const config = rule.config as Record<string, unknown> | undefined
  if (typeof config !== 'object' || config === null) {
    throw new CfEditError('A data bar rule has no configuration.')
  }
  const showValue = rule.isShowValue === false ? ' showValue="0"' : ''
  // The base schema stores one bar color; the negative color is x14-only and
  // is dropped.
  return (
    `<cfRule ${attributes('dataBar', priority, stopIfTrue)}>` +
    `<dataBar${showValue}>${serializeCfvo(config.min)}${serializeCfvo(config.max)}` +
    `<color rgb="${toArgb(requireColor(config.positiveColor ?? '#638EC6', 'data bar'))}"/></dataBar></cfRule>`
  )
}

function iconSetRule(rule: Record<string, unknown>, priority: number, stopIfTrue: boolean): string {
  const config = rule.config
  if (!Array.isArray(config) || config.length < 2) {
    throw new CfEditError('An icon set needs at least two thresholds.')
  }
  const iconTypes = new Set(config.map((entry) => String(entry?.iconType)))
  if (iconTypes.size !== 1) {
    throw new CfEditError('Mixed icon sets are extended-format-only and cannot be saved.')
  }
  const iconSet = [...iconTypes][0] ?? ''
  if (!OOXML_ICON_SETS.has(iconSet)) {
    throw new CfEditError(`The "${iconSet}" icon set cannot be saved to xlsx.`)
  }
  // Univer's config runs best-icon-first with descending thresholds (the
  // inverse of the file's ascending cfvo order — see the read-side install).
  const ascending = [...config].reverse()
  const iconIds = ascending.map((entry) => String(entry?.iconId))
  const upIds = ascending.map((_, index) => String(index))
  const downIds = ascending.map((_, index) => String(ascending.length - 1 - index))
  const [naturalIds, reversedIds] = WORST_FIRST_ICON_SETS.has(iconSet)
    ? [upIds, downIds]
    : [downIds, upIds]
  let reverse = ''
  if (iconIds.every((id, index) => id === reversedIds[index])) reverse = ' reverse="1"'
  else if (!iconIds.every((id, index) => id === naturalIds[index])) {
    throw new CfEditError('Custom icon orderings are extended-format-only and cannot be saved.')
  }
  const cfvos = ascending
    .map((entry, index) => {
      const gte = entry?.operator === 'greaterThan' ? ' gte="0"' : ''
      // The first threshold is the catch-all minimum.
      if (index === 0) return '<cfvo type="percent" val="0"/>'
      return serializeCfvo(entry?.value, gte)
    })
    .join('')
  const showValue = rule.isShowValue === false ? ' showValue="0"' : ''
  return (
    `<cfRule ${attributes('iconSet', priority, stopIfTrue)}>` +
    `<iconSet iconSet="${escapeXmlAttribute(iconSet)}"${showValue}${reverse}>${cfvos}</iconSet></cfRule>`
  )
}

function serializeCfvo(value: unknown, extra = ''): string {
  const config = value as { type?: unknown; value?: unknown } | undefined
  const type = String(config?.type ?? '')
  if (type === 'min' || type === 'max') return `<cfvo type="${type}"${extra}/>`
  if (!['num', 'percent', 'percentile', 'formula'].includes(type)) {
    throw new CfEditError(`Unsupported threshold type "${type}".`)
  }
  const raw = config?.value
  const val = type === 'formula' ? String(raw ?? '0') : String(Number(raw ?? 0))
  return `<cfvo type="${type}" val="${escapeXmlAttribute(val)}"${extra}/>`
}

/// Univer IStyleBase highlight style → dxf XML (font + solid fill).
export function buildDxfXml(style: unknown): string {
  const s = (typeof style === 'object' && style !== null ? style : {}) as Record<string, unknown>
  const fontParts: string[] = []
  if (s.bl === 1) fontParts.push('<b/>')
  if (s.it === 1) fontParts.push('<i/>')
  if (isLine(s.st)) fontParts.push('<strike/>')
  if (isLine(s.ul)) fontParts.push('<u/>')
  const fontColor = rgbOf(s.cl)
  if (fontColor) fontParts.push(`<color rgb="${toArgb(fontColor)}"/>`)
  const fillColor = rgbOf(s.bg)
  const font = fontParts.length === 0 ? '' : `<font>${fontParts.join('')}</font>`
  const fill =
    fillColor === undefined
      ? ''
      : `<fill><patternFill><bgColor rgb="${toArgb(fillColor)}"/></patternFill></fill>`
  return `<dxf>${font}${fill}</dxf>`
}

function isLine(value: unknown): boolean {
  return typeof value === 'object' && value !== null && (value as Record<string, unknown>).s === 1
}

function rgbOf(value: unknown): string | undefined {
  if (typeof value !== 'object' || value === null) return undefined
  const rgb = (value as Record<string, unknown>).rgb
  return typeof rgb === 'string' ? parseColor(rgb) : undefined
}

/// Univer stores colors as hex OR ColorKit's `rgb(r,g,b)` string form.
function parseColor(input: string): string | undefined {
  const value = input.trim()
  const hex = /^#?([0-9a-fA-F]{6})([0-9a-fA-F]{2})?$/.exec(value)
  if (hex?.[1]) return `#${hex[1].toUpperCase()}`
  const rgb = /^rgba?\(\s*(\d+)\s*[, ]\s*(\d+)\s*[, ]\s*(\d+)/.exec(value)
  if (!rgb) return undefined
  const channels = rgb
    .slice(1, 4)
    .map((channel) => Math.min(255, Number(channel)).toString(16).padStart(2, '0'))
    .join('')
  return `#${channels}`.toUpperCase()
}

function requireColor(input: unknown, what: string): string {
  const parsed = typeof input === 'string' ? parseColor(input) : undefined
  if (parsed === undefined) {
    throw new CfEditError(`The ${what} color "${String(input)}" cannot be saved.`)
  }
  return parsed
}

function toRef(range: CfCellArea): string {
  return range.startRow === range.endRow && range.startColumn === range.endColumn
    ? `${columnToLetters(range.startColumn)}${range.startRow + 1}`
    : `${columnToLetters(range.startColumn)}${range.startRow + 1}` +
        `:${columnToLetters(range.endColumn)}${range.endRow + 1}`
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

function toArgb(hexColor: string): string {
  return `FF${hexColor.replace('#', '').toUpperCase()}`
}

function escapeXmlText(input: string): string {
  return input.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

function escapeXmlAttribute(input: string): string {
  return input
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}

// ── Read side (EXCEL-024) ──────────────────────────────────────────────────
//
// parseConditionalFormatting is the exact inverse of applyCfRules: it turns
// a worksheet's `<conditionalFormatting>` sections into the same
// CfWireRule[] the writer consumes (the Univer conditional-formatting model
// shape, with the dxf style PRE-RESOLVED into rule.style). The browser
// receives ready-to-install rules and never parses style XML.
//
// Fail closed PER SHEET (the parseDataValidations recipe): x14 extensions,
// time periods, unknown rule types/operators, malformed sqref, unresolvable
// dxf styling, or guard-rail overruns throw CfReadError — the caller
// surfaces NO cfRules for the sheet, the workbook still opens, and a no-op
// save preserves the file's XML byte-for-byte.

export class CfReadError extends Error {}

const CF_MAX_RULES = 500
const CF_MAX_RANGES_PER_RULE = 100
/// Formula CF costs one dependency tree per covered cell; above this, huge
/// (e.g. whole-column) rules keep the cheaper native number condition
/// (desktop univer-sync.ts parity — CELLIS_FORMULA_CELL_LIMIT).
const CF_CELLIS_FORMULA_CELL_LIMIT = 20_000

const CF_CELLIS_OPERATORS = new Set([
  'equal',
  'notEqual',
  'greaterThan',
  'lessThan',
  'greaterThanOrEqual',
  'lessThanOrEqual',
  'between',
  'notBetween',
])

/// OOXML iconSet names → icon count (the whitelist mirrors OOXML_ICON_SETS).
const ICON_SET_SIZES: Readonly<Record<string, number>> = {
  '3Arrows': 3,
  '3ArrowsGray': 3,
  '3Flags': 3,
  '3TrafficLights1': 3,
  '3TrafficLights2': 3,
  '3Signs': 3,
  '3Symbols': 3,
  '3Symbols2': 3,
  '4Arrows': 4,
  '4ArrowsGray': 4,
  '4RedToBlack': 4,
  '4Rating': 4,
  '4TrafficLights': 4,
  '5Arrows': 5,
  '5ArrowsGray': 5,
  '5Quarters': 5,
  '5Rating': 5,
}

/**
 * Parse a worksheet's `<conditionalFormatting>` sections into the canonical
 * CfWireRule[] read model. Rules come back priority-ascending (the install
 * order the desktop uses — Univer applies rules in insertion order, and
 * lower xlsx priority = higher precedence).
 *
 * `dxfAt(dxfId)` resolves a dxf index to the raw `<dxf>` XML from
 * styles.xml (the StylesheetReader owns the dxfs list); only the
 * writer-round-trippable subset is accepted — anything else fails closed.
 */
export function parseConditionalFormatting(
  worksheetXml: string,
  dxfAt: (dxfId: number) => string | undefined,
): readonly CfWireRule[] {
  // x14 twins (pure extension rules in the worksheet extLst, or base halves
  // linked through a cfRule extLst) cannot be edited through the typed
  // model — the browser must not render a CF surface it cannot save.
  if (/<x14:conditionalFormatting\b/.test(worksheetXml)) {
    throw new CfReadError(
      'This sheet has extended (x14) conditional formatting — it cannot be represented yet.',
    )
  }
  if (!/<conditionalFormatting\b/.test(worksheetXml)) return []
  const collected: Array<{ priority: number; rule: CfWireRule }> = []
  // Self-closing alternatives come FIRST: a trailing `/>` form must never
  // fall through to the paired form and span a later element.
  const blockPattern =
    /<conditionalFormatting\b([^>]*?)\/>|<conditionalFormatting\b([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g
  let block: RegExpExecArray | null
  while ((block = blockPattern.exec(worksheetXml)) !== null) {
    const attributes = (block[1] ?? block[2] ?? '').trim()
    const inner = block[3] ?? ''
    const ranges = parseSqref(readXmlAttribute(attributes, 'sqref'))
    if (/<extLst\b/.test(inner)) {
      throw new CfReadError(
        'This range has extended conditional formatting (x14) that cannot be represented yet',
      )
    }
    const rulePattern = /<cfRule\b([^>]*?)\/>|<cfRule\b([^>]*)>([\s\S]*?)<\/cfRule>/g
    let ruleMatch: RegExpExecArray | null
    while ((ruleMatch = rulePattern.exec(inner)) !== null) {
      const ruleAttributes = (ruleMatch[1] ?? ruleMatch[2] ?? '').trim()
      const body = ruleMatch[3] ?? ''
      const priority = readPriority(ruleAttributes)
      collected.push({ priority, rule: parseCfRule(ruleAttributes, body, ranges, dxfAt) })
      if (collected.length > CF_MAX_RULES) {
        throw new CfReadError(
          `Worksheet carries more than ${CF_MAX_RULES} conditional-formatting rules.`,
        )
      }
    }
  }
  collected.sort((left, right) => left.priority - right.priority)
  return collected.map((entry) => entry.rule)
}

function parseCfRule(
  attributes: string,
  body: string,
  ranges: readonly CfCellArea[],
  dxfAt: (dxfId: number) => string | undefined,
): CfWireRule {
  const type = readXmlAttribute(attributes, 'type')
  const stopIfTrue = readXmlAttribute(attributes, 'stopIfTrue') === '1'
  if (type === 'colorScale') {
    return { ranges, stopIfTrue, rule: colorScaleConfig(body) }
  }
  if (type === 'dataBar') {
    return { ranges, stopIfTrue, rule: dataBarConfig(body) }
  }
  if (type === 'iconSet') {
    return { ranges, stopIfTrue, rule: iconSetConfig(body) }
  }
  // Every remaining representable family is a highlightCell rule carrying a
  // dxf-resolved style.
  const style = resolveRuleStyle(attributes, dxfAt)
  if (type === 'cellIs') return cellIsConfig(attributes, body, ranges, style)
  if (type === 'expression') {
    const formula = extractFormulas(body)[0]
    if (formula === undefined) {
      throw new CfReadError('An expression rule has no formula.')
    }
    return {
      ranges,
      stopIfTrue,
      rule: { type: 'highlightCell', subType: 'formula', value: `=${formula}`, style },
    }
  }
  const textOperators: Readonly<Record<string, string>> = {
    containsText: 'containsText',
    notContainsText: 'notContainsText',
    beginsWith: 'beginsWith',
    endsWith: 'endsWith',
    containsBlanks: 'containsBlanks',
    notContainsBlanks: 'notContainsBlanks',
    containsErrors: 'containsErrors',
    notContainsErrors: 'notContainsErrors',
  }
  if (type !== undefined && textOperators[type] !== undefined) {
    const text = readXmlAttribute(attributes, 'text') ?? ''
    return {
      ranges,
      stopIfTrue,
      rule: {
        type: 'highlightCell',
        subType: 'text',
        operator: textOperators[type]!,
        value: text,
        style,
      },
    }
  }
  if (type === 'duplicateValues' || type === 'uniqueValues') {
    return { ranges, stopIfTrue, rule: { type: 'highlightCell', subType: type, style } }
  }
  if (type === 'top10') {
    const rankText = readXmlAttribute(attributes, 'rank')
    const rank = rankText === undefined ? Number.NaN : Number(rankText)
    if (!Number.isInteger(rank) || rank < 1 || rank > 1000) {
      throw new CfReadError('A top-10 rule needs a rank between 1 and 1000.')
    }
    return {
      ranges,
      stopIfTrue,
      rule: {
        type: 'highlightCell',
        subType: 'rank',
        value: rank,
        isPercent: readXmlAttribute(attributes, 'percent') === '1',
        isBottom: readXmlAttribute(attributes, 'bottom') === '1',
        style,
      },
    }
  }
  if (type === 'aboveAverage') {
    const below = readXmlAttribute(attributes, 'aboveAverage') === '0'
    const equal = readXmlAttribute(attributes, 'equalAverage') === '1'
    const operator = below
      ? equal
        ? 'lessThanOrEqual'
        : 'lessThan'
      : equal
        ? 'greaterThanOrEqual'
        : 'greaterThan'
    return {
      ranges,
      stopIfTrue,
      rule: { type: 'highlightCell', subType: 'average', operator, style },
    }
  }
  if (type === 'timePeriod') {
    throw new CfReadError('Date-occurring rules cannot be represented yet.')
  }
  throw new CfReadError(`Unsupported conditional-formatting rule type "${String(type)}".`)
}

// ── highlightCell: cellIs ───────────────────────────────────────────────────
//
// Excel cellIs operands are formulas: a numeric literal, a quoted string, or
// a reference/expression. Numeric operands become Univer number rules (the
// writer's exact inverse); quoted strings become text equality rules (also
// an exact inverse); anything else becomes a formula rule anchored at the
// rule's top-left-sorted first range — desktop buildCellIsNonNumeric parity.
// Where Excel's blank-as-zero semantics would paint blanks differently from
// Univer's native number conditions, small ranges swap to the equivalent
// formula so the rendered outcome matches Excel (desktop parity).
function cellIsConfig(
  attributes: string,
  body: string,
  ranges: readonly CfCellArea[],
  style: Record<string, unknown>,
): CfWireRule {
  const operator = readXmlAttribute(attributes, 'operator')
  if (operator === undefined || !CF_CELLIS_OPERATORS.has(operator)) {
    throw new CfReadError(`Unsupported cellIs operator "${String(operator)}".`)
  }
  const stopIfTrue = readXmlAttribute(attributes, 'stopIfTrue') === '1'
  const formulas = extractFormulas(body)
  const first = formulas[0]
  if (first === undefined) {
    throw new CfReadError('A cellIs rule has no formula operand.')
  }
  const second = formulas[1]
  if ((operator === 'between' || operator === 'notBetween') && second === undefined) {
    throw new CfReadError('A between rule needs two formula operands.')
  }
  const firstNumber = Number(first)
  const secondNumber = Number(second ?? first)
  if (Number.isFinite(firstNumber) && (second === undefined || Number.isFinite(secondNumber))) {
    const coveredCells = ranges.reduce(
      (sum, range) =>
        sum + (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1),
      0,
    )
    if (
      coveredCells > 0 &&
      coveredCells <= CF_CELLIS_FORMULA_CELL_LIMIT &&
      cellIsBlankDiverges(operator, firstNumber, secondNumber)
    ) {
      return {
        ranges,
        stopIfTrue,
        rule: {
          type: 'highlightCell',
          subType: 'formula',
          value: cellIsFormula(operator, anchorOf(ranges), firstNumber, secondNumber),
          style,
        },
      }
    }
    return {
      ranges,
      stopIfTrue,
      rule: {
        type: 'highlightCell',
        subType: 'number',
        operator,
        ...(operator === 'between' || operator === 'notBetween'
          ? { value: [firstNumber, secondNumber] }
          : { value: firstNumber }),
        style,
      },
    }
  }
  // Quoted string operands: text equality round-trips exactly through the
  // writer (cellIs equal/notEqual + quoted formula).
  const quoted = /^"([\s\S]*)"$/.exec(first)
  if (quoted && (operator === 'equal' || operator === 'notEqual')) {
    return {
      ranges,
      stopIfTrue,
      rule: {
        type: 'highlightCell',
        subType: 'text',
        operator,
        value: quoted[1]!.replaceAll('""', '"'),
        style,
      },
    }
  }
  const anchor = anchorOf(ranges)
  const wrap = (operand: string): string => (/^"[\s\S]*"$/.test(operand) ? operand : `(${operand})`)
  const formula =
    operator === 'equal'
      ? `=${anchor}=(${first})`
      : operator === 'notEqual'
        ? `=${anchor}<>${wrap(first)}`
        : operator === 'greaterThan'
          ? `=${anchor}>${wrap(first)}`
          : operator === 'greaterThanOrEqual'
            ? `=${anchor}>=${wrap(first)}`
            : operator === 'lessThan'
              ? `=${anchor}<${wrap(first)}`
              : operator === 'lessThanOrEqual'
                ? `=${anchor}<=${wrap(first)}`
                : operator === 'between'
                  ? `=AND(${anchor}>=${wrap(first!)},${anchor}<=${wrap(second!)})`
                  : `=NOT(AND(${anchor}>=${wrap(first!)},${anchor}<=${wrap(second!)}))`
  return {
    ranges,
    stopIfTrue,
    rule: { type: 'highlightCell', subType: 'formula', value: formula, style },
  }
}

/// Univer offsets relative CF formulas from the top-left-sorted first range
/// (not the file's sqref order) — desktop parity.
function anchorOf(ranges: readonly CfCellArea[]): string {
  const first = [...ranges].sort(
    (left, right) => left.startRow - right.startRow || left.startColumn - right.startColumn,
  )[0]
  return first === undefined ? 'A1' : `${columnToLetters(first.startColumn)}${first.startRow + 1}`
}

/// True when Excel (blank = 0) and Univer's native conditions (blanks skip,
/// except notEqual/notBetween) paint blanks differently — desktop parity.
function cellIsBlankDiverges(operator: string, first: number, second: number): boolean {
  let excelBlank: boolean
  switch (operator) {
    case 'greaterThan':
      excelBlank = 0 > first
      break
    case 'greaterThanOrEqual':
      excelBlank = 0 >= first
      break
    case 'lessThan':
      excelBlank = 0 < first
      break
    case 'lessThanOrEqual':
      excelBlank = 0 <= first
      break
    case 'equal':
      excelBlank = first === 0
      break
    case 'notEqual':
      excelBlank = first !== 0
      break
    case 'between':
      excelBlank = Math.min(first, second) <= 0 && 0 <= Math.max(first, second)
      break
    case 'notBetween':
      excelBlank = !(Math.min(first, second) <= 0 && 0 <= Math.max(first, second))
      break
    default:
      return false
  }
  const univerBlank = operator === 'notEqual' || operator === 'notBetween'
  return excelBlank !== univerBlank
}

function cellIsFormula(operator: string, anchor: string, first: number, second: number): string {
  switch (operator) {
    case 'equal':
      return `=${anchor}=${first}`
    case 'notEqual':
      return `=${anchor}<>${first}`
    case 'greaterThan':
      return `=${anchor}>${first}`
    case 'greaterThanOrEqual':
      return `=${anchor}>=${first}`
    case 'lessThan':
      return `=${anchor}<${first}`
    case 'lessThanOrEqual':
      return `=${anchor}<=${first}`
    case 'between':
      return `=AND(${anchor}>=${Math.min(first, second)},${anchor}<=${Math.max(first, second)})`
    default:
      return `=NOT(AND(${anchor}>=${Math.min(first, second)},${anchor}<=${Math.max(first, second)}))`
  }
}

// ── Visual families ─────────────────────────────────────────────────────────

function colorScaleConfig(body: string): Record<string, unknown> {
  const colorScale = /<colorScale\b[^>]*>([\s\S]*?)<\/colorScale>/.exec(body)?.[1] ?? ''
  const cfvos = extractCfvos(colorScale)
  const colors = [...colorScale.matchAll(/<color\b([^>]*)\/?>/g)].map((match) =>
    readXmlAttribute(match[1] ?? '', 'rgb'),
  )
  if (cfvos.length < 2 || cfvos.length > 5) {
    throw new CfReadError('A color scale needs two to five stops.')
  }
  if (colors.length !== cfvos.length) {
    throw new CfReadError('A color scale needs one color per stop.')
  }
  const config = cfvos.map((cfvo, index) => {
    const color = toHexColor(colors[index])
    if (color === undefined) {
      throw new CfReadError('A color scale stop color cannot be represented.')
    }
    return { index, color, value: cfvoValue(cfvo) }
  })
  return { type: 'colorScale', config }
}

function dataBarConfig(body: string): Record<string, unknown> {
  const dataBar = /<dataBar\b([^>]*)>([\s\S]*?)<\/dataBar>/.exec(body)
  if (dataBar === null) {
    throw new CfReadError('A data bar rule has no dataBar element.')
  }
  const showValue = readXmlAttribute(dataBar[1] ?? '', 'showValue')
  const cfvos = extractCfvos(dataBar[2] ?? '')
  if (cfvos.length !== 2) {
    throw new CfReadError('A data bar needs exactly two thresholds.')
  }
  const color = toHexColor(
    readXmlAttribute(/<color\b([^>]*)\/?>/.exec(dataBar[2] ?? '')?.[1] ?? '', 'rgb'),
  )
  if (color === undefined) {
    throw new CfReadError('A data bar color cannot be represented.')
  }
  return {
    type: 'dataBar',
    isShowValue: showValue !== '0',
    config: {
      min: cfvoValue(cfvos[0]!),
      max: cfvoValue(cfvos[1]!),
      isGradient: true,
      positiveColor: color,
      nativeColor: '#FF0000',
    },
  }
}

function iconSetConfig(body: string): Record<string, unknown> {
  const iconSetElement = /<iconSet\b([^>]*)>([\s\S]*?)<\/iconSet>/.exec(body)
  if (iconSetElement === null) {
    throw new CfReadError('An icon set rule has no iconSet element.')
  }
  const attributes = iconSetElement[1] ?? ''
  const iconSet = readXmlAttribute(attributes, 'iconSet') ?? ''
  const reverse = readXmlAttribute(attributes, 'reverse') === '1'
  const showValue = readXmlAttribute(attributes, 'showValue')
  if (!OOXML_ICON_SETS.has(iconSet)) {
    throw new CfReadError(`The "${iconSet}" icon set cannot be represented.`)
  }
  const count = ICON_SET_SIZES[iconSet] ?? 0
  const cfvos = extractCfvos(iconSetElement[2] ?? '')
  if (cfvos.length !== count) {
    throw new CfReadError(`The "${iconSet}" icon set needs ${count} thresholds.`)
  }
  const worstFirst = WORST_FIRST_ICON_SETS.has(iconSet)
  const flipped = reverse !== worstFirst
  // File cfvos run ascending (worst → best); Univer's config runs
  // best-first with per-icon thresholds — the writer's exact inverse.
  const configs: Array<Record<string, unknown>> = []
  for (let position = 0; position < count; position += 1) {
    const ascendingIndex = count - 1 - position
    const cfvo = cfvos[ascendingIndex]!
    configs.push({
      iconType: iconSet,
      iconId: String(flipped ? count - 1 - position : position),
      operator: ascendingIndex > 0 && cfvo.gte === false ? 'greaterThan' : 'greaterThanOrEqual',
      value: ascendingIndex === 0 ? { type: 'min' } : cfvoValue(cfvo),
    })
  }
  return { type: 'iconSet', isShowValue: showValue !== '0', config: configs }
}

interface ParsedCfvo {
  readonly kind: string
  readonly value?: string
  readonly gte?: boolean
}

function extractCfvos(xml: string): ParsedCfvo[] {
  const cfvos: ParsedCfvo[] = []
  for (const match of xml.matchAll(/<cfvo\b([^>]*?)\/>|<cfvo\b([^>]*)>([\s\S]*?)<\/cfvo>/g)) {
    const attributes = (match[1] ?? match[2] ?? '').trim()
    const kind = readXmlAttribute(attributes, 'type')
    if (kind === undefined) {
      throw new CfReadError('A threshold has no type attribute.')
    }
    if (!['min', 'max', 'num', 'percent', 'percentile', 'formula'].includes(kind)) {
      throw new CfReadError(`Unsupported threshold type "${kind}".`)
    }
    // The value normally rides the val attribute; a few producers put it
    // in the element body instead (seen with formula thresholds).
    const text = (match[3] ?? '').trim()
    const value = readXmlAttribute(attributes, 'val') ?? (text === '' ? undefined : text)
    cfvos.push({
      kind,
      value,
      gte: readXmlAttribute(attributes, 'gte') !== '0',
    })
  }
  return cfvos
}

function cfvoValue(cfvo: ParsedCfvo): Record<string, unknown> {
  switch (cfvo.kind) {
    case 'min':
    case 'max':
      return { type: cfvo.kind }
    case 'percent':
    case 'percentile': {
      const value = Number(cfvo.value ?? 0)
      if (!Number.isFinite(value)) {
        throw new CfReadError('A percent threshold needs a numeric value.')
      }
      return { type: cfvo.kind, value }
    }
    case 'formula':
      return { type: 'formula', value: `=${cfvo.value ?? '0'}` }
    default: {
      const numeric = Number(cfvo.value ?? 0)
      return Number.isFinite(numeric)
        ? { type: 'num', value: numeric }
        : { type: 'formula', value: `=${cfvo.value ?? '0'}` }
    }
  }
}

// ── dxf style resolution ────────────────────────────────────────────────────
//
// The accepted dxf subset is the exact inverse of buildDxfXml: font marks
// (b/i/strike/u), an rgb font color, a solid bgColor fill, and a numFmt
// pattern. Anything else (borders, theme/indexed colors, font sizes …)
// would not survive a CF-dirty rewrite, so the sheet fails closed instead
// of silently dropping styling.

function resolveRuleStyle(
  attributes: string,
  dxfAt: (dxfId: number) => string | undefined,
): Record<string, unknown> {
  const dxfIdText = readXmlAttribute(attributes, 'dxfId')
  if (dxfIdText === undefined) return {}
  const dxfId = Number(dxfIdText)
  if (!Number.isInteger(dxfId) || dxfId < 0) {
    throw new CfReadError(`A rule references malformed dxfId "${dxfIdText}".`)
  }
  const dxf = dxfAt(dxfId)
  if (dxf === undefined) {
    throw new CfReadError(`A rule references dxfId ${dxfId}, which styles.xml does not define.`)
  }
  return parseDxfStyle(dxf)
}

function parseDxfStyle(dxfXml: string): Record<string, unknown> {
  const inner = /<dxf\b([^>]*?)\/>|<dxf\b([^>]*)>([\s\S]*?)<\/dxf>/.exec(dxfXml)
  if (inner === null) {
    throw new CfReadError('A dxf style entry cannot be read.')
  }
  const body = inner[3] ?? ''
  const style: Record<string, unknown> = {}
  const font = /<font\b([^>]*?)\/>|<font\b([^>]*)>([\s\S]*?)<\/font>/.exec(body)
  if (font !== null) {
    const fontBody = font[3] ?? ''
    for (const element of extractTopLevelElements(fontBody)) {
      if (/^<b\b/.test(element)) {
        style.bl = 1
      } else if (/^<i\b/.test(element)) {
        style.it = 1
      } else if (/^<strike\b/.test(element)) {
        style.st = { s: 1 }
      } else if (/^<u\b/.test(element)) {
        const val = readXmlAttribute(/^<u\b([^>]*)/.exec(element)?.[1] ?? '', 'val')
        if (val !== undefined && val !== 'single') {
          throw new CfReadError(`A dxf underline style "${val}" cannot be represented.`)
        }
        style.ul = { s: 1 }
      } else if (/^<color\b/.test(element)) {
        const color = colorFromElement(element, 'font color')
        if (color === undefined) {
          throw new CfReadError('A dxf font color cannot be represented.')
        }
        style.cl = { rgb: color }
      } else {
        throw new CfReadError('A dxf font carries properties the model cannot represent.')
      }
    }
  }
  const fill = /<fill\b([^>]*?)\/>|<fill\b([^>]*)>([\s\S]*?)<\/fill>/.exec(body)
  if (fill !== null) {
    const fillBody = fill[3] ?? ''
    const pattern =
      /<patternFill\b([^>]*?)\/>|<patternFill\b([^>]*)>([\s\S]*?)<\/patternFill>/.exec(fillBody)
    if (pattern === null) {
      throw new CfReadError('A dxf fill cannot be represented.')
    }
    const patternAttributes = pattern[1] ?? pattern[2] ?? ''
    const patternBody = pattern[3] ?? ''
    const patternType = readXmlAttribute(patternAttributes, 'patternType')
    const bgColor = /<bgColor\b([^>]*)\/?>/.exec(patternBody)
    const fgColor = /<fgColor\b([^>]*)\/?>/.exec(patternBody)
    if (fgColor !== null && bgColor === null) {
      throw new CfReadError('A dxf foreground-only fill cannot be represented.')
    }
    if (bgColor !== null) {
      if (patternType !== undefined && patternType !== 'solid') {
        throw new CfReadError(`A dxf fill pattern "${patternType}" cannot be represented.`)
      }
      const color = colorFromElement(bgColor[0] ?? '', 'fill color')
      if (color === undefined) {
        throw new CfReadError('A dxf fill color cannot be represented.')
      }
      style.bg = { rgb: color }
    }
  }
  const numFmt = /<numFmt\b([^>]*)\/?>/.exec(body)
  if (numFmt !== null) {
    const pattern = readXmlAttribute(numFmt[1] ?? '', 'formatCode')
    if (pattern === undefined || pattern === '') {
      throw new CfReadError('A dxf number format cannot be represented.')
    }
    style.n = { pattern }
  }
  // Any other top-level dxf child (border, alignment, protection …) would
  // be dropped by the writer's dxf serializer — refuse instead.
  for (const element of extractTopLevelElements(body)) {
    if (!/^<(font|fill|numFmt)\b/.test(element)) {
      throw new CfReadError('A dxf style carries properties the model cannot represent.')
    }
  }
  return style
}

function extractTopLevelElements(xml: string): string[] {
  const elements: string[] = []
  const pattern = /<([A-Za-z][\w.-]*)((?:[^>"']|"[^"]*"|'[^']*')*?)(?:\/>|>([\s\S]*?)<\/\1\s*>)/g
  let match: RegExpExecArray | null
  while ((match = pattern.exec(xml)) !== null) {
    elements.push(match[0])
  }
  return elements
}

function colorFromElement(element: string, what: string): string | undefined {
  const rgb = readXmlAttribute(element, 'rgb')
  if (rgb !== undefined) return toHexColor(rgb)
  const theme = readXmlAttribute(element, 'theme')
  if (theme !== undefined || readXmlAttribute(element, 'indexed') !== undefined) {
    throw new CfReadError(`A dxf ${what} uses theme/indexed colors, which cannot be represented.`)
  }
  return undefined
}

/// "FFRRGGBB" / "RRGGBB" / "#RRGGBB" → "#RRGGBB" (the Univer color form).
function toHexColor(rgb: string | undefined): string | undefined {
  if (rgb === undefined) return undefined
  const match = /^#?([0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/.exec(rgb.trim())
  if (match === null) return undefined
  const hex = match[1]!.toUpperCase()
  return `#${hex.length === 8 ? hex.slice(2) : hex}`
}

// ── shared read helpers (the xlsx-dv.ts recipe) ─────────────────────────────

function readXmlAttribute(attributes: string, name: string): string | undefined {
  const match = new RegExp(`(?:^|\\s)${name}="([^"]*)"`).exec(attributes)
  return match === null ? undefined : decodeXmlText(match[1] ?? '')
}

function readPriority(attributes: string): number {
  const priority = readXmlAttribute(attributes, 'priority')
  const value = priority === undefined ? Number.NaN : Number(priority)
  if (!Number.isInteger(value) || value < 1 || value > 1_000_000) {
    throw new CfReadError('A conditional-formatting rule has no readable priority.')
  }
  return value
}

function parseSqref(sqref: string | undefined): readonly CfCellArea[] {
  if (sqref === undefined || sqref.trim() === '') {
    throw new CfReadError('conditionalFormatting has no sqref attribute.')
  }
  const ranges: CfCellArea[] = []
  for (const ref of sqref.split(/\s+/)) {
    if (ref === '') continue
    const area = parseSqrefPart(ref)
    if (area === null) {
      throw new CfReadError(`conditionalFormatting sqref "${ref}" is not a readable range.`)
    }
    ranges.push(area)
  }
  if (ranges.length === 0 || ranges.length > CF_MAX_RANGES_PER_RULE) {
    throw new CfReadError(
      `conditionalFormatting sqref must carry 1..${CF_MAX_RANGES_PER_RULE} ranges.`,
    )
  }
  return ranges
}

function parseSqrefPart(ref: string): CfCellArea | null {
  const [startRef, endRef] = ref.split(':')
  const start = parseA1(startRef ?? '')
  if (!start) return null
  const end = endRef === undefined ? start : parseA1(endRef)
  if (!end) return null
  const area = {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startColumn: Math.min(start.column, end.column),
    endColumn: Math.max(start.column, end.column),
  }
  if (
    area.startRow < 0 ||
    area.startColumn < 0 ||
    area.endRow > 1_048_575 ||
    area.endColumn > 16_383
  ) {
    return null
  }
  return area
}

function parseA1(address: string): { row: number; column: number } | null {
  const match = /^\$?([A-Z]{1,3})\$?([0-9]+)$/.exec(address)
  if (!match) return null
  return { column: lettersToColumn(match[1]!), row: Number(match[2]!) - 1 }
}

function lettersToColumn(letters: string): number {
  let column = 0
  for (const letter of letters) {
    column = column * 26 + (letter.charCodeAt(0) - 64)
  }
  return column - 1
}

function extractFormulas(body: string): string[] {
  const formulas: string[] = []
  for (const match of body.matchAll(/<formula>([\s\S]*?)<\/formula>/g)) {
    const text = decodeXmlText(match[1] ?? '')
    if (text !== '') formulas.push(text)
  }
  return formulas
}

function decodeXmlText(input: string): string {
  return input
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&#10;', '\n')
    .replaceAll('&amp;', '&')
}
