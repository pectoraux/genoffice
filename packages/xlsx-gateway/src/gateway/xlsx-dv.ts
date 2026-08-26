/// Declarative data-validation save: the renderer snapshots the full Univer
/// rule set of a dirty sheet and this module rewrites the worksheet's
/// `<dataValidations>` section from it (mirroring the CF/filter recipe).
/// Mappings are the exact inverse of the read-side install in App.tsx.

export class DvEditError extends Error {}

/// Reading `<dataValidations>` back failed closed: the section carries
/// constructs the canonical model cannot represent (x14 extensions, unknown
/// types/operators/error styles, malformed sqref). The sheet's DV state is
/// NOT surfaced — the browser never renders an unfaithful rule, and a no-op
/// save preserves the file's XML byte-for-byte.
export class DvReadError extends Error {}

export interface DvCellArea {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

/// One rule in the Univer data-validation model shape (validated structurally
/// here — unknown shapes fail the save rather than guess).
export interface DvWireRule {
  readonly ranges: readonly DvCellArea[]
  readonly rule: Record<string, unknown>
}

const DV_TYPES = new Set(['whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom'])
const DV_OPERATORS = new Set([
  'between',
  'notBetween',
  'equal',
  'notEqual',
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
])
/// Univer DataValidationErrorStyle: INFO=0, STOP=1 (OOXML default), WARNING=2.
const DV_ERROR_STYLE_NAMES: Record<number, string | undefined> = {
  0: 'information',
  1: undefined,
  2: 'warning',
}
/// Read-side inverse: OOXML errorStyle name → Univer number. `stop` (the
/// OOXML default, emitted attribute-less) is absent — Univer's STOP is also
/// the enum default, so an absent attribute maps to an absent field.
const DV_ERROR_STYLE_NUMBERS: Record<string, number | undefined> = {
  information: 0,
  warning: 2,
}

/// Guard rails mirroring the desktop's shared wire schema
/// (workbookDvStateSchema): bounded rule/range counts.
const DV_MAX_RULES = 500
const DV_MAX_RANGES_PER_RULE = 100

export function applyDvRules(worksheetXml: string, rules: readonly DvWireRule[]): string {
  if (/<x14:dataValidation\b/.test(worksheetXml)) {
    throw new DvEditError(
      'This sheet has extended (x14) data validation — editing its rules is not ' +
        'supported yet.',
    )
  }
  const xml = worksheetXml.replace(
    /<dataValidations\b[^>]*>[\s\S]*?<\/dataValidations>|<dataValidations\b[^>]*\/>/g,
    '',
  )
  if (rules.length === 0) return xml

  const body = rules.map(serializeRule).join('')
  const section = `<dataValidations count="${rules.length}">${body}</dataValidations>`
  const anchor =
    /<hyperlinks\b|<printOptions\b|<pageMargins\b|<pageSetup\b|<headerFooter\b|<rowBreaks\b|<colBreaks\b|<drawing\b|<legacyDrawing\b|<picture\b|<oleObjects\b|<tableParts\b|<extLst\b/.exec(
      xml,
    )
  if (anchor) {
    return xml.slice(0, anchor.index) + section + xml.slice(anchor.index)
  }
  const end = xml.lastIndexOf('</worksheet>')
  if (end === -1) throw new DvEditError('Worksheet has no closing element.')
  return xml.slice(0, end) + section + xml.slice(end)
}

/**
 * Parse a worksheet's `<dataValidations>` section into the canonical
 * DvWireRule[] read model — the exact inverse of serializeRule():
 *
 *   type            → rule.type ('none' → 'any', the messages-only form)
 *   operator        → rule.operator (absent = OOXML default 'between' for the
 *                     operator-carrying types; serializeRule re-omits it)
 *   allowBlank="1"  → rule.allowBlank = true
 *   showDropDown="1"→ rule.showDropDown = false (OOXML INVERTS the name: the
 *                     attribute SUPPRESSES the dropdown; only meaningful for
 *                     list rules, but preserved verbatim for all types so the
 *                     write side re-emits the identical attribute)
 *   showInputMessage / showErrorMessage → true
 *   errorStyle      → Univer number (information=0, warning=2; absent = stop
 *                     = 1 = enum default, field omitted)
 *   errorTitle/error/promptTitle/prompt → string fields
 *   sqref           → ranges[] (space-separated A1 refs)
 *   formula1/formula2 → rule.formula1/formula2 with the install-side
 *                     transforms: list literals keep the quoted form ONLY as
 *                     a marker (the browser install converts `"a,b"` → `a,b`
 *                     and `ref` → `=ref` exactly like the desktop's
 *                     toUniverDvRule); custom keeps the raw body. Formulas
 *                     are NOT untransformed here — the rule stays in the
 *                     wire shape both sides already understand.
 *
 * Fail-closed: x14 extensions, types/operators/error styles outside the
 * canonical whitelists, unreadable sqref, empty ranges, or rule counts past
 * the guard rails throw DvReadError — the caller surfaces NO dvRules for the
 * sheet, so the browser never renders a validation it cannot save faithfully
 * and a no-op save leaves the XML untouched.
 */
export function parseDataValidations(worksheetXml: string): readonly DvWireRule[] {
  const section =
    /<dataValidations\b[^>]*>[\s\S]*?<\/dataValidations>|<dataValidations\b[^>]*\/>/.exec(
      worksheetXml,
    )
  if (!section) return []
  if (/<x14:dataValidation\b/.test(worksheetXml)) {
    throw new DvReadError(
      'This sheet has extended (x14) data validation — it cannot be represented yet.',
    )
  }
  const inner = section[0].includes('</dataValidations>')
    ? section[0].slice(section[0].indexOf('>') + 1, section[0].length - '</dataValidations>'.length)
    : ''
  const rules: DvWireRule[] = []
  for (const match of inner.matchAll(
    /<dataValidation\b([^>]*)\/>|<dataValidation\b([^>]*)>([\s\S]*?)<\/dataValidation>/g,
  )) {
    const attributes = (match[1] ?? match[2] ?? '').trim()
    const body = match[3] ?? ''
    rules.push(parseRule(attributes, body))
    if (rules.length > DV_MAX_RULES) {
      throw new DvReadError(`Worksheet carries more than ${DV_MAX_RULES} validation rules.`)
    }
  }
  return rules
}

function parseRule(attributes: string, body: string): DvWireRule {
  const attr = (name: string): string | undefined =>
    new RegExp(`(?:^|\\s)${name}="([^"]*)"`).exec(attributes)?.[1]

  // sqref → ranges. Empty or malformed refs fail closed — a rule whose
  // target cannot be located must not silently apply elsewhere.
  const sqref = attr('sqref')
  if (sqref === undefined || sqref.trim() === '') {
    throw new DvReadError('dataValidation has no sqref attribute.')
  }
  const ranges: DvCellArea[] = []
  for (const ref of sqref.split(/\s+/)) {
    if (ref === '') continue
    const area = parseSqrefPart(ref)
    if (area === null) {
      throw new DvReadError(`dataValidation sqref "${ref}" is not a readable range.`)
    }
    ranges.push(area)
  }
  if (ranges.length === 0 || ranges.length > DV_MAX_RANGES_PER_RULE) {
    throw new DvReadError(`dataValidation sqref must carry 1..${DV_MAX_RANGES_PER_RULE} ranges.`)
  }

  // type → whitelist; 'none' maps to Univer 'any' (messages-only).
  const rawType = attr('type')
  let type: string | undefined
  if (rawType === undefined || rawType === 'none') {
    type = 'any'
  } else {
    if (!DV_TYPES.has(rawType)) {
      throw new DvReadError(`Unsupported data-validation type "${rawType}".`)
    }
    type = rawType
  }

  // operator → whitelist. Absent means the OOXML default 'between' for the
  // operator-carrying types; serializeRule omits 'between' on write, so an
  // explicit between and an absent attribute are the same state.
  const rawOperator = attr('operator')
  let operator: string | undefined
  if (rawOperator !== undefined) {
    if (!DV_OPERATORS.has(rawOperator)) {
      throw new DvReadError(`Unsupported data-validation operator "${rawOperator}".`)
    }
    operator = rawOperator
  } else if (type !== 'any' && type !== 'list' && type !== 'custom') {
    operator = 'between'
  }

  const errorStyleName = attr('errorStyle')
  let errorStyle: number | undefined
  if (errorStyleName !== undefined) {
    const mapped = DV_ERROR_STYLE_NUMBERS[errorStyleName]
    if (mapped === undefined) {
      throw new DvReadError(`Unsupported data-validation error style "${errorStyleName}".`)
    }
    errorStyle = mapped
  }

  // Formulas: preserved verbatim in the wire shape (the browser install
  // applies the list/custom transforms; the save side inverts them).
  const rule: Record<string, unknown> = { type }
  const formula1 = extractFormula(body, 'formula1')
  const formula2 = extractFormula(body, 'formula2')
  if (attr('allowBlank') === '1') rule.allowBlank = true
  // OOXML showDropDown="1" SUPPRESSES the dropdown (inverted attribute name).
  const rawShowDropDown = attr('showDropDown')
  if (rawShowDropDown !== undefined) rule.showDropDown = rawShowDropDown !== '1'
  if (attr('showInputMessage') === '1') rule.showInputMessage = true
  if (attr('showErrorMessage') === '1') rule.showErrorMessage = true
  if (operator !== undefined) rule.operator = operator
  if (formula1 !== undefined) rule.formula1 = formula1
  if (formula2 !== undefined) rule.formula2 = formula2
  if (errorStyle !== undefined) rule.errorStyle = errorStyle
  for (const key of ['errorTitle', 'error', 'promptTitle', 'prompt'] as const) {
    const value = attr(key)
    if (value !== undefined && value !== '') rule[key] = decodeXmlText(value)
  }
  return { ranges, rule }
}

/// One space-separated sqref token: "A1" or "A1:B4".
function parseSqrefPart(ref: string): DvCellArea | null {
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

function extractFormula(body: string, tag: 'formula1' | 'formula2'): string | undefined {
  const match = new RegExp(`<${tag}>([\\s\\S]*?)</${tag}>`).exec(body)
  if (!match) return undefined
  const text = decodeXmlText(match[1] ?? '')
  return text === '' ? undefined : text
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

function serializeRule(wireRule: DvWireRule): string {
  if (wireRule.ranges.length === 0) {
    throw new DvEditError('A data-validation rule has no ranges.')
  }
  let rule = wireRule.rule
  let rawType = String(rule.type ?? '')
  if (rawType === 'listMultiple') {
    throw new DvEditError(
      'Multi-select list rules are Univer-only and cannot be saved to xlsx — ' +
        'delete the rule before saving.',
    )
  }
  if (rawType === 'checkbox') {
    // Checkbox is Univer-only; degrade to a two-value list so Excel keeps the
    // constraint. The default 1/0 pair round-trips back to a checkbox on load.
    const checked =
      rule.formula1 === undefined || rule.formula1 === '' ? '1' : String(rule.formula1)
    const unchecked =
      rule.formula2 === undefined || rule.formula2 === '' ? '0' : String(rule.formula2)
    rule = {
      ...rule,
      type: 'list',
      operator: undefined,
      formula1: `${checked},${unchecked}`,
      formula2: undefined,
    }
    rawType = 'list'
  }
  // 'any' is the read-side mapping of OOXML type="none" (no constraint, just
  // messages); it round-trips back to the default, attribute-less form.
  const type = rawType === 'any' || rawType === 'none' ? undefined : rawType
  if (type !== undefined && !DV_TYPES.has(type)) {
    throw new DvEditError(`Unsupported data-validation type "${rawType}".`)
  }

  const attrs: string[] = []
  if (type !== undefined) attrs.push(`type="${type}"`)
  const operator = rule.operator === undefined ? undefined : String(rule.operator)
  if (operator !== undefined && operator !== '') {
    if (!DV_OPERATORS.has(operator)) {
      throw new DvEditError(`Unsupported data-validation operator "${operator}".`)
    }
    // "between" is the OOXML default; only operator-carrying types keep it.
    if (operator !== 'between' && type !== undefined && type !== 'list' && type !== 'custom') {
      attrs.push(`operator="${operator}"`)
    }
  }
  if (rule.allowBlank === true) attrs.push('allowBlank="1"')
  // OOXML's showDropDown="1" SUPPRESSES the in-cell dropdown (inverted name);
  // Univer's showDropDown means what it says.
  if (type === 'list' && rule.showDropDown === false) attrs.push('showDropDown="1"')
  if (rule.showInputMessage === true) attrs.push('showInputMessage="1"')
  if (rule.showErrorMessage === true) attrs.push('showErrorMessage="1"')
  const errorStyle = errorStyleName(rule.errorStyle)
  if (errorStyle !== undefined) attrs.push(`errorStyle="${errorStyle}"`)
  for (const [key, attribute] of [
    ['errorTitle', 'errorTitle'],
    ['error', 'error'],
    ['promptTitle', 'promptTitle'],
    ['prompt', 'prompt'],
  ] as const) {
    const value = rule[key]
    if (typeof value === 'string' && value.length > 0) {
      attrs.push(`${attribute}="${escapeXmlAttribute(value)}"`)
    }
  }
  attrs.push(`sqref="${wireRule.ranges.map(toRef).join(' ')}"`)

  const formulas = serializeFormulas(type, rule)
  return formulas === ''
    ? `<dataValidation ${attrs.join(' ')}/>`
    : `<dataValidation ${attrs.join(' ')}>${formulas}</dataValidation>`
}

function errorStyleName(value: unknown): string | undefined {
  if (value === undefined || value === null) return undefined
  const style = Number(value)
  if (!(style in DV_ERROR_STYLE_NAMES)) {
    throw new DvEditError(`Unsupported data-validation error style "${String(value)}".`)
  }
  return DV_ERROR_STYLE_NAMES[style]
}

function serializeFormulas(type: string | undefined, rule: Record<string, unknown>): string {
  const formula1 = formulaText(type, rule.formula1)
  const formula2 = formulaText(type, rule.formula2)
  return (
    (formula1 === undefined ? '' : `<formula1>${escapeXmlText(formula1)}</formula1>`) +
    (formula2 === undefined ? '' : `<formula2>${escapeXmlText(formula2)}</formula2>`)
  )
}

/// Inverse of the install-side formula transforms: list literals regain their
/// quotes, `=`-prefixed references/formulas lose the prefix, and panel-edited
/// date/time strings become Excel serial numbers.
function formulaText(type: string | undefined, raw: unknown): string | undefined {
  if (raw === undefined || raw === null) return undefined
  const text = String(raw)
  if (text === '') return undefined
  if (type === 'list') {
    if (text.startsWith('=')) return text.slice(1)
    // Univer's panel builder serializes list values as a JSON array
    // ('["A","B"]'); the file format carries a quoted CSV literal
    // ('"A,B"'). Normalize the JSON form to CSV so Excel reads either
    // source identically.
    if (text.startsWith('[')) {
      try {
        const parsed = JSON.parse(text)
        if (Array.isArray(parsed) && parsed.every((v) => typeof v === 'string')) {
          return `"${parsed.join(',')}"`
        }
      } catch {
        // Fall through to the quoted-literal form.
      }
    }
    return `"${text}"`
  }
  if (type === 'custom') {
    return text.startsWith('=') ? text.slice(1) : text
  }
  if (type === 'date') {
    const serial = dateToSerial(text)
    if (serial !== undefined) return String(serial)
  }
  if (type === 'time') {
    const fraction = timeToFraction(text)
    if (fraction !== undefined) return String(fraction)
  }
  return text.startsWith('=') ? text.slice(1) : text
}

/// 'YYYY-MM-DD[ HH:mm[:ss]]' (or slashes) → Excel serial (days since
/// 1899-12-30). Plain numbers and references pass through untouched.
function dateToSerial(text: string): number | undefined {
  const match =
    /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T ](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/.exec(
      text.trim(),
    )
  if (!match) return undefined
  const [, year, month, day, hour, minute, second] = match
  const days =
    (Date.UTC(Number(year), Number(month) - 1, Number(day)) - Date.UTC(1899, 11, 30)) / 86_400_000
  const seconds = Number(hour ?? 0) * 3600 + Number(minute ?? 0) * 60 + Number(second ?? 0)
  return seconds === 0 ? days : days + seconds / 86_400
}

function timeToFraction(text: string): number | undefined {
  const match = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/.exec(text.trim())
  if (!match) return undefined
  const [, hour, minute, second] = match
  return (Number(hour) * 3600 + Number(minute) * 60 + Number(second ?? 0)) / 86_400
}

function toRef(range: DvCellArea): string {
  const start = `${columnToLetters(range.startColumn)}${range.startRow + 1}`
  return range.startRow === range.endRow && range.startColumn === range.endColumn
    ? start
    : `${start}:${columnToLetters(range.endColumn)}${range.endRow + 1}`
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

function escapeXmlText(input: string): string {
  return input.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

function escapeXmlAttribute(input: string): string {
  return input
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll('\n', '&#10;')
}
