/// Declarative defined-names save: rewrites workbook.xml's `<definedNames>`
/// from the editor's model. Entries the editor never models — `_xlnm.*`
/// built-ins, hidden names, and names it failed to install (`preserveNames`)
/// — stay byte-verbatim.

export class DefinedNameError extends Error {}

/// A `<definedNames>` section the reader could not classify at all (XML it
/// cannot structurally parse). The whole family fails closed: the snapshot
/// carries no `definedNames` state plus a `namesLocked` marker, so the
/// browser refuses name edits and a no-op save preserves the bytes.
export class DefinedNamesReadError extends Error {}

export interface DefinedNameEntry {
  readonly name: string
  readonly formula: string
  /// localSheetId (position in workbook sheet order) for sheet-scoped names.
  readonly sheetIndex?: number | undefined
}

export interface DefinedNamesState {
  readonly names: readonly DefinedNameEntry[]
  readonly preserveNames: readonly string[]
}

/// Excel name rules (simplified): starts with a letter, `_`, or `\`;
/// continues with letters, digits, `_`, `.`, or `\`; must not look like an
/// A1 or R1C1 cell reference. Letters are Unicode — Excel accepts CJK names,
/// and Create from Selection builds them from localized headers, so this
/// must accept everything definedNameFromLabel produces.
const NAME_PATTERN = /^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u
const CELL_REF_PATTERN = /^(?:[A-Za-z]{1,3}[0-9]+|[Rr][0-9]*[Cc][0-9]*)$/

export function applyDefinedNamesState(workbookXml: string, state: DefinedNamesState): string {
  const preserved = new Set(state.preserveNames)
  const seen = new Set<string>()
  for (const entry of state.names) {
    validateName(entry.name)
    if (preserved.has(entry.name)) {
      throw new DefinedNameError(
        `The name "${entry.name}" also exists in a form the editor cannot model — ` +
          'saving would duplicate it.',
      )
    }
    // Same name may repeat across different scopes, never within one —
    // case-insensitively, the way Excel resolves names ('Total' and 'TOTAL'
    // at one scope are the same name).
    const key = `${entry.name.toLowerCase()}\u0000${entry.sheetIndex ?? -1}`
    if (seen.has(key)) {
      throw new DefinedNameError(`The name "${entry.name}" is defined twice.`)
    }
    seen.add(key)
  }

  // Drop every modeled entry from the existing section, keeping built-ins,
  // hidden names, and preserve-listed ones in place.
  const xml = workbookXml.replace(
    /<definedName\b[^>]*>[\s\S]*?<\/definedName>|<definedName\b[^>]*\/>/g,
    (element) => {
      const name = /\bname="([^"]*)"/.exec(element)?.[1] ?? ''
      const unescaped = unescapeXml(name)
      const keep =
        unescaped.startsWith('_xlnm') ||
        /\bhidden="(?:1|true)"/.test(element) ||
        preserved.has(unescaped)
      return keep ? element : ''
    },
  )

  const additions = state.names
    .map(
      (entry) =>
        `<definedName name="${escapeXmlAttribute(entry.name)}"` +
        (entry.sheetIndex === undefined ? '' : ` localSheetId="${entry.sheetIndex}"`) +
        `>${escapeXmlText(entry.formula.replace(/^=/, ''))}</definedName>`,
    )
    .join('')

  const section = /<definedNames\b[^>]*>([\s\S]*?)<\/definedNames>|<definedNames\b[^>]*\/>/.exec(
    xml,
  )
  if (section) {
    const inner = (section[1] ?? '') + additions
    const replacement = inner === '' ? '' : `<definedNames>${inner}</definedNames>`
    return xml.slice(0, section.index) + replacement + xml.slice(section.index + section[0].length)
  }
  if (additions === '') return xml
  // Schema order: definedNames follows sheets (and functionGroups/externalReferences).
  const anchor = /<\/sheets>|<sheets\b[^>]*\/>/.exec(xml)
  if (!anchor) throw new DefinedNameError('workbook.xml has no sheets element.')
  const externals =
    /<externalReferences\b[^>]*>[\s\S]*?<\/externalReferences>|<externalReferences\b[^>]*\/>/.exec(
      xml,
    )
  const at = externals ? externals.index + externals[0].length : anchor.index + anchor[0].length
  return `${xml.slice(0, at)}<definedNames>${additions}</definedNames>${xml.slice(at)}`
}

function validateName(name: string): void {
  if (
    name.length === 0 ||
    name.length > 255 ||
    !NAME_PATTERN.test(name) ||
    CELL_REF_PATTERN.test(name) ||
    name.toLowerCase() === 'true' ||
    name.toLowerCase() === 'false'
  ) {
    throw new DefinedNameError(`"${name}" is not a valid defined name.`)
  }
  if (name.startsWith('_xlnm')) {
    throw new DefinedNameError('Names starting with "_xlnm" are reserved by Excel.')
  }
}

function unescapeXml(input: string): string {
  return input
    .replaceAll('&quot;', '"')
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&amp;', '&')
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

/// True when the entry is one the writer's keep-rules retain verbatim with
/// no help from `preserveNames` (Excel built-ins and hidden names).
function autoPreserved(name: string, element: string): boolean {
  return name.startsWith('_xlnm') || /\bhidden="(?:1|true)"/.test(element)
}

/// True when a name is valid per the WRITER's rules — the exact predicate
/// `validateName` enforces at save time. A file-native name that fails this
/// must never enter the model: the writer would reject the save (or worse,
/// a rewrite would drop the entry), so the reader routes it to
/// `preserveNames` instead. Exported so the office route's wire validation
/// can enforce the same predicate before bytes reach the engine — one
/// canonical rule set, no browser-side copy.
export function definedNameIsSaveable(name: string): boolean {
  if (
    name.length === 0 ||
    name.length > 255 ||
    !NAME_PATTERN.test(name) ||
    CELL_REF_PATTERN.test(name) ||
    name.toLowerCase() === 'true' ||
    name.toLowerCase() === 'false'
  ) {
    return false
  }
  return !name.startsWith('_xlnm')
}

function nameIsSaveable(name: string): boolean {
  return definedNameIsSaveable(name)
}

/// Parses workbook.xml's `<definedNames>` into the editable model plus the
/// preserve list — the exact inverse of `applyDefinedNamesState`, mirroring
/// the desktop's Rust reader (which filters `_xlnm`, hidden, empty-name, and
/// empty-formula entries) with two additions the web needs:
///
/// - Writer-mirrored name validation (anything `validateName` would reject
///   goes to `preserveNames`, so editing another name can never drop it).
/// - Scope bounds + duplicate ranking per (case-insensitive name, scope) —
///   the canonical writer's exact uniqueness key. A name defined at BOTH
///   workbook scope and sheet scope is TWO legitimate Excel definitions
///   that resolve differently by context: both are modeled (the desktop
///   collapses them because its engine table is name-keyed — the canonical
///   model must not). Only a GENUINE same-scope duplicate (which Excel
///   itself never writes) collapses: the live definition outranks a `#REF!`
///   residue; the winner is modeled, the losers are preserved verbatim.
///
/// Throws `DefinedNamesReadError` when the section's XML cannot be
/// structurally parsed (unclosed elements, duplicate name attributes on one
/// element, non-numeric localSheetId) — the whole family fails closed and
/// the workbook still opens.
export function parseDefinedNamesState(workbookXml: string, sheetCount: number): DefinedNamesState {
  const section = /<definedNames\b[^>]*>([\s\S]*?)<\/definedNames>|<definedNames\b[^>]*\/>/.exec(
    workbookXml,
  )
  if (section === null) return { names: [], preserveNames: [] }
  const inner = section[1] ?? ''
  const names: DefinedNameEntry[] = []
  const preserveNames: string[] = []
  /// Names preserved because an ELEMENT could not be modeled at all
  /// (dangling scope, empty body, unsaveable name). preserveNames is
  /// name-granular — the writer keeps every element carrying a preserved
  /// name verbatim and rejects ANY modeled entry with it — so these names
  /// must never carry a modeled entry either (a live sibling at another
  /// scope stays file-only rather than poisoning every names-dirty save).
  /// Duplicate-group losers below are deliberately NOT in this set: their
  /// modeled winner + preserved loser is the architect-endorsed fail-closed
  /// combination the writer rejects at save time.
  const poisonedNames = new Set<string>()
  const seenElements = new Set<string>()
  const elementPattern = /<definedName\b([^>]*?)(?:\/>|>([\s\S]*?)<\/definedName>)/g
  let match: RegExpExecArray | null
  while ((match = elementPattern.exec(inner)) !== null) {
    const [element, attributeText = '', body = ''] = match
    if (seenElements.has(element)) {
      throw new DefinedNamesReadError('Duplicate <definedName> element parsed.')
    }
    seenElements.add(element)
    const rawName = new RegExp('(?:^|\\s)name="([^"]*)"').exec(attributeText)?.[1] ?? ''
    if (rawName === '' && !/\bname=/.test(attributeText)) {
      throw new DefinedNamesReadError('A <definedName> element carries no name attribute.')
    }
    const name = unescapeXml(rawName)
    if (autoPreserved(name, attributeText)) continue
    const localSheetIdText = /(?:^|\s)localSheetId="([^"]*)"/.exec(attributeText)?.[1]
    let sheetIndex: number | undefined
    if (localSheetIdText !== undefined) {
      if (!/^\d+$/.test(localSheetIdText)) {
        throw new DefinedNamesReadError(
          `The defined name "${name}" carries a non-numeric localSheetId.`,
        )
      }
      const parsed = Number(localSheetIdText)
      if (parsed >= sheetCount) {
        // Scoped to a sheet the workbook does not contain (Excel leaves such
        // names behind when a sheet is deleted). Not modelable — preserve.
        preserveNames.push(name)
        poisonedNames.add(name)
        continue
      }
      sheetIndex = parsed
    }
    const formula = unescapeXml(body).trim()
    if (formula === '') {
      // The desktop reader skips empty bodies; the writer round-trips text
      // verbatim, but an empty element body cannot survive the declarative
      // rewrite as a modeled entry — preserve it verbatim instead.
      preserveNames.push(name)
      poisonedNames.add(name)
      continue
    }
    if (!nameIsSaveable(name)) {
      preserveNames.push(name)
      poisonedNames.add(name)
      continue
    }
    names.push({ name, formula, ...(sheetIndex === undefined ? {} : { sheetIndex }) })
  }
  // The element scan must account for every element the section carries —
  // anything left (a nested or malformed construct) fails the whole family
  // closed rather than risking a rewrite that drops it.
  const elementCount = (inner.match(/<definedName\b/g) ?? []).length
  if (elementCount !== seenElements.size) {
    throw new DefinedNamesReadError('The <definedNames> section could not be fully parsed.')
  }
  // Duplicate ranking per (case-insensitive name, scope) — the canonical
  // writer's exact uniqueness key (case-insensitive because Excel resolves
  // names case-insensitively). The same name at workbook scope and at sheet
  // scope is two LEGITIMATE definitions Excel writes and resolves by
  // context: each forms its own group and both are modeled. Only a GENUINE
  // same-scope duplicate (Excel never writes one; hand-crafted files can)
  // collapses — the live definition outranks a #REF! residue, the winner is
  // modeled, and the losers are preserved verbatim (fail-closed: the name
  // becomes uneditable rather than lost).
  const rank = (entry: DefinedNameEntry): number => (entry.formula.includes('#REF!') ? 1 : 0)
  const byKey = new Map<string, DefinedNameEntry[]>()
  for (const entry of names) {
    const key = `${entry.name.toLowerCase()}\u0000${entry.sheetIndex ?? -1}`
    const list = byKey.get(key) ?? []
    list.push(entry)
    byKey.set(key, list)
  }
  const modeled: DefinedNameEntry[] = []
  for (const list of byKey.values()) {
    // Stable sort: equal ranks keep the file's element order.
    const sorted = [...list].sort((a, b) => rank(a) - rank(b))
    modeled.push(sorted[0]!)
    for (const loser of sorted.slice(1)) preserveNames.push(loser.name)
  }
  // A name preserved for an UNMODELABLE element never carries a modeled
  // entry either (see poisonedNames) — modeling one element of such a name
  // would make every names-dirty save fail the writer's collision guard.
  // The whole name stays file-only, byte-preserved, and uneditable.
  for (let i = modeled.length - 1; i >= 0; i -= 1) {
    if (poisonedNames.has(modeled[i]!.name)) {
      preserveNames.push(modeled[i]!.name)
      modeled.splice(i, 1)
    }
  }
  // Stable order: keep the file's element order among the modeled entries.
  modeled.sort((a, b) => names.indexOf(a) - names.indexOf(b))
  return { names: modeled, preserveNames: dedupeNames(preserveNames) }
}

function dedupeNames(values: readonly string[]): string[] {
  const seen = new Set<string>()
  const out: string[] = []
  for (const value of values) {
    if (seen.has(value)) continue
    seen.add(value)
    out.push(value)
  }
  return out
}
