/**
 * Increment 15A — Final architecture assertions.
 *
 * Verifies each of the 7 architectural invariants from the user directive:
 *
 *   runtime-contracts:
 *     ZERO Electron
 *     ZERO node:*
 *     ZERO apps/sheets
 *
 *   services-sheets:
 *     ZERO Electron
 *     ZERO filesystem
 *     ZERO raw sidecar protocol
 *     ZERO Promise<unknown> for pivot
 *
 *   coordinator:
 *     ZERO raw sidecar command construction
 *     ZERO global caller state
 *
 *   migrated handlers:
 *     ZERO pivot parser
 *     ZERO sidecar client
 *     ZERO filesystem implementation
 *     ZERO type assertions
 *
 * Run with: node /home/z/my-project/scripts/final-arch-assertions.mjs
 */
import { readFileSync, readdirSync, statSync } from 'node:fs'
import { join } from 'node:path'

const ROOT = '/home/z/my-project'
let failures = 0
let passes = 0

function listSourceFiles(rootDir) {
  const out = []
  const walk = (dir) => {
    for (const name of readdirSync(dir)) {
      if (name === 'node_modules' || name === 'dist' || name === '.git' || name === 'target') continue
      const full = join(dir, name)
      const st = statSync(full)
      if (st.isDirectory()) walk(full)
      else if (st.isFile() && (full.endsWith('.ts') || full.endsWith('.tsx'))) out.push(full)
    }
  }
  walk(rootDir)
  return out
}

function stripComments(text) {
  return text
    .replace(/\/\*\*?[\s\S]*?\*\//g, '') // block comments
    .replace(/^\s*\/\/.*$/gm, '')         // line comments
}

function check(name, condition, detail = '') {
  if (condition) {
    passes++
    console.log(`  ✓ ${name}`)
  } else {
    failures++
    console.error(`  ✗ ${name}${detail ? ' — ' + detail : ''}`)
  }
}

// ── runtime-contracts ─────────────────────────────────────────────────
console.log('\n=== runtime-contracts ===')
const rcSrc = join(ROOT, 'packages/runtime-contracts/src')
const rcFiles = listSourceFiles(rcSrc)

let rcElectron = 0
let rcNode = 0
let rcAppsSheets = 0
for (const file of rcFiles) {
  const text = readFileSync(file, 'utf8')
  const stripped = stripComments(text)
  // Match `from 'electron'` / `from "electron"` (only as a module specifier)
  if (/(?:from\s+|require\s*\(\s*|import\s*\(\s*)['"]electron['"]/.test(stripped)) rcElectron++
  if (/(?:from\s+|require\s*\(\s*|import\s*\(\s*)['"]node:/.test(stripped)) rcNode++
  if (/(?:from\s+|require\s*\(\s*|import\s*\(\s*)['"][^'"]*apps\/sheets/.test(stripped)) rcAppsSheets++
}
check('runtime-contracts: ZERO Electron imports', rcElectron === 0, `found ${rcElectron}`)
check('runtime-contracts: ZERO node:* imports', rcNode === 0, `found ${rcNode}`)
check('runtime-contracts: ZERO apps/sheets imports', rcAppsSheets === 0, `found ${rcAppsSheets}`)

// ── services-sheets ───────────────────────────────────────────────────
console.log('\n=== services-sheets ===')
const ssSrc = join(ROOT, 'packages/services-sheets/src')
const ssFiles = listSourceFiles(ssSrc)
const serviceFile = join(ssSrc, 'spreadsheet-service.ts')
const serviceSrc = readFileSync(serviceFile, 'utf8')
const serviceStripped = stripComments(serviceSrc)

let ssElectron = 0
let ssFs = 0
let ssSidecarCmd = 0
for (const file of ssFiles) {
  const text = readFileSync(file, 'utf8')
  const stripped = stripComments(text)
  if (/(?:from\s+|require\s*\(\s*|import\s*\(\s*)['"]electron['"]/.test(stripped)) ssElectron++
  if (/(?:from\s+|require\s*\(\s*|import\s*\(\s*)['"]node:/.test(stripped)) ssFs++
  if (/^import.*node:fs/.test(stripped) || /readFileSync|writeFileSync|mkdirSync|rmSync|existsSync|renameSync/.test(stripped)) ssFs++
  if (/command:\s*['"](?:read_entries|read_range|read_formula_cells|read_media|open|close|save_archive|recalc_cells|archive_manifest|scan_entries|convert_workbook)['"]/.test(stripped)) ssSidecarCmd++
}
check('services-sheets: ZERO Electron imports', ssElectron === 0, `found ${ssElectron}`)
check('services-sheets: ZERO filesystem references', ssFs === 0, `found ${ssFs}`)
check('services-sheets: ZERO raw sidecar protocol construction', ssSidecarCmd === 0, `found ${ssSidecarCmd}`)
check(
  'services-sheets: ZERO Promise<unknown> for pivot',
  !/readPivotDefinition[\s\S]*?:\s*Promise<unknown>/.test(serviceStripped),
  'readPivotDefinition still returns Promise<unknown>',
)
check(
  'services-sheets: readPivotDefinition returns Promise<WorkbookPivotDefinition>',
  /readPivotDefinition[\s\S]*?:\s*Promise<WorkbookPivotDefinition>/.test(serviceStripped),
  'readPivotDefinition does not return Promise<WorkbookPivotDefinition>',
)
check(
  'services-sheets: ZERO xlsx-gateway imports (static AND dynamic)',
  !/(?:from\s+|require\s*\(\s*)['"]@genoffice\/xlsx-gateway/.test(serviceStripped) &&
    !/import\s*\(\s*['"]@genoffice\/xlsx-gateway/.test(serviceStripped),
  'spreadsheet-service.ts still references xlsx-gateway',
)

// ── coordinator ────────────────────────────────────────────────────────
console.log('\n=== coordinator ===')
const coordFile = join(ROOT, 'apps/sheets/src/main/sheets-shell-coordinator.ts')
const coordSrc = readFileSync(coordFile, 'utf8')
const coordStripped = stripComments(coordSrc)

check(
  'coordinator: ZERO raw sidecar command construction',
  !/command:\s*['"](?:read_entries|read_range|read_formula_cells|read_media|open|close|save_archive|recalc_cells|archive_manifest|scan_entries|convert_workbook)['"]/.test(coordStripped),
  'coordinator constructs sidecar commands directly',
)
check(
  'coordinator: ZERO global caller state (no currentWcId/activeSession/globalSession)',
  !/^(let|var|const)\s+(currentWcId|activeSession|globalSession|currentRenderer)\b/m.test(coordStripped),
  'coordinator has module-level global caller state',
)
check(
  'coordinator: ZERO sidecarClient references in deps',
  !/sidecarClient\??\s*:/.test(coordStripped),
  'coordinator deps interface still references sidecarClient',
)
check(
  'coordinator: readPivotDefinition returns Promise<WorkbookPivotDefinition>',
  /readPivotDefinition[\s\S]*?:\s*Promise<WorkbookPivotDefinition>/.test(coordStripped),
)
check(
  'coordinator: onWorkbookRenamed callback declared in deps',
  /onWorkbookRenamed\??\s*:\s*\(wcId:\s*number,\s*oldPath:\s*string,\s*newPath:\s*string\)\s*=>\s*void/.test(coordSrc),
)
check(
  'coordinator: renameWorkbook invokes onWorkbookRenamed',
  /this\.deps\.onWorkbookRenamed/.test(coordStripped) &&
    /onWorkbookRenamed\(wcId,\s*oldPath,\s*target\)/.test(coordStripped),
)

// ── migrated handlers ─────────────────────────────────────────────────
console.log('\n=== migrated handlers ===')
const handlerFile = join(ROOT, 'apps/sheets/src/main/sheets-migrated-handlers.ts')
const handlerSrc = readFileSync(handlerFile, 'utf8')
const handlerStripped = stripComments(handlerSrc)

check(
  'handlers: ZERO pivot parser (no parsePivotDefinition)',
  !/parsePivotDefinition/.test(handlerStripped),
  'handler still references parsePivotDefinition',
)
check(
  'handlers: ZERO sidecar client references',
  !/sidecarClient/.test(handlerStripped),
  'handler still references sidecarClient',
)
check(
  'handlers: ZERO filesystem implementation',
  !/^import.*node:fs/.test(handlerStripped) &&
    !/readFileSync|writeFileSync|mkdirSync|rmSync|existsSync/.test(handlerStripped),
  'handler performs filesystem operations',
)
check(
  'handlers: ZERO type assertions (as unknown as / as any / as never)',
  !/\bas\s+unknown\s+as\b/.test(handlerStripped) &&
    !/\bas\s+any\b/.test(handlerStripped) &&
    !/\bas\s+never\b/.test(handlerStripped),
  'handler uses type assertions',
)
check(
  'handlers: ZERO xlsx-gateway imports (static AND dynamic)',
  !/(?:from\s+|require\s*\(\s*)['"]@genoffice\/xlsx-gateway/.test(handlerStripped) &&
    !/import\s*\(\s*['"]@genoffice\/xlsx-gateway/.test(handlerStripped),
  'handler imports xlsx-gateway',
)

// ── engine contract ───────────────────────────────────────────────────
console.log('\n=== engine contract (extra) ===')
const engineContract = join(ROOT, 'packages/runtime-contracts/src/services/spreadsheet-engine.ts')
const engineSrc = readFileSync(engineContract, 'utf8')
const engineStripped = stripComments(engineSrc)

check(
  'engine contract: ZERO readArchiveEntry method declaration',
  !/\breadArchiveEntry\s*\(/.test(engineStripped),
  'engine contract still declares readArchiveEntry',
)
check(
  'engine contract: declares readPivotDefinition',
  /readPivotDefinition[\s\S]*?:\s*Promise<WorkbookPivotDefinition>/.test(engineSrc),
)
check(
  'engine contract: WorkbookPivotDefinition imported',
  /import type \{ WorkbookPivotDefinition \} from ['"]\.\/pivot-definition\.js['"]/.test(engineSrc),
)

// ── engine impl ───────────────────────────────────────────────────────
console.log('\n=== engine impl (extra) ===')
const engineImpl = join(ROOT, 'packages/platform-electron/src/capabilities/electron-xlsx-sidecar-engine.ts')
const engineImplSrc = readFileSync(engineImpl, 'utf8')
const engineImplStripped = stripComments(engineImplSrc)

check(
  'engine impl: ZERO unchecked as Record / as Array / as unknown as casts on sidecar response',
  !/\bas\s+Record<string,\s*unknown>\b/.test(engineImplStripped) &&
    !/\bas\s+Array<Record<string,\s*unknown>>\b/.test(engineImplStripped) &&
    !/\bas\s+unknown\s+as\b/.test(engineImplStripped) &&
    !/\bas\s+any\b/.test(engineImplStripped),
)
check(
  'engine impl: ZERO readArchiveEntry method declaration',
  !/async readArchiveEntry\s*\(/.test(engineImplStripped),
)
check(
  'engine impl: readPivotDefinition returns Promise<WorkbookPivotDefinition>',
  /readPivotDefinition[\s\S]*?:\s*Promise<WorkbookPivotDefinition>/.test(engineImplStripped),
)
check(
  'engine impl: readPivotDefinition uses mkdtempSync + rmSync(workDir)',
  /async readPivotDefinition[\s\S]*?mkdtempSync[\s\S]*?rmSync\(workDir/.test(engineImplStripped),
)
check(
  'engine impl: readPivotDefinition delegates to validateReadEntriesResponse',
  /validateReadEntriesResponse\(raw\)/.test(engineImplStripped),
)

// ── Summary ───────────────────────────────────────────────────────────
console.log('\n=== Summary ===')
console.log(`  ${passes} assertions PASSED`)
console.log(`  ${failures} assertions FAILED`)
if (failures > 0) {
  console.error('\n❌ ARCHITECTURE ASSERTIONS FAILED')
  process.exit(1)
} else {
  console.log('\n✅ ALL ARCHITECTURE ASSERTIONS PASSED')
  process.exit(0)
}
