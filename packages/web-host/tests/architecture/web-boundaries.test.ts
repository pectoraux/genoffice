/**
 * Architecture tests for the browser layer (apps/web) + web-host.
 * (Phase 2C.1 §18, §19; Phase 2C.3.7 path-resolution fix)
 *
 * Enforces:
 *  - apps/web cannot import electron / electron-utils / shell / persistence / pg / pglite
 *  - apps/web cannot contain SQL / pricing formulas / audit mutation / tenant-authority construction
 *  - web-host contains no raw SQL (all persistence SQL in repositories)
 *  - web-host contains no pricing formulas (derived values from domain/service)
 *  - web-host contains no direct audit mutation (audit inside service transactions)
 *  - web-host is the composition root — repository imports ARE allowed (it wires repos into services/CoreApi)
 *  - web-host transport files delegate commercial /api/* to CoreApi.handle()
 *
 * Reads actual source files. Skips comment lines to avoid false positives.
 *
 * Phase 2C.3.7: fixed REPO_ROOT resolution. The test file is at
 * packages/web-host/tests/architecture/ — 4 levels below the repo root
 * (architecture → tests → web-host → packages → repo root).
 * The old code went up only 3 levels, resolving to packages/ instead of
 * the repo root, causing all source scans to return empty arrays.
 * Tests passed vacuously (zero files = zero violations). Now fixed with
 * explicit existence assertions so future path breakage cannot silently
 * produce an empty scan.
 */

import { describe, it, expect } from 'vitest'
import { readdirSync, readFileSync, statSync, existsSync } from 'node:fs'
import { join, relative, dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = dirname(fileURLToPath(import.meta.url))
// The test file is at packages/web-host/tests/architecture/web-boundaries.test.ts
// Going up 4 levels: architecture → tests → web-host → packages → REPO_ROOT
const REPO_ROOT = resolve(__dirname, '..', '..', '..', '..')

function walkTs(dir: string): string[] {
  const files: string[] = []
  if (!existsSync(dir)) return files
  for (const entry of readdirSync(dir)) {
    const full = join(dir, entry)
    if (statSync(full).isDirectory()) {
      files.push(...walkTs(full))
    } else if (entry.endsWith('.ts') || entry.endsWith('.tsx')) {
      files.push(full)
    }
  }
  return files
}
function readFiles(dir: string): { rel: string; content: string }[] {
  return walkTs(dir).map((p) => ({ rel: relative(REPO_ROOT, p), content: readFileSync(p, 'utf8') }))
}
function nonCommentLines(content: string): string[] {
  return content.split('\n').filter((line) => {
    const t = line.trim()
    return !t.startsWith('//') && !t.startsWith('*') && !t.startsWith('/*')
  })
}

// ── Path-resolution guards (Phase 2C.3.7) ──────────────────────────────────
// These assertions prove the source directories exist and contain files.
// Without them, a path-resolution bug would silently produce empty scans
// and all "no violations" tests would pass vacuously.

describe('architecture: source directories are non-empty (path resolution guard)', () => {
  it('apps/web/src exists and contains at least one source file', () => {
    const webDir = join(REPO_ROOT, 'apps', 'web', 'src')
    const files = walkTs(webDir)
    expect(
      files.length,
      `apps/web/src should contain source files (found 0 — path resolution broken?)`,
    ).toBeGreaterThan(0)
  })

  it('packages/web-host/src exists and contains at least one source file', () => {
    const hostDir = join(REPO_ROOT, 'packages', 'web-host', 'src')
    const files = walkTs(hostDir)
    expect(
      files.length,
      `packages/web-host/src should contain source files (found 0 — path resolution broken?)`,
    ).toBeGreaterThan(0)
  })

  it('REPO_ROOT contains the expected top-level directories', () => {
    expect(existsSync(join(REPO_ROOT, 'apps'))).toBe(true)
    expect(existsSync(join(REPO_ROOT, 'packages'))).toBe(true)
    expect(existsSync(join(REPO_ROOT, 'architecture'))).toBe(true)
  })
})

// ── apps/web boundary ──────────────────────────────────────────────────────

describe('architecture: apps/web cannot import Electron / persistence / DB drivers', () => {
  it('apps/web does NOT import electron, @genoffice/electron-utils, apps/shell, pg, @electric-sql/pglite, @contractor/core/persistence, @contractor/core/service, @contractor/core/storage', () => {
    // Scan apps/web/src — the browser bundle surface. Files under apps/web/tests
    // are node-run test infrastructure (vitest node env / Playwright specs) that
    // never enter the bundle; scanning them false-positives on the architecture
    // suite's own regex self-checks (e.g. the string literal
    // "import { ipcRenderer } from 'electron'" used to verify the guard regex).
    // The in-repo apps/web architecture suite enforces the same rules over src/.
    const webDir = join(REPO_ROOT, 'apps', 'web', 'src')
    const webFiles = readFiles(webDir)
    expect(webFiles.length, 'apps/web/src should have source files to scan').toBeGreaterThan(0)
    const forbidden = [
      /from\s+['"]electron['"]/,
      /from\s+['"]@genoffice\/electron-utils['"]/,
      /from\s+['"]@genoffice\/project-store['"]/,
      /from\s+['"]apps\/shell/,
      /from\s+['"]pg['"]/,
      /from\s+['"]@electric-sql\/pglite['"]/,
      /from\s+['"]@contractor\/core\/persistence['"]/,
      /from\s+['"]@contractor\/core\/service['"]/,
      /from\s+['"]@contractor\/core\/storage['"]/,
    ]
    const violations = webFiles.filter((f) => forbidden.some((re) => re.test(f.content)))
    expect(violations.map((v) => v.rel)).toEqual([])
  })
})

describe('architecture: apps/web contains no SQL / pricing / audit / tenant authority', () => {
  it('apps/web does NOT contain INSERT/UPDATE/DELETE/SELECT SQL, pricing formulas, audit mutation, or tenantId-as-authority', () => {
    const webDir = join(REPO_ROOT, 'apps', 'web')
    const webFiles = readFiles(webDir)
    expect(webFiles.length, 'apps/web should have source files to scan').toBeGreaterThan(0)
    const sqlRe =
      /['"`]\s*(INSERT\s+INTO|UPDATE\s+\w+\s+SET|DELETE\s+FROM|SELECT\s+[\s\S]*?\s+FROM)/i
    const pricingRe =
      /\b(totalCost|sellPrice|grossProfit|grossMargin|overhead|contingency)\s*=\s*[^=]/
    const auditRe = /AuditRepository|\.append\s*\(\s*\)|audit\.record\s*\(/
    const tenantAuthorityRe = /\btenantId\s*=\s*[^=]/
    const violations = webFiles.filter((f) => {
      const lines = nonCommentLines(f.content)
      return lines.some(
        (line) =>
          sqlRe.test(line) ||
          pricingRe.test(line) ||
          auditRe.test(line) ||
          tenantAuthorityRe.test(line),
      )
    })
    expect(violations.map((v) => v.rel)).toEqual([])
  })
})

// ── web-host boundary ──────────────────────────────────────────────────────
//
// The web-host is the composition root + HTTP transport adapter. It wires
// repositories into application services and CoreApi, then delegates HTTP
// requests to CoreApi.handle(). Repository imports are ALLOWED in web-host
// because it is the composition root (same as test setup). What is FORBIDDEN:
//  - raw SQL (all persistence SQL must be in repository methods)
//  - pricing formulas (derived values come from the domain/service layer)
//  - direct audit mutation from HTTP handlers (audit is emitted inside service transactions)
//  - bypassing CoreApi for commercial routes (all /api/* routes go through CoreApi)

describe('architecture: web-host contains no SQL / pricing / direct audit mutation', () => {
  it('web-host does NOT contain raw SQL, pricing formulas, or direct audit mutation', () => {
    const hostDir = join(REPO_ROOT, 'packages', 'web-host', 'src')
    const hostFiles = readFiles(hostDir)
    expect(hostFiles.length, 'web-host/src should have source files to scan').toBeGreaterThan(0)
    const pricingRe =
      /\b(totalCost|sellPrice|grossProfit|grossMargin|overhead|contingency)\s*=\s*[^=]/
    // AuditRepository constructor calls in composition-root wiring (new AuditRepository(db))
    // are ALLOWED — the repo is constructed and passed to services/CoreApi. What is
    // forbidden is calling .append() directly from a request handler (bypassing the
    // service transaction). The regex matches .append( calls that are NOT inside
    // the password-auth.ts service (which legitimately calls audit.append inside db.tx).
    const auditRe = /AuditRepository\s*\(\s*\)/
    const directAuditMutationRe = /\.append\s*\(\s*\{/
    const sqlRe =
      /['"`]\s*(INSERT\s+INTO|UPDATE\s+\w+\s+SET|DELETE\s+FROM|SELECT\s+[\s\S]*?\s+FROM)/i
    const violations = hostFiles.filter((f) => {
      const lines = nonCommentLines(f.content)
      return lines.some(
        (line) =>
          pricingRe.test(line) ||
          (auditRe.test(line) && !/new\s+AuditRepository/.test(line)) ||
          (directAuditMutationRe.test(line) &&
            !f.rel.includes('password-auth.ts') &&
            !f.rel.includes('magic-link.ts')) ||
          sqlRe.test(line),
      )
    })
    expect(violations.map((v) => v.rel)).toEqual([])
  })
})

/**
 * Regression: a future web-host module that reintroduces raw SQL should be caught.
 */
describe('architecture: web-host SQL boundary regression', () => {
  it('would catch a future module with raw SQL', () => {
    const sqlRe =
      /['"`]\s*(INSERT\s+INTO|UPDATE\s+\w+\s+SET|DELETE\s+FROM|SELECT\s+[\s\S]*?\s+FROM)/i
    const fakeModule = 'await db.execute(`INSERT INTO users VALUES (...)`)'
    const lines = nonCommentLines(fakeModule)
    expect(lines.some((line) => sqlRe.test(line))).toBe(true)
  })
  it('would NOT flag a repository method call (no SQL string literal)', () => {
    const sqlRe =
      /['"`]\s*(INSERT\s+INTO|UPDATE\s+\w+\s+SET|DELETE\s+FROM|SELECT\s+[\s\S]*?\s+FROM)/i
    const fakeModule = 'await users.createWithPassword(user, hash)'
    const lines = nonCommentLines(fakeModule)
    expect(lines.some((line) => sqlRe.test(line))).toBe(false)
  })
})

// ── CoreApi delegation boundary (Phase 2C.3.5 / 2C.3.6 / 2C.3.7) ──────────
//
// The web-host is the composition root: it wires repositories into services
// and CoreApi, then delegates commercial /api/* requests to CoreApi.handle().
// What is FORBIDDEN in the HTTP transport layer (server.ts, vercel-handler.ts):
//  - direct commercial service invocation (EstimateService, BidService, etc.)
//  - direct createTenantContext construction
//  - repository mutation calls from request handlers (mutations go through CoreApi → service → tx)
//  - a second commercial router before the coreApi.handle() fallback
//
// Auth routes (/api/auth/*) are handled directly — they establish/manage sessions.
// ALL other /api/* routes must fall through to coreApi.handle().

describe('architecture: web-host commercial request handlers delegate to CoreApi', () => {
  // The two HTTP transport files that handle requests
  const transportFiles = ['server.ts', 'vercel-handler.ts']

  // Resolve the source dir relative to the test file (reliable path)
  const srcDir = resolve(__dirname, '..', '..', 'src')

  function getTransportFile(fname: string): string {
    const fullPath = join(srcDir, fname)
    expect(existsSync(fullPath), `Transport file ${fname} not found at ${fullPath}`).toBe(true)
    return readFileSync(fullPath, 'utf8')
  }

  it('transport files contain exactly one coreApi.handle() call for non-auth /api/*', () => {
    for (const fname of transportFiles) {
      const content = getTransportFile(fname)
      // Count occurrences of coreApi.handle( — there must be exactly 1
      const matches = content.match(/coreApi\.handle\s*\(/g)
      expect(matches, `${fname} should have exactly 1 coreApi.handle() call`).toHaveLength(1)
    }
  })

  it('non-auth /api/* routes fall through to coreApi.handle() (no second commercial router)', () => {
    // The pattern is: a series of /api/auth/* if-blocks, then a single
    // if (path.startsWith('/api/')) { ... coreApi.handle() ... } fallback.
    // We verify there's no commercial route handler between the auth blocks
    // and the coreApi.handle() fallback.
    for (const fname of transportFiles) {
      const content = getTransportFile(fname)
      const lines = nonCommentLines(content)
      // Find the line index of the coreApi.handle() call
      const handleIdx = lines.findIndex((l) => /coreApi\.handle\s*\(/.test(l))
      expect(handleIdx).toBeGreaterThan(-1)
      // Check that every /api/ route reference BEFORE coreApi.handle() is an auth route
      const beforeHandle = lines.slice(0, handleIdx)
      const apiRouteRefs = beforeHandle.filter((l) => /\/api\//.test(l))
      const nonAuthApiRefs = apiRouteRefs.filter((l) => !/\/api\/auth\//.test(l))
      // Non-auth /api/ references before coreApi.handle() are NOT route handlers
      // (they can be comments, the startsWith check, or the slice — all OK)
      // The key invariant: no second commercial router (e.g. if path === '/api/projects')
      const commercialRouteHandlers = nonAuthApiRefs.filter((l) =>
        /path\s*(===|startsWith\()\s*['"`]\/api\/(projects|estimates|bids|boqs|boq-items|measurements)/.test(
          l,
        ),
      )
      expect(
        commercialRouteHandlers,
        `${fname} must not have a commercial route handler before coreApi.handle()`,
      ).toEqual([])
    }
  })

  it('transport files do NOT import commercial services directly', () => {
    const commercialServiceRe = /\b(EstimateService|BidService|BOQService|PlanMeasurementService)\b/
    for (const fname of transportFiles) {
      const content = getTransportFile(fname)
      const lines = nonCommentLines(content)
      const violations = lines.filter(
        (line) => commercialServiceRe.test(line) && /from\s+['"]/.test(line),
      )
      expect(violations).toEqual([])
    }
  })

  it('transport files do NOT invoke commercial services directly (no .createEstimateDraft, .submitBid, etc.)', () => {
    // Check for direct commercial service method calls in transport files
    const commercialMethodRe =
      /\.(createEstimateDraft|finalizeEstimate|supersedeEstimate|replayEstimate|updateEstimateDraft|createBid|submitBid|recordBidOutcome|withdrawBid|createBOQ|addBOQItem|updateBOQItemQuantity|createMeasurement)\s*\(/
    for (const fname of transportFiles) {
      const content = getTransportFile(fname)
      const lines = nonCommentLines(content)
      const violations = lines.filter((line) => commercialMethodRe.test(line))
      expect(violations, `${fname} must not call commercial service methods directly`).toEqual([])
    }
  })

  it('transport files do NOT call repository mutation methods directly', () => {
    // Repository mutation methods that should go through CoreApi → service → tx
    const repoMutationRe =
      /\.(create|update|delete|approve|submit|finalize|supersede|withdraw|addItem|updateItemQuantity)\s*\(/
    for (const fname of transportFiles) {
      const content = getTransportFile(fname)
      const lines = nonCommentLines(content)
      // Filter out: constructor calls (new X(...)), non-mutation context (e.g. .createBinding)
      // and composition-root wiring (where repos are constructed, not mutated in request handling)
      // We check for repo mutation calls that are NOT inside the composition-root (getDeps) function
      // The transport files handle requests in the handler function, not in getDeps
      // So we look for mutation calls AFTER the handler function starts
      const handlerIdx = lines.findIndex((l) =>
        /export default|async function handleRequest|async function handle/.test(l),
      )
      if (handlerIdx === -1) continue // server.ts uses startWebHost which delegates
      const handlerLines = lines.slice(handlerIdx)
      const violations = handlerLines.filter(
        (line) =>
          repoMutationRe.test(line) &&
          !/createBinding|setDemoFlag|createDemoUser|createWithPassword|updatePasswordHash/.test(
            line,
          ) && // auth-infrastructure methods
          !/new\s+\w+Repository/.test(line) && // constructor calls
          !/deps\.(coreApi|resolver|passwordAuth|magicLinkAuth|users|memberships|organizations|magicLinks|waitlist|audit|config|magicLinkConfig)\b/.test(
            line,
          ), // deps access
      )
      expect(violations, `${fname} handler must not call repository mutation methods`).toEqual([])
    }
  })

  it('transport files do NOT use createTenantContext', () => {
    for (const fname of transportFiles) {
      const content = getTransportFile(fname)
      const lines = nonCommentLines(content)
      const violations = lines.filter((line) => /createTenantContext/.test(line))
      expect(violations).toEqual([])
    }
  })

  it('auth routes (/api/auth/*) may be handled directly — this is allowed', () => {
    const handlerContent = getTransportFile('vercel-handler.ts')
    expect(handlerContent).toMatch(/\/api\/auth\//)
  })
})

/**
 * Regression: patterns that would catch future bypasses.
 */
describe('architecture: CoreApi delegation boundary regression', () => {
  it('would catch a direct EstimateService import', () => {
    const re = /\b(EstimateService|BidService|BOQService|PlanMeasurementService)\b/
    const fakeImport = "import { EstimateService } from '@contractor/core/service'"
    expect(re.test(fakeImport)).toBe(true)
  })

  it('would catch a direct createTenantContext call', () => {
    expect(
      /createTenantContext/.test('const ctx = createTenantContext(orgId, userId, membership)'),
    ).toBe(true)
  })

  it('would catch a second commercial router before coreApi.handle()', () => {
    const re =
      /path\s*(===|startsWith\()\s*['"`]\/api\/(projects|estimates|bids|boqs|boq-items|measurements)/
    expect(re.test("if (path === '/api/projects') { return ... }")).toBe(true)
  })

  it('would catch a direct boqs.addBOQItem(...) call', () => {
    const re =
      /\.(createEstimateDraft|finalizeEstimate|supersedeEstimate|replayEstimate|updateEstimateDraft|createBid|submitBid|recordBidOutcome|withdrawBid|createBOQ|addBOQItem|updateBOQItemQuantity|createMeasurement)\s*\(/
    expect(re.test('await boqs.addBOQItem(item, boqId, ctx.tenantId)')).toBe(true)
  })

  it('would NOT flag a coreApi.handle() delegation', () => {
    const re = /\b(EstimateService|BidService|BOQService|PlanMeasurementService)\b/
    expect(re.test('const apiRes = await deps.coreApi.handle(apiReq)')).toBe(false)
  })

  it('would NOT flag an auth route handler', () => {
    const re = /\b(EstimateService|BidService|BOQService|PlanMeasurementService)\b/
    expect(re.test("if (path === '/api/auth/password-login')")).toBe(false)
  })

  it('would NOT flag a composition-root constructor call', () => {
    const re =
      /\.(create|update|delete|approve|submit|finalize|supersede|withdraw|addItem|updateItemQuantity)\s*\(/
    expect(re.test('const users = new UserRepository(db)')).toBe(false)
  })
})
