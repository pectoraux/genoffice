/**
 * Vercel function adapter — wraps CoreApi.handle() as a Vercel serverless
 * function (ADR-0009 D1).
 *
 * This is the PRODUCTION entry point. The dev server (`dev-server.ts`) is
 * unchanged for local development. Both delegate to the same CoreApi +
 * WebSessionResolver + repositories; only the HTTP-entry shape differs.
 *
 * The adapter:
 *  - reads DATABASE_URL (Neon Postgres pooled connection string)
 *  - reads CG_SESSION_SECRET (signed-cookie secret)
 *  - reads CG_MAGIC_LINK_SECRET (magic-link token signing)
 *  - reads CG_APP_BASE_URL (for magic-link URL generation)
 *  - lazily constructs a module-global PostgresClient + CoreApi + resolver
 *    (reused across cold-start invocations)
 *  - applies migrations on first invocation (idempotent)
 *  - routes /api/auth/* to auth endpoints (dev-login if DEV auth; magic-link
 *    request/verify for production)
 *  - routes /api/* to CoreApi.handle()
 *  - serves the built browser bundle from apps/web/dist for non-/api routes
 *
 * No business logic. No SQL. No pricing. No audit mutation. Pure transport.
 */

import type { IncomingMessage, ServerResponse } from 'node:http'
import { Pool } from 'pg'
// Import from individual source files to avoid the persistence barrel export
// which re-exports PgLiteClient (and @electric-sql/pglite WASM).
// The Vercel serverless bundler (esbuild) cannot handle the pglite WASM module.
// All repository imports are from their direct source files; only PostgresClient
// (not PgLiteClient) is imported at runtime.
import { PostgresClient } from '@contractor/core/persistence/postgres-client.js'
import { PgLiteClient } from '@contractor/core/persistence/pglite-client.js'
import { applyMigration } from '@contractor/core/persistence/db-client.js'
import type { DbClient } from '@contractor/core/persistence/db-client.js'
import type { Membership } from '@contractor/core/domain'
import { FOUNDATION_MIGRATION_SQL, COMMERCIAL_MIGRATION_SQL, MAGIC_LINKS_MIGRATION_SQL, AUTH_MIGRATION_SQL } from '@contractor/core/persistence/migrations-loader.js'
import { OrganizationRepository } from '@contractor/core/persistence/repositories/organization.repository.js'
import { UserRepository } from '@contractor/core/persistence/repositories/user.repository.js'
import { MembershipRepository } from '@contractor/core/persistence/repositories/membership.repository.js'
import { WorkspaceRepository } from '@contractor/core/persistence/repositories/workspace.repository.js'
import { ProjectRepository } from '@contractor/core/persistence/repositories/project.repository.js'
import { AuditRepository } from '@contractor/core/persistence/repositories/audit.repository.js'
import { RevisionRepository } from '@contractor/core/persistence/repositories/revision.repository.js'
import { PlanMeasurementRepository } from '@contractor/core/persistence/repositories/plan-measurement.repository.js'
import { BOQRepository } from '@contractor/core/persistence/repositories/boq.repository.js'
import { EstimateRevisionRepository } from '@contractor/core/persistence/repositories/estimate-revision.repository.js'
import { BidRepository } from '@contractor/core/persistence/repositories/bid.repository.js'
import { MagicLinkRepository } from '@contractor/core/persistence/repositories/magic-link.repository.js'
import { WaitlistRepository } from '@contractor/core/persistence/repositories/waitlist.repository.js'
import type { AuditRepository as AuditRepoType } from '@contractor/core/persistence/repositories/audit.repository.js'
import {
  IdentityService, OrganizationService, WorkspaceService, ProjectService,
  AuditService, RevisionService, PlanMeasurementService, BOQService, EstimateService, BidService,
} from '@contractor/core/service'
import { CoreApi, type ApiRequest, type ApiResponse, routeOffice } from '@contractor/core/api'
import {
  loadSessionConfigFromEnv, WebSessionResolver,
  signSession, sessionCookieHeader, clearSessionCookieHeader,
} from './index.js'
import { MagicLinkAuthService } from './magic-link.js'
import { PasswordAuthService } from './password-auth.js'

// ── Module-global singletons (reused across cold-start invocations) ──────────
// Vercel reuses the module across invocations within a warm instance.
// The PostgresClient holds a Pool (not a connection); connections are checked
// out per-tx and released in finally (see postgres-client.ts).
let cachedDeps: CachedDeps | null = null

/**
 * Inject deps for testing (bypasses the lazy getDeps() initialization).
 * Allows tests to share a single PGlite instance with the handler.
 */
export function setCachedDepsForTesting(deps: CachedDeps): void {
  cachedDeps = deps
}

interface CachedDeps {
  coreApi: CoreApi
  resolver: WebSessionResolver
  users: UserRepository
  memberships: MembershipRepository
  organizations: OrganizationRepository
  magicLinks: MagicLinkRepository
  magicLinkAuth: MagicLinkAuthService
  passwordAuth: PasswordAuthService
  waitlist: WaitlistRepository
  audit: AuditRepoType
  config: ReturnType<typeof loadSessionConfigFromEnv>
  magicLinkConfig: { linkSecret: string; linkTtlSeconds: number; appBaseUrl: string }
}

async function getDeps(): Promise<CachedDeps> {
  if (cachedDeps) return cachedDeps
  const db = createDb()
  // Migrations are applied lazily on first invocation (idempotent).
  // In production, prefer running migrations as a deploy step instead.
  await applyMigrations(db)
  const users = new UserRepository(db)
  const memberships = new MembershipRepository(db)
  const organizations = new OrganizationRepository(db)
  const workspaces = new WorkspaceRepository(db)
  // Seed deterministic demo data when running in ephemeral PGlite mode.
  if (demoMode) {
    await seedDemoData({ users, memberships, organizations, workspaces })
  }
  const projects = new ProjectRepository(db)
  const audit = new AuditRepository(db)
  const revisions = new RevisionRepository(db)
  const pm = new PlanMeasurementRepository(db)
  const boq = new BOQRepository(db)
  const estRev = new EstimateRevisionRepository(db)
  const bids = new BidRepository(db)
  const magicLinks = new MagicLinkRepository(db)
  const waitlist = new WaitlistRepository(db)

  const identity = new IdentityService(users, memberships)
  const orgService = new OrganizationService(organizations, memberships, audit)
  const wsService = new WorkspaceService(workspaces, audit)
  const projService = new ProjectService(projects, workspaces, audit)
  const auditService = new AuditService(audit)
  const revService = new RevisionService(revisions, projects, audit)
  const measurements = new PlanMeasurementService(db, pm, projects, audit)
  const boqs = new BOQService(db, boq, projects, audit)
  const estimates = new EstimateService(db, estRev, projects, audit)
  const bidService = new BidService(db, bids, estRev, audit)
  const config = loadSessionConfigFromEnv()
  const resolver = new WebSessionResolver({ users, memberships, config })
  const coreApi = new CoreApi(
    { identity, organizations: orgService, workspaces: wsService, projects: projService,
      audit: auditService, revisions: revService, measurements, boqs, estimates, bids: bidService },
    resolver,
  )
  const magicLinkConfig = {
    linkSecret: process.env.CG_MAGIC_LINK_SECRET ?? '',
    linkTtlSeconds: Number(process.env.CG_MAGIC_LINK_TTL_SECONDS ?? 900),
    appBaseUrl: process.env.CG_APP_BASE_URL ?? '',
  }
  if (!magicLinkConfig.linkSecret || magicLinkConfig.linkSecret.length < 32) {
    // In production (no DEV auth), magic-link is the auth path — require the secret.
    if (!config.devAuthEnabled) {
      throw new Error('CG_MAGIC_LINK_SECRET must be set (min 32 bytes) for production magic-link auth')
    }
  }
  const magicLinkAuth = new MagicLinkAuthService(users, magicLinks, magicLinkConfig)
  const passwordAuth = new PasswordAuthService({ db, users, memberships, organizations, waitlist, audit })
  cachedDeps = {
    coreApi, resolver, users, memberships, organizations, magicLinks, magicLinkAuth, passwordAuth,
    waitlist, audit, config, magicLinkConfig,
  }
  return cachedDeps
}

// Demo mode flag: true when running on an ephemeral in-memory PGlite instance
// (no DATABASE_URL configured). Demo users are seeded with deterministic IDs so
// demo-login survives serverless instance recycles.
let demoMode = false

function createDb(): DbClient {
  // Production: real PostgreSQL via DATABASE_URL (persistent).
  // Only honor a real postgres:// connection string; ignore non-postgres
  // values (e.g. a stray SQLite file: URL) so the demo PGlite fallback still
  // kicks in and the app stays runnable.
  if (process.env.DATABASE_URL && /^postgres(ql)?:\/\//.test(process.env.DATABASE_URL)) {
    const pool = new Pool({
      connectionString: process.env.DATABASE_URL,
      max: 10,
      idleTimeoutMillis: 30000,
      connectionTimeoutMillis: 10000,
      ssl: process.env.DATABASE_SSL === '1' ? { rejectUnauthorized: false } : undefined,
    })
    return new PostgresClient(pool)
  }
  // Demo fallback: PGlite (in-memory WASM PostgreSQL). Lets the deployed app
  // run immediately without a configured Postgres. Data is ephemeral per
  // serverless instance; demo users are re-seeded with deterministic IDs on
  // each cold start so demo-login persists across instance recycles. For
  // persistent production data, set DATABASE_URL to a real Postgres URL.
  demoMode = true
  return new PgLiteClient()
}

// Deterministic IDs so demo sessions survive serverless instance recycles.
const DEMO_ORG_ID = 'org_demo_0001'
const DEMO_ORG_SLUG = 'genoffice-demo'
const DEMO_WS_ID = 'ws_demo_default'
const DEMO_USERS = [
  { id: 'usr_demo_owner', role: 'owner', email: 'demo-owner@contractor.dev', name: 'Demo Owner' },
  { id: 'usr_demo_member', role: 'member', email: 'demo-member@contractor.dev', name: 'Demo Member' },
  { id: 'usr_demo_viewer', role: 'viewer', email: 'demo-viewer@contractor.dev', name: 'Demo Viewer' },
] as const

async function seedDemoData(deps: {
  users: UserRepository
  memberships: MembershipRepository
  organizations: OrganizationRepository
  workspaces: WorkspaceRepository
}): Promise<void> {
  const now = new Date().toISOString()
  // Demo org (idempotent — self-tenant: tenantId === org id).
  const existingOrg = await deps.organizations.getById(DEMO_ORG_ID, DEMO_ORG_ID)
  if (!existingOrg) {
    await deps.organizations.create({
      id: DEMO_ORG_ID, tenantId: DEMO_ORG_ID, name: 'GenOffice Demo', slug: DEMO_ORG_SLUG,
      status: 'active', createdAt: now,
    })
  }
  // Default workspace for the demo org (idempotent) so project creation works
  // out of the box (the Projects screen requires at least one workspace).
  if (!(await deps.workspaces.getById(DEMO_WS_ID, DEMO_ORG_ID))) {
    await deps.workspaces.create({
      id: DEMO_WS_ID, tenantId: DEMO_ORG_ID, organizationId: DEMO_ORG_ID,
      name: 'Default Workspace', createdAt: now,
    })
  }
  // Demo users + memberships (idempotent).
  for (const u of DEMO_USERS) {
    if (await deps.users.getByEmail(u.email)) continue
    await deps.users.createDemoUser(
      { id: u.id, email: u.email, displayName: u.name, status: 'active', createdAt: now },
    )
    // Email binding (password-auth path).
    await deps.users.createBinding({
      id: `auth_${u.id}`, userId: u.id, provider: 'email', subject: u.email,
      createdAt: now, lastUsedAt: null,
    })
    // Web binding (provider='web', subject=userId) — required by the session
    // resolver's resolveTenantContext so demo sessions are authorized for
    // Core API calls (projects, workspaces, etc.).
    await deps.users.createBinding({
      id: `web_${u.id}`, userId: u.id, provider: 'web', subject: u.id,
      createdAt: now, lastUsedAt: null,
    })
    const membership: Membership = {
      id: `mbr_${u.id}`, userId: u.id, organizationId: DEMO_ORG_ID,
      role: u.role, status: 'active', createdAt: now,
    }
    await deps.memberships.create(membership)
  }
}

async function applyMigrations(db: DbClient): Promise<void> {
  await applyMigration(db, FOUNDATION_MIGRATION_SQL)
  await applyMigration(db, COMMERCIAL_MIGRATION_SQL)
  await applyMigration(db, MAGIC_LINKS_MIGRATION_SQL)
  await applyMigration(db, AUTH_MIGRATION_SQL)
}

// ── Vercel function handler ─────────────────────────────────────────────────

export default async function handler(req: IncomingMessage, res: ServerResponse): Promise<void> {
  try {
    const url = new URL(req.url ?? '/', 'http://localhost')
    const path = url.pathname
    const method = req.method ?? 'GET'

    // Stateless routes — do NOT require a DB connection.
    // dev-mode just reports whether DEV auth is enabled (env flag).
    // logout just clears the session cookie.
    // These must work even when DATABASE_URL is unset so the frontend can boot
    // and report its auth mode.
    if (path === '/api/auth/dev-mode' && method === 'GET') {
      const devAuthEnabled = process.env.CONTRACTOR_DEV_AUTH === '1'
      return sendJson(res, 200, { devAuth: devAuthEnabled })
    }
    if (path === '/api/auth/logout' && method === 'POST') {
      res.setHeader('Set-Cookie', clearSessionCookieHeader(process.env.NODE_ENV === "production"))
      return sendJson(res, 200, { ok: true })
    }

    // All other routes require the full deps (DB + services).
    const deps = await getDeps()

    // Office file routes — pure (no DB), but we still need getDeps() to be ready
    // before any subsequent CoreApi call. Bypass auth entirely: office routes
    // are stateless file transforms. The 14MB JSON envelope cap accommodates
    // the 10MB-decoded base64 fileBytes limit (base64 inflates by ~33%).
    if (path.startsWith('/api/office/')) {
      let officeBody: unknown
      try {
        officeBody = method === 'GET' || method === 'HEAD' ? null : await readLargeJsonBody(req)
      } catch (e) {
        const msg = e instanceof Error ? e.message : 'invalid_body'
        return sendJson(res, 400, {
          error: 'validation',
          message: msg === 'payload_too_large' ? 'Payload too large' : 'Invalid JSON body',
        })
      }
      const officeRes = await routeOffice({ method, path, body: officeBody })
      if (officeRes) return sendApiResponse(res, officeRes)
      // Fall through to 404 if routeOffice returned null (unknown office route)
      return sendJson(res, 404, { error: 'not_found', message: 'Unknown office route' })
    }

    // Auth routes (stateful — require DB)
    if (path === '/api/auth/password-login' && method === 'POST') {
      return handlePasswordLogin(req, res, deps)
    }
    if (path === '/api/auth/signup' && method === 'POST') {
      return handleSignup(req, res, deps)
    }
    if (path === '/api/auth/demo-login' && method === 'POST') {
      return handleDemoLogin(req, res, deps)
    }
    if (path === '/api/auth/waitlist' && method === 'GET') {
      return handleListWaitlist(req, res, deps)
    }
    if (path === '/api/auth/waitlist' && method === 'POST') {
      return handleApproveWaitlist(req, res, deps)
    }
    if (path === '/api/auth/dev-login' && method === 'POST') {
      return handleDevLogin(req, res, deps)
    }
    if (path === '/api/auth/request-link' && method === 'POST') {
      return handleRequestLink(req, res, deps)
    }
    if (path === '/api/auth/verify' && method === 'GET') {
      return handleVerify(req, res, deps, url)
    }
    if (path === '/api/auth/memberships' && method === 'GET') {
      return handleListMemberships(req, res, deps)
    }
    if (path === '/api/auth/select-tenant' && method === 'POST') {
      return handleSelectTenant(req, res, deps)
    }
    if (path === '/api/auth/session' && method === 'GET') {
      return handleSession(req, res, deps)
    }

    // Core API routes
    if (path.startsWith('/api/')) {
      let body: unknown
      try {
        body = method === 'GET' || method === 'HEAD' ? null : await readJsonBody(req)
      } catch {
        return sendJson(res, 400, { error: 'validation', message: 'Invalid JSON body' })
      }
      const cookieHeader = req.headers.cookie
      const apiReq: ApiRequest = {
        method,
        path: path.slice('/api'.length),
        headers: { authorization: `Bearer ${cookieHeader ?? ''}` },
        body,
      }
      const apiRes = await deps.coreApi.handle(apiReq)
      return sendApiResponse(res, apiRes)
    }

    // Non-/api routes: the built browser bundle is served by Vercel's static
    // hosting (configured in vercel.json). This function only handles /api/*.
    return sendJson(res, 404, { error: 'not_found', message: 'Not found' })
  } catch (e) {
    console.error('[vercel-handler] internal error:', e)
    return sendJson(res, 500, { error: 'internal_error', message: 'Internal server error' })
  }
}

// ── Auth route handlers (shared with dev server) ────────────────────────────

async function handleDevLogin(req: IncomingMessage, res: ServerResponse, deps: CachedDeps): Promise<void> {
  if (!deps.config.devAuthEnabled) {
    return sendJson(res, 404, { error: 'not_found', message: 'Dev auth is not enabled' })
  }
  const body = await readJsonBody(req)
  if (!body || typeof body !== 'object') return sendJson(res, 400, { error: 'validation', message: 'Invalid body' })
  const credential = (body as Record<string, unknown>).credential
  if (typeof credential !== 'string' || credential.length === 0) {
    return sendJson(res, 400, { error: 'validation', message: 'credential required' })
  }
  if (credential !== deps.config.devCredential) {
    return sendJson(res, 401, { error: 'unauthenticated', message: 'Invalid dev credential' })
  }
  const devUserEmail = process.env.CG_DEV_USER_EMAIL
  if (!devUserEmail) return sendJson(res, 500, { error: 'internal_error', message: 'Dev user email not configured' })
  const user = await deps.users.getByEmail(devUserEmail)
  if (!user || user.status !== 'active') {
    return sendJson(res, 401, { error: 'unauthenticated', message: 'Dev user not found or inactive' })
  }
  const existingBinding = await deps.users.getBindingBySubject('web', user.id)
  if (!existingBinding) {
    await deps.users.createBinding({
      id: 'dev-binding-' + user.id, userId: user.id, provider: 'web', subject: user.id,
      createdAt: new Date().toISOString(), lastUsedAt: null,
    })
  }
  const exp = Math.floor(Date.now() / 1000) + deps.config.sessionTtlSeconds
  const token = signSession({ userId: user.id, selectedMembershipId: null, exp }, deps.config.sessionSecret)
  res.setHeader('Set-Cookie', sessionCookieHeader(token, deps.config.sessionTtlSeconds, process.env.NODE_ENV === "production"))
  return sendJson(res, 200, { userId: user.id, email: user.email, displayName: user.displayName })
}

async function handleRequestLink(req: IncomingMessage, res: ServerResponse, deps: CachedDeps): Promise<void> {
  const body = await readJsonBody(req)
  if (!body || typeof body !== 'object') return sendJson(res, 400, { error: 'validation', message: 'Invalid body' })
  const email = (body as Record<string, unknown>).email
  if (typeof email !== 'string' || email.length === 0) {
    return sendJson(res, 400, { error: 'validation', message: 'email required' })
  }
  try {
    const result = await deps.magicLinkAuth.requestLink(email)
    // For dev: log the link (no email service). For production: an email
    // provider would be wired here (EMAIL_API_KEY + Resend/SendGrid).
    if (deps.config.devAuthEnabled) {
      console.log(`[magic-link] Dev mode — link for ${result.email}: ${result.linkUrl}`)
    }
    // Never return the link/token in the response body (security: the link
    // must be delivered out-of-band via email). Return only a confirmation.
    return sendJson(res, 200, { sent: true, email: result.email })
  } catch (e) {
    return sendJson(res, 400, { error: 'validation', message: e instanceof Error ? e.message : 'Failed' })
  }
}

async function handleVerify(req: IncomingMessage, res: ServerResponse, deps: CachedDeps, url: URL): Promise<void> {
  const token = url.searchParams.get('token')
  if (!token) return sendJson(res, 400, { error: 'validation', message: 'token required' })
  try {
    const result = await deps.magicLinkAuth.verifyLink(token)
    const exp = Math.floor(Date.now() / 1000) + deps.config.sessionTtlSeconds
    const sessionToken = signSession({ userId: result.userId, selectedMembershipId: null, exp }, deps.config.sessionSecret)
    res.setHeader('Set-Cookie', sessionCookieHeader(sessionToken, deps.config.sessionTtlSeconds, process.env.NODE_ENV === "production"))
    // Redirect to the app root (tenant selection follows)
    res.writeHead(302, { Location: '/' })
    res.end()
    return
  } catch (e) {
    return sendJson(res, 401, { error: 'unauthenticated', message: e instanceof Error ? e.message : 'Invalid token' })
  }
}

async function handleListMemberships(req: IncomingMessage, res: ServerResponse, deps: CachedDeps): Promise<void> {
  const payload = deps.resolver.resolvePayload(req.headers.cookie)
  if (!payload) return sendJson(res, 401, { error: 'unauthenticated', message: 'Not authenticated' })
  const user = await deps.users.getById(payload.userId)
  if (!user || user.status !== 'active') return sendJson(res, 401, { error: 'unauthenticated', message: 'User not found' })
  const memberships = await deps.memberships.listTenantsForUser(payload.userId)
  const result = []
  for (const m of memberships) {
    const org = await deps.organizations.getById(m.organizationId, m.organizationId)
    result.push({ membershipId: m.id, organizationId: m.organizationId, organizationName: org?.name ?? m.organizationId, role: m.role })
  }
  return sendJson(res, 200, { memberships: result })
}

async function handleSelectTenant(req: IncomingMessage, res: ServerResponse, deps: CachedDeps): Promise<void> {
  const payload = deps.resolver.resolvePayload(req.headers.cookie)
  if (!payload) return sendJson(res, 401, { error: 'unauthenticated', message: 'Not authenticated' })
  const body = await readJsonBody(req)
  if (!body || typeof body !== 'object') return sendJson(res, 400, { error: 'validation', message: 'Invalid body' })
  const membershipId = (body as Record<string, unknown>).membershipId
  if (typeof membershipId !== 'string' || membershipId.length === 0) {
    return sendJson(res, 400, { error: 'validation', message: 'membershipId required' })
  }
  const userMemberships = await deps.memberships.listTenantsForUser(payload.userId)
  const found = userMemberships.find((m) => m.id === membershipId)
  if (!found) return sendJson(res, 403, { error: 'forbidden', message: 'Membership not found or not yours' })
  const exp = Math.floor(Date.now() / 1000) + deps.config.sessionTtlSeconds
  const token = signSession({ userId: payload.userId, selectedMembershipId: membershipId, exp }, deps.config.sessionSecret)
  res.setHeader('Set-Cookie', sessionCookieHeader(token, deps.config.sessionTtlSeconds, process.env.NODE_ENV === "production"))
  return sendJson(res, 200, { tenantId: found.organizationId, membershipId: found.id, role: found.role })
}

async function handleSession(req: IncomingMessage, res: ServerResponse, deps: CachedDeps): Promise<void> {
  const payload = deps.resolver.resolvePayload(req.headers.cookie)
  if (!payload) return sendJson(res, 200, { authenticated: false })
  const user = await deps.users.getById(payload.userId)
  if (!user || user.status !== 'active') return sendJson(res, 200, { authenticated: false })
  return sendJson(res, 200, {
    authenticated: true, userId: user.id, email: user.email, displayName: user.displayName,
    tenantSelected: payload.selectedMembershipId !== null,
  })
}

// ── Password auth + waitlist handlers (Phase 2C.3) ──────────────────────────

async function handlePasswordLogin(req: IncomingMessage, res: ServerResponse, deps: CachedDeps): Promise<void> {
  const body = await readJsonBody(req)
  if (!body || typeof body !== 'object') return sendJson(res, 400, { error: 'validation', message: 'Invalid body' })
  const b = body as Record<string, unknown>
  const email = typeof b.email === 'string' ? b.email : ''
  const password = typeof b.password === 'string' ? b.password : ''
  if (!email || !password) return sendJson(res, 400, { error: 'validation', message: 'email and password required' })
  const result = await deps.passwordAuth.login(email, password)
  if (!result) return sendJson(res, 401, { error: 'unauthenticated', message: 'Invalid email or password' })
  const exp = Math.floor(Date.now() / 1000) + deps.config.sessionTtlSeconds
  const token = signSession({ userId: result.userId, selectedMembershipId: null, exp }, deps.config.sessionSecret)
  res.setHeader('Set-Cookie', sessionCookieHeader(token, deps.config.sessionTtlSeconds, process.env.NODE_ENV === "production"))
  return sendJson(res, 200, { userId: result.userId })
}

async function handleSignup(req: IncomingMessage, res: ServerResponse, deps: CachedDeps): Promise<void> {
  const body = await readJsonBody(req)
  if (!body || typeof body !== 'object') return sendJson(res, 400, { error: 'validation', message: 'Invalid body' })
  const b = body as Record<string, unknown>
  const email = typeof b.email === 'string' ? b.email : ''
  const displayName = typeof b.displayName === 'string' ? b.displayName : null
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return sendJson(res, 400, { error: 'validation', message: 'valid email required' })
  }
  const entry = await deps.passwordAuth.joinWaitlist(email, displayName)
  return sendJson(res, 200, { id: entry.id, email: entry.email, status: entry.status, message: 'You are on the waitlist. An admin will review your request.' })
}

async function handleDemoLogin(req: IncomingMessage, res: ServerResponse, deps: CachedDeps): Promise<void> {
  const body = await readJsonBody(req)
  if (!body || typeof body !== 'object') return sendJson(res, 400, { error: 'validation', message: 'Invalid body' })
  const role = (body as Record<string, unknown>).role
  if (role !== 'owner' && role !== 'member' && role !== 'viewer') {
    return sendJson(res, 400, { error: 'validation', message: 'role must be owner, member, or viewer' })
  }
  const result = await deps.passwordAuth.demoLogin(role)
  if (!result) return sendJson(res, 401, { error: 'unauthenticated', message: 'Demo user not found. Run the bootstrap script.' })
  const exp = Math.floor(Date.now() / 1000) + deps.config.sessionTtlSeconds
  const token = signSession({ userId: result.userId, selectedMembershipId: null, exp }, deps.config.sessionSecret)
  res.setHeader('Set-Cookie', sessionCookieHeader(token, deps.config.sessionTtlSeconds, process.env.NODE_ENV === "production"))
  return sendJson(res, 200, { userId: result.userId, role })
}

async function handleListWaitlist(req: IncomingMessage, res: ServerResponse, deps: CachedDeps): Promise<void> {
  const payload = deps.resolver.resolvePayload(req.headers.cookie)
  if (!payload) return sendJson(res, 401, { error: 'unauthenticated', message: 'Not authenticated' })
  // H4: use the session's SELECTED membership for admin verification
  if (!payload.selectedMembershipId) {
    return sendJson(res, 403, { error: 'forbidden', message: 'No tenant selected' })
  }
  const userMemberships = await deps.memberships.listTenantsForUser(payload.userId)
  const selectedM = userMemberships.find((m) => m.id === payload.selectedMembershipId)
  if (!selectedM || selectedM.status !== 'active' || (selectedM.role !== 'admin' && selectedM.role !== 'owner')) {
    return sendJson(res, 403, { error: 'forbidden', message: 'Admin or owner role required for the selected tenant' })
  }
  const entries = await deps.waitlist.listAll()
  return sendJson(res, 200, { entries })
}

async function handleApproveWaitlist(req: IncomingMessage, res: ServerResponse, deps: CachedDeps): Promise<void> {
  const payload = deps.resolver.resolvePayload(req.headers.cookie)
  if (!payload) return sendJson(res, 401, { error: 'unauthenticated', message: 'Not authenticated' })
  // H4 FIX: use the session's SELECTED membership, NOT the first admin/owner membership.
  // This ensures the approval creates the user in the tenant the admin actually selected.
  if (!payload.selectedMembershipId) {
    return sendJson(res, 403, { error: 'forbidden', message: 'No tenant selected' })
  }
  // Verify the selected membership belongs to the authenticated user + is active + is admin/owner.
  const userMemberships = await deps.memberships.listTenantsForUser(payload.userId)
  const selectedM = userMemberships.find((m) => m.id === payload.selectedMembershipId)
  if (!selectedM) {
    return sendJson(res, 403, { error: 'forbidden', message: 'Selected membership not found or revoked' })
  }
  if (selectedM.status !== 'active') {
    return sendJson(res, 403, { error: 'forbidden', message: 'Selected membership is not active' })
  }
  if (selectedM.role !== 'admin' && selectedM.role !== 'owner') {
    return sendJson(res, 403, { error: 'forbidden', message: 'Admin or owner role required for the selected tenant' })
  }
  const body = await readJsonBody(req)
  if (!body || typeof body !== 'object') return sendJson(res, 400, { error: 'validation', message: 'Invalid body' })
  const b = body as Record<string, unknown>
  const waitlistId = typeof b.waitlistId === 'string' ? b.waitlistId : ''
  const password = typeof b.password === 'string' ? b.password : ''
  if (!waitlistId || !password || password.length < 6) {
    return sendJson(res, 400, { error: 'validation', message: 'waitlistId and password (min 6 chars) required' })
  }
  try {
    // H4: tenantId comes from the SELECTED membership — never client-supplied.
    const result = await deps.passwordAuth.approveWaitlistEntry(
      waitlistId, payload.userId, selectedM.organizationId, password,
    )
    return sendJson(res, 200, { userId: result.userId, email: result.email, message: 'User created. They can now login with their email + the password you set.' })
  } catch (e) {
    // H6 FIX: map expected validation/conflict errors appropriately;
    // unexpected DB/internal failures → 500 with sanitized message (no SQL/stack/schema leak).
    const msg = e instanceof Error ? e.message : 'Failed to approve'
    if (msg.includes('not found') || msg.includes('already') || msg.includes('could not be approved')) {
      return sendJson(res, 409, { error: 'conflict', message: msg })
    }
    // Unexpected error — do not leak details
    return sendJson(res, 500, { error: 'internal_error', message: 'Failed to approve waitlist entry' })
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────

async function readJsonBody(req: IncomingMessage): Promise<unknown> {
  const chunks: Buffer[] = []
  let total = 0
  for await (const chunk of req) {
    total += chunk.length
    if (total > 1024 * 1024) throw new Error('payload_too_large')
    chunks.push(chunk as Buffer)
  }
  if (chunks.length === 0) return null
  const text = Buffer.concat(chunks).toString('utf8')
  if (text.length === 0) return null
  return JSON.parse(text)
}

/**
 * Larger body reader for office file requests. The 10MB decoded-fileBytes cap
 * inflates to ~13.4MB base64; we allow 14MB total envelope size to give the
 * JSON wrapper some headroom while still resisting oversized payloads.
 */
async function readLargeJsonBody(req: IncomingMessage): Promise<unknown> {
  const chunks: Buffer[] = []
  let total = 0
  const MAX = 14 * 1024 * 1024
  for await (const chunk of req) {
    total += chunk.length
    if (total > MAX) throw new Error('payload_too_large')
    chunks.push(chunk as Buffer)
  }
  if (chunks.length === 0) return null
  const text = Buffer.concat(chunks).toString('utf8')
  if (text.length === 0) return null
  try {
    return JSON.parse(text)
  } catch {
    throw new Error('invalid_json')
  }
}

function sendJson(res: ServerResponse, status: number, body: unknown): void {
  const json = JSON.stringify(body)
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Content-Length': Buffer.byteLength(json) })
  res.end(json)
}

function sendApiResponse(res: ServerResponse, apiRes: ApiResponse): void {
  const body = typeof apiRes.body === 'string' ? apiRes.body : JSON.stringify(apiRes.body)
  const headers: Record<string, string> = {
    'Content-Type': typeof apiRes.body === 'string' ? 'text/plain; charset=utf-8' : 'application/json; charset=utf-8',
    'Content-Length': String(Buffer.byteLength(body)),
  }
  res.writeHead(apiRes.status, headers)
  res.end(body)
}
