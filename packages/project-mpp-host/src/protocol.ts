/**
 * PROJECT-018 — MPXJ sidecar wire protocol (host side).
 *
 * Protocol v1 (one-shot conversion, per the PROJECT-017 spike):
 *
 *   invocation (argument array — NEVER a shell string):
 *     java [-Djava.awt.headless=true] [-Xmx…] -cp <mpxj.jar:lib/*> \
 *          MppSidecar.java <input.mpp> <output.mspdi> <requestId>
 *
 *   stdout:  EXACTLY ONE JSON status line (the frame). The MSPDI payload
 *            itself travels ONLY via the output file, so no arbitrary
 *            sidecar stdout can ever contaminate MSPDI parsing. When the
 *            JVM prints unexpected leading/trailing noise (e.g. a banner),
 *            the LAST non-empty line is authoritative.
 *   stderr:  diagnostic noise (Log4j provider notice, stack traces) —
 *            never part of the protocol; capped and attached to failure
 *            diagnostics only.
 *   exit 0:  the frame is the truth (ok true, or ok false with a
 *            protocol error code such as MPP_UNSUPPORTED_FORMAT).
 *   exit ≠0: unexpected failure (crash/IO) — the frame, if any, was not
 *            produced; the launcher reports MPP_SIDECAR_EXIT with the
 *            stderr tail.
 *
 * The frame validator is PURE — no I/O, no `as` casts on untrusted data.
 */
import {
  MPP_SIDECAR_PROTOCOL_VERSION,
  type MppSidecarCounts,
  type MppSidecarFrame,
} from '@genoffice/project-file'

/**
 * Parse and runtime-validate the sidecar status frame from raw stdout.
 * Returns `null` when no valid frame is present (the caller reports
 * `MPP_SIDECAR_RESPONSE_INVALID`).
 */
export function parseSidecarFrame(stdoutText: string): MppSidecarFrame | null {
  const lines = stdoutText.split('\n').filter((line) => line.trim() !== '')
  if (lines.length === 0) return null
  // The LAST non-empty line is authoritative (tolerates leading JVM noise).
  const last = lines[lines.length - 1]
  let parsed: unknown
  try {
    parsed = JSON.parse(last)
  } catch {
    return null
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) return null
  const obj = parsed as Record<string, unknown>

  if (obj.version !== MPP_SIDECAR_PROTOCOL_VERSION) return null
  if (typeof obj.requestId !== 'string') return null
  if (typeof obj.ok !== 'boolean') return null

  let format: string | undefined
  if (obj.format !== undefined) {
    if (typeof obj.format !== 'string') return null
    format = obj.format
  }

  let counts: MppSidecarCounts | undefined
  if (obj.counts !== undefined) {
    if (typeof obj.counts !== 'object' || obj.counts === null) return null
    const c = obj.counts as Record<string, unknown>
    for (const key of ['tasks', 'resources', 'calendars', 'predecessorLinks', 'assignments']) {
      if (typeof c[key] !== 'number' || !Number.isFinite(c[key])) return null
    }
    counts = {
      tasks: c.tasks as number,
      resources: c.resources as number,
      calendars: c.calendars as number,
      predecessorLinks: c.predecessorLinks as number,
      assignments: c.assignments as number,
    }
  }

  let error: { code: string; message: string } | undefined
  if (obj.error !== undefined) {
    if (typeof obj.error !== 'object' || obj.error === null) return null
    const e = obj.error as Record<string, unknown>
    if (typeof e.code !== 'string' || typeof e.message !== 'string') return null
    error = { code: e.code, message: e.message }
  }

  return {
    version: obj.version,
    requestId: obj.requestId,
    ok: obj.ok,
    format,
    counts,
    error,
  }
}
