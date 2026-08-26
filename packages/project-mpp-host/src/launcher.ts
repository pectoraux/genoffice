/**
 * PROJECT-018 — MPXJ conversion sidecar launcher (host side).
 *
 * One-shot process model: every conversion spawns a FRESH JVM (java single-
 * file source-launcher mode — runs on a JRE, no javac required, exactly as
 * proven in the PROJECT-017 spike). A fresh process per file is the
 * strongest failure-containment posture: a hostile/corrupt MPP can never
 * poison the state of a long-lived sidecar, and memory limits
 * (`-Xmx`) apply per conversion. The JVM start cost (~1–3 s) is the price
 * of isolation; amortization (a resident sidecar pool) is explicitly
 * deferred to PROJECT-048 performance work.
 *
 * Security invariants (all enforced here, all tested):
 *   - DIRECT argument arrays: the MPP path/requestId are passed as argv
 *     entries, never interpolated into a command string; `shell` is never
 *     enabled. No shell injection surface exists.
 *   - OS-ENFORCED NETWORK ISOLATION (default policy 'required'): the JVM is
 *     wrapped in a fresh kernel network namespace (unshare --net
 *     --map-root-user — see network-isolation.ts); a host that cannot
 *     provide the mechanism FAILS CLOSED with
 *     MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE and no process is started.
 *     'off' is an explicit operator opt-out.
 *   - headless JVM (`-Djava.awt.headless=true`) and a hard `-Xmx` cap.
 *   - Wall-clock timeout with SIGTERM→SIGKILL escalation.
 *   - stdout/stderr accumulation caps.
 *   - Output-size cap checked before the bytes are handed onward.
 *   - The launcher NEVER writes to the input file and never touches any
 *     path other than the caller-provided input/output paths.
 */
import { spawn } from 'node:child_process'
import { existsSync, readFileSync, statSync, unlinkSync } from 'node:fs'
import { randomUUID } from 'node:crypto'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import {
  MPP_MAX_MSPDI_OUTPUT_BYTES,
  MPP_OUTPUT_TOO_LARGE,
  MPP_SIDECAR_EXIT,
  MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE,
  MPP_SIDECAR_RESPONSE_INVALID,
  MPP_SIDECAR_TIMEOUT,
  MPP_SIDECAR_UNAVAILABLE,
  MPP_UNSUPPORTED_FORMAT,
  type MppDiagnostic,
  type MppSidecarFrame,
} from '@genoffice/project-file'
import { parseSidecarFrame } from './protocol.js'
import {
  DEFAULT_UNSHARE_EXECUTABLE,
  probeNetworkIsolation,
  wrapNetworkIsolated,
  type NetworkIsolationPolicy,
} from './network-isolation.js'

/** Default conversion timeout (30 s — the PROJECT-017 spike converted every
 * corpus file in the sub-second range once the JVM was warm). */
export const MPP_DEFAULT_TIMEOUT_MS = 30_000

/** Default stderr accumulation cap (matches the xlsx-sidecar precedent). */
export const DEFAULT_MAX_STDERR_LENGTH = 8_192

/** Default JVM memory cap for the conversion process. */
export const DEFAULT_MAX_HEAP = '-Xmx512m'

export interface MppSidecarLauncherConfig {
  /** Directory of the extracted MPXJ distribution (contains mpxj.jar and
   * lib/). Managed by scripts/fetch-sidecar-deps.mjs; never committed. */
  readonly mpxjHome: string
  /** Absolute path of the Java sidecar source (single-file source-launcher
   * mode). Defaults to the in-repo java/MppSidecar.java. */
  readonly sidecarSource?: string
  /** Java executable. Defaults to `'java'` (PATH). */
  readonly javaExecutable?: string
  /** Extra JVM arguments (headless + heap defaults are always included
   * unless overridden here). */
  readonly javaArgs?: readonly string[]
  /** Conversion timeout in milliseconds (default 30 000). */
  readonly timeoutMs?: number
  /** stderr accumulation cap (default 8 192). */
  readonly maxStderrLength?: number
  /** Maximum MSPDI output size in bytes (default MPP_MAX_MSPDI_OUTPUT_BYTES). */
  readonly maxOutputBytes?: number
  /**
   * Test seam: replaces the default java command construction. Production
   * callers never set this; the failure-injection suite uses it to drive
   * the REAL process-management logic (timeouts, exit codes, frame
   * validation) with a deterministic stand-in executable. The command is
   * still spawned with a direct argument array and no shell — and still
   * wrapped by the network-isolation policy below.
   */
  readonly commandBuilder?: (
    inputPath: string,
    outputPath: string,
    requestId: string,
  ) => { command: string; args: readonly string[] }
  /** OS network-isolation policy for the sidecar process. DEFAULT
   * 'required': the conversion only runs inside the kernel-enforced
   * network-isolated context (Linux user + network namespace); a host
   * that cannot provide it fails closed with
   * MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE and no process is started.
   * 'off' is an explicit operator opt-out (local development only). */
  readonly networkIsolation?: NetworkIsolationPolicy
  /** The isolation wrapper executable (default 'unshare'); also the probe
   * target. Exposed as config so the behavioral isolation tests can point
   * it at a recording stand-in. */
  readonly unshareExecutable?: string
}

/** A successful conversion: the MSPDI bytes plus the validated frame. */
export interface MppConversionSuccess {
  readonly ok: true
  readonly mspdiBytes: Uint8Array
  readonly frame: MppSidecarFrame
}

/** A failed conversion: sidecar-stage error diagnostics (full provenance —
 * the caller maps them straight into the staged MppImportResult). */
export interface MppConversionFailure {
  readonly ok: false
  readonly diagnostics: MppDiagnostic[]
}

export type MppConversionResult = MppConversionSuccess | MppConversionFailure

/** Spawn options contract: the launcher always uses a direct argument array
 * with the shell explicitly disabled (the architecture test enforces that
 * this file never enables shell execution and never calls exec). */
export const SPAWN_OPTIONS = { shell: false, windowsHide: true } as const

/**
 * Build the sidecar command as a DIRECT argument array (the default java
 * invocation). Exported as a pure function so the argument form itself is
 * unit-tested: the input/output/request paths are argv entries, never
 * interpolated into a command string.
 */
export function buildSidecarCommand(
  config: MppSidecarLauncherConfig,
  inputPath: string,
  outputPath: string,
  requestId: string,
): { command: string; args: string[] } {
  const javaExecutable = config.javaExecutable ?? 'java'
  const javaArgs = config.javaArgs ?? ['-Djava.awt.headless=true', DEFAULT_MAX_HEAP]
  const sidecarSource = config.sidecarSource ?? defaultSidecarSource()
  const classpath = `${config.mpxjHome}/mpxj.jar:${config.mpxjHome}/lib/*`
  return {
    command: javaExecutable,
    args: [...javaArgs, '-cp', classpath, sidecarSource, inputPath, outputPath, requestId],
  }
}

/** Default sidecar source: the in-repo java/MppSidecar.java (resolved from
 * this module's location — no working-directory dependence). */
function defaultSidecarSource(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '..', 'java', 'MppSidecar.java')
}

export class MppSidecarLauncher {
  private readonly config: Required<
    Pick<
      MppSidecarLauncherConfig,
      | 'javaExecutable'
      | 'timeoutMs'
      | 'maxStderrLength'
      | 'maxOutputBytes'
      | 'networkIsolation'
      | 'unshareExecutable'
    >
  > &
    MppSidecarLauncherConfig

  constructor(config: MppSidecarLauncherConfig) {
    this.config = {
      javaExecutable: config.javaExecutable ?? 'java',
      timeoutMs: config.timeoutMs ?? MPP_DEFAULT_TIMEOUT_MS,
      maxStderrLength: config.maxStderrLength ?? DEFAULT_MAX_STDERR_LENGTH,
      maxOutputBytes: config.maxOutputBytes ?? MPP_MAX_MSPDI_OUTPUT_BYTES,
      networkIsolation: config.networkIsolation ?? 'required',
      unshareExecutable: config.unshareExecutable ?? DEFAULT_UNSHARE_EXECUTABLE,
      javaArgs: config.javaArgs,
      sidecarSource: config.sidecarSource,
      mpxjHome: config.mpxjHome,
      commandBuilder: config.commandBuilder,
    }
  }

  /** The configured network-isolation policy (exposed for tests asserting
   * the production default posture). */
  get networkIsolationPolicy(): NetworkIsolationPolicy {
    return this.config.networkIsolation
  }

  /**
   * Convert one MPP file to MSPDI. `inputPath` and `outputPath` are passed
   * as direct argv entries (no shell, no interpolation). `requestId`, when
   * omitted, is a fresh random UUID used purely for frame correlation.
   */
  convert(
    inputPath: string,
    outputPath: string,
    requestId: string = randomUUID(),
  ): Promise<MppConversionResult> {
    const built = this.config.commandBuilder
      ? this.config.commandBuilder(inputPath, outputPath, requestId)
      : buildSidecarCommand(this.config, inputPath, outputPath, requestId)
    return this.runIsolated(built, requestId, outputPath)
  }

  /** Apply the network-isolation policy to the built sidecar command, then
   * run it: 'required' (default) wraps the command in the kernel network
   * namespace and FAILS CLOSED — refusing the conversion, starting no
   * sidecar process — when the host cannot provide the mechanism; 'off'
   * runs the command unwrapped by explicit operator decision. */
  private async runIsolated(
    built: { command: string; args: readonly string[] },
    requestId: string,
    outputPath: string,
  ): Promise<MppConversionResult> {
    if (this.config.networkIsolation === 'off') {
      return this.run(built.command, built.args, requestId, outputPath, false)
    }
    const capability = await probeNetworkIsolation(this.config.unshareExecutable)
    if (!capability.supported) {
      return {
        ok: false,
        diagnostics: [
          sidecarError(
            MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE,
            `the sidecar requires OS-enforced network isolation, but this host cannot provide it (${capability.reason ?? 'unknown reason'}); the conversion failed closed — the sidecar was not started`,
          ),
        ],
      }
    }
    const wrapped = wrapNetworkIsolated(this.config.unshareExecutable, built.command, built.args)
    return this.run(wrapped.command, wrapped.args, requestId, outputPath, true)
  }

  /** Run the sidecar command (argument array) and interpret the outcome.
   * `wrapped` records whether the command is the isolation wrapper — under
   * the wrapper a missing/unstartable sidecar executable surfaces as the
   * wrapper's exit 127 + "failed to execute" stderr, which is mapped back
   * to the semantically-correct MPP_SIDECAR_UNAVAILABLE (never a misleading
   * "conversion exited nonzero"). */
  private run(
    command: string,
    args: readonly string[],
    requestId: string,
    outputPath: string,
    wrapped: boolean,
  ): Promise<MppConversionResult> {
    return new Promise((resolve) => {
      let child: ReturnType<typeof spawn>
      try {
        child = spawn(command, args as string[], {
          stdio: ['ignore', 'pipe', 'pipe'],
          ...SPAWN_OPTIONS,
        })
      } catch (error) {
        resolve({
          ok: false,
          diagnostics: [
            sidecarError(
              MPP_SIDECAR_UNAVAILABLE,
              `sidecar process could not be spawned: ${errorMessage(error)}`,
            ),
          ],
        })
        return
      }

      let stdout = ''
      let stderr = ''
      let stdoutOverflow = false
      let timedOut = false
      let settled = false

      const timeout = setTimeout(() => {
        timedOut = true
        child.kill('SIGTERM')
        // Escalate to SIGKILL after a grace period.
        setTimeout(() => {
          if (!child.killed) child.kill('SIGKILL')
        }, 2_000)
      }, this.config.timeoutMs)

      child.stdout?.setEncoding('utf8')
      child.stdout?.on('data', (chunk: string) => {
        if (stdout.length + chunk.length > 1_000_000) {
          stdoutOverflow = true
          return
        }
        stdout += chunk
      })
      child.stderr?.setEncoding('utf8')
      child.stderr?.on('data', (chunk: string) => {
        if (stderr.length + chunk.length > this.config.maxStderrLength) {
          stderr = `${stderr}${chunk}`.slice(-this.config.maxStderrLength)
          return
        }
        stderr += chunk
      })

      const finish = (result: MppConversionResult) => {
        if (settled) return
        settled = true
        clearTimeout(timeout)
        resolve(result)
      }

      child.once('error', (error) => {
        finish({
          ok: false,
          diagnostics: [
            sidecarError(
              MPP_SIDECAR_UNAVAILABLE,
              `sidecar process could not be started: ${errorMessage(error)}`,
            ),
          ],
        })
      })

      child.once('close', (code) => {
        if (timedOut) {
          finish({
            ok: false,
            diagnostics: [
              sidecarError(
                MPP_SIDECAR_TIMEOUT,
                `sidecar exceeded the ${this.config.timeoutMs} ms conversion timeout and was terminated`,
              ),
            ],
          })
          return
        }
        if (code !== 0) {
          // Under the isolation wrapper, a missing/unstartable sidecar
          // executable surfaces as the wrapper's exit 127 with a
          // "failed to execute" stderr — map it back to the precise
          // MPP_SIDECAR_UNAVAILABLE (the sidecar could not be started):
          const execFailure = wrapped && code === 127 && /failed to execute/i.test(stderr)
          finish({
            ok: false,
            diagnostics: [
              sidecarError(
                execFailure ? MPP_SIDECAR_UNAVAILABLE : MPP_SIDECAR_EXIT,
                `sidecar ${execFailure ? 'executable could not be started' : `exited with code ${String(code)}`}${stderr.trim() ? `: ${truncate(stderr.trim(), 512)}` : ''}`,
              ),
            ],
          })
          return
        }
        const frame = parseSidecarFrame(stdoutOverflow ? '' : stdout)
        if (frame === null || frame.requestId !== requestId) {
          finish({
            ok: false,
            diagnostics: [
              sidecarError(
                MPP_SIDECAR_RESPONSE_INVALID,
                `sidecar stdout is not a valid protocol frame${stdoutOverflow ? ' (stdout exceeded the accumulation cap)' : ''}: ${truncate(stdout, 512)}`,
              ),
            ],
          })
          return
        }
        if (!frame.ok) {
          const protocolCode = frame.error?.code ?? MPP_SIDECAR_RESPONSE_INVALID
          const message =
            frame.error?.message ?? 'sidecar reported a conversion failure without a message'
          finish({ ok: false, diagnostics: [sidecarError(protocolCode, message)] })
          return
        }
        if (!existsSync(outputPath)) {
          finish({
            ok: false,
            diagnostics: [
              sidecarError(
                MPP_SIDECAR_RESPONSE_INVALID,
                'sidecar reported success but wrote no MSPDI output file',
              ),
            ],
          })
          return
        }
        let mspdiBytes: Uint8Array
        try {
          const size = statSync(outputPath).size
          if (size > this.config.maxOutputBytes) {
            safeUnlink(outputPath)
            finish({
              ok: false,
              diagnostics: [
                sidecarError(
                  MPP_OUTPUT_TOO_LARGE,
                  `sidecar MSPDI output is ${size} bytes (limit ${this.config.maxOutputBytes})`,
                ),
              ],
            })
            return
          }
          mspdiBytes = readFileSync(outputPath)
        } catch (error) {
          finish({
            ok: false,
            diagnostics: [
              sidecarError(
                MPP_SIDECAR_RESPONSE_INVALID,
                `sidecar MSPDI output could not be read: ${errorMessage(error)}`,
              ),
            ],
          })
          return
        }
        finish({ ok: true, mspdiBytes, frame })
      })
    })
  }
}

/** Map a protocol error code to its foundation diagnostic name (the Java
 * sidecar uses the raw protocol spelling `UNSUPPORTED_MPP_FORMAT`; unknown
 * codes pass through verbatim — open set, full provenance). */
function sidecarError(code: string, message: string): MppDiagnostic {
  const normalized = code === 'UNSUPPORTED_MPP_FORMAT' ? MPP_UNSUPPORTED_FORMAT : code
  return {
    code: normalized,
    severity: 'error',
    message,
    stage: 'sidecar',
  }
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function truncate(text: string, max: number): string {
  return text.length <= max ? text : `${text.slice(0, max)}…`
}

function safeUnlink(path: string): void {
  try {
    unlinkSync(path)
  } catch {
    // best-effort cleanup — the caller's temp-dir sweep is authoritative
  }
}
