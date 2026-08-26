/**
 * PROJECT-018 corrections — OS-enforced network isolation for the JVM sidecar.
 *
 * The independent architect review rejected "the sidecar performs no network
 * access" as a purely static claim: untrusted MPP input requires the network
 * denial to be a RUNTIME invariant of the process boundary, not a property
 * of today's source code. This module provides that invariant on Linux:
 *
 *   unshare --net --map-root-user <command> <args…>
 *
 *   - `--map-root-user` creates a fresh USER namespace with the current user
 *     mapped to root — this is what makes unprivileged network-namespace
 *     creation possible (CAP_SYS_ADMIN inside the new namespace only; NO
 *     extra privileges over files: the kernel uid is unchanged, so the
 *     wrapped process keeps exactly the caller's filesystem access).
 *   - `--net` creates a fresh NETWORK namespace whose only interface is a
 *     DOWN loopback: no routes, no DNS, no connectable sockets — enforced
 *     by the kernel, impossible for the wrapped process (or any library it
 *     loads, MPXJ included) to bypass.
 *   - `unshare` executes the wrapped command IN PLACE (same PID), so the
 *     launcher's process management — signals, SIGTERM→SIGKILL escalation,
 *     exit-code propagation — is completely unchanged.
 *
 * The wrapper is a DIRECT argument array (no shell, ever). The mechanism is
 * probed once per executable (cached); under the launcher's default
 * `'required'` policy an unavailable mechanism FAILS CLOSED (the conversion
 * is refused, no sidecar process is started) — `'off'` is an explicit
 * operator opt-out. Runtime enforcement is proven by a real test: a process
 * launched through the wrapper cannot connect even to a live loopback
 * listener in the parent namespace, while the identical unwrapped control
 * connects (see tests/network-isolation.test.ts).
 *
 * Non-Linux hosts currently have no OS mechanism here: under the default
 * policy they fail closed; opting out there is an explicit, documented
 * operator decision (spec/project/requirements.md, PROJECT-018 security
 * model).
 */
import { spawn } from 'node:child_process'

/** Whether the sidecar must run inside the OS-enforced isolated context
 * ('required', the production default — fail closed when unavailable) or
 * the operator has explicitly opted out ('off', local development only). */
export type NetworkIsolationPolicy = 'required' | 'off'

/** The mechanism actually in effect. */
export type NetworkIsolationMechanism = 'linux-unshare-netns' | 'none'

/** Result of probing whether this host can create the isolated context. */
export interface NetworkIsolationCapability {
  readonly supported: boolean
  readonly mechanism: NetworkIsolationMechanism
  /** When unsupported: why (missing wrapper binary, restricted namespaces…). */
  readonly reason?: string
}

/** The isolation wrapper executable (util-linux unshare). */
export const DEFAULT_UNSHARE_EXECUTABLE = 'unshare'

/** Probe timeout: the wrapper probe is a trivial exec; anything slower than
 * this is treated as a broken/ hung mechanism (unsupported). */
export const PROBE_TIMEOUT_MS = 5_000

/** Spawn options contract for this module: always a direct argument array
 * with the shell explicitly disabled (mirroring the launcher contract; the
 * architecture test enforces the no-shell discipline here too). */
const ISOLATION_SPAWN_OPTIONS = { shell: false, windowsHide: true } as const

/**
 * Wrap a command so it runs inside a fresh, kernel-enforced network-isolated
 * context (Linux user + network namespace). Pure function — the returned
 * form is itself unit-tested, and the launcher applies it to whatever sidecar
 * command it built (including commandBuilder results).
 */
export function wrapNetworkIsolated(
  unshareExecutable: string,
  command: string,
  args: readonly string[],
): { command: string; args: string[] } {
  return {
    command: unshareExecutable,
    args: ['--net', '--map-root-user', command, ...args],
  }
}

const probeCache = new Map<string, Promise<NetworkIsolationCapability>>()

/**
 * Probe (once per executable path, cached for the process lifetime — host
 * capability does not change mid-process) whether the isolation wrapper can
 * actually create the isolated context here: spawn a trivial command
 * (`node -e 0`) through the wrapper and require a clean exit 0.
 */
export function probeNetworkIsolation(
  unshareExecutable: string = DEFAULT_UNSHARE_EXECUTABLE,
): Promise<NetworkIsolationCapability> {
  const cached = probeCache.get(unshareExecutable)
  if (cached) return cached
  const probe = runIsolationProbe(unshareExecutable)
  probeCache.set(unshareExecutable, probe)
  return probe
}

function runIsolationProbe(unshareExecutable: string): Promise<NetworkIsolationCapability> {
  return new Promise((resolve) => {
    const unsupported = (reason: string): NetworkIsolationCapability => ({
      supported: false,
      mechanism: 'none',
      reason,
    })
    let child: ReturnType<typeof spawn>
    try {
      child = spawn(unshareExecutable, ['--net', '--map-root-user', process.execPath, '-e', '0'], {
        stdio: ['ignore', 'ignore', 'ignore'],
        ...ISOLATION_SPAWN_OPTIONS,
      })
    } catch (error) {
      resolve(unsupported(`isolation wrapper could not be spawned: ${errorMessage(error)}`))
      return
    }
    let settled = false
    const finish = (capability: NetworkIsolationCapability) => {
      if (settled) return
      settled = true
      clearTimeout(timeout)
      resolve(capability)
    }
    const timeout = setTimeout(() => {
      child.kill('SIGKILL')
      finish(unsupported(`isolation wrapper probe exceeded ${PROBE_TIMEOUT_MS} ms`))
    }, PROBE_TIMEOUT_MS)
    child.once('error', (error) => {
      finish(unsupported(`isolation wrapper is unavailable: ${errorMessage(error)}`))
    })
    child.once('close', (code) => {
      if (code === 0) {
        finish({ supported: true, mechanism: 'linux-unshare-netns' })
        return
      }
      finish(
        unsupported(
          `isolation wrapper probe exited with code ${String(code)} (unprivileged user/network namespaces may be restricted on this host)`,
        ),
      )
    })
  })
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
