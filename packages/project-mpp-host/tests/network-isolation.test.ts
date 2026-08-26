/**
 * PROJECT-018 corrections — OS-enforced network isolation evidence.
 *
 * The independent architect review rejected "the sidecar performs no network
 * access" as a purely static claim and required a runtime invariant. This
 * suite provides the evidence chain for the implemented mechanism (Linux
 * user + network namespace via `unshare --net --map-root-user`, direct
 * argv, fail-closed 'required' default, explicit 'off' opt-out):
 *
 *   1. wrapper form — the wrapped command is a direct argument array
 *      (`unshare --net --map-root-user <command> <args…>`);
 *   2. REAL runtime enforcement — a process launched through the wrapper
 *      cannot connect even to a LIVE loopback listener in the parent
 *      network namespace, while the identical unwrapped control CAN
 *      (proving both the isolation and the validity of the test apparatus);
 *   3. fail-closed — the default 'required' policy + an unavailable
 *      mechanism yields the deterministic
 *      MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE refusal, atomically, with
 *      no sidecar process started;
 *   4. explicit opt-out — policy 'off' runs without the mechanism;
 *   5. behavioral wrap proof — a recording fake `unshare` shows the
 *      launcher actually probes with, and wraps the sidecar command in,
 *      the configured wrapper executable;
 *   6. the production default policy is 'required'.
 */
import { spawn } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import net from 'node:net'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  MppSidecarLauncher,
  probeNetworkIsolation,
  wrapNetworkIsolated,
  importMppFromBytes,
} from '../src/index.js'

import {
  MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE,
  MPP_SIDECAR_RESPONSE_INVALID,
  emptyProjectDocument,
} from '@genoffice/project-file'

const FIXTURES = join(import.meta.dirname, 'fixtures')
let workspace: string

beforeEach(() => {
  workspace = mkdtempSync(join(tmpdir(), 'netiso-test-'))
})
afterEach(() => {
  rmSync(workspace, { recursive: true, force: true })
  delete process.env.UNSHARE_RECORD_PATH
})

/** Spawn a direct argument array (no shell) and collect stdout. */
function runCapture(
  command: string,
  args: readonly string[],
): Promise<{ stdout: string; code: number | null }> {
  return new Promise((resolve) => {
    const child = spawn(command, args as string[], {
      stdio: ['ignore', 'pipe', 'pipe'],
      shell: false,
      windowsHide: true,
    })
    let stdout = ''
    child.stdout?.setEncoding('utf8')
    child.stdout?.on('data', (chunk: string) => {
      stdout += chunk
    })
    child.once('close', (code) => resolve({ stdout: stdout.trim(), code }))
    child.once('error', () => resolve({ stdout: '', code: null }))
  })
}

describe('wrapper form', () => {
  it('wrapNetworkIsolated produces the direct-argv isolation wrapper', () => {
    const wrapped = wrapNetworkIsolated('/usr/bin/unshare', 'java', [
      '-Djava.awt.headless=true',
      '-Xmx512m',
      '-cp',
      'mpxj.jar:lib/*',
      'MppSidecar.java',
      'in.mpp',
      'out.mspdi',
      'req-1',
    ])
    // Direct argument array: the whole original command survives as argv
    // entries — nothing is interpolated into a shell string:
    expect(wrapped.command).toBe('/usr/bin/unshare')
    expect(wrapped.args).toEqual([
      '--net',
      '--map-root-user',
      'java',
      '-Djava.awt.headless=true',
      '-Xmx512m',
      '-cp',
      'mpxj.jar:lib/*',
      'MppSidecar.java',
      'in.mpp',
      'out.mspdi',
      'req-1',
    ])
  })
})

describe('runtime enforcement (real network namespace)', () => {
  it('a process launched through the wrapper cannot connect; the unwrapped control can', async () => {
    const capability = await probeNetworkIsolation()
    if (!capability.supported) {
      // This host cannot provide the mechanism at all — the honest branch:
      // assert the capability report AND that the default REQUIRED policy
      // fails closed on exactly this real condition (the deterministic
      // missing-binary variant is proven separately below).
      expect(capability.mechanism).toBe('none')
      expect(capability.reason).toBeTruthy()
      const result = await new MppSidecarLauncher({
        mpxjHome: '/nonexistent',
      }).convert(join(workspace, 'in.mpp'), join(workspace, 'out.xml'))
      expect(result.ok).toBe(false)
      if (!result.ok) {
        expect(result.diagnostics[0].code).toBe(MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE)
      }
      return
    }
    expect(capability.mechanism).toBe('linux-unshare-netns')

    // A LIVE listener on loopback in THIS (parent) network namespace:
    const server = net.createServer()
    await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve))
    const address = server.address() as net.AddressInfo
    const probeScript =
      'import net from "node:net"; ' +
      `const s = net.connect(${address.port}, "127.0.0.1"); ` +
      's.on("connect", () => { console.log("CONNECTED"); process.exit(0); }); ' +
      's.on("error", (e) => { console.log("CONNECT_FAILED:" + e.code); process.exit(0); }); ' +
      'setTimeout(() => { console.log("PROBE_TIMEOUT"); process.exit(1); }, 8000);'
    try {
      // CONTROL — identical probe, NO isolation wrapper: must connect. This
      // proves the listener is live and the probe is a valid connectivity
      // test (the only variable left is the wrapper itself):
      const control = await runCapture(process.execPath, ['-e', probeScript])
      expect(control.stdout).toBe('CONNECTED')

      // ISOLATED — identical probe through the real wrapper: must FAIL
      // (kernel-enforced: the fresh namespace's only interface is a down
      // loopback; even 127.0.0.1 of the parent namespace is unreachable):
      const wrapped = wrapNetworkIsolated('unshare', process.execPath, ['-e', probeScript])
      const isolated = await runCapture(wrapped.command, wrapped.args)
      expect(isolated.stdout).toMatch(/^CONNECT_FAILED:/)
      expect(isolated.stdout).not.toContain('CONNECTED')
    } finally {
      server.close()
    }
  })
})

describe('fail-closed policy (deterministic on every host)', () => {
  it("the DEFAULT policy is 'required'", () => {
    const launcher = new MppSidecarLauncher({ mpxjHome: '/nonexistent' })
    expect(launcher.networkIsolationPolicy).toBe('required')
  })

  it('MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE: required policy + missing mechanism fails closed', async () => {
    const launcher = new MppSidecarLauncher({
      mpxjHome: '/nonexistent',
      unshareExecutable: '/definitely/not/a/real/unshare',
    })
    const result = await launcher.convert(join(workspace, 'in.mpp'), join(workspace, 'out.xml'))
    expect(result.ok).toBe(false)
    if (!result.ok) {
      expect(result.diagnostics).toHaveLength(1)
      expect(result.diagnostics[0].code).toBe(MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE)
      expect(result.diagnostics[0].severity).toBe('error')
      expect(result.diagnostics[0].stage).toBe('sidecar')
      expect(result.diagnostics[0].message).toContain('failed closed')
    }
  })

  it('the fail-closed refusal is atomic: empty document, sidecar stage only, no sidecar ran', async () => {
    // The fake sidecar would SUCCEED if it were ever spawned; the observed
    // isolation refusal therefore proves it never ran.
    const launcher = new MppSidecarLauncher({
      mpxjHome: '/nonexistent',
      unshareExecutable: '/definitely/not/a/real/unshare',
      commandBuilder: (inputPath, outputPath, requestId) => ({
        command: process.execPath,
        args: [join(FIXTURES, 'fake-ok.mjs'), inputPath, outputPath, requestId],
      }),
    })
    const result = await importMppFromBytes(new Uint8Array([1, 2, 3]), { launcher })
    expect(result.document).toEqual(emptyProjectDocument())
    expect(result.schedule).toBeUndefined()
    expect(result.diagnostics).toHaveLength(1)
    expect(result.diagnostics[0].code).toBe(MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE)
    expect(result.diagnostics.every((d) => d.stage === 'sidecar')).toBe(true)
  })

  it("the explicit 'off' opt-out runs without the mechanism", async () => {
    const launcher = new MppSidecarLauncher({
      mpxjHome: '/nonexistent',
      networkIsolation: 'off',
      unshareExecutable: '/definitely/not/a/real/unshare',
      commandBuilder: (inputPath, outputPath, requestId) => ({
        command: process.execPath,
        args: [join(FIXTURES, 'fake-ok.mjs'), inputPath, outputPath, requestId],
      }),
    })
    const result = await launcher.convert(join(workspace, 'in.mpp'), join(workspace, 'out.xml'))
    expect(result.ok).toBe(true)
  })
})

describe('behavioral wrap proof (recording fake unshare)', () => {
  it('the launcher probes with, and wraps the sidecar command in, the configured wrapper executable', async () => {
    const record = join(workspace, 'unshare-argv.jsonl')
    process.env.UNSHARE_RECORD_PATH = record
    const fakeUnshare = join(FIXTURES, 'fake-unshare.mjs')
    let builtCommand: { command: string; args: readonly string[] } | undefined
    const launcher = new MppSidecarLauncher({
      mpxjHome: '/nonexistent',
      unshareExecutable: fakeUnshare,
      commandBuilder: (inputPath, outputPath, requestId) => {
        builtCommand = {
          command: process.execPath,
          args: [join(FIXTURES, 'fake-ok.mjs'), inputPath, outputPath, requestId],
        }
        return builtCommand
      },
    })
    const result = await launcher.convert(
      join(workspace, 'in.mpp'),
      join(workspace, 'out.xml'),
      'req-iso-1',
    )
    // The fake unshare exits 0 without running the real sidecar, so the
    // launcher correctly reports an invalid frame — the WRAP FORM is what
    // this test observes:
    expect(result.ok).toBe(false)
    if (!result.ok) {
      expect(result.diagnostics[0].code).toBe(MPP_SIDECAR_RESPONSE_INVALID)
    }
    const lines = readLines(record)
    expect(lines).toHaveLength(2)
    // Invocation 1 — the capability probe, through the wrapper executable:
    expect(lines[0]).toEqual(['--net', '--map-root-user', process.execPath, '-e', '0'])
    // Invocation 2 — the conversion command, WRAPPED as a direct argv:
    expect(lines[1]).toEqual([
      '--net',
      '--map-root-user',
      builtCommand!.command,
      ...builtCommand!.args,
    ])
  })
})

function readLines(path: string): string[][] {
  const text = readFileSync(path, 'utf8')
  return text
    .split('\n')
    .filter((line) => line.trim() !== '')
    .map((line) => JSON.parse(line) as string[])
}
