/**
 * PROJECT-018 — architecture/security discipline suite.
 *
 * Static guards over the increment's own surfaces:
 *   - the launcher and the network-isolation module never enable a shell
 *     and never use exec-style calls;
 *   - the host library never imports renderer/Electron surfaces;
 *   - the foundation package (project-file) never references the host
 *     package (dependency direction is host → foundation only);
 *   - the Java sidecar performs no application-level network access and
 *     executes no subprocess of its own (the OS-level network denial is
 *     enforced by the launcher's wrapper, tested in
 *     network-isolation.test.ts);
 *   - the enforced network-isolation semantics stay documented in the
 *     specification (requirements/implementation lockstep);
 *   - licensing artifacts and version pins are present and consistent.
 */
import { readFileSync, existsSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { MPXJ_PINNED_VERSION, MPP_SIDECAR_PROTOCOL_VERSION } from '@genoffice/project-file'

const PKG_ROOT = join(import.meta.dirname, '..')
const REPO_ROOT = join(PKG_ROOT, '..', '..')

function read(rel: string): string {
  return readFileSync(join(PKG_ROOT, rel), 'utf8')
}

describe('launcher process discipline', () => {
  it('the launcher source never enables a shell and never uses exec()', () => {
    const source = read('src/launcher.ts')
    expect(source).not.toContain('shell: true')
    expect(source).not.toMatch(/\bexec(?:Sync|File)?\s*\(/)
    expect(source).not.toContain('execSync')
    // spawn is used with the shared SPAWN_OPTIONS contract (shell: false):
    expect(source).toContain('shell: false')
  })

  it('the network-isolation source never enables a shell and never uses exec()', () => {
    const source = read('src/network-isolation.ts')
    expect(source).not.toContain('shell: true')
    expect(source).not.toMatch(/\bexec(?:Sync|File)?\s*\(/)
    expect(source).not.toContain('execSync')
    expect(source).toContain('shell: false')
  })

  it('the import pipeline source never spawns processes itself', () => {
    const source = read('src/import-mpp.ts')
    expect(source).not.toMatch(/spawn|child_process/)
  })
})

describe('host-library boundaries', () => {
  it('src/ contains no renderer/Electron imports', () => {
    for (const file of [
      'src/index.ts',
      'src/launcher.ts',
      'src/protocol.ts',
      'src/import-mpp.ts',
      'src/network-isolation.ts',
    ]) {
      const source = read(file)
      expect(source).not.toMatch(
        /from ['"](react|electron|@genoffice\/ui|@genoffice\/renderer-bridge)/,
      )
    }
  })

  it('the foundation package never references the host package (direction: host → foundation)', () => {
    const foundationIndex = readFileSync(
      join(REPO_ROOT, 'packages/project-file/src/index.ts'),
      'utf8',
    )
    expect(foundationIndex).not.toContain('project-mpp-host')
    const foundationMpp = readFileSync(
      join(REPO_ROOT, 'packages/project-file/src/mpp/index.ts'),
      'utf8',
    )
    expect(foundationMpp).not.toContain('project-mpp-host')
    // And the foundation mpp surface carries no process imports at all
    // (also enforced by the project-foundation.yml boundary grep):
    for (const file of ['contract.ts', 'normalize.ts', 'types.ts', 'diagnostics.ts', 'index.ts']) {
      const source = readFileSync(join(REPO_ROOT, 'packages/project-file/src/mpp', file), 'utf8')
      expect(source).not.toMatch(/from\s+['"]node:|child_process/)
    }
  })

  it('the Java sidecar performs no application-level network access and spawns no subprocess', () => {
    const source = read('java/MppSidecar.java')
    expect(source).not.toMatch(
      /java\.net|HttpURLConnection|Socket|Runtime\.getRuntime|ProcessBuilder/,
    )
    // The source-level claim is explicitly framed as APPLICATION-level; the
    // OS-level denial is the launcher's job (behaviorally proven in
    // network-isolation.test.ts):
    expect(source).toContain('no network access at the application level')
  })
})

describe('specification / implementation lockstep', () => {
  it('requirements.md documents the enforced network-isolation semantics (Option A, not a known limitation)', () => {
    const requirements = readFileSync(join(REPO_ROOT, 'spec/project/requirements.md'), 'utf8')
    const start = requirements.indexOf('## PROJECT-018')
    expect(start).toBeGreaterThan(-1)
    const section = requirements.slice(
      start,
      requirements.indexOf('\n## ', start + 1) === -1
        ? undefined
        : requirements.indexOf('\n## ', start + 1),
    )
    expect(section).toContain('network namespace')
    expect(section).toContain("'required'")
    expect(section).toContain("'off'")
    expect(section).toContain('MPP_SIDECAR_NETWORK_ISOLATION_UNAVAILABLE')
    expect(section).toContain('MPP_INPUT_UNREADABLE')
    // The old static-only claim must be GONE from the security model:
    expect(section).not.toContain('is NOT enforced and is recorded as a known limitation')
  })
})

describe('licensing and version pinning', () => {
  it('LICENSE-THIRD-PARTY.md exists and documents the LGPL posture', () => {
    const text = read('LICENSE-THIRD-PARTY.md')
    expect(text).toContain('GNU Lesser General Public License')
    expect(text).toContain('16.7.0')
    expect(text).toContain('2a149f3ae9f6ac1034d0b0fadeca8f98e3f1f900764640c6bd16188ec2f078bd')
  })

  it('the corpus manifest pins the same MPXJ version as the foundation constant', () => {
    const manifest = JSON.parse(read('corpus/corpus-manifest.json')) as {
      mpxj: { version: string; sha256: string }
      source: { pinnedCommit: string }
      corpus: Array<{ filename: string; format: string; sha256: string }>
    }
    expect(manifest.mpxj.version).toBe(MPXJ_PINNED_VERSION)
    expect(manifest.source.pinnedCommit).toBe('abdbf6ef85654e3eff35c11c5e76cf08da842dce')
    // All four supported format families are represented by real files:
    for (const version of ['MPP8', 'MPP9', 'MPP12', 'MPP14']) {
      expect(manifest.corpus.some((entry) => entry.format.startsWith(version))).toBe(true)
    }
    // Every corpus entry carries a SHA-256:
    expect(manifest.corpus.every((entry) => /^[0-9a-f]{64}$/.test(entry.sha256))).toBe(true)
  })

  it('the sidecar-deps workspace is gitignored (LGPL artifacts stay external)', () => {
    const gitignore = read('.gitignore')
    expect(gitignore).toContain('.sidecar-deps/')
  })

  it('the protocol version is pinned at 1 (foundation and Java sidecar agree)', () => {
    expect(MPP_SIDECAR_PROTOCOL_VERSION).toBe(1)
    const java = read('java/MppSidecar.java')
    expect(java).toContain('protocol v1')
    expect(java).toContain('System.out.println(sb.toString())')
  })

  it('the fetch script exists and verifies checksums', () => {
    expect(existsSync(join(PKG_ROOT, 'scripts/fetch-sidecar-deps.mjs'))).toBe(true)
    const script = read('scripts/fetch-sidecar-deps.mjs')
    expect(script).toContain('sha256')
    expect(script).toContain('checksum mismatch')
  })
})
