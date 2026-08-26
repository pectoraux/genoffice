/**
 * PROJECT-018 — corpus manifest verification (test item 27).
 *
 * Verifies the downloaded corpus against the in-repo manifest: every file
 * present, every SHA-256 matching, and the pinned MPXJ distribution in
 * place. If the sidecar dependencies are missing the test FAILS with a
 * pointer to the fetch script (the Project gate runs the fetch step before
 * the test step — zero skipped tests, zero silent degradation).
 */
import { createHash } from 'node:crypto'
import { existsSync, readFileSync, statSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const PKG_ROOT = join(import.meta.dirname, '..')
const DEPS = join(PKG_ROOT, '.sidecar-deps')

interface ManifestEntry {
  filename: string
  format: string
  sha256: string
  sourcePath: string
  provenance: string
  expectedNormalizations: Record<string, number>
  expectedErrors: number
  canonical: Record<string, unknown>
  coverage: string[]
}

interface Manifest {
  source: { repository: string; pinnedCommit: string; pinnedTag: string; license: string }
  mpxj: { version: string; url: string; sha256: string; license: string }
  corpus: ManifestEntry[]
}

const manifest: Manifest = JSON.parse(
  readFileSync(join(PKG_ROOT, 'corpus', 'corpus-manifest.json'), 'utf8'),
)

function sha256(path: string): string {
  return createHash('sha256').update(readFileSync(path)).digest('hex')
}

describe('corpus manifest verification', () => {
  it('the pinned MPXJ distribution is in place (mpxj.jar + lib/)', () => {
    const mpxjHome = join(DEPS, `mpxj-${manifest.mpxj.version}`)
    expect(existsSync(join(mpxjHome, 'mpxj.jar'))).toBe(true)
    expect(statSync(join(mpxjHome, 'mpxj.jar')).size).toBeGreaterThan(1_000_000)
    expect(existsSync(join(mpxjHome, 'lib'))).toBe(true)
  })

  it('every corpus file is present with the pinned SHA-256', () => {
    expect(existsSync(join(DEPS, 'corpus'))).toBe(true)
    for (const entry of manifest.corpus) {
      const path = join(DEPS, 'corpus', entry.filename)
      expect(
        existsSync(path),
        `${entry.filename} missing — run: npm run fetch-sidecar-deps -w @genoffice/project-mpp-host`,
      ).toBe(true)
      expect(sha256(path)).toBe(entry.sha256)
      // Every corpus file is a CFB container (the MPP storage format):
      const magic = readFileSync(path).subarray(0, 8)
      expect([...magic]).toEqual([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
    }
  })

  it('the manifest documents provenance and expected findings for every file', () => {
    for (const entry of manifest.corpus) {
      expect(entry.provenance.length).toBeGreaterThan(10)
      expect(entry.sourcePath).toMatch(/^junit\/data\//)
      expect(Object.keys(entry.expectedNormalizations).sort()).toEqual([
        'N1',
        'N2',
        'N3',
        'N4',
        'N5',
      ])
      expect(entry.coverage.length).toBeGreaterThan(0)
      expect(typeof entry.expectedErrors).toBe('number')
    }
  })

  it('the corpus covers all four supported MPP format versions with real files', () => {
    const formats = manifest.corpus.map((entry) => entry.format)
    expect(formats.some((f) => f.startsWith('MPP8'))).toBe(true)
    expect(formats.filter((f) => f.startsWith('MPP9')).length).toBeGreaterThanOrEqual(2)
    expect(formats.some((f) => f.startsWith('MPP12'))).toBe(true)
    expect(formats.filter((f) => f.startsWith('MPP14')).length).toBeGreaterThanOrEqual(2)
  })

  it('the manifest pins the MPXJ distribution URL and license', () => {
    expect(manifest.mpxj.url).toBe(
      'https://github.com/joniles/mpxj/releases/download/v16.7.0/mpxj-16.7.0.zip',
    )
    expect(manifest.mpxj.sha256).toBe(
      '2a149f3ae9f6ac1034d0b0fadeca8f98e3f1f900764640c6bd16188ec2f078bd',
    )
    expect(manifest.mpxj.license).toContain('LGPL')
    expect(manifest.source.pinnedCommit).toBe('abdbf6ef85654e3eff35c11c5e76cf08da842dce')
  })
})
