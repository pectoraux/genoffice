#!/usr/bin/env node
/**
 * PROJECT-018 corrections — behavioral test fixture: a fake `unshare`.
 *
 * Stands in for the real isolation-wrapper executable so the launcher tests
 * can OBSERVE (not just infer) that the launcher actually wraps the sidecar
 * command. Behavior:
 *   - every invocation appends its argv (as one JSON line) to the file named
 *     by the UNSHARE_RECORD_PATH environment variable;
 *   - it exits 0 for the launcher's capability probe (argv = --net
 *     --map-root-user <node> -e 0) — reporting "mechanism supported";
 *   - for a conversion invocation it records the wrapped argv and exits 0
 *     WITHOUT running the real sidecar (so the launcher reports
 *     MPP_SIDECAR_RESPONSE_INVALID — the wrap form is what this fixture
 *     proves, not the conversion).
 */
import { appendFileSync } from 'node:fs'

const argv = process.argv.slice(2)
if (process.env.UNSHARE_RECORD_PATH) {
  appendFileSync(process.env.UNSHARE_RECORD_PATH, `${JSON.stringify(argv)}\n`)
}
process.exit(0)
