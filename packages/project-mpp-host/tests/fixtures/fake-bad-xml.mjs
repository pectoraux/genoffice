// Fake sidecar: MALFORMED MSPDI — writes garbage XML and reports ok=true.
import { writeFileSync } from 'node:fs'
const [, , , outputPath, requestId] = process.argv
writeFileSync(outputPath, '<?xml version="1.0"?><Project><Unclosed')
console.log(
  JSON.stringify({
    version: 1,
    requestId,
    ok: true,
    counts: { tasks: 0, resources: 0, calendars: 0, predecessorLinks: 0, assignments: 0 },
  }),
)
