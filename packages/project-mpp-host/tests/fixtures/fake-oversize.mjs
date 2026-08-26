// Fake sidecar: OVERSIZED OUTPUT — writes a file larger than any sane cap.
import { writeFileSync } from 'node:fs'
const [, , , outputPath, requestId] = process.argv
writeFileSync(outputPath, 'x'.repeat(65_536))
console.log(
  JSON.stringify({
    version: 1,
    requestId,
    ok: true,
    counts: { tasks: 0, resources: 0, calendars: 0, predecessorLinks: 0, assignments: 0 },
  }),
)
