// Fake sidecar: REQUEST-ID MISMATCH — echoes a different correlation id.
import { writeFileSync } from 'node:fs'
const [, , , outputPath] = process.argv
writeFileSync(outputPath, '<Project/>')
console.log(JSON.stringify({ version: 1, requestId: 'not-the-one-you-sent', ok: true }))
