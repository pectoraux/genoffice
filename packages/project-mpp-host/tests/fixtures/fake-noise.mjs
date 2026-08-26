// Fake sidecar: JVM-STYLE NOISE — prints banner lines before the ok frame.
import { writeFileSync } from 'node:fs'
const [, , , outputPath, requestId] = process.argv
writeFileSync(outputPath, '<Project/>')
console.log('WARNING: sun.misc.Unsafe::objectFieldOffset is deprecated')
console.log('Log4j API could not find a logging provider.')
console.log(JSON.stringify({ version: 1, requestId, ok: true }))
