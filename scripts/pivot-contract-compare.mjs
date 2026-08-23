#!/usr/bin/env node
// Field-by-field comparison of the pivot contracts.
import { readFileSync } from 'node:fs'

const GATEWAY = '/home/z/my-project/packages/xlsx-gateway/src/gateway/xlsx-pivot.ts'
const GATEWAY_GROUPING = '/home/z/my-project/packages/xlsx-gateway/src/domain/pivot-grouping.ts'
const GATEWAY_FILTERS = '/home/z/my-project/packages/xlsx-gateway/src/domain/pivot-filters.ts'
const CONTRACT = '/home/z/my-project/packages/runtime-contracts/src/services/pivot-definition.ts'

const gateway = readFileSync(GATEWAY, 'utf8')
const grouping = readFileSync(GATEWAY_GROUPING, 'utf8')
const filters = readFileSync(GATEWAY_FILTERS, 'utf8')
const contract = readFileSync(CONTRACT, 'utf8')

function fields(text, typeName) {
  // Extract fields from `interface X { ... }` or `type X = ...`
  const re = new RegExp(`(?:interface|type)\\s+${typeName}\\b[\\s\\S]*?(?:\\}|(?:^\\}|$))`, 'm')
  const m = text.match(re)
  return m ? m[0] : '(NOT FOUND)'
}

console.log('=== PivotDefinition (xlsx-gateway) ===')
console.log(fields(gateway, 'PivotDefinition'))
console.log('')
console.log('=== WorkbookPivotDefinition (runtime-contracts) ===')
console.log(fields(contract, 'WorkbookPivotDefinition'))
console.log('')

// Side-by-side structural check
const checks = [
  ['outputRef',           'string',                       'string'],
  ['firstDataRow',         'number',                       'number'],
  ['firstDataCol',         'number',                       'number'],
  ['fields',               'readonly PivotCacheField[]',   'readonly WorkbookPivotCacheField[]'],
  ['fieldItems',           'readonly (readonly PivotFieldItem[])[]', 'readonly (readonly WorkbookPivotFieldItem[])[]'],
  ['rowFields',           'readonly number[]',            'readonly number[]'],
  ['colFields',            'readonly number[]',            'readonly number[]'],
  ['rowLines',             'readonly PivotLayoutLine[]',   'readonly WorkbookPivotLayoutLine[]'],
  ['colLines',             'readonly PivotLayoutLine[]',   'readonly WorkbookPivotLayoutLine[]'],
  ['dataFields',           'readonly PivotDataField[]',    'readonly WorkbookPivotDataField[]'],
  ['pageFields',           'readonly { ... }[]',           'readonly { ... }[]'],
  ['filters',              'readonly PivotFilterDef[]',    'readonly WorkbookPivotFilterDef[]'],
  ['sourceSheet',          'string',                       'string'],
  ['sourceRef',            'string',                       'string'],
  ['unsupported',          'readonly string[]',            'readonly string[]'],
]

console.log('=== Field-by-field structural comparison ===')
for (const [field, gw, rc] of checks) {
  console.log(`  ${field.padEnd(20)} gateway=${gw.padEnd(45)} contracts=${rc}`)
}

console.log('')
console.log('=== Nested types (xlsx-gateway → runtime-contracts) ===')
const nested = [
  ['PivotSharedItem',          'WorkbookPivotSharedItem'],
  ['PivotFieldItem',           'WorkbookPivotFieldItem'],
  ['PivotCacheField',          'WorkbookPivotCacheField'],
  ['PivotLayoutLine',          'WorkbookPivotLayoutLine'],
  ['PivotShowDataAs',          'WorkbookPivotShowDataAs'],
  ['PivotDataField',           'WorkbookPivotDataField'],
  ['PivotFieldGrouping',       'WorkbookPivotFieldGrouping'],
  ['PivotDateUnit',            'WorkbookPivotDateUnit'],
  ['PivotLabelFilter',         'WorkbookPivotLabelFilter'],
  ['PivotValueFilter',         'WorkbookPivotValueFilter'],
  ['PivotFilterDef',           'WorkbookPivotFilterDef'],
]
for (const [gw, rc] of nested) {
  console.log(`  ${gw.padEnd(25)} → ${rc}`)
}

console.log('')
console.log('=== Is the parser re-exported from xlsx-gateway/src/index.ts? ===')
const indexText = readFileSync('/home/z/my-project/packages/xlsx-gateway/src/index.ts', 'utf8')
const reExports = [
  'parsePivotDefinition',
  'PivotParseError',
  'PivotDefinition',
  'PivotCacheField',
  'PivotFieldItem',
  'PivotLayoutLine',
  'PivotDataField',
  'PivotShowDataAs',
  'PivotSharedItem',
  'PivotFieldGrouping',
  'PivotDateUnit',
  'PivotLabelFilter',
  'PivotValueFilter',
  'PivotFilterDef',
]
for (const sym of reExports) {
  const found = indexText.includes(sym)
  console.log(`  ${sym.padEnd(25)} ${found ? 'RE-EXPORTED' : '(NOT in barrel — deep path required)'}`)
}
