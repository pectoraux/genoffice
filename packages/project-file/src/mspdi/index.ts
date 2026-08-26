/**
 * PROJECT-015 / PROJECT-016 — MSPDI public surface.
 *
 * Exposed alongside the native `.gproj` adapter (PROJECT-014). The MSPDI
 * adapter reuses the same `ProjectFileAdapter` boundary, `ImportDiagnostic`
 * contract, brand-promotion helpers, and `validateProjectDocument` pipeline —
 * it does NOT introduce a second file-adapter abstraction. It exposes
 * `inspect` + `import` (PROJECT-015) and `export` (PROJECT-016 — canonical
 * `ProjectDocument` → deterministic MSPDI XML).
 */
export {
  mspdiFileAdapter,
  importMspdi,
  inspectMspdi,
  type MspdiFileAdapter,
  type MspdiImportResult,
} from './importer.js'
export { exportMspdi, type MspdiExportResult } from './exporter.js'
export {
  MSPDI_FORMAT,
  MSPDI_FORMAT_VERSION,
  MSPDI_SUPPORTED_SAVE_VERSIONS,
  MSPDI_ROOT_ELEMENT,
  MSPDI_NS,
  isSupportedMspdiSaveVersion,
} from './envelope.js'
export {
  MSPDI_MAX_INPUT_BYTES,
  MSPDI_MAX_PARSE_DEPTH,
  MSPDI_MAX_DECODED_BYTES,
  parseXml,
  decodeEntities,
  XmlParseError,
  type XmlNode,
  type XmlAttribute,
  childrenNamed,
  firstChild,
  childText,
} from './xml-parser.js'
export {
  uidToTaskId,
  uidToResourceId,
  uidToAssignmentId,
  uidToCalendarId,
  baselineIndexToId,
  dependencyId,
} from './identity.js'
export {
  isoDurationToMinutes,
  lagToMinutes,
  normalizeMspdiDate,
  isValidExceptionDate,
  mspdiTimeToMinutes,
  DEFAULT_LAG_FACTORS,
  LAG_FORMAT_MINUTE,
  LAG_FORMAT_HOUR,
  LAG_FORMAT_DAY,
  LAG_FORMAT_WEEK,
  LAG_FORMAT_MONTH,
  type LagFactors,
} from './conversions.js'
export {
  INVALID_MSPDI,
  UNSUPPORTED_MSPDI_VERSION,
  UNSUPPORTED_MSPDI_FEATURE,
  INVALID_MSPDI_REFERENCE,
  INVALID_MSPDI_DATE,
  INVALID_MSPDI_DURATION,
  INVALID_MSPDI_CALENDAR,
  INVALID_MSPDI_RESOURCE,
  INVALID_MSPDI_ASSIGNMENT,
  INVALID_MSPDI_CONSTRAINT,
  MISSING_MSPDI_FIELD,
  MSPDI_READ,
  MSPDI_DIAGNOSTIC_CODES,
} from './diagnostics.js'
export {
  INVALID_MSPDI_EXPORT,
  INVALID_MSPDI_EXPORT_LAG,
  UNREPRESENTABLE_MSPDI_VALUE,
  UNSUPPORTED_MSPDI_EXPORT_FEATURE,
  MSPDI_EXPORT_NORMALIZED,
  MSPDI_WRITTEN,
  MSPDI_EXPORT_DIAGNOSTIC_CODES,
} from './diagnostics.js'
export { escapeXmlText, escapeXmlAttribute, isValidXmlName, XmlWriter } from './xml-writer.js'
