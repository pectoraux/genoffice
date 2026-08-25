import type {
  ImportDiagnostic,
  ProjectDocument,
  ProjectFileMetadata,
  ProjectSavePlan,
} from '@genoffice/project-contracts'

export interface ProjectFileAdapter {
  readonly format: ProjectSavePlan['format']
  inspect(input: Uint8Array, metadata?: ProjectFileMetadata): ProjectFileMetadata
  import(
    input: Uint8Array,
    metadata?: ProjectFileMetadata,
  ): { document: ProjectDocument; diagnostics: ImportDiagnostic[] }
  export(document: ProjectDocument): { bytes: Uint8Array; diagnostics: ImportDiagnostic[] }
}
