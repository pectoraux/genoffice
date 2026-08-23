import { LocaleType, Univer } from '@univerjs/core'
import { FUniver } from '@univerjs/core/lib/facade'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting'
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing'
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation'
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter'
import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace'
import { UniverSheetsNotePreset } from '@univerjs/preset-sheets-note'
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort'
import { UniverSheetsTablePreset } from '@univerjs/preset-sheets-table'
import { greenTheme } from '@univerjs/themes'

import '@univerjs/preset-sheets-core/lib/index.css'
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css'
import '@univerjs/preset-sheets-drawing/lib/index.css'
import '@univerjs/preset-sheets-data-validation/lib/index.css'
import '@univerjs/preset-sheets-filter/lib/index.css'
import '@univerjs/preset-sheets-find-replace/lib/index.css'
import '@univerjs/preset-sheets-note/lib/index.css'
import '@univerjs/preset-sheets-sort/lib/index.css'
import '@univerjs/preset-sheets-table/lib/index.css'

export interface BrowserUniverRuntime {
  readonly univer: Univer
  readonly univerAPI: FUniver
}

export function createBrowserUniver(container: string): BrowserUniverRuntime {
  const univer = new Univer({
    theme: greenTheme,
    locale: LocaleType.EN_US,
  })
  univer.registerPlugin(
    UniverSheetsCorePreset,
    {
      container,
      header: true,
      toolbar: true,
      contextMenu: true,
      formulaBar: true,
      footer: {
        sheetBar: true,
        statisticBar: true,
        menus: true,
        zoomSlider: true,
      },
    },
  )
  univer.registerPlugin(UniverSheetsDrawingPreset)
  univer.registerPlugin(UniverSheetsConditionalFormattingPreset)
  univer.registerPlugin(UniverSheetsFilterPreset)
  univer.registerPlugin(UniverSheetsDataValidationPreset)
  univer.registerPlugin(UniverSheetsNotePreset)
  univer.registerPlugin(UniverSheetsFindReplacePreset)
  univer.registerPlugin(UniverSheetsSortPreset)
  univer.registerPlugin(UniverSheetsTablePreset)

  return { univer, univerAPI: FUniver.newAPI(univer) }
}
