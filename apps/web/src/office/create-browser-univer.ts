import { LogLevel, LocaleType, Univer } from '@univerjs/core'
import type { Plugin, PluginCtor } from '@univerjs/core'
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

type PluginEntry = PluginCtor<Plugin> | [PluginCtor<Plugin>, ConstructorParameters<PluginCtor<Plugin>>[0]]
type BrowserPreset = { plugins: PluginEntry[] }

export interface BrowserUniverRuntime {
  readonly univer: Univer
  readonly univerAPI: FUniver
}

export function createBrowserUniver(container: string): BrowserUniverRuntime {
  const univer = new Univer({ logLevel: LogLevel.WARN, theme: greenTheme, locale: LocaleType.EN_US })
  const presets: BrowserPreset[] = [
    UniverSheetsCorePreset({
      container,
      header: true,
      toolbar: true,
      contextMenu: true,
      formulaBar: true,
      footer: { sheetBar: true, statisticBar: true, menus: true, zoomSlider: true },
    }),
    UniverSheetsDrawingPreset(),
    UniverSheetsConditionalFormattingPreset(),
    UniverSheetsFilterPreset(),
    UniverSheetsDataValidationPreset(),
    UniverSheetsNotePreset(),
    UniverSheetsFindReplacePreset(),
    UniverSheetsSortPreset(),
    UniverSheetsTablePreset(),
  ]
  const registry = new Map<string, { plugin: PluginCtor<Plugin>; options: unknown }>()
  for (const preset of presets) {
    for (const entry of preset.plugins) {
      const [plugin, options] = Array.isArray(entry) ? entry : [entry, undefined]
      registry.delete(plugin.pluginName)
      registry.set(plugin.pluginName, { plugin, options })
    }
  }
  for (const { plugin, options } of registry.values()) univer.registerPlugin(plugin, options)
  return { univer, univerAPI: FUniver.newAPI(univer) }
}
