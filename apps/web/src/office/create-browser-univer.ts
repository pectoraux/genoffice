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

// Locale packs — every preset contributes UI strings. Without these, the
// Univer React UI (Ribbon, formula bar, sheet bar) throws
// "[LocaleService]: Locale not initialized" and never mounts.
import sheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import sheetsConditionalFormattingEnUS from '@univerjs/preset-sheets-conditional-formatting/locales/en-US'
import sheetsDrawingEnUS from '@univerjs/preset-sheets-drawing/locales/en-US'
import sheetsDataValidationEnUS from '@univerjs/preset-sheets-data-validation/locales/en-US'
import sheetsFilterEnUS from '@univerjs/preset-sheets-filter/locales/en-US'
import sheetsFindReplaceEnUS from '@univerjs/preset-sheets-find-replace/locales/en-US'
import sheetsNoteEnUS from '@univerjs/preset-sheets-note/locales/en-US'
import sheetsSortEnUS from '@univerjs/preset-sheets-sort/locales/en-US'
import sheetsTableEnUS from '@univerjs/preset-sheets-table/locales/en-US'

import '@univerjs/preset-sheets-core/lib/index.css'
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css'
import '@univerjs/preset-sheets-drawing/lib/index.css'
import '@univerjs/preset-sheets-data-validation/lib/index.css'
import '@univerjs/preset-sheets-filter/lib/index.css'
import '@univerjs/preset-sheets-find-replace/lib/index.css'
import '@univerjs/preset-sheets-note/lib/index.css'
import '@univerjs/preset-sheets-sort/lib/index.css'
import '@univerjs/preset-sheets-table/lib/index.css'

type PluginEntry =
  PluginCtor<Plugin> | [PluginCtor<Plugin>, ConstructorParameters<PluginCtor<Plugin>>[0]]
type BrowserPreset = { plugins: PluginEntry[] }

export interface BrowserUniverRuntime {
  readonly univer: Univer
  readonly univerAPI: FUniver
}

/** Merged en-US locale messages from every registered preset. */
const EN_US_LOCALES = {
  ...sheetsCoreEnUS,
  ...sheetsConditionalFormattingEnUS,
  ...sheetsDrawingEnUS,
  ...sheetsDataValidationEnUS,
  ...sheetsFilterEnUS,
  ...sheetsFindReplaceEnUS,
  ...sheetsNoteEnUS,
  ...sheetsSortEnUS,
  ...sheetsTableEnUS,
}

export function createBrowserUniver(container: string): BrowserUniverRuntime {
  const univer = new Univer({
    logLevel: LogLevel.WARN,
    theme: greenTheme,
    locale: LocaleType.EN_US,
    locales: { [LocaleType.EN_US]: EN_US_LOCALES },
  })
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
