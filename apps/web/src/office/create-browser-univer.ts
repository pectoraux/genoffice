import { LogLevel, LocaleType, ThemeService, IUndoRedoService, Univer } from '@univerjs/core'
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
  | PluginCtor<Plugin>
  | [PluginCtor<Plugin>, ConstructorParameters<PluginCtor<Plugin>>[0]]
type BrowserPreset = { plugins: PluginEntry[] }

export interface BrowserUniverRuntime {
  readonly univer: Univer
  readonly univerAPI: FUniver
  /** Univer's theme service — ExcelEditor mirrors <html data-theme> into it. */
  readonly themeService: ThemeService
  /** Undo/redo stack occupancy — drives the QAT button greying. */
  readonly undoRedoService: IUndoRedoService
}

/**
 * Read the resolved theme from <html data-theme> (set by useTheme in
 * theme.ts). Falls back to the OS appearance when no attribute is set.
 * Mirrors the desktop's isDarkTheme() in App.tsx.
 */
function isDarkTheme(): boolean {
  if (typeof document === 'undefined') return false
  const attr = document.documentElement.getAttribute('data-theme')
  if (attr === 'dark') return true
  if (attr === 'light') return false
  return (
    typeof window !== 'undefined' &&
    typeof window.matchMedia === 'function' &&
    window.matchMedia('(prefers-color-scheme: dark)').matches
  )
}

/**
 * Merged en-US locale messages from every registered preset, plus the
 * sheets-ui compatibility patch. Univer 0.25.1's sheets-ui references two
 * info keys (error, forceStringInfo) the shipped pack lacks — without the
 * patch the raw keys pop up for users (same fix the desktop applies in
 * App.tsx). The existing sheets-ui namespace is spread first so the patch
 * only adds the missing entries instead of overwriting the whole namespace.
 */
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
  'sheets-ui': {
    ...(sheetsCoreEnUS as Record<string, Record<string, unknown>>)['sheets-ui'],
    info: {
      ...((sheetsCoreEnUS as Record<string, Record<string, Record<string, string>>>)['sheets-ui']?.info),
      error: 'Number stored as text',
      forceStringInfo:
        'The value in this cell is stored as text — it will not be treated as a ' +
        'number in formulas.',
    },
  },
}

export function createBrowserUniver(container: string): BrowserUniverRuntime {
  const univer = new Univer({
    logLevel: LogLevel.WARN,
    // green selection/highlight instead of Univer's default blue — matches
    // the desktop and the --accent token in theme.css.
    theme: greenTheme,
    darkMode: isDarkTheme(),
    locale: LocaleType.EN_US,
    locales: { [LocaleType.EN_US]: EN_US_LOCALES },
  })
  const presets: BrowserPreset[] = [
    // SheetsCore preset — Phase D parity with the desktop (App.tsx:1245-1264):
    //   toolbar: false  → hide Univer's preset ribbon; the web shell renders
    //                     its own 7-tab Ribbon so the chrome matches Electron
    //                     (no duplicate ribbon).
    //   zoomSlider: false → zoom lives in the custom status bar.
    //   statusBarStatistic: true → opt into the preset's statistic bar.
    //   sheets.isRowStylePrecedeColumnStyle: true → OOXML resolves style
    //     defaults as cell-over-row-over-column; without this flag Univer
    //     defaults to column-wins and inherits the wrong style for cells
    //     that rely on row/column defaults. NOT cosmetic.
    // NB: toolbar:false requires that React NOT double-mount Univer (StrictMode
    //   double-invokes effects in dev, which leaves the grid canvas unsized).
    //   main.tsx therefore omits <StrictMode> — this is dev-only (StrictMode
    //   has no effect on the production build) and the deployed app is
    //   unaffected. The custom Name Box + Formula Bar row render above the
    //   grid; Univer's preset header (name box + formula bar) is also
    //   visible inside the container — that pair is a known cosmetic
    //   duplicate (the desktop has the same: custom name box + Univer's
    //   formula bar). Hiding Univer's header via CSS breaks grid
    //   interactivity, so it is left visible.
    UniverSheetsCorePreset({
      container,
      header: true,
      toolbar: false,
      contextMenu: true,
      formulaBar: true,
      footer: { sheetBar: true, statisticBar: true, menus: true, zoomSlider: false },
      statusBarStatistic: true,
      sheets: { isRowStylePrecedeColumnStyle: true },
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
  // Dedupe plugins by pluginName across presets (last registration wins),
  // then register each survivor with its options — same algorithm the
  // desktop's create-univer.ts uses.
  const registry = new Map<string, { plugin: PluginCtor<Plugin>; options: unknown }>()
  for (const preset of presets) {
    for (const entry of preset.plugins) {
      const [plugin, options] = Array.isArray(entry) ? entry : [entry, undefined]
      registry.delete(plugin.pluginName)
      registry.set(plugin.pluginName, { plugin, options })
    }
  }
  for (const { plugin, options } of registry.values()) {
    univer.registerPlugin(plugin, options)
  }
  // Expose the theme + undo/redo services so the shell can mirror the DOM
  // theme into Univer's canvas and grey out the QAT buttons correctly.
  const injector = univer.__getInjector()
  const themeService = injector.get(ThemeService)
  const undoRedoService = injector.get(IUndoRedoService)
  return { univer, univerAPI: FUniver.newAPI(univer), themeService, undoRedoService }
}
