/**
 * GenOffice web Sheets — Ribbon.
 *
 * Seven top-level tabs matching the Electron Sheets ribbon (Home, Insert,
 * Page Layout, Formulas, Data, Review, View). The Home, Insert and View
 * tabs carry the highest-value commands wired through the runtime API;
 * the Page Layout / Formulas / Data / Review tabs are visual-structure
 * placeholders whose controls are clearly disabled (per spec: "Disabled
 * controls must be clearly disabled"). Nothing is faked — a disabled
 * control does nothing.
 *
 * The desktop's ExcelShell.tsx ribbon (3286 lines) is a frozen surface and
 * is NOT imported; this is a fresh web implementation using the desktop
 * only as the visual/interaction reference.
 */
import { useState, type ReactNode } from 'react'
import type { ExcelRuntimeApi, ExcelRuntimeState } from './useExcelRuntime'
import {
  UndoIcon,
  RedoIcon,
  BoldIcon,
  ItalicIcon,
  UnderlineIcon,
  FontColorIcon,
  FillColorIcon,
  AlignLeftIcon,
  AlignCenterIcon,
  AlignRightIcon,
  AlignTopIcon,
  AlignMiddleIcon,
  AlignBottomIcon,
  WrapIcon,
  MergeIcon,
  TableIcon,
  ChartIcon,
  ImageIcon,
  GridlinesIcon,
  FreezePaneIcon,
  FunctionIcon,
} from './RibbonIcons'

type TabId = 'home' | 'insert' | 'page' | 'formulas' | 'data' | 'review' | 'view'

const TABS: ReadonlyArray<{ id: TabId; label: string }> = [
  { id: 'home', label: 'Home' },
  { id: 'insert', label: 'Insert' },
  { id: 'page', label: 'Page Layout' },
  { id: 'formulas', label: 'Formulas' },
  { id: 'data', label: 'Data' },
  { id: 'review', label: 'Review' },
  { id: 'view', label: 'View' },
]

const FONT_FAMILIES = [
  'Calibri',
  'Cambria',
  'Arial',
  'Comic Sans MS',
  'Courier New',
  'Georgia',
  'Tahoma',
  'Times New Roman',
  'Trebuchet MS',
  'Verdana',
]
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72]
const NUMBER_FORMATS: ReadonlyArray<{ label: string; pattern: string }> = [
  { label: 'General', pattern: 'General' },
  { label: 'Number', pattern: '#,##0' },
  { label: 'Currency', pattern: '$#,##0' },
  { label: 'Percentage', pattern: '0%' },
  { label: 'Date', pattern: 'yyyy-mm-dd' },
  { label: 'Time', pattern: 'hh:mm' },
  { label: 'Text', pattern: '@' },
]

function Group({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="excel-ribbon-group">
      <div className="excel-ribbon-group-controls">{children}</div>
      <div className="excel-ribbon-group-label">{label}</div>
    </div>
  )
}

function RibbonButton({
  label,
  title,
  icon,
  active,
  disabled,
  onClick,
}: {
  label?: string
  title: string
  icon?: ReactNode
  active?: boolean
  disabled?: boolean
  onClick?: () => void
}) {
  return (
    <button
      type="button"
      className={`rb-btn${active ? ' active' : ''}`}
      title={title}
      aria-label={title}
      aria-pressed={active ? true : undefined}
      data-tip={title}
      disabled={disabled}
      onClick={onClick}
    >
      {icon}
      {label ? <span>{label}</span> : null}
    </button>
  )
}

/** Fallback state shown before the runtime is ready, so the shell layout is
 * stable from first paint (the grid gets its final height before Univer
 * measures it — otherwise the canvas mounts at 0×0). */
const NULL_STATE: ExcelRuntimeState = {
  ready: false,
  activeCellA1: '',
  selectionFormat: {
    bold: false,
    italic: false,
    underline: false,
    fontFamily: null,
    fontSize: null,
    fontColor: null,
    fillColor: null,
    hAlign: null,
    vAlign: null,
    wrap: false,
  },
  zoomPercent: 100,
  gridlinesVisible: true,
  canUndo: false,
  canRedo: false,
  seq: 0,
}

export function Ribbon({ api }: { api: ExcelRuntimeApi | null }) {
  const [tab, setTab] = useState<TabId>('home')
  const s = api?.state ?? NULL_STATE
  const disabled = !s.ready

  return (
    <div className="excel-ribbon" data-testid="excel-ribbon">
      <div className="excel-ribbon-tabs" role="tablist" aria-label="Ribbon tabs">
        {TABS.map((t) => (
          <button
            key={t.id}
            role="tab"
            aria-selected={tab === t.id}
            className={`excel-ribbon-tab${tab === t.id ? ' active' : ''}`}
            onClick={() => setTab(t.id)}
          >
            {t.label}
          </button>
        ))}
      </div>
      <div className="excel-ribbon-panel">
        {tab === 'home' && (
          <>
            <Group label="Undo">
              <RibbonButton title="Undo" icon={<UndoIcon />} disabled={!s.canUndo} onClick={() => api?.undo()} />
              <RibbonButton title="Redo" icon={<RedoIcon />} disabled={!s.canRedo} onClick={() => api?.redo()} />
            </Group>
            <Group label="Font">
              <select
                className="rb-select"
                aria-label="Font family"
                value={s.selectionFormat.fontFamily ?? 'Calibri'}
                disabled={disabled}
                onChange={(e) => api?.setFontFamily(e.target.value)}
              >
                {FONT_FAMILIES.map((f) => (
                  <option key={f} value={f}>
                    {f}
                  </option>
                ))}
              </select>
              <select
                className="rb-select"
                aria-label="Font size"
                style={{ maxWidth: 64 }}
                value={s.selectionFormat.fontSize ?? 11}
                disabled={disabled}
                onChange={(e) => api?.setFontSize(Number(e.target.value))}
              >
                {FONT_SIZES.map((n) => (
                  <option key={n} value={n}>
                    {n}
                  </option>
                ))}
              </select>
              <RibbonButton title="Bold" icon={<BoldIcon />} active={s.selectionFormat.bold} disabled={disabled} onClick={() => api?.toggleBold()} />
              <RibbonButton title="Italic" icon={<ItalicIcon />} active={s.selectionFormat.italic} disabled={disabled} onClick={() => api?.toggleItalic()} />
              <RibbonButton title="Underline" icon={<UnderlineIcon />} active={s.selectionFormat.underline} disabled={disabled} onClick={() => api?.toggleUnderline()} />
              <label className="rb-btn" title="Font color" aria-label="Font color" style={{ position: 'relative', padding: 0 }}>
                <FontColorIcon />
                <input
                  type="color"
                  aria-label="Font color picker"
                  className="rb-color"
                  style={{ position: 'absolute', inset: 0, opacity: 0, width: '100%', height: '100%', padding: 0, border: 'none', background: 'transparent', cursor: 'pointer' }}
                  value={s.selectionFormat.fontColor ?? '#000000'}
                  disabled={disabled}
                  onChange={(e) => api?.setFontColor(e.target.value)}
                />
              </label>
              <label className="rb-btn" title="Fill color" aria-label="Fill color" style={{ position: 'relative', padding: 0 }}>
                <FillColorIcon />
                <input
                  type="color"
                  aria-label="Fill color picker"
                  className="rb-color"
                  style={{ position: 'absolute', inset: 0, opacity: 0, width: '100%', height: '100%', padding: 0, border: 'none', background: 'transparent', cursor: 'pointer' }}
                  value={s.selectionFormat.fillColor ?? '#ffffff'}
                  disabled={disabled}
                  onChange={(e) => api?.setFillColor(e.target.value)}
                />
              </label>
            </Group>
            <Group label="Alignment">
              <RibbonButton title="Align left" icon={<AlignLeftIcon />} active={s.selectionFormat.hAlign === 'left'} disabled={disabled} onClick={() => api?.setHAlign(1)} />
              <RibbonButton title="Align center" icon={<AlignCenterIcon />} active={s.selectionFormat.hAlign === 'center'} disabled={disabled} onClick={() => api?.setHAlign(2)} />
              <RibbonButton title="Align right" icon={<AlignRightIcon />} active={s.selectionFormat.hAlign === 'right'} disabled={disabled} onClick={() => api?.setHAlign(3)} />
              <RibbonButton title="Align top" icon={<AlignTopIcon />} active={s.selectionFormat.vAlign === 'top'} disabled={disabled} onClick={() => api?.setVAlign(1)} />
              <RibbonButton title="Align middle" icon={<AlignMiddleIcon />} active={s.selectionFormat.vAlign === 'middle'} disabled={disabled} onClick={() => api?.setVAlign(2)} />
              <RibbonButton title="Align bottom" icon={<AlignBottomIcon />} active={s.selectionFormat.vAlign === 'bottom'} disabled={disabled} onClick={() => api?.setVAlign(3)} />
              <RibbonButton title="Wrap text" icon={<WrapIcon />} active={s.selectionFormat.wrap} disabled={disabled} onClick={() => api?.toggleWrap()} />
              <RibbonButton title="Merge & center" icon={<MergeIcon />} disabled={disabled} onClick={() => api?.toggleMerge()} />
            </Group>
            <Group label="Number">
              <select
                className="rb-select"
                aria-label="Number format"
                defaultValue="General"
                disabled={disabled}
                onChange={(e) => api?.setNumberFormat(e.target.value)}
              >
                {NUMBER_FORMATS.map((n) => (
                  <option key={n.label} value={n.pattern}>
                    {n.label}
                  </option>
                ))}
              </select>
            </Group>
          </>
        )}

        {tab === 'insert' && (
          <>
            <Group label="Tables">
              <RibbonButton label="Table" title="Table" icon={<TableIcon />} disabled />
            </Group>
            <Group label="Charts">
              <RibbonButton label="Chart" title="Chart" icon={<ChartIcon />} disabled />
            </Group>
            <Group label="Illustrations">
              <RibbonButton label="Picture" title="Image" icon={<ImageIcon />} disabled />
            </Group>
          </>
        )}

        {tab === 'page' && (
          <>
            <Group label="Page Setup">
              <RibbonButton label="Margins" title="Margins" disabled />
              <RibbonButton label="Orientation" title="Orientation" disabled />
              <RibbonButton label="Size" title="Size" disabled />
            </Group>
            <Group label="Scale">
              <RibbonButton label="Width" title="Width" disabled />
              <RibbonButton label="Height" title="Height" disabled />
            </Group>
          </>
        )}

        {tab === 'formulas' && (
          <>
            <Group label="Function Library">
              <RibbonButton label="ƒx" title="Insert function" icon={<FunctionIcon />} disabled />
              <RibbonButton label="AutoSum" title="AutoSum" disabled />
            </Group>
            <Group label="Defined Names">
              <RibbonButton label="Name Manager" title="Name Manager" disabled />
            </Group>
          </>
        )}

        {tab === 'data' && (
          <>
            <Group label="Sort & Filter">
              <RibbonButton label="Sort" title="Sort" disabled />
              <RibbonButton label="Filter" title="Filter" disabled />
            </Group>
            <Group label="Data Tools">
              <RibbonButton label="Remove Duplicates" title="Remove Duplicates" disabled />
              <RibbonButton label="Data Validation" title="Data Validation" disabled />
            </Group>
          </>
        )}

        {tab === 'review' && (
          <>
            <Group label="Comments">
              <RibbonButton label="New Comment" title="New Comment" disabled />
            </Group>
            <Group label="Protection">
              <RibbonButton label="Protect Sheet" title="Protect Sheet" disabled />
            </Group>
          </>
        )}

        {tab === 'view' && (
          <>
            <Group label="Show">
              <RibbonButton
                label="Gridlines"
                title="Toggle gridlines"
                icon={<GridlinesIcon />}
                active={s.gridlinesVisible}
                disabled={disabled}
                onClick={() => api?.toggleGridlines()}
              />
            </Group>
            <Group label="Zoom">
              <RibbonButton title="Zoom out" label="−" disabled={disabled} onClick={() => api?.zoomOut()} />
              <input
                className="rb-select"
                type="range"
                min={50}
                max={400}
                step={10}
                aria-label="Zoom slider"
                value={Math.min(400, Math.max(50, s.zoomPercent))}
                disabled={disabled}
                onChange={(e) => api?.setZoom(Number(e.target.value) / 100)}
                style={{ width: 100, accentColor: 'var(--accent)' }}
              />
              <RibbonButton title="Zoom in" label="+" disabled={disabled} onClick={() => api?.zoomIn()} />
              <span style={{ fontSize: 11, color: 'var(--text-secondary)', minWidth: 38, textAlign: 'right' }}>
                {s.zoomPercent}%
              </span>
            </Group>
            <Group label="Window">
              <RibbonButton label="Freeze Panes" title="Freeze panes" icon={<FreezePaneIcon />} disabled />
            </Group>
          </>
        )}
      </div>
    </div>
  )
}
