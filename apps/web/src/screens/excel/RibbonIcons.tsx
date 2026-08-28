/**
 * GenOffice web Sheets — compact ribbon icon set.
 *
 * Self-contained 24×24 SVG glyphs (1.5-unit strokes, round caps/joins),
 * matching the desktop's ribbon-icons standard. Only the icons the web
 * Ribbon actually uses are defined here; the desktop's full glyph catalog
 * (apps/sheets/src/renderer/ribbon-icons.tsx) is a frozen surface and is
 * NOT imported.
 */
import type { ReactElement } from 'react'

function Icon({ children }: { readonly children: ReactElement | ReactElement[] }): ReactElement {
  return (
    <svg
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.5}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      {children}
    </svg>
  )
}

export function UndoIcon(): ReactElement {
  return (
    <Icon>
      <path d="M5.9 4 2.5 7.15 5.9 10.82" />
      <path d="M3.96 7.41h11.2c3.35 0 6.2 2.74 6.33 6.1.14 3.54-2.79 6.57-6.33 6.57H6.88" />
    </Icon>
  )
}

export function RedoIcon(): ReactElement {
  return (
    <Icon>
      <path d="M18.1 4 21.5 7.15 18.1 10.82" />
      <path d="M20.04 7.41H8.84C5.49 7.41 2.64 10.15 2.5 13.51 2.37 17.05 5.29 20.08 8.84 20.08h12.28" />
    </Icon>
  )
}

export function BoldIcon(): ReactElement {
  return (
    <Icon>
      <path d="M7 5h6.5a3.5 3.5 0 0 1 0 7H7zM7 12h7.5a3.5 3.5 0 0 1 0 7H7z" />
    </Icon>
  )
}

export function ItalicIcon(): ReactElement {
  return (
    <Icon>
      <path d="M10 5h8M6 19h8M14 5l-4 14" />
    </Icon>
  )
}

export function UnderlineIcon(): ReactElement {
  return (
    <Icon>
      <path d="M7 4v7a5 5 0 0 0 10 0V4M5 20h14" />
    </Icon>
  )
}

/** "A" with a colored bar underneath — the font-color affordance. */
export function FontColorIcon(): ReactElement {
  return (
    <Icon>
      <path d="M5.75 7V4.75h12.5V7M12 4.75v14.5M9.5 19.25h5" />
      <path d="M4 21.5h16" strokeWidth={2.5} />
    </Icon>
  )
}

/** Paint-bucket — the fill-color affordance. */
export function FillColorIcon(): ReactElement {
  return (
    <Icon>
      <path d="M11 4.75 4.35 11.4a1.5 1.5 0 0 0 0 2.12l4.9 4.9a1.5 1.5 0 0 0 2.12 0l6.66-6.66z" />
      <path d="M19.4 13.9c.75 1 1.15 1.85 1.15 2.5a1.45 1.45 0 0 1-2.9 0c0-.65.4-1.5 1.15-2.5z" />
    </Icon>
  )
}

export function AlignLeftIcon(): ReactElement {
  return (
    <Icon>
      <path d="M4.75 5.75h14.5M4.75 10.5h9.5M4.75 15.25h14.5M4.75 20h9.5" />
    </Icon>
  )
}

export function AlignCenterIcon(): ReactElement {
  return (
    <Icon>
      <path d="M4.75 4.75h14.5M7.25 9.5h9.5M4.75 14.25h14.5M7.25 19h9.5" />
    </Icon>
  )
}

export function AlignRightIcon(): ReactElement {
  return (
    <Icon>
      <path d="M4.75 5.75h14.5M9.75 10.5h9.5M4.75 15.25h14.5M9.75 20h9.5" />
    </Icon>
  )
}

export function AlignTopIcon(): ReactElement {
  return (
    <Icon>
      <path d="M5 4.75h14" />
      <path d="M12 19.25V9.5M8.5 13 12 9.5l3.5 3.5" />
    </Icon>
  )
}

export function AlignMiddleIcon(): ReactElement {
  return (
    <Icon>
      <path d="M4 12h16" />
      <path d="M12 19.25V4.75M8.5 8.25 12 4.75l3.5 3.5M8.5 15.75 12 19.25l3.5-3.5" />
    </Icon>
  )
}

export function AlignBottomIcon(): ReactElement {
  return (
    <Icon>
      <path d="M5 19.25h14" />
      <path d="M12 4.75v9.75M8.5 11l3.5 3.5 3.5-3.5" />
    </Icon>
  )
}

export function WrapIcon(): ReactElement {
  return (
    <Icon>
      <path d="M4.75 5.75h14.5M4.75 18.25h14.5" />
      <path d="M4.75 12h7.5a3.5 3.5 0 0 1 0 7H8.5M11.5 17 8.5 19l3 2" />
    </Icon>
  )
}

export function MergeIcon(): ReactElement {
  return (
    <Icon>
      <rect x="4.5" y="4.5" width="15" height="15" rx="1.5" />
      <path d="M9.5 4.5v15M14.5 4.5v15" />
      <path d="M4.5 12h15" strokeDasharray="2 2" />
    </Icon>
  )
}

export function TableIcon(): ReactElement {
  return (
    <Icon>
      <rect x="4.5" y="4.5" width="15" height="15" rx="1.5" />
      <path d="M4.5 9.5h15M4.5 14.5h15M9.5 4.5v15M14.5 4.5v15" />
    </Icon>
  )
}

export function ChartIcon(): ReactElement {
  return (
    <Icon>
      <path d="M4.5 19.5h15" />
      <path d="M8 19.5V12M12.5 19.5V8.5M17 19.5V15" />
    </Icon>
  )
}

export function ImageIcon(): ReactElement {
  return (
    <Icon>
      <rect x="3.9" y="5" width="16.2" height="14" rx="2.16" />
      <circle cx="8.76" cy="9.57" r="1.51" />
      <path d="m4.44 16.86 4.21-4.21 3.51 3.51 3.62-3.62 3.78 3.78" />
    </Icon>
  )
}

export function GridlinesIcon(): ReactElement {
  return (
    <Icon>
      <rect x="4.5" y="4.5" width="15" height="15" rx="1.5" />
      <path d="M4.5 9.5h15M4.5 14.5h15M9.5 4.5v15M14.5 4.5v15" />
    </Icon>
  )
}

export function FreezePaneIcon(): ReactElement {
  return (
    <Icon>
      <rect x="4.5" y="4.5" width="15" height="15" rx="1.5" />
      <path d="M12 4.5v15" strokeDasharray="3 2.5" />
      <path d="M4.5 12h15" strokeDasharray="3 2.5" />
    </Icon>
  )
}

export function FunctionIcon(): ReactElement {
  return (
    <Icon>
      <path d="M17.75 7.5V4.75H6.25L12.25 12l-6 7.25h11.5V16.5" />
    </Icon>
  )
}

/** Shaded grid cells — the conditional-formatting affordance (EXCEL-024). */
export function CfIcon(): ReactElement {
  return (
    <Icon>
      <rect x="3.75" y="3.75" width="6" height="6" rx="1" />
      <rect x="14.25" y="3.75" width="6" height="6" rx="1" />
      <rect x="9" y="14.25" width="6" height="6" rx="1" />
      <path d="M3.75 14.25h2.5M17.75 14.25h2.5M3.75 17.75h2.5M17.75 17.75h2.5" />
    </Icon>
  )
}

/** A name tag over a grid — the Name Manager affordance (EXCEL-025). */
export function NamesIcon(): ReactElement {
  return (
    <Icon>
      <path d="M3.5 8.25 8 3.75h12.5v12.5L16 20.75" />
      <path d="M8 3.75v4.5h4.5" />
      <path d="M3.5 8.25h9v12.5h-9z" />
      <path d="M6 12h4M6 15h4" />
    </Icon>
  )
}
