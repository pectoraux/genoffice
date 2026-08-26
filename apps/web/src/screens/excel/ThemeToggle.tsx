/**
 * GenOffice web Sheets — Theme toggle (Light / Dark / System).
 *
 * Single source of truth: writes the chosen mode back into useTheme (which
 * sets <html data-theme>); ExcelEditor's effect mirrors the resolved theme
 * into Univer's ThemeService.setDarkMode. No parallel theme state.
 */
import type { ThemeMode } from '../../theme'

export function ThemeToggle({
  mode,
  setMode,
}: {
  mode: ThemeMode
  setMode: (m: ThemeMode) => void
}) {
  return (
    <select
      className="theme-toggle"
      aria-label="Theme"
      data-testid="theme-toggle"
      value={mode}
      onChange={(e) => setMode(e.target.value as ThemeMode)}
    >
      <option value="light">Light</option>
      <option value="dark">Dark</option>
      <option value="system">System</option>
    </select>
  )
}
