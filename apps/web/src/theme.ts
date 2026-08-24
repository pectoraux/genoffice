/**
 * GenOffice web — single theme source.
 *
 * One hook owns the light/dark/system mode, persists it to localStorage,
 * sets <html data-theme="..."> for the CSS token layer, and subscribes to
 * the OS appearance media query so "system" stays live.
 *
 * ExcelEditor reads the resolved `effective` value and mirrors it into
 * Univer's ThemeService.setDarkMode() — no parallel theme state.
 */
import { useCallback, useEffect, useState } from 'react'

export type ThemeMode = 'light' | 'dark' | 'system'
export type EffectiveTheme = 'light' | 'dark'

const STORAGE_KEY = 'genoffice-theme'

function readStoredMode(): ThemeMode {
  try {
    const v = localStorage.getItem(STORAGE_KEY)
    if (v === 'light' || v === 'dark' || v === 'system') return v
  } catch {
    /* localStorage unavailable */
  }
  return 'system'
}

function systemPrefersDark(): boolean {
  return (
    typeof window !== 'undefined' &&
    typeof window.matchMedia === 'function' &&
    window.matchMedia('(prefers-color-scheme: dark)').matches
  )
}

function resolveEffective(mode: ThemeMode): EffectiveTheme {
  if (mode === 'light') return 'light'
  if (mode === 'dark') return 'dark'
  return systemPrefersDark() ? 'dark' : 'light'
}

export interface ThemeState {
  readonly mode: ThemeMode
  readonly effective: EffectiveTheme
  readonly setMode: (mode: ThemeMode) => void
}

/**
 * Mount this once near the root. It mutates <html data-theme> globally so
 * every screen (login, office home, editors) shares one theme.
 */
export function useTheme(): ThemeState {
  const [mode, setModeState] = useState<ThemeMode>(readStoredMode)
  const [effective, setEffective] = useState<EffectiveTheme>(() =>
    resolveEffective(readStoredMode()),
  )

  // Apply mode → <html data-theme> + recompute on OS appearance change.
  useEffect(() => {
    const apply = () => {
      const eff = resolveEffective(mode)
      setEffective(eff)
      document.documentElement.setAttribute('data-theme', eff)
    }
    apply()
    if (mode !== 'system') return
    const mq = window.matchMedia('(prefers-color-scheme: dark)')
    mq.addEventListener('change', apply)
    return () => mq.removeEventListener('change', apply)
  }, [mode])

  const setMode = useCallback((next: ThemeMode) => {
    try {
      localStorage.setItem(STORAGE_KEY, next)
    } catch {
      /* storage unavailable — keep in-memory */
    }
    setModeState(next)
  }, [])

  return { mode, effective, setMode }
}
