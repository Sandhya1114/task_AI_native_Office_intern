// ─────────────────────────────────────────────────────────────
//  useLocalStorage.js
//  Auto-saves spreadsheet state to localStorage with 500ms debounce.
//  Restores state on page reload.
//
//  What is persisted:
//    - cell values and formulas (via engine snapshot)
//    - cell styles (bold, italic, colors, etc.)
//    - grid dimensions (rows, cols)
//
//  What is NOT persisted:
//    - undo/redo history (intentional — per spec)
//    - sort/filter state (resets on reload — clean UX)
// ─────────────────────────────────────────────────────────────

import { useEffect, useMemo } from 'react'
import { debounce } from '../utils/debounce.js'

const STORAGE_KEY = 'spreadsheet_state'
const SAVE_DELAY_MS = 500

/**
 * Reads all non-empty cell raw values out of the engine.
 * Returns a plain object: { "A1": "hello", "B2": "=A1+1", ... }
 */
function snapshotCells(engine) {
  const data = {}
  for (let r = 0; r < engine.rows; r++) {
    for (let c = 0; c < engine.cols; c++) {
      const cell = engine.getCell(r, c)
      // Only store cells that have content — keeps storage small
      if (cell.raw && cell.raw !== '') {
        data[`${r},${c}`] = cell.raw
      }
    }
  }
  return data
}

/**
 * Saves the full spreadsheet state to localStorage.
 * Wrapped in try/catch to handle QuotaExceededError gracefully.
 */
function saveToStorage(engine, cellStyles) {
  try {
    const state = {
      version: 1,                        // schema version for future migrations
      rows: engine.rows,
      cols: engine.cols,
      cells: snapshotCells(engine),
      styles: cellStyles,
      savedAt: Date.now(),
    }
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state))
  } catch (err) {
    // QuotaExceededError — storage is full, silently skip
    // This is intentional: we never want a storage error to crash the app
    console.warn('LocalStorage save failed:', err.message)
  }
}

/**
 * Reads and validates saved state from localStorage.
 * Returns null if nothing is saved or data is corrupted.
 */
export function loadFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (!raw) return null

    const state = JSON.parse(raw)

    // Basic schema validation — if these fields are missing, treat as corrupt
    if (!state || typeof state !== 'object') return null
    if (!state.cells || !state.rows || !state.cols) return null
    if (state.version !== 1) return null   // unknown schema version

    return state
  } catch {
    // JSON.parse failed — corrupted data, clear it and start fresh
    console.warn('LocalStorage data corrupted — clearing')
    localStorage.removeItem(STORAGE_KEY)
    return null
  }
}

/**
 * Restores saved cell values into the engine.
 * Called once on mount before first render.
 */
export function restoreCells(engine, savedCells) {
  if (!savedCells) return
  for (const [key, raw] of Object.entries(savedCells)) {
    const [r, c] = key.split(',').map(Number)
    // Bounds check — grid may be smaller than saved state
    if (r < engine.rows && c < engine.cols) {
      engine.setCell(r, c, raw)
    }
  }
}

/**
 * React hook: sets up debounced auto-save.
 * Call this in App with the current engine and cellStyles.
 * It watches `version` (the forceRerender counter) to know when to save.
 */
export function useLocalStorage({ engine, cellStyles, version }) {
  // Create the debounced save function once — stable across renders
  // useMemo is used instead of useRef().current to satisfy the
  // react-hooks/refs lint rule (no ref access during render)
  const debouncedSave = useMemo(
    () => debounce((eng, styles) => saveToStorage(eng, styles), SAVE_DELAY_MS),
    [] // empty deps — create once, never recreate
  )

  useEffect(() => {
    // Trigger a save whenever version or cellStyles changes
    debouncedSave(engine, cellStyles)
  }, [version, cellStyles, engine, debouncedSave])
}