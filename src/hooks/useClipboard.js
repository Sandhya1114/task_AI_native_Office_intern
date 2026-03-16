// ─────────────────────────────────────────────────────────────
//  useClipboard.js
//  Ctrl+C  → copies computed value to system clipboard
//  Ctrl+V  → uses native paste event (no permission needed)
// ─────────────────────────────────────────────────────────────

import { useEffect, useRef } from 'react'

export function useClipboard({ engine, selectedCell, onAfterPaste }) {
  const selectedCellRef = useRef(selectedCell)
  useEffect(() => {
    selectedCellRef.current = selectedCell
  }, [selectedCell])

  useEffect(() => {
    // ── Ctrl+C: copy computed value ───────────────────────────
    function handleKeyDown(e) {
      const ctrl = e.ctrlKey || e.metaKey
      if (ctrl && e.key === 'c') {
        const cell = selectedCellRef.current
        if (!cell) return
        const cellData = engine.getCell(cell.r, cell.c)
        const value = cellData.error
          ? cellData.error
          : (cellData.computed !== null && cellData.computed !== ''
              ? String(cellData.computed)
              : cellData.raw)
        navigator.clipboard.writeText(value).catch(() => {})
      }
    }

    // ── Paste event ───────────────────────────────────────────
    function handlePaste(e) {
      const cell = selectedCellRef.current
      if (!cell) return

      // If the user is actively typing inside a cell input, let the
      // browser handle the paste normally (inserts text into the input).
      // We only intercept paste when a cell is selected but NOT being edited.
      const active = document.activeElement
      const isCellInput = active && active.classList.contains('cell-input')
      if (isCellInput) return  // let browser paste into the input as normal

      e.preventDefault()

      const text = e.clipboardData?.getData('text/plain')
      if (!text) return

      // Parse tab-separated rows (Excel / Google Sheets format)
      const rows = text.trimEnd().split('\n').map(row => row.split('\t'))

      rows.forEach((rowData, rowOffset) => {
        rowData.forEach((cellValue, colOffset) => {
          const targetRow = cell.r + rowOffset
          const targetCol = cell.c + colOffset
          if (targetRow >= engine.rows || targetCol >= engine.cols) return
          engine.setCell(targetRow, targetCol, cellValue.trim())
        })
      })

      onAfterPaste()
    }

    window.addEventListener('keydown', handleKeyDown)
    window.addEventListener('paste', handlePaste)
    return () => {
      window.removeEventListener('keydown', handleKeyDown)
      window.removeEventListener('paste', handlePaste)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])
}