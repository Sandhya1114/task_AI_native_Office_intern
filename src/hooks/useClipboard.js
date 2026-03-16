// ─────────────────────────────────────────────────────────────
//  useClipboard.js
//
//  Ctrl+C  — copies selected range (or single cell) as
//            tab-separated text to system clipboard
//  Ctrl+V  — pastes tab-separated data starting at selectedCell
//            uses engine.batchSetCells so the ENTIRE paste is
//            undone with a single Ctrl+Z
// ─────────────────────────────────────────────────────────────

import { useEffect, useRef } from 'react'

export function useClipboard({ engine, selectedCell, selectionRange, onAfterPaste }) {
  // Always read the latest values without re-registering listeners
  const selectedCellRef  = useRef(selectedCell)
  const selectionRangeRef = useRef(selectionRange)

  useEffect(() => { selectedCellRef.current  = selectedCell  }, [selectedCell])
  useEffect(() => { selectionRangeRef.current = selectionRange }, [selectionRange])

  useEffect(() => {

    // ── Ctrl+C: copy range or single cell ─────────────────────
    function handleKeyDown(e) {
      const ctrl = e.ctrlKey || e.metaKey
      if (!ctrl || e.key !== 'c') return

      const cell  = selectedCellRef.current
      const range = selectionRangeRef.current
      if (!cell) return

      let text = ''

      if (range) {
        // Multi-cell: build tab-separated grid string
        const r1 = Math.min(range.start.r, range.end.r)
        const r2 = Math.max(range.start.r, range.end.r)
        const c1 = Math.min(range.start.c, range.end.c)
        const c2 = Math.max(range.start.c, range.end.c)

        const rowStrings = []
        for (let r = r1; r <= r2; r++) {
          const cols = []
          for (let c = c1; c <= c2; c++) {
            const cellData = engine.getCell(r, c)
            // Copy computed value so pasting gives the number, not the formula
            const val = cellData.error
              ? cellData.error
              : (cellData.computed !== null && cellData.computed !== ''
                  ? String(cellData.computed)
                  : cellData.raw)
            cols.push(val)
          }
          rowStrings.push(cols.join('\t'))
        }
        text = rowStrings.join('\n')
      } else {
        // Single cell
        const cellData = engine.getCell(cell.r, cell.c)
        text = cellData.error
          ? cellData.error
          : (cellData.computed !== null && cellData.computed !== ''
              ? String(cellData.computed)
              : cellData.raw)
      }

      navigator.clipboard.writeText(text).catch(() => {})
    }

    // ── Paste: parse and write as one batch ───────────────────
    function handlePaste(e) {
      const cell = selectedCellRef.current
      if (!cell) return

      // If user is typing inside a cell input, let browser handle it normally
      if (document.activeElement?.classList.contains('cell-input')) return

      e.preventDefault()

      const text = e.clipboardData?.getData('text/plain')
      if (!text) return

      // Parse tab-separated rows (Excel / Google Sheets format)
      const rows = text.trimEnd().split('\n').map(row => row.split('\t'))

      // Build the writes array
      const writes = []
      rows.forEach((rowData, rowOffset) => {
        rowData.forEach((cellValue, colOffset) => {
          const targetRow = cell.r + rowOffset
          const targetCol = cell.c + colOffset
          // Skip out-of-bounds cells
          if (targetRow >= engine.rows || targetCol >= engine.cols) return
          writes.push({ r: targetRow, c: targetCol, value: cellValue.trim() })
        })
      })

      if (writes.length === 0) return

      // ONE batch call = ONE undo entry → single Ctrl+Z undoes everything
      engine.batchSetCells(writes)

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