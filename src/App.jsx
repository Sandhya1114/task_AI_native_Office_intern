import { useState, useRef, useCallback, useMemo, useEffect } from 'react'
import './App.css'
import { createEngine } from './engine/core.js'
import { createSortEngine } from './engine/sortEngine.js'
import { createFilterEngine } from './engine/filterEngine.js'
import FilterDropdown from './components/FilterDropdown.jsx'
import { useClipboard } from './hooks/useClipboard.js'
import { useLocalStorage, loadFromStorage, restoreCells } from './hooks/useLocalStorage.js'

const TOTAL_ROWS = 100
const TOTAL_COLS = 26

export default function App() {
  // ── Engine init (restore from localStorage on first load) ──
  const [engine] = useState(() => {
    const eng = createEngine(TOTAL_ROWS, TOTAL_COLS)
    const saved = loadFromStorage()
    if (saved) restoreCells(eng, saved.cells)
    return eng
  })
  const [sortEngine]   = useState(() => createSortEngine(engine))
  const [filterEngine] = useState(() => createFilterEngine(engine))

  // ── Core state ─────────────────────────────────────────────
  const [version, setVersion]               = useState(0)
  const [selectedCell, setSelectedCell]     = useState(null)
  const [selectionRange, setSelectionRange] = useState(null)
  const [editingCell, setEditingCell]       = useState(null)
  const [editValue, setEditValue]           = useState('')
  const [cellStyles, setCellStyles]         = useState(() => loadFromStorage()?.styles ?? {})
  const [viewRows, setViewRows]             = useState(() => Array.from({ length: TOTAL_ROWS }, (_, i) => i))
  const [openFilter, setOpenFilter]         = useState(null)
  const [colWidths, setColWidths]           = useState({})
  const [rowHeights, setRowHeights]         = useState({})
  const [saveStatus, setSaveStatus]         = useState(null)
  const [frozenRows, setFrozenRows]         = useState(0)
  const [frozenCols, setFrozenCols]         = useState(0)
  const [contextMenu, setContextMenu]       = useState(null)
  const [showFindReplace, setShowFindReplace] = useState(false)
  const [findText, setFindText]             = useState('')
  const [replaceText, setReplaceText]       = useState('')
  const [findResults, setFindResults]       = useState([])
  const [findIndex, setFindIndex]           = useState(0)
  const [sheetName, setSheetName]           = useState('Sheet1')
  const [editingSheetName, setEditingSheetName] = useState(false)
  // Number formatting: { "row,col": "number" | "percent" | "currency" | "date" }
  const [cellFormats, setCellFormats]       = useState({})
  // Cell comments: { "row,col": "comment text" }
  const [cellComments, setCellComments]     = useState({})
  // Zoom level: 0.7 – 1.5
  const [zoom, setZoom]                     = useState(1)
  // Autocomplete suggestion
  const [autoSuggest, setAutoSuggest]       = useState(null) // { value: string }
  // Active tooltip (comment hover): { row, col, x, y }
  const [commentTooltip, setCommentTooltip] = useState(null)

  // ── Refs ───────────────────────────────────────────────────
  const isDragging    = useRef(false)
  const dragStart     = useRef(null)
  const cellInputRef  = useRef(null)
  const resizingCol   = useRef(null)
  const resizingRow   = useRef(null)
  const gridScrollRef = useRef(null)

  const forceRerender = useCallback(() => setVersion(v => v + 1), [])

  // ── View helpers ───────────────────────────────────────────
  const refreshViewRows = useCallback(() => {
    setViewRows(sortEngine.computeViewRows(filterEngine.getFilteredRows()))
  }, [sortEngine, filterEngine])

  const handleColumnSort = useCallback((col) => {
    setViewRows(sortEngine.cycleSort(col, filterEngine.getFilteredRows()).viewRows)
  }, [sortEngine, filterEngine])

  const handleFilterApply = useCallback((col, vals) => {
    filterEngine.setFilter(col, vals)
    refreshViewRows()
  }, [filterEngine, refreshViewRows])

  // ── Column label helper ────────────────────────────────────
  const getColumnLabel = useCallback((col) => {
    let label = '', num = col + 1
    while (num > 0) { num--; label = String.fromCharCode(65 + (num % 26)) + label; num = Math.floor(num / 26) }
    return label
  }, [])

  // ── Clipboard ──────────────────────────────────────────────
  useClipboard({
    engine, selectedCell,
    onAfterPaste: useCallback(() => { refreshViewRows(); forceRerender() }, [refreshViewRows, forceRerender]),
  })

  // ── Local storage ──────────────────────────────────────────
  useLocalStorage({
    engine, cellStyles, version,
    onSave: useCallback(() => { setSaveStatus('saved'); setTimeout(() => setSaveStatus(null), 2000) }, []),
  })

  // ── Column resize (drag header edge) ──────────────────────
  const handleResizeMouseDown = useCallback((e, colIndex) => {
    e.preventDefault(); e.stopPropagation()
    // Capture everything in local variables — never read from ref inside setState
    const startWidth = colWidths[colIndex] ?? 100
    const startX = e.clientX
    resizingCol.current = { col: colIndex, startX, startWidth }

    const onMove = ev => {
      if (!resizingCol.current) return          // guard: already cleaned up
      const { col, startX: sx, startWidth: sw } = resizingCol.current
      const w = Math.max(40, sw + ev.clientX - sx)
      setColWidths(p => ({ ...p, [col]: w }))  // only captured primitives used
    }
    const onUp = () => {
      resizingCol.current = null
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
  }, [colWidths])

  // ── Row resize (drag row header edge) ─────────────────────
  const handleRowResizeMouseDown = useCallback((e, rowIndex) => {
    e.preventDefault(); e.stopPropagation()
    const startHeight = rowHeights[rowIndex] ?? 24
    const startY = e.clientY
    resizingRow.current = { row: rowIndex, startY, startHeight }

    const onMove = ev => {
      if (!resizingRow.current) return
      const { row, startY: sy, startHeight: sh } = resizingRow.current
      const h = Math.max(20, sh + ev.clientY - sy)
      setRowHeights(p => ({ ...p, [row]: h }))
    }
    const onUp = () => {
      resizingRow.current = null
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
  }, [rowHeights])

  // ── Range helpers ──────────────────────────────────────────
  const getRangeLabel = useCallback((range) => {
    if (!range) return null
    const { start, end } = range
    const r1 = Math.min(start.r, end.r), r2 = Math.max(start.r, end.r)
    const c1 = Math.min(start.c, end.c), c2 = Math.max(start.c, end.c)
    if (r1 === r2 && c1 === c2) return null
    return `${getColumnLabel(c1)}${r1+1}:${getColumnLabel(c2)}${r2+1}`
  }, [getColumnLabel])

  const isInRange = useCallback((row, col) => {
    if (!selectionRange) return false
    const { start, end } = selectionRange
    return row >= Math.min(start.r, end.r) && row <= Math.max(start.r, end.r) &&
           col >= Math.min(start.c, end.c) && col <= Math.max(start.c, end.c)
  }, [selectionRange])

  // Range stats for status bar (SUM / AVG / COUNT)
  const rangeStats = useMemo(() => {
    if (!selectionRange) return null
    const { start, end } = selectionRange
    const r1 = Math.min(start.r, end.r), r2 = Math.max(start.r, end.r)
    const c1 = Math.min(start.c, end.c), c2 = Math.max(start.c, end.c)
    if (r1 === r2 && c1 === c2) return null
    const nums = []
    for (let r = r1; r <= r2; r++)
      for (let c = c1; c <= c2; c++) {
        const cell = engine.getCell(r, c)
        const v = cell.computed !== null ? cell.computed : cell.raw
        const n = parseFloat(v)
        if (!isNaN(n)) nums.push(n)
      }
    if (nums.length === 0) return null
    const sum = nums.reduce((a, b) => a + b, 0)
    return { sum: sum.toFixed(2), avg: (sum / nums.length).toFixed(2), count: nums.length }
  }, [selectionRange, version, engine]) // eslint-disable-line

  // ── Number formatting ──────────────────────────────────────
  const formatValue = useCallback((raw, computed, fmt) => {
    if (!fmt) return null
    const n = parseFloat(computed ?? raw)
    if (isNaN(n)) return null
    if (fmt === 'number')   return n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })
    if (fmt === 'currency') return n.toLocaleString(undefined, { style: 'currency', currency: 'USD' })
    if (fmt === 'percent')  return (n / 100).toLocaleString(undefined, { style: 'percent', minimumFractionDigits: 1 })
    if (fmt === 'date')     return new Date(n).toLocaleDateString()
    return null
  }, [])

  // ── Formula reference highlighting ─────────────────────────
  const formulaRefs = useMemo(() => {
    if (!editingCell || !editValue.startsWith('=')) return new Set()
    const refs = new Set()
    for (const m of editValue.toUpperCase().matchAll(/\b([A-Z]+)(\d+)\b/g))
      refs.add(`${parseInt(m[2]) - 1},${m[1].charCodeAt(0) - 65}`)
    return refs
  }, [editingCell, editValue])

  // ── Cell style helpers ─────────────────────────────────────
  const getCellStyle = useCallback((row, col) =>
    cellStyles[`${row},${col}`] || { bold: false, italic: false, underline: false, bg: '', color: '#202124', align: 'left', fontSize: 13 }
  , [cellStyles])

  const updateCellStyle = useCallback((row, col, updates) => {
    setCellStyles(p => ({ ...p, [`${row},${col}`]: { ...getCellStyle(row, col), ...updates } }))
  }, [getCellStyle])

  // Apply style to entire selection range
  const applyStyleToRange = useCallback((updates) => {
    if (!selectionRange) {
      if (selectedCell) updateCellStyle(selectedCell.r, selectedCell.c, updates)
      return
    }
    const { start, end } = selectionRange
    const r1 = Math.min(start.r, end.r), r2 = Math.max(start.r, end.r)
    const c1 = Math.min(start.c, end.c), c2 = Math.max(start.c, end.c)
    setCellStyles(prev => {
      const next = { ...prev }
      for (let r = r1; r <= r2; r++)
        for (let c = c1; c <= c2; c++) {
          const key = `${r},${c}`
          next[key] = { ...(prev[key] || getCellStyle(r, c)), ...updates }
        }
      return next
    })
  }, [selectionRange, selectedCell, updateCellStyle, getCellStyle])

  // ── Cell editing ───────────────────────────────────────────
  const startEditing = useCallback((row, col) => {
    setSelectedCell({ r: row, c: col })
    setSelectionRange(null)
    setEditingCell({ r: row, c: col })
    setEditValue(engine.getCell(row, col).raw)
    setTimeout(() => cellInputRef.current?.focus(), 0)
  }, [engine])

  const commitEdit = useCallback((row, col) => {
    const current = engine.getCell(row, col)
    if (current.raw !== editValue) { engine.setCell(row, col, editValue); refreshViewRows(); forceRerender() }
    setEditingCell(null)
    setAutoSuggest(null)
  }, [engine, editValue, forceRerender, refreshViewRows])

  // ── Mouse handlers (click + drag-select) ──────────────────
  const handleCellMouseDown = useCallback((e, row, col) => {
    e.preventDefault()
    if (editingCell) commitEdit(editingCell.r, editingCell.c)
    document.activeElement?.blur()
    if (e.shiftKey && selectedCell) {
      setSelectionRange({ start: selectedCell, end: { r: row, c: col } })
    } else {
      setSelectedCell({ r: row, c: col })
      setSelectionRange({ start: { r: row, c: col }, end: { r: row, c: col } })
      isDragging.current = true
      dragStart.current = { r: row, c: col }
    }
    setContextMenu(null)
    setOpenFilter(null)
  }, [editingCell, commitEdit, selectedCell])

  const handleCellMouseEnter = useCallback((row, col) => {
    if (isDragging.current && dragStart.current) {
      setSelectionRange({ start: dragStart.current, end: { r: row, c: col } })
    }
  }, [])

  const handleMouseUp = useCallback(() => { isDragging.current = false }, [])
  useEffect(() => {
    window.addEventListener('mouseup', handleMouseUp)
    return () => window.removeEventListener('mouseup', handleMouseUp)
  }, [handleMouseUp])

  const handleCellDoubleClick = useCallback((row, col) => startEditing(row, col), [startEditing])

  // ── Context menu ───────────────────────────────────────────
  const handleCellContextMenu = useCallback((e, row, col) => {
    e.preventDefault()
    setSelectedCell({ r: row, c: col })
    setContextMenu({ x: e.clientX, y: e.clientY, row, col })
  }, [])

  const closeContextMenu = useCallback(() => setContextMenu(null), [])

  const contextMenuAction = useCallback((action) => {
    if (!contextMenu) return
    const { row, col } = contextMenu
    closeContextMenu()
    const reset = () => { sortEngine.resetSort(); filterEngine.clearAllFilters(); refreshViewRows(); forceRerender() }
    if (action === 'insertRowAbove') { engine.insertRow(row);     reset() }
    if (action === 'insertRowBelow') { engine.insertRow(row + 1); reset() }
    if (action === 'deleteRow')      { engine.deleteRow(row);     reset() }
    if (action === 'insertColLeft')  { engine.insertColumn(col);     reset() }
    if (action === 'insertColRight') { engine.insertColumn(col + 1); reset() }
    if (action === 'deleteCol')      { engine.deleteColumn(col);     reset() }
    if (action === 'clearCell')      { engine.setCell(row, col, ''); forceRerender() }
    if (action === 'freezeRow')      { setFrozenRows(r => r === row + 1 ? 0 : row + 1) }
    if (action === 'freezeCol')      { setFrozenCols(c => c === col + 1 ? 0 : col + 1) }
    // Number formatting from context menu
    if (['fmt-number','fmt-currency','fmt-percent','fmt-date','fmt-none'].includes(action)) {
      const fmt = action === 'fmt-none' ? undefined : action.replace('fmt-', '')
      setCellFormats(p => { const n = {...p}; const k = `${row},${col}`; if (fmt) n[k] = fmt; else delete n[k]; return n })
    }
  }, [contextMenu, engine, forceRerender, refreshViewRows, sortEngine, filterEngine, closeContextMenu])

  // ── Find & Replace ─────────────────────────────────────────
  const runFind = useCallback(() => {
    if (!findText) return
    const results = []
    for (let r = 0; r < engine.rows; r++)
      for (let c = 0; c < engine.cols; c++) {
        const raw = engine.getCell(r, c).raw
        if (raw.toLowerCase().includes(findText.toLowerCase())) results.push({ r, c })
      }
    setFindResults(results)
    setFindIndex(0)
    if (results.length > 0) setSelectedCell(results[0])
  }, [findText, engine])

  const findNext = useCallback(() => {
    if (findResults.length === 0) return
    const next = (findIndex + 1) % findResults.length
    setFindIndex(next); setSelectedCell(findResults[next])
  }, [findResults, findIndex])

  const findPrev = useCallback(() => {
    if (findResults.length === 0) return
    const prev = (findIndex - 1 + findResults.length) % findResults.length
    setFindIndex(prev); setSelectedCell(findResults[prev])
  }, [findResults, findIndex])

  const runReplace = useCallback(() => {
    if (!findText || findResults.length === 0) return
    const { r, c } = findResults[findIndex]
    engine.setCell(r, c, engine.getCell(r, c).raw.replace(new RegExp(findText, 'gi'), replaceText))
    forceRerender(); runFind()
  }, [findText, replaceText, findResults, findIndex, engine, forceRerender, runFind])

  const runReplaceAll = useCallback(() => {
    if (!findText) return
    for (let r = 0; r < engine.rows; r++)
      for (let c = 0; c < engine.cols; c++) {
        const raw = engine.getCell(r, c).raw
        if (raw.toLowerCase().includes(findText.toLowerCase()))
          engine.setCell(r, c, raw.replace(new RegExp(findText, 'gi'), replaceText))
      }
    forceRerender(); setFindResults([])
  }, [findText, replaceText, engine, forceRerender])

  // ── Export CSV ────────────────────────────────────────────
  const exportCSV = useCallback(() => {
    const rows = []
    for (let r = 0; r < engine.rows; r++) {
      const row = []
      let hasData = false
      for (let c = 0; c < engine.cols; c++) {
        const cell = engine.getCell(r, c)
        const val = cell.computed !== null && cell.computed !== '' ? String(cell.computed) : cell.raw
        if (val) hasData = true
        // Wrap in quotes if contains comma or newline
        row.push(val.includes(',') || val.includes('\n') ? `"${val}"` : val)
      }
      if (hasData) rows.push(row.join(','))
    }
    const csv = rows.join('\n')
    const blob = new Blob([csv], { type: 'text/csv' })
    const url  = URL.createObjectURL(blob)
    const a    = document.createElement('a')
    a.href = url; a.download = `${sheetName}.csv`; a.click()
    URL.revokeObjectURL(url)
  }, [engine, sheetName])

  // ── Fill Down (Ctrl+D) ─────────────────────────────────────
  const fillDown = useCallback(() => {
    if (!selectionRange) return
    const { start, end } = selectionRange
    const r1 = Math.min(start.r, end.r), r2 = Math.max(start.r, end.r)
    const c1 = Math.min(start.c, end.c), c2 = Math.max(start.c, end.c)
    for (let c = c1; c <= c2; c++) {
      const srcVal = engine.getCell(r1, c).raw
      for (let r = r1 + 1; r <= r2; r++) engine.setCell(r, c, srcVal)
    }
    forceRerender()
  }, [selectionRange, engine, forceRerender])

  // ── Fill Right (Ctrl+R) ────────────────────────────────────
  const fillRight = useCallback(() => {
    if (!selectionRange) return
    const { start, end } = selectionRange
    const r1 = Math.min(start.r, end.r), r2 = Math.max(start.r, end.r)
    const c1 = Math.min(start.c, end.c), c2 = Math.max(start.c, end.c)
    for (let r = r1; r <= r2; r++) {
      const srcVal = engine.getCell(r, c1).raw
      for (let c = c1 + 1; c <= c2; c++) engine.setCell(r, c, srcVal)
    }
    forceRerender()
  }, [selectionRange, engine, forceRerender])

  // ── Autocomplete: collect unique values from column ────────
  const updateAutoSuggest = useCallback((value, col) => {
    if (!value || value.startsWith('=')) { setAutoSuggest(null); return }
    const seen = new Set()
    for (let r = 0; r < engine.rows; r++) {
      const raw = engine.getCell(r, col).raw
      if (raw && raw !== value && raw.toLowerCase().startsWith(value.toLowerCase())) seen.add(raw)
    }
    setAutoSuggest(seen.size > 0 ? { value: [...seen][0] } : null)
  }, [engine])

  // ── Comment helpers ────────────────────────────────────────
  const setComment = useCallback((row, col, text) => {
    setCellComments(p => {
      const n = { ...p }
      if (text.trim()) n[`${row},${col}`] = text.trim()
      else delete n[`${row},${col}`]
      return n
    })
  }, [])

  // ── Keyboard: in-cell ──────────────────────────────────────
  const handleKeyDown = useCallback((e, row, col) => {
    const vi = viewRows.indexOf(row)
    if (e.key === 'Enter')      { e.preventDefault(); commitEdit(row, col); startEditing(viewRows[Math.min(vi+1, viewRows.length-1)], col) }
    else if (e.key === 'Tab')   { e.preventDefault(); commitEdit(row, col); startEditing(row, Math.min(col+1, engine.cols-1)) }
    else if (e.key === 'Escape'){ setEditValue(engine.getCell(row, col).raw); setEditingCell(null) }
    else if (e.key === 'ArrowDown') { e.preventDefault(); commitEdit(row, col); startEditing(viewRows[Math.min(vi+1, viewRows.length-1)], col) }
    else if (e.key === 'ArrowUp')   { e.preventDefault(); commitEdit(row, col); startEditing(viewRows[Math.max(vi-1, 0)], col) }
    else if (e.key === 'ArrowLeft') { e.preventDefault(); commitEdit(row, col); if (col > 0) startEditing(row, col-1) }
    else if (e.key === 'ArrowRight'){ e.preventDefault(); commitEdit(row, col); startEditing(row, Math.min(col+1, engine.cols-1)) }
  }, [engine, commitEdit, startEditing, viewRows])

  // ── Keyboard: global ───────────────────────────────────────
  useEffect(() => {
    function onKey(e) {
      const ctrl = e.ctrlKey || e.metaKey

      if (ctrl && e.key === 'f') { e.preventDefault(); setShowFindReplace(p => !p); return }
      if (e.key === 'Escape' && showFindReplace) { setShowFindReplace(false); return }
      if (ctrl && e.key === 'z') { e.preventDefault(); if (engine.undo()) { refreshViewRows(); forceRerender() }; return }
      if (ctrl && e.key === 'y') { e.preventDefault(); if (engine.redo()) { refreshViewRows(); forceRerender() }; return }
      if (ctrl && e.key === 'b') { e.preventDefault(); if (selectedCell) applyStyleToRange({ bold: !getCellStyle(selectedCell.r, selectedCell.c).bold }); return }
      if (ctrl && e.key === 'i') { e.preventDefault(); if (selectedCell) applyStyleToRange({ italic: !getCellStyle(selectedCell.r, selectedCell.c).italic }); return }
      if (ctrl && e.key === 'u') { e.preventDefault(); if (selectedCell) applyStyleToRange({ underline: !getCellStyle(selectedCell.r, selectedCell.c).underline }); return }

      // Ctrl+D — fill down
      if (ctrl && e.key === 'd') { e.preventDefault(); fillDown(); return }
      // Ctrl+R — fill right
      if (ctrl && e.key === 'r') { e.preventDefault(); fillRight(); return }
      // Ctrl+A — select all
      if (ctrl && e.key === 'a') {
        e.preventDefault()
        if (selectedCell) {
          setSelectionRange({ start: { r: 0, c: 0 }, end: { r: engine.rows-1, c: engine.cols-1 } })
        }
        return
      }

      if (contextMenu) { closeContextMenu(); return }

      if (!editingCell && selectedCell) {
        const { r, c } = selectedCell
        const vi = viewRows.indexOf(r)

        // Arrow keys with optional Shift for range extension
        if (e.key === 'ArrowDown') {
          e.preventDefault()
          const nextR = viewRows[Math.min(vi+1, viewRows.length-1)]
          if (e.shiftKey) setSelectionRange(p => ({ start: p?.start ?? {r,c}, end: { r: nextR, c } }))
          else { setSelectedCell({ r: nextR, c }); setSelectionRange({ start: { r: nextR, c }, end: { r: nextR, c } }) }
        }
        if (e.key === 'ArrowUp') {
          e.preventDefault()
          const prevR = viewRows[Math.max(vi-1, 0)]
          if (e.shiftKey) setSelectionRange(p => ({ start: p?.start ?? {r,c}, end: { r: prevR, c } }))
          else { setSelectedCell({ r: prevR, c }); setSelectionRange({ start: { r: prevR, c }, end: { r: prevR, c } }) }
        }
        if (e.key === 'ArrowRight') {
          e.preventDefault()
          const nextC = Math.min(c+1, engine.cols-1)
          if (e.shiftKey) setSelectionRange(p => ({ start: p?.start ?? {r,c}, end: { r, c: nextC } }))
          else { setSelectedCell({ r, c: nextC }); setSelectionRange({ start: { r, c: nextC }, end: { r, c: nextC } }) }
        }
        if (e.key === 'ArrowLeft') {
          e.preventDefault()
          const prevC = Math.max(c-1, 0)
          if (e.shiftKey) setSelectionRange(p => ({ start: p?.start ?? {r,c}, end: { r, c: prevC } }))
          else { setSelectedCell({ r, c: prevC }); setSelectionRange({ start: { r, c: prevC }, end: { r, c: prevC } }) }
        }

        // Ctrl+Home / Ctrl+End — jump to corners
        if (ctrl && e.key === 'Home') { e.preventDefault(); setSelectedCell({ r: viewRows[0], c: 0 }); setSelectionRange(null) }
        if (ctrl && e.key === 'End')  { e.preventDefault(); setSelectedCell({ r: viewRows[viewRows.length-1], c: engine.cols-1 }); setSelectionRange(null) }

        // Tab — move right
        if (e.key === 'Tab') { e.preventDefault(); setSelectedCell({ r, c: Math.min(c+1, engine.cols-1) }); setSelectionRange(null) }

        // Delete / Backspace — clear selection
        if (e.key === 'Delete' || e.key === 'Backspace') {
          if (selectionRange) {
            const { start, end } = selectionRange
            for (let rr = Math.min(start.r,end.r); rr <= Math.max(start.r,end.r); rr++)
              for (let cc = Math.min(start.c,end.c); cc <= Math.max(start.c,end.c); cc++)
                engine.setCell(rr, cc, '')
          } else { engine.setCell(r, c, '') }
          forceRerender()
        }

        if (e.key === 'Enter' || e.key === 'F2') { e.preventDefault(); startEditing(r, c) }
        // Start typing directly into selected cell
        if (!ctrl && e.key.length === 1 && !e.altKey) { startEditing(r, c); setEditValue(e.key) }
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [editingCell, selectedCell, selectionRange, viewRows, engine, forceRerender, refreshViewRows, showFindReplace, contextMenu, applyStyleToRange, getCellStyle, closeContextMenu, startEditing]) // eslint-disable-line

  // ── Formatting actions ─────────────────────────────────────
  const toggleBold          = () => { if (selectedCell) applyStyleToRange({ bold:          !getCellStyle(selectedCell.r, selectedCell.c).bold }) }
  const toggleStrikethrough = () => { if (selectedCell) applyStyleToRange({ strikethrough: !getCellStyle(selectedCell.r, selectedCell.c).strikethrough }) }
  const toggleWrap          = () => { if (selectedCell) applyStyleToRange({ wrap:          !getCellStyle(selectedCell.r, selectedCell.c).wrap }) }
  const toggleItalic    = () => { if (selectedCell) applyStyleToRange({ italic:    !getCellStyle(selectedCell.r, selectedCell.c).italic }) }
  const toggleUnderline = () => { if (selectedCell) applyStyleToRange({ underline: !getCellStyle(selectedCell.r, selectedCell.c).underline }) }
  const changeFontSize  = (s) => { if (selectedCell) applyStyleToRange({ fontSize: s }) }
  const changeAlignment = (a) => { if (selectedCell) applyStyleToRange({ align: a }) }
  const changeFontColor = (c) => { if (selectedCell) applyStyleToRange({ color: c }) }
  const changeBg        = (b) => { if (selectedCell) applyStyleToRange({ bg: b }) }

  // ── Clear ──────────────────────────────────────────────────
  const clearSelectedCell = useCallback(() => {
    if (!selectedCell) return
    if (selectionRange) {
      const { start, end } = selectionRange
      for (let r = Math.min(start.r,end.r); r <= Math.max(start.r,end.r); r++)
        for (let c = Math.min(start.c,end.c); c <= Math.max(start.c,end.c); c++)
          engine.setCell(r, c, '')
    } else { engine.setCell(selectedCell.r, selectedCell.c, '') }
    forceRerender(); setEditValue('')
  }, [selectedCell, selectionRange, engine, forceRerender])

  const clearAllCells = useCallback(() => {
    for (let r = 0; r < engine.rows; r++)
      for (let c = 0; c < engine.cols; c++)
        engine.setCell(r, c, '')
    forceRerender(); setCellStyles({}); setCellFormats({})
    setSelectedCell(null); setSelectionRange(null); setEditingCell(null); setEditValue('')
    sortEngine.resetSort(); filterEngine.clearAllFilters(); refreshViewRows()
  }, [engine, forceRerender, sortEngine, filterEngine, refreshViewRows])

  // ── Row / Col operations ───────────────────────────────────
  const insertRow    = () => { if (!selectedCell) return; engine.insertRow(selectedCell.r);       sortEngine.resetSort(); filterEngine.clearAllFilters(); refreshViewRows(); forceRerender() }
  const deleteRow    = () => { if (!selectedCell) return; engine.deleteRow(selectedCell.r);       sortEngine.resetSort(); filterEngine.clearAllFilters(); refreshViewRows(); forceRerender() }
  const insertColumn = () => { if (!selectedCell) return; engine.insertColumn(selectedCell.c);   sortEngine.resetSort(); filterEngine.clearAllFilters(); refreshViewRows(); forceRerender() }
  const deleteColumn = () => { if (!selectedCell) return; engine.deleteColumn(selectedCell.c);   sortEngine.resetSort(); filterEngine.clearAllFilters(); refreshViewRows(); forceRerender() }

  // ── Undo / Redo ────────────────────────────────────────────
  const handleUndo = () => { if (engine.undo()) { refreshViewRows(); forceRerender() } }
  const handleRedo = () => { if (engine.redo()) { refreshViewRows(); forceRerender() } }

  // ── Formula bar ────────────────────────────────────────────
  const handleFormulaBarFocus   = () => { if (selectedCell && !editingCell) { setEditingCell(selectedCell); setEditValue(engine.getCell(selectedCell.r, selectedCell.c).raw) } }
  const handleFormulaBarChange  = (v) => { if (!editingCell && selectedCell) setEditingCell(selectedCell); setEditValue(v) }
  const handleFormulaBarKeyDown = (e) => { if (editingCell) handleKeyDown(e, editingCell.r, editingCell.c) }

  // ── Derived state ──────────────────────────────────────────
  const selectedCellStyle  = useMemo(() => selectedCell ? getCellStyle(selectedCell.r, selectedCell.c) : null, [selectedCell, getCellStyle])
  const rangeLabel         = getRangeLabel(selectionRange)
  const selectedCellLabel  = rangeLabel || (selectedCell ? `${getColumnLabel(selectedCell.c)}${selectedCell.r+1}` : '—')
  const formulaBarValue    = editingCell ? editValue : (selectedCell ? engine.getCell(selectedCell.r, selectedCell.c).raw : '')

  // ── Render ─────────────────────────────────────────────────
  return (
    <div className="app-wrapper" onClick={() => { closeContextMenu(); setOpenFilter(null) }}>

      {/* ── Header ── */}
      <div className="app-header">
        <div className="app-header-left">
          <span className="app-logo">📊</span>
          {editingSheetName
            ? <input className="sheet-name-input" autoFocus value={sheetName}
                onChange={e => setSheetName(e.target.value)}
                onBlur={() => setEditingSheetName(false)}
                onKeyDown={e => (e.key === 'Enter' || e.key === 'Escape') && setEditingSheetName(false)} />
            : <h2 className="app-title" onDoubleClick={() => setEditingSheetName(true)} title="Double-click to rename">{sheetName}</h2>
          }
        </div>
        <div className="app-header-center">
          <button className="header-btn" onClick={() => setShowFindReplace(p => !p)} title="Find & Replace (Ctrl+F)">🔍 Find</button>
          <button className="header-btn" onClick={() => setFrozenRows(r => r > 0 ? 0 : (selectedCell ? selectedCell.r + 1 : 1))}>
            {frozenRows > 0 ? `❄ Unfreeze` : '❄ Freeze Row'}
          </button>
        </div>
        <div className="app-header-right">
          {saveStatus === 'saved' && <span className="save-indicator saved">✓ Saved</span>}
        </div>
      </div>

      {/* ── Find & Replace ── */}
      {showFindReplace && (
        <div className="find-replace-bar" onClick={e => e.stopPropagation()}>
          <span className="find-replace-title">🔍 Find &amp; Replace</span>
          <input className="find-input" placeholder="Find…" value={findText}
            onChange={e => setFindText(e.target.value)}
            onKeyDown={e => e.key === 'Enter' && runFind()} />
          <input className="find-input" placeholder="Replace with…" value={replaceText}
            onChange={e => setReplaceText(e.target.value)} />
          <button className="find-btn" onClick={runFind}>Find</button>
          <button className="find-btn" onClick={findPrev}  disabled={findResults.length === 0}>◀</button>
          <button className="find-btn" onClick={findNext}  disabled={findResults.length === 0}>▶</button>
          <button className="find-btn" onClick={runReplace} disabled={findResults.length === 0}>Replace</button>
          <button className="find-btn" onClick={runReplaceAll}>Replace All</button>
          {findResults.length > 0 && <span className="find-count">{findIndex+1} / {findResults.length}</span>}
          <button className="find-close" onClick={() => setShowFindReplace(false)}>✕</button>
        </div>
      )}

      <div className="main-content">
        {/* ── Toolbar ── */}
        <div className="toolbar">
          <div className="toolbar-group">
            <button className={`toolbar-btn bold-btn      ${selectedCellStyle?.bold      ? 'active' : ''}`} onClick={toggleBold}      title="Bold (Ctrl+B)">B</button>
            <button className={`toolbar-btn italic-btn    ${selectedCellStyle?.italic    ? 'active' : ''}`} onClick={toggleItalic}    title="Italic (Ctrl+I)">I</button>
            <button className={`toolbar-btn underline-btn ${selectedCellStyle?.underline ? 'active' : ''}`} onClick={toggleUnderline} title="Underline (Ctrl+U)">U</button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Size</span>
            <select className="toolbar-select" value={selectedCellStyle?.fontSize || 13} onChange={e => changeFontSize(parseInt(e.target.value))}>
              {[8,10,11,12,13,14,16,18,20,24].map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>

          <div className="toolbar-group">
            <button className={`align-btn ${selectedCellStyle?.align === 'left'   ? 'active':''}`} onClick={() => changeAlignment('left')}   title="Left">≡←</button>
            <button className={`align-btn ${selectedCellStyle?.align === 'center' ? 'active':''}`} onClick={() => changeAlignment('center')} title="Center">≡</button>
            <button className={`align-btn ${selectedCellStyle?.align === 'right'  ? 'active':''}`} onClick={() => changeAlignment('right')}  title="Right">≡→</button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Text</span>
            <input type="color" value={selectedCellStyle?.color || '#202124'} onChange={e => changeFontColor(e.target.value)} title="Font color" className="color-picker" />
            <span className="toolbar-label">Fill</span>
            <input type="color" value={selectedCellStyle?.bg || '#ffffff'} onChange={e => changeBg(e.target.value)} title="Fill color" className="color-picker" />
          </div>

          <div className="toolbar-divider" />

          <div className="toolbar-group">
            <span className="toolbar-label">Format</span>
            <select className="toolbar-select"
              value={selectedCell ? (cellFormats[`${selectedCell.r},${selectedCell.c}`] || '') : ''}
              onChange={e => {
                if (!selectedCell) return
                const fmt = e.target.value
                setCellFormats(p => {
                  const n = {...p}
                  const k = `${selectedCell.r},${selectedCell.c}`
                  if (fmt) n[k] = fmt; else delete n[k]
                  return n
                })
                forceRerender()
              }}>
              <option value="">General</option>
              <option value="number">Number (1,234.00)</option>
              <option value="currency">Currency ($)</option>
              <option value="percent">Percent (%)</option>
              <option value="date">Date</option>
            </select>
          </div>

          <div className="toolbar-divider" />

          <div className="toolbar-group">
            <button className="toolbar-btn icon-btn" onClick={handleUndo} disabled={!engine.canUndo()} title="Undo (Ctrl+Z)">↶</button>
            <button className="toolbar-btn icon-btn" onClick={handleRedo} disabled={!engine.canRedo()} title="Redo (Ctrl+Y)">↷</button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn" onClick={insertRow}    title="Insert row">+ Row</button>
            <button className="toolbar-btn" onClick={deleteRow}    title="Delete row">− Row</button>
            <button className="toolbar-btn" onClick={insertColumn} title="Insert column">+ Col</button>
            <button className="toolbar-btn" onClick={deleteColumn} title="Delete column">− Col</button>
          </div>

          <div className="toolbar-group">
            <button className={`toolbar-btn ${selectedCellStyle?.strikethrough ? 'active':''}`} onClick={toggleStrikethrough} title="Strikethrough (S̶)"><s>S</s></button>
            <button className={`toolbar-btn ${selectedCellStyle?.wrap ? 'active':''}`} onClick={toggleWrap} title="Wrap text">⏎</button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Zoom</span>
            <button className="toolbar-btn icon-btn" onClick={() => setZoom(z => Math.max(0.6, +(z-0.1).toFixed(1)))} title="Zoom out">−</button>
            <span className="zoom-label">{Math.round(zoom*100)}%</span>
            <button className="toolbar-btn icon-btn" onClick={() => setZoom(z => Math.min(2, +(z+0.1).toFixed(1)))} title="Zoom in">+</button>
            <button className="toolbar-btn" onClick={() => setZoom(1)} title="Reset zoom">↺</button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn" onClick={exportCSV} title="Export as CSV">⬇ CSV</button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn danger" onClick={clearSelectedCell} title="Clear selection (Del)">✕ Clear</button>
            <button className="toolbar-btn danger" onClick={clearAllCells}     title="Clear all">✕ All</button>
          </div>
        </div>

        {/* ── Formula Bar ── */}
        <div className="formula-bar">
          <span className="formula-bar-label">{selectedCellLabel}</span>
          <span className="formula-bar-fx">fx</span>
          <input className="formula-bar-input" value={formulaBarValue}
            onChange={e => handleFormulaBarChange(e.target.value)}
            onKeyDown={handleFormulaBarKeyDown}
            onFocus={handleFormulaBarFocus}
            placeholder="Enter a value or formula — e.g. =SUM(A1:A10)" />
        </div>

        {/* ── Grid ── */}
        <div className="grid-scroll" ref={gridScrollRef} onClick={() => { closeContextMenu(); setOpenFilter(null) }}>
          <div className="grid-zoom-wrapper" style={{ transform: `scale(${zoom})`, transformOrigin: 'top left', width: `${100/zoom}%` }}>
          <table className="grid-table">
            <thead>
              <tr>
                {/* Corner cell — click to select all */}
                <th className="col-header-blank" onClick={() => {
                  setSelectedCell({ r: 0, c: 0 })
                  setSelectionRange({ start: { r: 0, c: 0 }, end: { r: engine.rows-1, c: engine.cols-1 } })
                }} title="Select all" />
                {Array.from({ length: engine.cols }, (_, ci) => (
                  <th key={ci}
                    className={`col-header${ci < frozenCols ? ' col-frozen' : ''}`}
                    style={{ width: colWidths[ci] ?? 100, minWidth: colWidths[ci] ?? 100 }}>
                    <div className="col-header-inner">
                      <span className="col-header-label" onClick={() => handleColumnSort(ci)} title="Click to sort">
                        {getColumnLabel(ci)}<span className="sort-icon">{sortEngine.getSortIcon(ci)}</span>
                      </span>
                      <span className={`filter-icon-btn ${filterEngine.hasFilter(ci) ? 'active' : ''}`}
                        onClick={e => { e.stopPropagation(); setOpenFilter(openFilter === ci ? null : ci) }}
                        title="Filter">▾</span>
                      {openFilter === ci && (
                        <FilterDropdown col={ci}
                          getUniqueValues={col => filterEngine.getUniqueValues(col)}
                          activeValues={filterEngine.getFilterValues(ci)}
                          onApply={handleFilterApply}
                          onClose={() => setOpenFilter(null)} />
                      )}
                      <span className="col-resize-handle" onMouseDown={e => handleResizeMouseDown(e, ci)} title="Drag to resize" />
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {viewRows.map((dataRow) => (
                <tr key={dataRow}
                  className={dataRow < frozenRows ? 'row-frozen' : ''}
                  style={{ height: rowHeights[dataRow] ?? 24 }}>
                  <td className="row-header" style={{ height: rowHeights[dataRow] ?? 24 }}>
                    {dataRow + 1}
                    {/* Row resize handle */}
                    <span className="row-resize-handle" onMouseDown={e => handleRowResizeMouseDown(e, dataRow)} />
                  </td>
                  {Array.from({ length: engine.cols }, (_, ci) => {
                    const isSelected   = selectedCell?.r === dataRow && selectedCell?.c === ci
                    const isEditing    = editingCell?.r  === dataRow && editingCell?.c  === ci
                    const isRef        = formulaRefs.has(`${dataRow},${ci}`)
                    const inRange      = isInRange(dataRow, ci)
                    const isFindMatch  = findResults.some(f => f.r === dataRow && f.c === ci)
                    const cellData     = engine.getCell(dataRow, ci)
                    const style        = cellStyles[`${dataRow},${ci}`] || {}
                    const fmt          = cellFormats[`${dataRow},${ci}`]

                    // Determine display value with number formatting applied
                    let displayVal
                    if (cellData.error) {
                      displayVal = cellData.error
                    } else {
                      const raw = cellData.computed !== null && cellData.computed !== '' ? String(cellData.computed) : cellData.raw
                      const formatted = fmt ? formatValue(cellData.raw, cellData.computed, fmt) : null
                      displayVal = formatted ?? raw
                    }

                    const cellClass = ['cell',
                      isSelected ? 'selected' : '',
                      inRange && !isSelected ? 'in-range' : '',
                      isRef ? 'formula-ref' : '',
                      isFindMatch ? 'find-match' : '',
                    ].filter(Boolean).join(' ')

                    return (
                      <td key={ci} className={cellClass}
                        style={{ background: style.bg || 'transparent', width: colWidths[ci] ?? 100 }}
                        onMouseDown={e => handleCellMouseDown(e, dataRow, ci)}
                        onMouseEnter={() => handleCellMouseEnter(dataRow, ci)}
                        onDoubleClick={() => handleCellDoubleClick(dataRow, ci)}
                        onContextMenu={e => handleCellContextMenu(e, dataRow, ci)}>
                        {cellComments[`${dataRow},${ci}`] && (
                          <span className="comment-dot"
                            onMouseEnter={e => setCommentTooltip({ row: dataRow, col: ci, x: e.clientX, y: e.clientY })}
                            onMouseLeave={() => setCommentTooltip(null)} />
                        )}
                        {isEditing ? (
                          <>
                          <input autoFocus className="cell-input" value={editValue}
                            onChange={e => setEditValue(e.target.value)}
                            onBlur={() => commitEdit(dataRow, ci)}
                            ref={isSelected ? cellInputRef : undefined}
                            onChange={e => { setEditValue(e.target.value); updateAutoSuggest(e.target.value, ci) }}
                            onKeyDown={e => {
                              // Tab accepts autocomplete suggestion
                              if (e.key === 'Tab' && autoSuggest) { e.preventDefault(); setEditValue(autoSuggest.value); setAutoSuggest(null); return }
                              handleKeyDown(e, dataRow, ci)
                            }}
                            style={{
                              fontWeight:     style.bold      ? 'bold'      : 'normal',
                              fontStyle:      style.italic    ? 'italic'    : 'normal',
                              textDecoration: style.underline ? 'underline' : 'none',
                              color:          style.color     || '#202124',
                              fontSize:       (style.fontSize || 13) + 'px',
                              textAlign:      style.align     || 'left',
                              background:     style.bg        || 'white',
                            }} />
                          {isSelected && autoSuggest && (
                            <span className="autocomplete-hint">{autoSuggest.value}</span>
                          )}
                          </>
                        ) : (
                          <div className={`cell-display align-${style.align || 'left'} ${cellData.error ? 'error' : ''}`}
                            style={{
                              fontWeight:     style.bold      ? 'bold'      : 'normal',
                              fontStyle:      style.italic    ? 'italic'    : 'normal',
                              textDecoration: [style.underline ? 'underline' : '', style.strikethrough ? 'line-through' : ''].filter(Boolean).join(' ') || 'none',
                              color:          cellData.error  ? '#d93025'   : (style.color || '#202124'),
                              fontSize:       (style.fontSize || 13) + 'px',
                              whiteSpace:     style.wrap      ? 'normal'    : 'nowrap',
                              overflow:       style.wrap      ? 'visible'   : 'hidden',
                            }}>
                            {displayVal}
                          </div>
                        )}
                      </td>
                    )
                  })}
                </tr>
              ))}
            </tbody>
          </table>
          </div>
        </div>

        {/* ── Status Bar ── */}
        <div className="status-bar">
          <span className="status-left">
            {selectedCell
              ? rangeLabel
                ? `📐 ${rangeLabel}  ·  Shift+Arrow to extend  ·  Del to clear`
                : `📍 ${getColumnLabel(selectedCell.c)}${selectedCell.r+1}  ·  F2 or double-click to edit  ·  Start typing to replace`
              : '👆 Click a cell  ·  Ctrl+F find  ·  Right-click menu  ·  Ctrl+D fill down  ·  Ctrl+R fill right'}
          </span>
          <span className="status-center">
            {rangeStats && (
              <span className="range-stats">
                Σ <b>{rangeStats.sum}</b> · x̄ <b>{rangeStats.avg}</b> · # <b>{rangeStats.count}</b>
              </span>
            )}
          </span>
          <span className="status-right">
            {frozenRows > 0 && <span className="status-tag">❄ {frozenRows}R frozen</span>}
            {frozenCols > 0 && <span className="status-tag">❄ {frozenCols}C frozen</span>}
            <span>{engine.rows} × {engine.cols}</span>
          </span>
        </div>
      </div>

      {/* ── Comment Tooltip ── */}
      {commentTooltip && (
        <div className="comment-tooltip" style={{ top: commentTooltip.y + 12, left: commentTooltip.x + 8 }}>
          <span className="comment-tooltip-text">{cellComments[`${commentTooltip.row},${commentTooltip.col}`]}</span>
        </div>
      )}

      {/* ── Context Menu ── */}
      {contextMenu && (
        <div className="context-menu"
          style={{ top: contextMenu.y, left: contextMenu.x }}
          onClick={e => e.stopPropagation()}>
          <div className="ctx-item" onClick={() => contextMenuAction('insertRowAbove')}>↑ Insert row above</div>
          <div className="ctx-item" onClick={() => contextMenuAction('insertRowBelow')}>↓ Insert row below</div>
          <div className="ctx-item ctx-danger" onClick={() => contextMenuAction('deleteRow')}>✕ Delete row</div>
          <div className="ctx-divider" />
          <div className="ctx-item" onClick={() => contextMenuAction('insertColLeft')}>← Insert column left</div>
          <div className="ctx-item" onClick={() => contextMenuAction('insertColRight')}>→ Insert column right</div>
          <div className="ctx-item ctx-danger" onClick={() => contextMenuAction('deleteCol')}>✕ Delete column</div>
          <div className="ctx-divider" />
          <div className="ctx-item" onClick={() => contextMenuAction('clearCell')}>⌫ Clear cell</div>
          <div className="ctx-divider" />
          <div className="ctx-item" onClick={() => contextMenuAction('freezeRow')}>
            {frozenRows === contextMenu.row + 1 ? '❄ Unfreeze rows' : `❄ Freeze to row ${contextMenu.row + 1}`}
          </div>
          <div className="ctx-item" onClick={() => contextMenuAction('freezeCol')}>
            {frozenCols === contextMenu.col + 1 ? '❄ Unfreeze cols' : `❄ Freeze to col ${getColumnLabel(contextMenu.col)}`}
          </div>
          <div className="ctx-divider" />
          <div className="ctx-item" onClick={() => {
            const existing = cellComments[`${contextMenu.row},${contextMenu.col}`] || ''
            const text = window.prompt('Add comment:', existing)
            if (text !== null) { setComment(contextMenu.row, contextMenu.col, text); closeContextMenu() }
          }}>💬 {cellComments[`${contextMenu.row},${contextMenu.col}`] ? 'Edit comment' : 'Add comment'}</div>
          <div className="ctx-divider" />
          <div className="ctx-label">Number Format</div>
          <div className="ctx-item" onClick={() => contextMenuAction('fmt-number')}>123 Number</div>
          <div className="ctx-item" onClick={() => contextMenuAction('fmt-currency')}>$ Currency</div>
          <div className="ctx-item" onClick={() => contextMenuAction('fmt-percent')}>% Percent</div>
          <div className="ctx-item" onClick={() => contextMenuAction('fmt-date')}>📅 Date</div>
          <div className="ctx-item" onClick={() => contextMenuAction('fmt-none')}>× Clear format</div>
        </div>
      )}
    </div>
  )
}