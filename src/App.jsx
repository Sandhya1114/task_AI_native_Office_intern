import { useState, useRef, useCallback, useMemo } from 'react'
import './App.css'
import { createEngine } from './engine/core.js'
import { createSortEngine } from './engine/sortEngine.js'
import { createFilterEngine } from './engine/filterEngine.js'
import FilterDropdown from './components/FilterDropdown.jsx'
import { useClipboard } from './hooks/useClipboard.js'
import { useLocalStorage, loadFromStorage, restoreCells } from './hooks/useLocalStorage.js'

const TOTAL_ROWS = 50
const TOTAL_COLS = 50

export default function App() {
  // Core spreadsheet engine — created once
  const [engine] = useState(() => {
    const eng = createEngine(TOTAL_ROWS, TOTAL_COLS)
    // Restore saved cells on first load — happens before first render
    const saved = loadFromStorage()
    if (saved) restoreCells(eng, saved.cells)
    return eng
  })

  // Sort and filter engines are also created once and share the core engine reference
  const [sortEngine] = useState(() => createSortEngine(engine))
  const [filterEngine] = useState(() => createFilterEngine(engine))

  const [version, setVersion] = useState(0)
  const [selectedCell, setSelectedCell] = useState(null)
  const [editingCell, setEditingCell] = useState(null)
  const [editValue, setEditValue] = useState('')
  const [cellStyles, setCellStyles] = useState(() => {
    // Restore saved styles on first load
    const saved = loadFromStorage()
    return saved?.styles ?? {}
  })

  // viewRows: the ordered list of data-row indices to render (view-layer sort/filter)
  const [viewRows, setViewRows] = useState(() =>
    Array.from({ length: TOTAL_ROWS }, (_, i) => i)
  )

  // openFilter: which column's filter dropdown is open, or null
  const [openFilter, setOpenFilter] = useState(null)

  const cellInputRef = useRef(null)

  const forceRerender = useCallback(() => setVersion(v => v + 1), [])

  // ────── View row helpers ──────

  /**
   * Recomputes viewRows after any sort or filter change.
   * Always reads current filter state then applies sort on top.
   */
  const refreshViewRows = useCallback(() => {
    const filteredRows = filterEngine.getFilteredRows()
    const newViewRows = sortEngine.computeViewRows(filteredRows)
    setViewRows(newViewRows)
  }, [sortEngine, filterEngine])

  // ────── Column sort handler ──────

  const handleColumnSort = useCallback((col) => {
    const filteredRows = filterEngine.getFilteredRows()
    const { viewRows: newViewRows } = sortEngine.cycleSort(col, filteredRows)
    setViewRows(newViewRows)
  }, [sortEngine, filterEngine])

  // ────── Filter handlers ──────

  const handleFilterApply = useCallback((col, selectedValues) => {
    filterEngine.setFilter(col, selectedValues)
    refreshViewRows()
  }, [filterEngine, refreshViewRows])

  // ────── Clipboard (Ctrl+C / Ctrl+V) ──────
  // Placed here so it has access to refreshViewRows and forceRerender
  useClipboard({
    engine,
    selectedCell,
    onAfterPaste: useCallback(() => {
      refreshViewRows()
      forceRerender()
    }, [refreshViewRows, forceRerender]),
  })

  // ────── Local storage persistence ──────
  // Watches version + cellStyles; debounced 500ms save on every change
  useLocalStorage({ engine, cellStyles, version })

  // ────── Cell style helpers ──────

  const getCellStyle = useCallback((row, col) => {
    const key = `${row},${col}`
    return cellStyles[key] || {
      bold: false, italic: false, underline: false,
      bg: 'white', color: '#202124', align: 'left', fontSize: 13
    }
  }, [cellStyles])

  const updateCellStyle = useCallback((row, col, updates) => {
    const key = `${row},${col}`
    setCellStyles(prev => ({
      ...prev,
      [key]: { ...getCellStyle(row, col), ...updates }
    }))
  }, [getCellStyle])

  // ────── Cell editing ──────

  const startEditing = useCallback((row, col) => {
    setSelectedCell({ r: row, c: col })
    setEditingCell({ r: row, c: col })
    const cellData = engine.getCell(row, col)
    setEditValue(cellData.raw)
    setTimeout(() => cellInputRef.current?.focus(), 0)
  }, [engine])

  const commitEdit = useCallback((row, col) => {
    const currentCell = engine.getCell(row, col)
    if (currentCell.raw !== editValue) {
      engine.setCell(row, col, editValue)
      // Refresh view rows because computed values may have changed (affects sort order display)
      refreshViewRows()
      forceRerender()
    }
    setEditingCell(null)
  }, [engine, editValue, forceRerender, refreshViewRows])

  // Single click: select only, no edit mode.
  // This prevents the cell input from stealing focus before Ctrl+V fires.
  const handleCellClick = useCallback((row, col) => {
    if (editingCell && (editingCell.r !== row || editingCell.c !== col)) {
      commitEdit(editingCell.r, editingCell.c)
    }
    setSelectedCell({ r: row, c: col })
    setEditingCell(null)
    document.activeElement?.blur()
  }, [editingCell, commitEdit])

  // Double click: enter edit mode
  const handleCellDoubleClick = useCallback((row, col) => {
    startEditing(row, col)
  }, [startEditing])

  // ────── Keyboard navigation ──────

  const handleKeyDown = useCallback((event, row, col) => {
    if (event.key === 'Enter') {
      event.preventDefault()
      commitEdit(row, col)
      // Navigate to next row in VIEW order, not data order
      const currentViewIdx = viewRows.indexOf(row)
      const nextViewIdx = Math.min(currentViewIdx + 1, viewRows.length - 1)
      startEditing(viewRows[nextViewIdx], col)
    } else if (event.key === 'Tab') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(row, Math.min(col + 1, engine.cols - 1))
    } else if (event.key === 'Escape') {
      setEditValue(engine.getCell(row, col).raw)
      setEditingCell(null)
    } else if (event.key === 'ArrowDown') {
      event.preventDefault()
      commitEdit(row, col)
      const currentViewIdx = viewRows.indexOf(row)
      const nextViewIdx = Math.min(currentViewIdx + 1, viewRows.length - 1)
      startEditing(viewRows[nextViewIdx], col)
    } else if (event.key === 'ArrowUp') {
      event.preventDefault()
      commitEdit(row, col)
      const currentViewIdx = viewRows.indexOf(row)
      const prevViewIdx = Math.max(currentViewIdx - 1, 0)
      startEditing(viewRows[prevViewIdx], col)
    } else if (event.key === 'ArrowLeft') {
      event.preventDefault()
      commitEdit(row, col)
      if (col > 0) {
        startEditing(row, col - 1)
      }
    } else if (event.key === 'ArrowRight') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(row, Math.min(col + 1, engine.cols - 1))
    }
  }, [engine, commitEdit, startEditing, viewRows])

  // ────── Formula bar handlers ──────

  const handleFormulaBarKeyDown = useCallback((event) => {
    if (!editingCell) return
    handleKeyDown(event, editingCell.r, editingCell.c)
  }, [editingCell, handleKeyDown])

  const handleFormulaBarFocus = useCallback(() => {
    if (selectedCell && !editingCell) {
      setEditingCell(selectedCell)
      setEditValue(engine.getCell(selectedCell.r, selectedCell.c).raw)
    }
  }, [selectedCell, editingCell, engine])

  const handleFormulaBarChange = useCallback((value) => {
    if (!editingCell && selectedCell) setEditingCell(selectedCell)
    setEditValue(value)
  }, [editingCell, selectedCell])

  // ────── Undo / Redo ──────

  const handleUndo = useCallback(() => {
    if (engine.undo()) {
      refreshViewRows()
      forceRerender()
    }
  }, [engine, forceRerender, refreshViewRows])

  const handleRedo = useCallback(() => {
    if (engine.redo()) {
      refreshViewRows()
      forceRerender()
    }
  }, [engine, forceRerender, refreshViewRows])

  // ────── Formatting toggles ──────

  const toggleBold = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { bold: !style.bold })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const toggleItalic = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { italic: !style.italic })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const toggleUnderline = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { underline: !style.underline })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const changeFontSize = useCallback((size) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { fontSize: size })
  }, [selectedCell, updateCellStyle])

  const changeAlignment = useCallback((align) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { align })
  }, [selectedCell, updateCellStyle])

  const changeFontColor = useCallback((color) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { color })
  }, [selectedCell, updateCellStyle])

  const changeBackgroundColor = useCallback((color) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { bg: color })
  }, [selectedCell, updateCellStyle])

  // ────── Clear operations ──────

  const clearSelectedCell = useCallback(() => {
    if (!selectedCell) return
    engine.setCell(selectedCell.r, selectedCell.c, '')
    forceRerender()
    const key = `${selectedCell.r},${selectedCell.c}`
    setCellStyles(prev => { const next = { ...prev }; delete next[key]; return next })
    setEditValue('')
  }, [selectedCell, engine, forceRerender])

  const clearAllCells = useCallback(() => {
    for (let r = 0; r < engine.rows; r++) {
      for (let c = 0; c < engine.cols; c++) {
        engine.setCell(r, c, '')
      }
    }
    forceRerender()
    setCellStyles({})
    setSelectedCell(null)
    setEditingCell(null)
    setEditValue('')
    // Reset sort and filter when clearing all
    sortEngine.resetSort()
    filterEngine.clearAllFilters()
    refreshViewRows()
  }, [engine, forceRerender, sortEngine, filterEngine, refreshViewRows])

  // ────── Row / Column operations ──────

  const insertRow = useCallback(() => {
    if (!selectedCell) return
    engine.insertRow(selectedCell.r)
    // Sort/filter state may be stale after structural change; reset to be safe
    sortEngine.resetSort()
    filterEngine.clearAllFilters()
    refreshViewRows()
    forceRerender()
    setSelectedCell({ r: selectedCell.r + 1, c: selectedCell.c })
  }, [selectedCell, engine, forceRerender, sortEngine, filterEngine, refreshViewRows])

  const deleteRow = useCallback(() => {
    if (!selectedCell) return
    engine.deleteRow(selectedCell.r)
    sortEngine.resetSort()
    filterEngine.clearAllFilters()
    refreshViewRows()
    forceRerender()
    if (selectedCell.r >= engine.rows) {
      setSelectedCell({ r: engine.rows - 1, c: selectedCell.c })
    }
  }, [selectedCell, engine, forceRerender, sortEngine, filterEngine, refreshViewRows])

  const insertColumn = useCallback(() => {
    if (!selectedCell) return
    engine.insertColumn(selectedCell.c)
    sortEngine.resetSort()
    filterEngine.clearAllFilters()
    refreshViewRows()
    forceRerender()
    setSelectedCell({ r: selectedCell.r, c: selectedCell.c + 1 })
  }, [selectedCell, engine, forceRerender, sortEngine, filterEngine, refreshViewRows])

  const deleteColumn = useCallback(() => {
    if (!selectedCell) return
    engine.deleteColumn(selectedCell.c)
    sortEngine.resetSort()
    filterEngine.clearAllFilters()
    refreshViewRows()
    forceRerender()
    if (selectedCell.c >= engine.cols) {
      setSelectedCell({ r: selectedCell.r, c: engine.cols - 1 })
    }
  }, [selectedCell, engine, forceRerender, sortEngine, filterEngine, refreshViewRows])

  // ────── Derived state ──────

  const selectedCellStyle = useMemo(() => {
    return selectedCell ? getCellStyle(selectedCell.r, selectedCell.c) : null
  }, [selectedCell, getCellStyle])

  const getColumnLabel = useCallback((col) => {
    let label = ''
    let num = col + 1
    while (num > 0) {
      num--
      label = String.fromCharCode(65 + (num % 26)) + label
      num = Math.floor(num / 26)
    }
    return label
  }, [])

  const selectedCellLabel = selectedCell
    ? `${getColumnLabel(selectedCell.c)}${selectedCell.r + 1}`
    : 'No cell'

  const formulaBarValue = editingCell
    ? editValue
    : (selectedCell ? engine.getCell(selectedCell.r, selectedCell.c).raw : '')

  // ────── Render ──────

  return (
    <div className="app-wrapper">
      <div className="app-header">
        <h2 className="app-title">📊 Spreadsheet App</h2>
      </div>

      <div className="main-content">

        {/* ── Toolbar ── */}
        <div className="toolbar">
          <div className="toolbar-group">
            <button className={`toolbar-btn bold-btn ${selectedCellStyle?.bold ? 'active' : ''}`} onClick={toggleBold} title="Bold">B</button>
            <button className={`toolbar-btn italic-btn ${selectedCellStyle?.italic ? 'active' : ''}`} onClick={toggleItalic} title="Italic">I</button>
            <button className={`toolbar-btn underline-btn ${selectedCellStyle?.underline ? 'active' : ''}`} onClick={toggleUnderline} title="Underline">U</button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Size:</span>
            <select className="toolbar-select" value={selectedCellStyle?.fontSize || 13} onChange={(e) => changeFontSize(parseInt(e.target.value))}>
              {[8, 10, 11, 12, 13, 14, 16, 18, 20, 24].map(s => (
                <option key={s} value={s}>{s}</option>
              ))}
            </select>
          </div>

          <div className="toolbar-group">
            <button className={`align-btn ${selectedCellStyle?.align === 'left' ? 'active' : ''}`} onClick={() => changeAlignment('left')} title="Align Left">⬤←</button>
            <button className={`align-btn ${selectedCellStyle?.align === 'center' ? 'active' : ''}`} onClick={() => changeAlignment('center')} title="Align Center">⬤</button>
            <button className={`align-btn ${selectedCellStyle?.align === 'right' ? 'active' : ''}`} onClick={() => changeAlignment('right')} title="Align Right">⬤→</button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Text:</span>
            <input
              type="color"
              value={selectedCellStyle?.color || '#000000'}
              onChange={(e) => changeFontColor(e.target.value)}
              title="Font color"
              style={{ width: '32px', height: '32px', border: '1px solid #dadce0', cursor: 'pointer', borderRadius: '4px' }}
            />
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Fill:</span>
            <select className="toolbar-select" value={selectedCellStyle?.bg || 'white'} onChange={(e) => changeBackgroundColor(e.target.value)}>
              <option value="white">White</option>
              <option value="#ffff99">Yellow</option>
              <option value="#99ffcc">Green</option>
              <option value="#ffcccc">Red</option>
              <option value="#cce5ff">Blue</option>
              <option value="#e0ccff">Purple</option>
              <option value="#ffd9b3">Orange</option>
              <option value="#f0f0f0">Gray</option>
            </select>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn" onClick={handleUndo} disabled={!engine.canUndo()} title="Undo">↶ Undo</button>
            <button className="toolbar-btn" onClick={handleRedo} disabled={!engine.canRedo()} title="Redo">↷ Redo</button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn" onClick={insertRow} title="Insert Row">+ Row</button>
            <button className="toolbar-btn" onClick={deleteRow} title="Delete Row">- Row</button>
            <button className="toolbar-btn" onClick={insertColumn} title="Insert Column">+ Col</button>
            <button className="toolbar-btn" onClick={deleteColumn} title="Delete Column">- Col</button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn danger" onClick={clearSelectedCell}>✕ Cell</button>
            <button className="toolbar-btn danger" onClick={clearAllCells}>✕ All</button>
          </div>
        </div>

        {/* ── Formula Bar ── */}
        <div className="formula-bar">
          <span className="formula-bar-label">{selectedCellLabel}</span>
          <input
            className="formula-bar-input"
            value={formulaBarValue}
            onChange={(e) => handleFormulaBarChange(e.target.value)}
            onKeyDown={handleFormulaBarKeyDown}
            onFocus={handleFormulaBarFocus}
            placeholder="Select a cell then type, or enter a formula like =SUM(A1:A5)"
          />
        </div>

        {/* ── Grid ── */}
        <div className="grid-scroll" onClick={() => setOpenFilter(null)}>
          <table className="grid-table">
            <thead>
              <tr>
                <th className="col-header-blank"></th>
                {Array.from({ length: engine.cols }, (_, colIndex) => (
                  <th key={colIndex} className="col-header">
                    <div className="col-header-inner">
                      {/* Sort button: clicking the label cycles sort */}
                      <span
                        className="col-header-label"
                        onClick={() => handleColumnSort(colIndex)}
                        title="Click to sort"
                      >
                        {getColumnLabel(colIndex)}
                        <span className="sort-icon">{sortEngine.getSortIcon(colIndex)}</span>
                      </span>

                      {/* Filter button: opens the dropdown */}
                      <span
                        className={`filter-icon-btn ${filterEngine.hasFilter(colIndex) ? 'active' : ''}`}
                        onClick={(e) => {
                          e.stopPropagation()
                          setOpenFilter(openFilter === colIndex ? null : colIndex)
                        }}
                        title="Filter"
                      >
                        ▾
                      </span>

                      {/* Filter dropdown (only shown for the open column) */}
                      {openFilter === colIndex && (
                        <FilterDropdown
                          col={colIndex}
                          getUniqueValues={(col) => filterEngine.getUniqueValues(col)}
                          activeValues={filterEngine.getFilterValues(colIndex)}
                          onApply={handleFilterApply}
                          onClose={() => setOpenFilter(null)}
                        />
                      )}
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {viewRows.map((dataRow) => (
                <tr key={dataRow}>
                  {/* Row header shows the ORIGINAL data row number for formula clarity */}
                  <td className="row-header">{dataRow + 1}</td>
                  {Array.from({ length: engine.cols }, (_, colIndex) => {
                    const isSelected = selectedCell?.r === dataRow && selectedCell?.c === colIndex
                    const isEditing = editingCell?.r === dataRow && editingCell?.c === colIndex
                    const cellData = engine.getCell(dataRow, colIndex)
                    const style = cellStyles[`${dataRow},${colIndex}`] || {}
                    const displayValue = cellData.error
                      ? cellData.error
                      : (cellData.computed !== null && cellData.computed !== '' ? String(cellData.computed) : cellData.raw)

                    return (
                      <td
                        key={colIndex}
                        className={`cell ${isSelected ? 'selected' : ''}`}
                        style={{ background: style.bg || 'white' }}
                        onMouseDown={(e) => { e.preventDefault(); handleCellClick(dataRow, colIndex) }}
                        onDoubleClick={() => handleCellDoubleClick(dataRow, colIndex)}
                      >
                        {isEditing ? (
                          <input
                            autoFocus
                            className="cell-input"
                            value={editValue}
                            onChange={(e) => setEditValue(e.target.value)}
                            onBlur={() => commitEdit(dataRow, colIndex)}
                            onKeyDown={(e) => handleKeyDown(e, dataRow, colIndex)}
                            ref={isSelected ? cellInputRef : undefined}
                            style={{
                              fontWeight: style.bold ? 'bold' : 'normal',
                              fontStyle: style.italic ? 'italic' : 'normal',
                              textDecoration: style.underline ? 'underline' : 'none',
                              color: style.color || '#202124',
                              fontSize: (style.fontSize || 13) + 'px',
                              textAlign: style.align || 'left',
                              background: style.bg || 'white',
                            }}
                          />
                        ) : (
                          <div
                            className={`cell-display align-${style.align || 'left'} ${cellData.error ? 'error' : ''}`}
                            style={{
                              fontWeight: style.bold ? 'bold' : 'normal',
                              fontStyle: style.italic ? 'italic' : 'normal',
                              textDecoration: style.underline ? 'underline' : 'none',
                              color: cellData.error ? '#d93025' : (style.color || '#202124'),
                              fontSize: (style.fontSize || 13) + 'px',
                            }}
                          >
                            {displayValue}
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

        <p className="footer-hint">
          Click column letter to sort (▲▼) · Click ▾ to filter · Ctrl+C copy · Ctrl+V paste (Excel/Sheets supported) · Formulas: =SUM(A1:A5) · =AVG() · =MAX() · =MIN()
        </p>
      </div>
    </div>
  )
}