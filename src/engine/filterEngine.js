// filterEngine.js
// Manages per-column filtering as a VIEW-LAYER concern only.
// ALL rows participate in filtering — no pinned header row.

export function createFilterEngine(engine) {
  const activeFilters = new Map()

  // Gets unique display values for a column across ALL rows
  function getUniqueValues(col) {
    const values = new Set()
    for (let r = 0; r < engine.rows; r++) {
      const cell = engine.getCell(r, col)
      const val = (cell.computed !== null && cell.computed !== '') ? String(cell.computed) : (cell.raw || '')
      values.add(val)
    }
    return [...values].sort((a, b) => {
      const na = parseFloat(a), nb = parseFloat(b)
      if (!isNaN(na) && !isNaN(nb)) return na - nb
      return a.localeCompare(b)
    })
  }

  function setFilter(col, selectedValues) {
    if (selectedValues === null) activeFilters.delete(col)
    else activeFilters.set(col, new Set(selectedValues))
  }

  function clearFilter(col) { activeFilters.delete(col) }
  function clearAllFilters() { activeFilters.clear() }

  // Returns the set of row indices that pass ALL active filters, or null if no filters
  function getFilteredRows() {
    if (activeFilters.size === 0) return null
    const result = new Set()
    for (let r = 0; r < engine.rows; r++) {
      let passes = true
      for (const [col, allowedValues] of activeFilters) {
        const cell = engine.getCell(r, col)
        const val = (cell.computed !== null && cell.computed !== '') ? String(cell.computed) : (cell.raw || '')
        if (!allowedValues.has(val)) { passes = false; break }
      }
      if (passes) result.add(r)
    }
    return result
  }

  function hasFilter(col) { return activeFilters.has(col) }
  function getFilterValues(col) { return activeFilters.get(col) ?? null }

  return { getUniqueValues, setFilter, clearFilter, clearAllFilters, getFilteredRows, hasFilter, getFilterValues }
}