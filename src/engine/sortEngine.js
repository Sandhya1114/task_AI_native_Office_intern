// sortEngine.js
// Manages column sorting as a VIEW-LAYER concern only.
// Empty cells always sort to the bottom regardless of direction.

export function createSortEngine(engine) {
  let sortState = { col: -1, direction: 'none' }

  function naturalOrder() {
    return Array.from({ length: engine.rows }, (_, i) => i)
  }

  function computeViewRows(filterRows = null) {
    let rows = naturalOrder()

    if (filterRows !== null) {
      rows = rows.filter(r => filterRows.has(r))
    }

    if (sortState.direction === 'none' || sortState.col < 0) {
      return rows
    }

    const col = sortState.col
    const dir = sortState.direction === 'asc' ? 1 : -1

    return [...rows].sort((rowA, rowB) => {
      const cellA = engine.getCell(rowA, col)
      const cellB = engine.getCell(rowB, col)

      const valA = (cellA.computed !== null && cellA.computed !== '') ? cellA.computed : cellA.raw
      const valB = (cellB.computed !== null && cellB.computed !== '') ? cellB.computed : cellB.raw

      const emptyA = valA === '' || valA === null || valA === undefined
      const emptyB = valB === '' || valB === null || valB === undefined

      // Empty cells always go to the bottom, regardless of sort direction
      if (emptyA && emptyB) return 0
      if (emptyA) return 1   // A is empty → push A down
      if (emptyB) return -1  // B is empty → push B down

      const numA = parseFloat(valA)
      const numB = parseFloat(valB)
      if (!isNaN(numA) && !isNaN(numB)) return (numA - numB) * dir

      const strA = String(valA).toLowerCase()
      const strB = String(valB).toLowerCase()
      if (strA < strB) return -1 * dir
      if (strA > strB) return 1 * dir
      return 0
    })
  }

  function cycleSort(col, filteredRows = null) {
    if (sortState.col !== col) {
      sortState = { col, direction: 'asc' }
    } else {
      const cycle = { asc: 'desc', desc: 'none', none: 'asc' }
      sortState = { col, direction: cycle[sortState.direction] }
    }
    return { viewRows: computeViewRows(filteredRows), sortState: { ...sortState } }
  }

  function resetSort() {
    sortState = { col: -1, direction: 'none' }
  }

  function getSortIcon(col) {
    if (sortState.col !== col || sortState.direction === 'none') return ''
    return sortState.direction === 'asc' ? ' ▲' : ' ▼'
  }

  return { computeViewRows, cycleSort, resetSort, getSortIcon, getSortState: () => ({ ...sortState }) }
}