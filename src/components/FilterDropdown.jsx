// ─────────────────────────────────────────────────────────────
//  FilterDropdown.jsx
//  Excel-style filter dropdown for a single column.
//  Shows a checklist of unique values; user can select/deselect.
// ─────────────────────────────────────────────────────────────

import { useState, useEffect, useRef } from 'react'

/**
 * @param {object} props
 * @param {number} props.col - column index
 * @param {function} props.getUniqueValues - (col) => string[]
 * @param {Set<string>|null} props.activeValues - currently selected values or null (all)
 * @param {function} props.onApply - (col, Set<string>|null) => void
 * @param {function} props.onClose - () => void
 */
export default function FilterDropdown({ col, getUniqueValues, activeValues, onApply, onClose }) {
  const allValues = getUniqueValues(col)

  // Local state: which values are checked
  const [checked, setChecked] = useState(() => {
    if (activeValues === null) return new Set(allValues)
    return new Set(activeValues)
  })

  const dropdownRef = useRef(null)

  // Close on outside click
  useEffect(() => {
    function handleClickOutside(e) {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target)) {
        onClose()
      }
    }
    document.addEventListener('mousedown', handleClickOutside)
    return () => document.removeEventListener('mousedown', handleClickOutside)
  }, [onClose])

  function toggleValue(val) {
    setChecked(prev => {
      const next = new Set(prev)
      if (next.has(val)) next.delete(val)
      else next.add(val)
      return next
    })
  }

  function selectAll() {
    setChecked(new Set(allValues))
  }

  function clearAll() {
    setChecked(new Set())
  }

  function handleApply() {
    // If all values are selected, treat as "no filter" (null)
    if (checked.size === allValues.length) {
      onApply(col, null)
    } else {
      onApply(col, new Set(checked))
    }
    onClose()
  }

  return (
    <div ref={dropdownRef} className="filter-dropdown" onClick={e => e.stopPropagation()}>
      <div className="filter-dropdown-header">
        <span className="filter-dropdown-title">Filter</span>
        <button className="filter-close-btn" onClick={onClose}>✕</button>
      </div>

      <div className="filter-dropdown-actions">
        <button className="filter-action-btn" onClick={selectAll}>All</button>
        <button className="filter-action-btn" onClick={clearAll}>None</button>
      </div>

      <div className="filter-dropdown-list">
        {allValues.length === 0 ? (
          <div className="filter-empty">No data</div>
        ) : (
          allValues.map(val => (
            <label key={val} className="filter-item">
              <input
                type="checkbox"
                checked={checked.has(val)}
                onChange={() => toggleValue(val)}
              />
              <span className="filter-item-label">{val === '' ? '(blank)' : val}</span>
            </label>
          ))
        )}
      </div>

      <div className="filter-dropdown-footer">
        <button className="filter-apply-btn" onClick={handleApply}>Apply</button>
        <button className="filter-clear-btn" onClick={() => { onApply(col, null); onClose() }}>Clear</button>
      </div>
    </div>
  )
}