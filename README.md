# 📊 SpreadsheetApp

A production-quality spreadsheet web application built with **React + Vite**, inspired by Google Sheets.  
Built as part of the **WorkElate AI Native Office Intern** assessment.

---

## ✨ Features

### ✅ Task 1 — Column Sort & Filter
- 3-state column sorting: **None → Ascending ▲ → Descending ▼ → None**
- Sorting works on **computed formula values**, not raw input
- Sort is **view-layer only** — original data is never mutated
- Empty cells always sort to the bottom regardless of direction
- Excel-style **filter dropdown** per column with checkbox checklist
- Filtering hides rows without deleting data
- Both sort and filter are fully **reversible**

### ✅ Task 2 — Multi-Cell Copy & Paste
- `Ctrl+C` copies the **computed value** (not raw formula) to clipboard
- `Ctrl+V` parses **tab-separated data** pasted from Excel or Google Sheets
- Supports **multi-row and multi-column** paste
- Bounds checking — never writes outside the grid
- All paste operations are **undoable** with `Ctrl+Z`
- Uses native browser `paste` event — no clipboard permission prompt needed

### ✅ Task 3 — Local Storage Persistence
- **Auto-saves** on every change with **500ms debounce**
- Restores **cell values, formulas, styles, and grid size** on page reload
- Undo/redo history intentionally **not persisted** (per spec)
- **Schema version** field for safe future migrations
- Handles corrupted data and `QuotaExceededError` gracefully
- Live **"✓ Saved"** indicator confirms persistence

---

## ⭐ Extra Features

| Feature | Description |
|---------|-------------|
| **Range Selection** | Click and drag or Shift+Arrow to select multiple cells |
| **Range Stats** | Live Sum / Avg / Count in status bar when numbers are selected |
| **Range Formatting** | Bold, color, fill — applies to entire selected range |
| **Range Delete** | Select range → Delete key clears all cells at once |
| **Right-click Context Menu** | Insert/delete rows and columns, freeze rows/cols |
| **Find & Replace** | `Ctrl+F` opens panel with navigation, Replace, Replace All |
| **Formula Highlighting** | Referenced cells glow green when editing a formula |
| **Column Resizing** | Drag right edge of any column header |
| **Row Resizing** | Drag bottom edge of any row header |
| **Freeze Rows** | Rows stay pinned while scrolling via button or right-click |
| **Keyboard Navigation** | Arrow keys, Tab, Enter, F2, Delete, Ctrl+Home/End |
| **Sheet Rename** | Double-click the sheet name in the header |
| **Cell Formatting** | Bold, Italic, Underline, Font Size, Color, Fill, Alignment |
| **Formula Engine** | SUM, AVG, MIN, MAX, arithmetic, circular ref detection |
| **Undo / Redo** | Full history via Ctrl+Z / Ctrl+Y |
| **Row & Col Management** | Insert/delete at selected position |

---

## 🏗️ Architecture & Key Decisions

```
src/
├── engine/
│   ├── core.js             # Pure JS engine — no React dependency
│   ├── sortEngine.js       # View-layer sort only
│   └── filterEngine.js     # View-layer filter only
├── components/
│   └── FilterDropdown.jsx  # Filter checklist UI
├── hooks/
│   ├── useClipboard.js     # Ctrl+C / Ctrl+V via native paste event
│   └── useLocalStorage.js  # Debounced auto-save and restore
├── utils/
│   └── debounce.js
├── App.jsx                 # Main component
└── App.css                 # CSS custom properties for theming
```

### Decision 1 — View-layer sort/filter
`viewRows` is an array of data-row indices e.g. `[3, 0, 2, 1]`.  
The grid renders `viewRows[i]` instead of `i`. The engine cell data is never touched.  
Formulas still reference original row numbers — sort never breaks `=A1+A2`.

### Decision 2 — Engine fully separated from React
`core.js` is a pure JavaScript module. It handles cells, formula parsing, dependency graph, and undo/redo.  
React only calls `engine.getCell()` and `engine.setCell()` — it never touches internal state.  
This makes the engine independently testable.

### Decision 3 — Native paste event over Clipboard API
`navigator.clipboard.readText()` requires explicit browser permission and fails silently.  
Listening to the native `paste` event gives us `e.clipboardData.getData('text')` directly — no permission needed, works in all browsers.

### Decision 4 — Single click selects, double click edits
If single click opened edit mode immediately, `Ctrl+V` would paste into the cell input instead of being caught by our paste handler — causing a one-row offset bug.  
Splitting click (select) from double-click (edit) fixes this cleanly.

### Decision 5 — Ref values captured before setState in resize
Column/row resize uses `useRef` for drag state. Inside `onMouseMove`, we destructure ref values into local variables *before* calling `setColWidths()`.  
This prevents a crash where React's async batching could fire the setState callback after `onMouseUp` had already set the ref to `null`.

### Decision 6 — useMemo for debounced save function
`useRef().current` accessed at component top level violates the `react-hooks/refs` lint rule.  
Using `useMemo(() => debounce(...), [])` creates the same stable function instance without the lint violation.

---

## 🛠️ Getting Started

```bash
git clone [https://github.com/Sandhya1114/task_AI_native_Office_intern.git]
cd spreadsheet-app
npm install
npm run dev
# → http://localhost:5173
```

```bash
npm run build    # Production build
npm run preview  # Preview build locally
npm run lint     # Run ESLint
```

---

## 🌿 Branch Structure

| Branch | Feature |
|--------|---------|
| `main` | Stable, production-ready |
| `feature/sort-and-filter` | Task 1 — Column sort & filter |
| `feature/clipboard-copy-paste` | Task 2 — Multi-cell Ctrl+C / Ctrl+V |
| `feature/local-storage-persistence` | Task 3 — Auto-save & restore |
| `feature/extras-and-polish` | Range select, find/replace, context menu, freeze, resize |

Commit messages follow **Conventional Commits**:
```
feat: implement view-layer column sort and filter
feat: add filter dropdown with checkbox checklist
feat: implement multi-cell clipboard copy and paste
feat: add local storage persistence with debounced auto-save
feat: add range selection with drag and Shift+arrow
feat: add find and replace panel (Ctrl+F)
feat: add right-click context menu
feat: add row and column resizing
feat: add freeze rows via context menu and header button
fix: push empty cells to bottom during sort
fix: resolve paste offset bug - split click and edit modes
fix: capture ref values before setState in resize handlers
fix: resolve lint errors - remove unused variables
```

---

## 🔧 Tech Stack

| | Purpose |
|--|---------|
| React 18 | UI and state management |
| Vite | Build tool |
| Vanilla JS | Spreadsheet engine (zero external spreadsheet libs) |
| CSS Custom Properties | Design tokens and theming |
| localStorage | Client-side persistence |
| Native paste event | Clipboard without permissions |
