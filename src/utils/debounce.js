// ─────────────────────────────────────────────────────────────
//  debounce.js
//  Returns a debounced version of fn that only fires after
//  `delay` ms of silence. Used to avoid saving on every keystroke.
// ─────────────────────────────────────────────────────────────

export function debounce(fn, delay) {
  let timer = null
  return function (...args) {
    clearTimeout(timer)
    timer = setTimeout(() => fn.apply(this, args), delay)
  }
}