# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

### Added
- **Custom number format input** — "Custom…" entry in the Number Format toolbar dropdown opens an inline text field for arbitrary format strings (e.g. `#,##0.00`, `0.00E+00`)
- **Date format rendering** — `formatNumber` now converts Excel serial dates to human-readable strings using `yyyy`, `mm`, `dd`, `hh`, `mi`, `ss` tokens
- **Scientific notation format** — `0.00E+00` format string and toolbar preset render numbers in scientific notation
- **Freeze pane rendering fixes** — `SelectionLayer` now uses freeze-aware screen coordinates so the selection highlight and header highlights are drawn correctly over frozen rows/columns
- **Scroll-to-cell aware of frozen panes** — `ScrollManager.scrollToCell` skips scroll adjustment for cells in the frozen zone (they are always visible)

### Fixed
- `ROW()` / `COLUMN()` now return the actual row/column number of the cell being evaluated instead of always returning 1 (`Evaluator.currentCell` context)
- `OFFSET` and `INDIRECT` were always returning `#VALUE!` — both are now fully implemented
- Column widths and row heights lost on ODS save — now persisted as `table-column`/`table-row` styles in `content.xml`
- Merged cells: covered cells (`mergedInto`) not restored on re-read of an ODS file

### Changed
- `rowHeights` and `columnWidths` in `SheetModel` replaced with sparse `Map<number, number>` — memory per sheet reduced from ~520 KB to near-zero for sheets with default sizes
- Volatile functions (`RAND`, `NOW`, `TODAY`) re-evaluate every 30 seconds without a full dependency-graph rebuild

---

## [0.1.0] — 2026-01-01

### Added
- Core ODS read/write (JSZip + fast-xml-parser) with round-trip integration tests
- Canvas renderer with cells, column/row headers, gridlines, selection, and scroll
- Formula engine: tokenizer → parser → evaluator covering ~40 functions across math, text, logical, lookup, statistical, and datetime categories
- Dependency graph with topological recalculation order
- Command stack with undo/redo (20-command history)
- Multi-sheet support with tab bar (add, rename, delete, reorder)
- Cell formatting: bold, italic, underline, font size/family, text/background colour, alignment, borders, wrap text, number formats
- Merge cells (merge/unmerge, visual spanning)
- Freeze panes (freeze at current cell / unfreeze toggle)
- Sort (ascending/descending) and column filter (dropdown with value checkboxes)
- Conditional formatting rules (highlight cells by condition + colour)
- Data validation (list, number range) with inline dropdown for list type
- Named ranges (define, navigate, use in formulas)
- Find & Replace toolbar (Ctrl+F) with prev/next navigation and replace-one/replace-all
- Auto-fill handle (drag selection corner to extend a series)
- Fill Down (Ctrl+D) and Fill Right (Ctrl+R)
- Paste Special (values-only / formats-only / all)
- Copy/paste with formula offset adjustment
- Column resize by dragging the header border
- Row resize by dragging the row header border
- ESLint config, Prettier config, Gitea Actions CI pipeline
- Vitest test suite — 367 tests across 20 files

[Unreleased]: https://github.com/example/ods-editor/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/example/ods-editor/releases/tag/v0.1.0
