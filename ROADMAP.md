# ODS Editor — Roadmap

[Русская версия](ROADMAP.ru.md)

## Status legend
- ✅ Done  •  🔄 In progress  •  🗓 Planned  •  💡 Idea

---

## Phase 1 — Foundation (done / in progress)

| # | Item | Status | Notes |
|---|------|--------|-------|
| 1 | Core ODS read/write (JSZip + fast-xml-parser) | ✅ | Round-trip verified by integration tests |
| 2 | Canvas renderer (cells, headers, selection, scroll) | ✅ | CellLayer, HeaderLayer, GridLayer, SelectionLayer |
| 3 | Formula engine (tokenizer → parser → evaluator) | ✅ | ~40 functions across 6 modules |
| 4 | Dependency graph + topological recalc | ✅ | DependencyGraph.ts |
| 5 | Command stack (undo/redo, 20-command history) | ✅ | EditCommandStack with EditCommand interface |
| 6 | ESLint config + 0 lint errors | ✅ | .eslintrc.json; 51 warnings remain (non-null assertions) |
| 7 | Prettier config | ✅ | .prettierrc.json, .prettierignore |
| 8 | Gitea Actions CI (lint → build → test) | ✅ | .gitea/workflows/ci.yml |
| 9 | Test suite: 360 tests across 20 files | ✅ | Vitest; coverage config added |

---

## Phase 2 — Quality & correctness

| # | Item | Status | Priority |
|---|------|--------|----------|
| 10 | **RAND / NOW volatile recalc** | ✅ | Timer-based recalc every 30s; `FunctionRegistry.markVolatile()` |
| 11 | Column/row size persistence in ODS round-trip | ✅ | Written as `table-column`/`table-row` styles in `content.xml` |
| 12 | Covered-cell merge restoration on read | ✅ | `mergedInto` set on covered cells; serializer emits `covered-table-cell` |
| 13 | Row/col header resize (drag) in renderer | ✅ | `hitTestColResize`/`hitTestRowResize` → drag in InputManager → `resizeColumn`/`resizeRow` message |
| 14 | Fix `rowHeights` array memory (replace with Map) | ✅ | Sparse `Map<number,number>` — only non-default values stored |
| 15 | Freeze pane rendering | ✅ | SelectionLayer freeze-aware coords; ScrollManager skips frozen cells |

---

## Phase 3 — Features

| # | Item | Status | Priority |
|---|------|--------|----------|
| 16 | **Array formula support** `{=...}` | ✅ | Element-wise binary ops on `RangeValue`; `{=...}` braces stripped before parse |
| 17 | Named range UI (define/navigate) | ✅ | "NR" toolbar button opens define/navigate dialog |
| 18 | OFFSET / INDIRECT full implementation | ✅ | OFFSET uses `currentCell` base; INDIRECT parses A1/Sheet1.A1 |
| 19 | ROW / COLUMN with actual cell context | ✅ | `Evaluator.currentCell` set before each evaluation |
| 20 | Cell format editor (number format strings) | ✅ | Custom format input + date/scientific support in `formatNumber` |
| 21 | Find & Replace dialog | ✅ | Already in toolbar (Ctrl+F); replace one / replace all |
| 22 | Chart rendering (basic bar/line/pie) | 💡 | Low |
| 23 | Conditional formatting | 💡 | Low |

---

## Phase 4 — Developer experience

| # | Item | Status | Priority |
|---|------|--------|----------|
| 24 | Prettier formatting pass on all source files | ✅ | `npx prettier --write "src/**/*.ts" "webview/**/*.ts" ...` |
| 25 | E2E tests (VSCode Extension Test Runner) | 🗓 | Medium |
| 26 | CHANGELOG.md (Keep a Changelog format) | ✅ | CHANGELOG.md with v0.1.0 baseline and Unreleased section |
| 27 | Coverage badge in README | 🗓 | Low — lcov reporter ready |
| 28 | Semantic versioning + release workflow | 🗓 | Low |

---

## Known issues (bugs)

| # | Description | Severity |
|---|-------------|----------|
| B1 | ~~`RAND()` / `NOW()` / `TODAY()` not re-evaluated on each recalc~~ | ~~High~~ Fixed |
| B2 | ~~Column widths and row heights lost on ODS save~~ | ~~High~~ Fixed |
| B3 | ~~Merged cells: covered cells not restored on re-read~~ | ~~Medium~~ Fixed |
| B4 | ~~`OFFSET` / `INDIRECT` always return `#VALUE!`~~ | ~~Medium~~ Fixed |
| B5 | ~~`ROW()` / `COLUMN()` always return 1 (no cell context)~~ | ~~Low~~ Fixed |
