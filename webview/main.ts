import { WebviewState } from './state/WebviewState';
import { messageBridge } from './state/MessageBridge';
import { CanvasRenderer } from './renderer/CanvasRenderer';
import { SelectionManager } from './interaction/SelectionManager';
import { CellEditor } from './interaction/CellEditor';
import { InputManager } from './interaction/InputManager';
import { ContextMenu } from './interaction/ContextMenu';
import { colToLetter, letterToCol } from '../shared/types';
import type { DataValidation } from '../shared/types';

// ── State ───────────────────────────────────────────────────────────────────

const state = new WebviewState();

// ── DOM Elements ────────────────────────────────────────────────────────────

const canvas = document.getElementById('spreadsheet-canvas') as HTMLCanvasElement;
const cellAddressInput = document.getElementById('cell-address') as HTMLInputElement;
const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
const cellEditorEl = document.getElementById('cell-editor') as HTMLTextAreaElement;
const sheetTabsEl = document.getElementById('sheet-tabs') as HTMLElement;
const contextMenuEl = document.getElementById('context-menu') as HTMLElement;
const statusBarEl = document.getElementById('status-bar') as HTMLElement;

// ── Renderer ────────────────────────────────────────────────────────────────

const renderer = new CanvasRenderer(canvas, state);

// ── Selection ───────────────────────────────────────────────────────────────

const selection = new SelectionManager(state, () => {
  updateAddressBar();
  updateFormulaBar();
  updateStatusBar();
  renderer.markDirty();
  // Show DV dropdown if cell has list validation (deferred so renderer positions correctly)
  const isSingleCell =
    state.selectedCol === state.selectionEndCol && state.selectedRow === state.selectionEndRow;
  if (isSingleCell) {
    requestAnimationFrame(() => showDvDropdown(state.selectedCol, state.selectedRow));
  } else {
    dvDropdown?.classList.remove('open');
  }
});

// ── Cell Editor ─────────────────────────────────────────────────────────────

const editor = new CellEditor(
  cellEditorEl,
  state,
  renderer,
  (value: string) => {
    const sheet = state.activeSheet;
    if (!sheet) return;
    messageBridge.postMessage({
      type: 'editCell',
      sheet: sheet.name,
      col: state.selectedCol,
      row: state.selectedRow,
      value,
    });
    selection.moveSelection(0, 1);
    renderer.markDirty();
  },
  () => {
    renderer.markDirty();
  },
);

// ── Input Manager ───────────────────────────────────────────────────────────

const inputManager = new InputManager(canvas, state, renderer, selection, editor);

// ── Context Menu ────────────────────────────────────────────────────────────

new ContextMenu(contextMenuEl, state, renderer, {
  pasteValuesOnly: () => inputManager.pasteValuesOnly(),
  pasteFormatsOnly: () => inputManager.pasteFormatsOnly(),
  hasClipboard: () => inputManager.hasClipboard(),
});

// ── Toolbar ─────────────────────────────────────────────────────────────────

function setupToolbar(): void {
  const btn = (id: string) => document.getElementById(id)!;

  btn('btn-bold').addEventListener('click', () => applyStyle({ bold: true }));
  btn('btn-italic').addEventListener('click', () => applyStyle({ italic: true }));
  btn('btn-underline').addEventListener('click', () => applyStyle({ underline: true }));

  (btn('btn-text-color') as HTMLInputElement).addEventListener('input', (e) => {
    applyStyle({ textColor: (e.target as HTMLInputElement).value });
  });
  (btn('btn-bg-color') as HTMLInputElement).addEventListener('input', (e) => {
    applyStyle({ backgroundColor: (e.target as HTMLInputElement).value });
  });

  btn('btn-align-left').addEventListener('click', () => applyStyle({ horizontalAlign: 'left' }));
  btn('btn-align-center').addEventListener('click', () =>
    applyStyle({ horizontalAlign: 'center' }),
  );
  btn('btn-align-right').addEventListener('click', () => applyStyle({ horizontalAlign: 'right' }));

  // Border dropdown
  setupDropdown('btn-borders', 'border-menu', (target) => {
    const type = target.dataset.border;
    if (type) applyBorder(type);
  });

  // Number format dropdown
  setupDropdown('btn-numfmt', 'numfmt-menu', (target) => {
    const fmt = target.dataset.fmt;
    if (!fmt) return;
    if (fmt === 'custom') {
      // Show inline custom format row without closing menu
      const customRow = document.getElementById('custom-fmt-row')!;
      const customInput = document.getElementById('custom-fmt-input') as HTMLInputElement;
      customRow.classList.toggle('visible');
      if (customRow.classList.contains('visible')) customInput.focus();
      return true; // keep menu open
    }
    applyNumberFormat(fmt);
  });

  // Custom format apply button
  (document.getElementById('custom-fmt-apply') as HTMLButtonElement).addEventListener(
    'click',
    () => {
      const input = document.getElementById('custom-fmt-input') as HTMLInputElement;
      const fmt = input.value.trim();
      if (fmt) applyStyle({ numberFormat: fmt });
      document.getElementById('numfmt-menu')!.classList.remove('open');
      document.getElementById('custom-fmt-row')!.classList.remove('visible');
      input.value = '';
    },
  );

  // Custom format live preview on input
  (document.getElementById('custom-fmt-input') as HTMLInputElement).addEventListener(
    'input',
    (e) => {
      const fmt = (e.target as HTMLInputElement).value.trim();
      const preview = document.getElementById('custom-fmt-preview')!;
      if (!fmt) {
        preview.textContent = '';
        return;
      }
      const cell = state.getCell(state.selectedCol, state.selectedRow);
      const val = cell?.computedValue ?? cell?.rawValue;
      if (typeof val === 'number') {
        try {
          preview.textContent = state.formatForPreview(val, fmt);
        } catch {
          preview.textContent = '';
        }
      } else {
        preview.textContent = '';
      }
    },
  );

  // Wrap text
  btn('btn-wrap').addEventListener('click', () => applyStyle({ wrapText: true }));

  // Freeze panes
  btn('btn-freeze').addEventListener('click', () => toggleFreeze());

  btn('btn-sort-asc').addEventListener('click', () => sortSelection(true));
  btn('btn-sort-desc').addEventListener('click', () => sortSelection(false));
  btn('btn-filter').addEventListener('click', () => toggleFilter());

  // Named Ranges dialog
  btn('btn-named-ranges').addEventListener('click', () => openNamedRangesDialog());
  // Data Validation dialog
  btn('btn-data-validation').addEventListener('click', () => openDvDialog());
}

function setupDropdown(
  btnId: string,
  menuId: string,
  onSelect: (target: HTMLElement) => boolean | void,
): void {
  const btnEl = document.getElementById(btnId)!;
  const menuEl = document.getElementById(menuId)!;
  btnEl.addEventListener('click', (e) => {
    e.stopPropagation();
    // Close other dropdowns
    document.querySelectorAll('.border-dropdown.open, .fmt-dropdown.open').forEach((el) => {
      if (el !== menuEl) el.classList.remove('open');
    });
    menuEl.classList.toggle('open');
  });
  menuEl.addEventListener('click', (e) => {
    const target = (e.target as HTMLElement).closest('button') as HTMLElement;
    if (!target) return;
    e.stopPropagation();
    const keepOpen = onSelect(target);
    if (!keepOpen) menuEl.classList.remove('open');
  });
}

// Close all dropdowns on outside click
document.addEventListener('click', () => {
  document
    .querySelectorAll('.border-dropdown.open, .fmt-dropdown.open')
    .forEach((el) => el.classList.remove('open'));
});

function applyBorder(type: string): void {
  const border = { color: '#000000', style: 'solid' as const, width: 'thin' as const };

  switch (type) {
    case 'all':
      applyStyle({
        borderTop: border,
        borderBottom: border,
        borderLeft: border,
        borderRight: border,
      });
      break;
    case 'outer':
      applyStyleToOuterBorders(border);
      break;
    case 'none':
      applyStyle({
        borderTop: undefined,
        borderBottom: undefined,
        borderLeft: undefined,
        borderRight: undefined,
      });
      break;
    case 'bottom':
      applyStyle({ borderBottom: border });
      break;
    case 'top':
      applyStyle({ borderTop: border });
      break;
    case 'left':
      applyStyle({ borderLeft: border });
      break;
    case 'right':
      applyStyle({ borderRight: border });
      break;
  }
}

function applyStyleToOuterBorders(border: { color: string; style: string; width: string }): void {
  const sheet = state.activeSheet;
  const range = state.selectionRange;
  if (!sheet || !range) return;

  for (let r = range.startRow; r <= range.endRow; r++) {
    for (let c = range.startCol; c <= range.endCol; c++) {
      const style: Record<string, unknown> = {};
      if (r === range.startRow) style.borderTop = border;
      if (r === range.endRow) style.borderBottom = border;
      if (c === range.startCol) style.borderLeft = border;
      if (c === range.endCol) style.borderRight = border;
      if (Object.keys(style).length > 0) {
        messageBridge.postMessage({
          type: 'setStyle',
          sheet: sheet.name,
          range: { sheet: sheet.name, startCol: c, startRow: r, endCol: c, endRow: r },
          style,
        });
      }
    }
  }
}

function applyNumberFormat(fmt: string): void {
  const formats: Record<string, string> = {
    general: '',
    number: '#,##0.00',
    currency: '$#,##0.00',
    percent: '0.00%',
    date: 'yyyy-mm-dd',
    int: '#,##0',
    sci: '0.00E+00',
  };
  applyStyle({ numberFormat: formats[fmt] ?? '' });
}

function toggleFreeze(): void {
  const sheet = state.activeSheet;
  if (!sheet) return;

  // Toggle: if already frozen at current cell, unfreeze. Otherwise freeze at current cell.
  if (sheet.frozenRows > 0 || sheet.frozenCols > 0) {
    messageBridge.postMessage({
      type: 'freezePanes',
      sheet: sheet.name,
      frozenRows: 0,
      frozenCols: 0,
    });
  } else {
    messageBridge.postMessage({
      type: 'freezePanes',
      sheet: sheet.name,
      frozenRows: state.selectedRow,
      frozenCols: state.selectedCol,
    });
  }
}

function applyStyle(style: Record<string, unknown>): void {
  const sheet = state.activeSheet;
  const range = state.selectionRange;
  if (!sheet || !range) return;
  messageBridge.postMessage({
    type: 'setStyle',
    sheet: sheet.name,
    range,
    style,
  });
}

function sortSelection(ascending: boolean): void {
  const sheet = state.activeSheet;
  const range = state.selectionRange;
  if (!sheet || !range) return;
  messageBridge.postMessage({
    type: 'sort',
    sheet: sheet.name,
    range,
    column: state.selectedCol,
    ascending,
  });
}

function toggleFilter(): void {
  const sheet = state.activeSheet;
  if (!sheet) return;
  // Request unique values for current column, then show filter dropdown
  filterColumnIndex = state.selectedCol;
  messageBridge.postMessage({
    type: 'getColumnValues',
    sheet: sheet.name,
    column: state.selectedCol,
  });
}

// ── Filter Dropdown ─────────────────────────────────────────────────────────

const filterDropdown = document.getElementById('filter-dropdown')!;
const filterValuesEl = document.getElementById('filter-values')!;
const filterSelectAll = document.getElementById('filter-select-all') as HTMLInputElement;
let filterColumnIndex = 0;
let filterCheckedValues = new Set<string>();

function showFilterDropdown(column: number, values: string[]): void {
  filterColumnIndex = column;
  const sheet = state.activeSheet;
  const existingFilter = sheet?.activeFilters.get(column);

  filterValuesEl.innerHTML = '';
  filterCheckedValues = new Set(existingFilter ?? values);

  values.forEach((val) => {
    const label = document.createElement('label');
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = filterCheckedValues.has(val);
    cb.dataset.val = val;
    cb.addEventListener('change', () => {
      if (cb.checked) filterCheckedValues.add(val);
      else filterCheckedValues.delete(val);
      updateSelectAllState(values.length);
    });
    label.appendChild(cb);
    label.appendChild(document.createTextNode(' ' + (val || '(empty)')));
    filterValuesEl.appendChild(label);
  });

  filterSelectAll.checked = filterCheckedValues.size === values.length;
  filterSelectAll.indeterminate =
    filterCheckedValues.size > 0 && filterCheckedValues.size < values.length;
  filterSelectAll.onchange = () => {
    const cbs = filterValuesEl.querySelectorAll(
      'input[type=checkbox]',
    ) as NodeListOf<HTMLInputElement>;
    cbs.forEach((cb) => {
      cb.checked = filterSelectAll.checked;
      if (filterSelectAll.checked) filterCheckedValues.add(cb.dataset.val!);
      else filterCheckedValues.delete(cb.dataset.val!);
    });
  };

  // Position near the column header
  const vp = renderer.viewport;
  const scrollX = renderer.scrollManager.scrollX;
  const colX = 50 + vp.colLeft(column) - scrollX;
  filterDropdown.style.left = colX + 'px';
  filterDropdown.style.top = '56px'; // below toolbar + formula bar
  filterDropdown.classList.add('open');
}

function updateSelectAllState(total: number): void {
  filterSelectAll.checked = filterCheckedValues.size === total;
  filterSelectAll.indeterminate = filterCheckedValues.size > 0 && filterCheckedValues.size < total;
}

function setupFilterDropdown(): void {
  document.getElementById('filter-apply')!.addEventListener('click', () => {
    const sheet = state.activeSheet;
    if (!sheet) return;
    messageBridge.postMessage({
      type: 'setFilter',
      sheet: sheet.name,
      column: filterColumnIndex,
      criteria: { column: filterColumnIndex, values: Array.from(filterCheckedValues) },
    });
    filterDropdown.classList.remove('open');
  });
  document.getElementById('filter-clear')!.addEventListener('click', () => {
    const sheet = state.activeSheet;
    if (!sheet) return;
    messageBridge.postMessage({
      type: 'setFilter',
      sheet: sheet.name,
      column: filterColumnIndex,
      criteria: null,
    });
    filterDropdown.classList.remove('open');
  });
  document.getElementById('filter-close')!.addEventListener('click', () => {
    filterDropdown.classList.remove('open');
  });
}

// ── Conditional Format Dialog ───────────────────────────────────────────────

const condDialog = document.getElementById('cond-fmt-dialog')!;

function openCondFmtDialog(): void {
  condDialog.classList.add('open');
  renderCondRules();
  // Show/hide value2 for "between"
  const condType = document.getElementById('cond-type') as HTMLSelectElement;
  condType.onchange = () => {
    const show = condType.value === 'between';
    document.getElementById('cond-value2-label')!.style.display = show ? 'flex' : 'none';
  };
}

function renderCondRules(): void {
  const list = document.getElementById('cond-rules-list')!;
  const sheet = state.activeSheet;
  list.innerHTML = '';
  if (!sheet || sheet.conditionalFormats.length === 0) return;

  sheet.conditionalFormats.forEach((rule) => {
    const div = document.createElement('div');
    div.className = 'cond-rule';
    const label = document.createElement('span');
    const swatch = rule.style.backgroundColor
      ? `<span style="display:inline-block;width:12px;height:12px;background:${rule.style.backgroundColor};border:1px solid #666;margin-right:4px;vertical-align:middle"></span>`
      : '';
    label.innerHTML =
      swatch + rule.condition + ' ' + rule.value1 + (rule.value2 ? ' - ' + rule.value2 : '');
    div.appendChild(label);
    const removeBtn = document.createElement('span');
    removeBtn.className = 'remove-rule';
    removeBtn.textContent = '\u00d7';
    removeBtn.addEventListener('click', () => {
      messageBridge.postMessage({
        type: 'removeConditionalFormat',
        sheet: sheet.name,
        ruleId: rule.id,
      });
    });
    div.appendChild(removeBtn);
    list.appendChild(div);
  });
}

function setupCondFmtDialog(): void {
  document.getElementById('btn-cond-fmt')!.addEventListener('click', () => openCondFmtDialog());
  document.getElementById('cond-apply')!.addEventListener('click', () => {
    const sheet = state.activeSheet;
    const range = state.selectionRange;
    if (!sheet || !range) return;

    const condType = (document.getElementById('cond-type') as HTMLSelectElement).value as any;
    const value1 = (document.getElementById('cond-value1') as HTMLInputElement).value;
    const value2 = (document.getElementById('cond-value2') as HTMLInputElement).value;
    const bgColor = (document.getElementById('cond-bg') as HTMLInputElement).value;
    const textColor = (document.getElementById('cond-text') as HTMLInputElement).value;
    const bold = (document.getElementById('cond-bold') as HTMLInputElement).checked;
    const italic = (document.getElementById('cond-italic') as HTMLInputElement).checked;

    const rule = {
      id: 'cf-' + Date.now(),
      range,
      condition: condType,
      value1,
      value2: condType === 'between' ? value2 : undefined,
      style: {
        backgroundColor: bgColor,
        textColor,
        ...(bold ? { bold: true } : {}),
        ...(italic ? { italic: true } : {}),
      },
    };

    messageBridge.postMessage({
      type: 'setConditionalFormat',
      sheet: sheet.name,
      rule,
    });
    condDialog.classList.remove('open');
  });
  document.getElementById('cond-cancel')!.addEventListener('click', () => {
    condDialog.classList.remove('open');
  });
}

// ── Named Ranges Dialog ──────────────────────────────────────────────────────

let namedRanges: Array<{
  name: string;
  sheet: string;
  range: { startCol: number; startRow: number; endCol: number; endRow: number };
}> = [];

const namedRangesDialog = document.getElementById('named-ranges-dialog')!;

function openNamedRangesDialog(): void {
  namedRangesDialog.classList.add('open');
  renderNamedRangesList();
  // Pre-fill range from selection
  const range = state.selectionRange;
  if (range) {
    const addr = `${colToLetter(range.startCol)}${range.startRow + 1}:${colToLetter(range.endCol)}${range.endRow + 1}`;
    (document.getElementById('nr-range') as HTMLInputElement).value = addr;
  }
}

function renderNamedRangesList(): void {
  const list = document.getElementById('nr-list')!;
  list.innerHTML = '';
  for (const nr of namedRanges) {
    const div = document.createElement('div');
    div.className = 'nr-item';
    const span = document.createElement('span');
    span.textContent = `${nr.name}  →  ${nr.sheet}!${colToLetter(nr.range.startCol)}${nr.range.startRow + 1}:${colToLetter(nr.range.endCol)}${nr.range.endRow + 1}`;
    div.appendChild(span);
    const rm = document.createElement('span');
    rm.className = 'remove-nr';
    rm.textContent = '\u00d7';
    rm.addEventListener('click', () => {
      messageBridge.postMessage({ type: 'deleteNamedRange', name: nr.name });
    });
    div.appendChild(rm);
    list.appendChild(div);
  }
}

function setupNamedRangesDialog(): void {
  document.getElementById('nr-define')!.addEventListener('click', () => {
    const sheet = state.activeSheet;
    if (!sheet) return;
    const name = (document.getElementById('nr-name') as HTMLInputElement).value.trim();
    const rangeStr = (document.getElementById('nr-range') as HTMLInputElement).value
      .trim()
      .toUpperCase();
    if (!name) {
      alert('Enter a name');
      return;
    }

    // Parse range string like A1:B10
    const match = rangeStr.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!match) {
      alert('Invalid range format. Use A1:B10');
      return;
    }

    const startCol = letterToCol(match[1]);
    const startRow = parseInt(match[2]) - 1;
    const endCol = letterToCol(match[3]);
    const endRow = parseInt(match[4]) - 1;

    messageBridge.postMessage({
      type: 'defineNamedRange',
      name,
      sheet: sheet.name,
      range: { sheet: sheet.name, startCol, startRow, endCol, endRow },
    });
    (document.getElementById('nr-name') as HTMLInputElement).value = '';
  });

  document.getElementById('nr-close')!.addEventListener('click', () => {
    namedRangesDialog.classList.remove('open');
  });
}

// ── Data Validation ──────────────────────────────────────────────────────────

let dataValidations: DataValidation[] = [];

const dvDialog = document.getElementById('dv-dialog')!;
const dvDropdown = document.getElementById('dv-dropdown')!;

function openDvDialog(): void {
  dvDialog.classList.add('open');
  // Pre-fill range from selection
  const sc = Math.min(state.selectedCol, state.selectionEndCol);
  const ec = Math.max(state.selectedCol, state.selectionEndCol);
  const sr = Math.min(state.selectedRow, state.selectionEndRow);
  const er = Math.max(state.selectedRow, state.selectionEndRow);
  (document.getElementById('dv-range') as HTMLInputElement).value =
    `${colToLetter(sc)}${sr + 1}:${colToLetter(ec)}${er + 1}`;
  renderDvList();
}

function renderDvList(): void {
  const list = document.getElementById('dv-list')!;
  list.innerHTML = '';
  const sheet = state.activeSheet;
  if (!sheet) return;
  const sheetValidations = dataValidations.filter((v) => v.range.sheet === sheet.name);
  if (sheetValidations.length === 0) {
    list.innerHTML = '<div style="color:#888;padding:4px 0">No validations on this sheet</div>';
    return;
  }
  for (const dv of sheetValidations) {
    const div = document.createElement('div');
    div.className = 'dv-item';
    const r = dv.range;
    const rangeStr = `${colToLetter(r.startCol)}${r.startRow + 1}:${colToLetter(r.endCol)}${r.endRow + 1}`;
    div.innerHTML = `<span>${rangeStr} — ${dv.listSource}</span><span class="remove-dv" data-id="${dv.id}">✕</span>`;
    div.querySelector('.remove-dv')!.addEventListener('click', () => {
      const sh = state.activeSheet;
      if (!sh) return;
      messageBridge.postMessage({ type: 'removeDataValidation', sheet: sh.name, id: dv.id });
    });
    list.appendChild(div);
  }
}

function setupDvDialog(): void {
  const dvTypeEl = document.getElementById('dv-type') as HTMLSelectElement;
  const dvSourceLabel = document.getElementById('dv-source-label')!;

  dvTypeEl.addEventListener('change', () => {
    dvSourceLabel.style.display = dvTypeEl.value === 'list' ? '' : 'none';
  });

  document.getElementById('dv-apply')!.addEventListener('click', () => {
    const sheet = state.activeSheet;
    if (!sheet) return;
    const rangeStr = (document.getElementById('dv-range') as HTMLInputElement).value.trim();
    const type = dvTypeEl.value as 'list' | 'none';
    const source = (document.getElementById('dv-source') as HTMLInputElement).value.trim();

    // Parse range
    const match = rangeStr.match(/^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/);
    if (!match) return;
    const startCol = letterToCol(match[1].toUpperCase());
    const startRow = parseInt(match[2], 10) - 1;
    const endCol = match[3] ? letterToCol(match[3].toUpperCase()) : startCol;
    const endRow = match[4] ? parseInt(match[4], 10) - 1 : startRow;

    const validation: DataValidation = {
      id: `dv_${Date.now()}`,
      range: { sheet: sheet.name, startCol, startRow, endCol, endRow },
      type,
      listSource: source,
    };
    messageBridge.postMessage({ type: 'setDataValidation', sheet: sheet.name, validation });
  });

  document.getElementById('dv-close')!.addEventListener('click', () => {
    dvDialog.classList.remove('open');
  });
}

/** Show inline dropdown for list validation at the active cell */
function showDvDropdown(col: number, row: number): void {
  const sheet = state.activeSheet;
  if (!sheet) return;
  const dv = dataValidations.find(
    (v) =>
      v.type === 'list' &&
      v.range.sheet === sheet.name &&
      col >= v.range.startCol &&
      col <= v.range.endCol &&
      row >= v.range.startRow &&
      row <= v.range.endRow,
  );
  if (!dv) {
    dvDropdown.classList.remove('open');
    return;
  }

  const items = dv.listSource
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
  dvDropdown.innerHTML = '';
  for (const item of items) {
    const div = document.createElement('div');
    div.className = 'dv-option';
    div.textContent = item;
    div.addEventListener('mousedown', (e) => {
      e.preventDefault();
      messageBridge.postMessage({
        type: 'editCell',
        sheet: sheet.name,
        col,
        row,
        value: item,
      });
      dvDropdown.classList.remove('open');
    });
    dvDropdown.appendChild(div);
  }

  // Position below the cell using viewport calculator
  const { viewport, scrollManager } = renderer;
  const container = document.getElementById('canvas-container')!;
  const containerRect = container.getBoundingClientRect();
  const cellX = 50 + viewport.colLeft(col) - scrollManager.scrollX; // ROW_HEADER_WIDTH = 50
  const cellY = 24 + viewport.rowTop(row) - scrollManager.scrollY; // COL_HEADER_HEIGHT = 24
  const cellW = viewport.colWidth(col);
  const cellH = viewport.rowHeight(row);
  dvDropdown.style.position = 'fixed';
  dvDropdown.style.left = `${containerRect.left + cellX}px`;
  dvDropdown.style.top = `${containerRect.top + cellY + cellH}px`;
  dvDropdown.style.minWidth = `${cellW}px`;
  dvDropdown.classList.add('open');
}

document.addEventListener('click', (e) => {
  if (!dvDropdown.contains(e.target as Node)) {
    dvDropdown.classList.remove('open');
  }
});

// ── Find & Replace ──────────────────────────────────────────────────────────

const findBar = document.getElementById('find-bar')!;
const findInput = document.getElementById('find-input') as HTMLInputElement;
const replaceInput = document.getElementById('replace-input') as HTMLInputElement;
const findCountEl = document.getElementById('find-count')!;
let findMatches: Array<{ col: number; row: number }> = [];
let findMatchIndex = -1;

function openFindBar(): void {
  findBar.classList.add('open');
  findInput.focus();
  findInput.select();
}

function closeFindBar(): void {
  findBar.classList.remove('open');
  findMatches = [];
  findMatchIndex = -1;
  findCountEl.textContent = '';
  canvas.focus();
  renderer.markDirty();
}

function performFind(): void {
  const query = findInput.value.toLowerCase();
  findMatches = [];
  findMatchIndex = -1;

  if (!query || !state.activeSheet) {
    findCountEl.textContent = '';
    renderer.markDirty();
    return;
  }

  const sheet = state.activeSheet;
  for (const [key, cell] of sheet.cells) {
    const val = cell.computedValue ?? cell.rawValue;
    if (val !== null && String(val).toLowerCase().includes(query)) {
      const idx = key.indexOf(',');
      findMatches.push({
        col: parseInt(key.substring(0, idx), 10),
        row: parseInt(key.substring(idx + 1), 10),
      });
    }
  }

  // Sort by row, then col
  findMatches.sort((a, b) => a.row - b.row || a.col - b.col);

  if (findMatches.length > 0) {
    findMatchIndex = 0;
    navigateToMatch();
  }
  updateFindCount();
  renderer.markDirty();
}

function navigateToMatch(): void {
  if (findMatchIndex < 0 || findMatchIndex >= findMatches.length) return;
  const m = findMatches[findMatchIndex];
  selection.selectCell(m.col, m.row);
  const activeSheet = state.activeSheet;
  renderer.scrollManager.scrollToCell(
    m.col,
    m.row,
    activeSheet?.frozenCols ?? 0,
    activeSheet?.frozenRows ?? 0,
  );
  renderer.markDirty();
  updateFindCount();
}

function findNext(): void {
  if (findMatches.length === 0) return;
  findMatchIndex = (findMatchIndex + 1) % findMatches.length;
  navigateToMatch();
}

function findPrev(): void {
  if (findMatches.length === 0) return;
  findMatchIndex = (findMatchIndex - 1 + findMatches.length) % findMatches.length;
  navigateToMatch();
}

function replaceOne(): void {
  if (findMatchIndex < 0 || findMatchIndex >= findMatches.length) return;
  const sheet = state.activeSheet;
  if (!sheet) return;
  const m = findMatches[findMatchIndex];
  messageBridge.postMessage({
    type: 'editCell',
    sheet: sheet.name,
    col: m.col,
    row: m.row,
    value: replaceInput.value,
  });
  findMatches.splice(findMatchIndex, 1);
  if (findMatchIndex >= findMatches.length) findMatchIndex = 0;
  if (findMatches.length > 0) navigateToMatch();
  updateFindCount();
}

function replaceAll(): void {
  const sheet = state.activeSheet;
  if (!sheet || findMatches.length === 0) return;
  const replacement = replaceInput.value;
  const query = findInput.value.toLowerCase();

  for (const m of findMatches) {
    const cell = state.getCell(m.col, m.row);
    const val = cell.computedValue ?? cell.rawValue;
    if (val === null) continue;
    const str = String(val);
    const newVal = str.replace(new RegExp(escapeRegex(query), 'gi'), replacement);
    messageBridge.postMessage({
      type: 'editCell',
      sheet: sheet.name,
      col: m.col,
      row: m.row,
      value: newVal,
    });
  }
  findMatches = [];
  findMatchIndex = -1;
  updateFindCount();
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function updateFindCount(): void {
  if (findMatches.length === 0) {
    findCountEl.textContent = findInput.value ? '0 results' : '';
  } else {
    findCountEl.textContent = `${findMatchIndex + 1} of ${findMatches.length}`;
  }
}

function setupFindBar(): void {
  findInput.addEventListener('input', () => performFind());
  findInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      e.shiftKey ? findPrev() : findNext();
    }
    if (e.key === 'Escape') {
      e.preventDefault();
      closeFindBar();
    }
  });
  replaceInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      replaceOne();
    }
    if (e.key === 'Escape') {
      e.preventDefault();
      closeFindBar();
    }
  });
  document.getElementById('find-next')!.addEventListener('click', findNext);
  document.getElementById('find-prev')!.addEventListener('click', findPrev);
  document.getElementById('replace-one')!.addEventListener('click', replaceOne);
  document.getElementById('replace-all')!.addEventListener('click', replaceAll);
  document.getElementById('find-close')!.addEventListener('click', closeFindBar);
}

// Ctrl+F / Ctrl+H — expose to InputManager via global keydown
document.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && (e.key === 'f' || e.key === 'h')) {
    e.preventDefault();
    openFindBar();
  }
});

// ── Formula Bar ─────────────────────────────────────────────────────────────

formulaInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    const sheet = state.activeSheet;
    if (!sheet) return;
    messageBridge.postMessage({
      type: 'editCell',
      sheet: sheet.name,
      col: state.selectedCol,
      row: state.selectedRow,
      value: formulaInput.value,
    });
    canvas.focus();
  } else if (e.key === 'Escape') {
    e.preventDefault();
    updateFormulaBar();
    canvas.focus();
  }
});

function updateAddressBar(): void {
  cellAddressInput.value = colToLetter(state.selectedCol) + (state.selectedRow + 1);
}

// Address bar navigation — type "B5" or "A1:C10" and press Enter
cellAddressInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    const raw = cellAddressInput.value.trim().toUpperCase();
    const rangeMatch = raw.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    const cellMatch = raw.match(/^([A-Z]+)(\d+)$/);
    if (rangeMatch) {
      const c1 = addressLetterToCol(rangeMatch[1]);
      const r1 = parseInt(rangeMatch[2], 10) - 1;
      const c2 = addressLetterToCol(rangeMatch[3]);
      const r2 = parseInt(rangeMatch[4], 10) - 1;
      if (c1 >= 0 && r1 >= 0 && c2 >= 0 && r2 >= 0) {
        selection.selectCell(c1, r1);
        selection.selectCell(c2, r2, true);
        renderer.scrollManager.scrollToCell(
          c1,
          r1,
          state.activeSheet?.frozenCols ?? 0,
          state.activeSheet?.frozenRows ?? 0,
        );
        renderer.markDirty();
      }
    } else if (cellMatch) {
      const c = addressLetterToCol(cellMatch[1]);
      const r = parseInt(cellMatch[2], 10) - 1;
      if (c >= 0 && r >= 0) {
        selection.selectCell(c, r);
        renderer.scrollManager.scrollToCell(
          c,
          r,
          state.activeSheet?.frozenCols ?? 0,
          state.activeSheet?.frozenRows ?? 0,
        );
        renderer.markDirty();
      }
    }
    canvas.focus();
  } else if (e.key === 'Escape') {
    updateAddressBar();
    canvas.focus();
  }
});

function addressLetterToCol(letters: string): number {
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col - 1;
}

function updateFormulaBar(): void {
  const cell = state.getCell(state.selectedCol, state.selectedRow);
  if (cell.formula) {
    formulaInput.value = '=' + cell.formula;
  } else {
    formulaInput.value = cell.rawValue !== null ? String(cell.rawValue) : '';
  }
}

// ── Status Bar (SUM / AVG / COUNT) ──────────────────────────────────────────

function updateStatusBar(): void {
  if (!statusBarEl) return;
  const range = state.selectionRange;
  if (!range) {
    statusBarEl.textContent = '';
    return;
  }

  const isSingle = range.startCol === range.endCol && range.startRow === range.endRow;
  if (isSingle) {
    statusBarEl.textContent = '';
    return;
  }

  let sum = 0;
  let count = 0;
  let numCount = 0;

  for (let r = range.startRow; r <= range.endRow; r++) {
    for (let c = range.startCol; c <= range.endCol; c++) {
      const cell = state.getCell(c, r);
      const val = cell.computedValue ?? cell.rawValue;
      if (val !== null && val !== '') {
        count++;
        const num = typeof val === 'number' ? val : Number(val);
        if (!isNaN(num) && typeof val !== 'boolean') {
          sum += num;
          numCount++;
        }
      }
    }
  }

  const parts: string[] = [];
  if (numCount > 0) {
    parts.push(`Sum: ${formatNum(sum)}`);
    parts.push(`Avg: ${formatNum(sum / numCount)}`);
  }
  parts.push(`Count: ${count}`);
  statusBarEl.textContent = parts.join('   ');
}

function formatNum(n: number): string {
  if (Number.isInteger(n)) return String(n);
  return n.toFixed(2);
}

// ── Sheet Tabs ──────────────────────────────────────────────────────────────

let dragSrcTabIndex: number | null = null;

function renderSheetTabs(): void {
  sheetTabsEl.innerHTML = '';
  state.sheets.forEach((sheet, index) => {
    const tab = document.createElement('div');
    tab.className = 'tab' + (index === state.activeSheetIndex ? ' active' : '');
    tab.textContent = sheet.name;
    tab.draggable = true;
    tab.dataset.index = String(index);

    tab.addEventListener('click', () => {
      messageBridge.postMessage({ type: 'switchSheet', index });
    });
    tab.addEventListener('dblclick', () => {
      const name = prompt('Rename sheet:', sheet.name);
      if (name) {
        messageBridge.postMessage({ type: 'renameSheet', index, name });
      }
    });
    tab.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      if (confirm(`Delete sheet "${sheet.name}"?`)) {
        messageBridge.postMessage({ type: 'removeSheet', index });
      }
    });

    // Drag-and-drop reorder
    tab.addEventListener('dragstart', (e) => {
      dragSrcTabIndex = index;
      tab.style.opacity = '0.5';
      e.dataTransfer!.effectAllowed = 'move';
    });
    tab.addEventListener('dragend', () => {
      tab.style.opacity = '';
      sheetTabsEl
        .querySelectorAll('.tab')
        .forEach((t) => (t as HTMLElement).classList.remove('drag-over'));
    });
    tab.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.dataTransfer!.dropEffect = 'move';
      sheetTabsEl
        .querySelectorAll('.tab')
        .forEach((t) => (t as HTMLElement).classList.remove('drag-over'));
      tab.classList.add('drag-over');
    });
    tab.addEventListener('drop', (e) => {
      e.preventDefault();
      tab.classList.remove('drag-over');
      if (dragSrcTabIndex !== null && dragSrcTabIndex !== index) {
        messageBridge.postMessage({
          type: 'reorderSheet',
          fromIndex: dragSrcTabIndex,
          toIndex: index,
        });
        dragSrcTabIndex = null;
      }
    });

    sheetTabsEl.appendChild(tab);
  });

  const addBtn = document.createElement('div');
  addBtn.className = 'add-sheet';
  addBtn.textContent = '+';
  addBtn.title = 'Add Sheet';
  addBtn.addEventListener('click', () => {
    const name = `Sheet${state.sheets.length + 1}`;
    messageBridge.postMessage({ type: 'addSheet', name });
  });
  sheetTabsEl.appendChild(addBtn);
}

// ── Message Handling ────────────────────────────────────────────────────────

messageBridge.onMessage((msg: any) => {
  switch (msg.type) {
    case 'init':
    case 'modelSnapshot':
      state.loadFromSerialized(msg.data);
      namedRanges = msg.data.namedRanges ?? [];
      dataValidations = msg.data.sheets.flatMap(
        (s: { dataValidations?: DataValidation[] }) => s.dataValidations ?? [],
      );
      renderer.updateViewport();
      renderer.markDirty();
      renderSheetTabs();
      updateAddressBar();
      updateFormulaBar();
      updateStatusBar();
      // Refresh named ranges list if dialog is open
      if (namedRangesDialog.classList.contains('open')) {
        renderNamedRangesList();
      }
      if (dvDialog.classList.contains('open')) {
        renderDvList();
      }
      break;

    case 'columnValues':
      showFilterDropdown(msg.column, msg.values);
      break;
  }
});

// ── Initialize ──────────────────────────────────────────────────────────────

function init(): void {
  setupToolbar();
  setupFindBar();
  setupFilterDropdown();
  setupCondFmtDialog();
  setupNamedRangesDialog();
  setupDvDialog();
  renderer.resize();

  window.addEventListener('resize', () => {
    renderer.resize();
  });

  canvas.tabIndex = 0;
  canvas.focus();

  messageBridge.postMessage({ type: 'ready' });
}

init();
