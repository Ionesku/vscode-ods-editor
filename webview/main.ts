import { WebviewState } from './state/WebviewState';
import { messageBridge } from './state/MessageBridge';
import { CanvasRenderer } from './renderer/CanvasRenderer';
import { SelectionManager } from './interaction/SelectionManager';
import { CellEditor } from './interaction/CellEditor';
import { InputManager } from './interaction/InputManager';
import { ContextMenu } from './interaction/ContextMenu';
import { getElement } from './ui/domUtils';
import { FormulaBar } from './ui/FormulaBar';
import { StatusBar } from './ui/StatusBar';
import { SheetTabs } from './ui/SheetTabs';
import { FindBar } from './ui/FindBar';
import { FilterDropdown } from './ui/FilterDropdown';
import { CondFmtDialog } from './ui/CondFmtDialog';
import { NamedRangesDialog } from './ui/NamedRangesDialog';
import { DvDialog } from './ui/DvDialog';
import { Toolbar } from './ui/Toolbar';
import { CommentDialog } from './ui/CommentDialog';
import type { DataValidation } from '../shared/types';
import type { ExtToWebviewMessage } from '../src/messages';

// ── State ───────────────────────────────────────────────────────────────────

const state = new WebviewState();

// ── Renderer ────────────────────────────────────────────────────────────────

const canvas = getElement<HTMLCanvasElement>('spreadsheet-canvas');
const renderer = new CanvasRenderer(canvas, state);

// ── Selection ───────────────────────────────────────────────────────────────
// NOTE: formulaBar, statusBar, dvDialog are referenced via closure — they are
// initialized below and will be set before any user interaction triggers callbacks.

const selection = new SelectionManager(state, () => {
  formulaBar.updateAddress();
  formulaBar.updateFormula();
  statusBar.update();
  renderer.markSelectionDirty();
  const isSingleCell =
    state.selectedCol === state.selectionEndCol && state.selectedRow === state.selectionEndRow;
  if (isSingleCell) {
    requestAnimationFrame(() => dvDialog.showCellDropdown(state.selectedCol, state.selectedRow));
  }
});

// ── Cell Editor ─────────────────────────────────────────────────────────────

const cellEditorEl = getElement<HTMLTextAreaElement>('cell-editor');
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
  () => { renderer.markDirty(); },
);

// ── Input Manager ───────────────────────────────────────────────────────────

const inputManager = new InputManager(canvas, state, renderer, selection, editor);

// ── Comment Dialog (declared before ContextMenu so the closure captures it) ──

const commentDialog = new CommentDialog(state, renderer, canvas);

// ── Context Menu ────────────────────────────────────────────────────────────

const contextMenuEl = getElement('context-menu');
new ContextMenu(
  contextMenuEl, state, renderer,
  {
    pasteValuesOnly: () => inputManager.pasteValuesOnly(),
    pasteFormatsOnly: () => inputManager.pasteFormatsOnly(),
    hasClipboard: () => inputManager.hasClipboard(),
  },
  { open: (col, row) => commentDialog.open(col, row) },
);

// ── UI Modules ──────────────────────────────────────────────────────────────

const formulaBar = new FormulaBar(state, renderer, selection);
const statusBar = new StatusBar(state);
const sheetTabs = new SheetTabs(state);
const findBar = new FindBar(state, renderer, selection);
const filterDropdown = new FilterDropdown(state, renderer);
const condFmtDialog = new CondFmtDialog(state);
const namedRangesDialog = new NamedRangesDialog(state);
const dvDialog = new DvDialog(state, renderer);
new Toolbar(state, renderer, condFmtDialog, namedRangesDialog, dvDialog, filterDropdown);

// ── Global keyboard shortcuts ────────────────────────────────────────────────

document.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && (e.key === 'f' || e.key === 'h')) {
    e.preventDefault();
    findBar.open();
  }
});

// ── Message Handling ─────────────────────────────────────────────────────────

messageBridge.onMessage((rawMsg: unknown) => {
  const msg = rawMsg as ExtToWebviewMessage;
  switch (msg.type) {
    case 'init':
    case 'modelSnapshot':
      state.loadFromSerialized(msg.data);
      namedRangesDialog.setNamedRanges(msg.data.namedRanges ?? []);
      dvDialog.setValidations(
        (msg.data.sheets as Array<{ dataValidations?: DataValidation[] }>).flatMap(
          (s) => s.dataValidations ?? [],
        ),
      );
      renderer.updateViewport();
      renderer.markDirty();
      sheetTabs.render();
      formulaBar.updateAddress();
      formulaBar.updateFormula();
      statusBar.update();
      break;

    case 'columnValues':
      filterDropdown.show(msg.column, msg.values);
      break;
  }
});

// ── Initialize ───────────────────────────────────────────────────────────────

function init(): void {
  renderer.resize();
  window.addEventListener('resize', () => renderer.resize());
  canvas.tabIndex = 0;
  canvas.focus();
  messageBridge.postMessage({ type: 'ready' });
}

init();
