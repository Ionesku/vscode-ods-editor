import * as vscode from 'vscode';
import { OdsDocument } from './odsDocument';
import { OdsReader } from './ods/OdsReader';
import { OdsWriter } from './ods/OdsWriter';
import { SpreadsheetModel } from './model/SpreadsheetModel';
import {
  SetCellValueCommand,
  SetCellStyleCommand,
  MergeCellsCommand,
  InsertRowsCommand,
  InsertColumnsCommand,
  DeleteRowsCommand,
  DeleteColumnsCommand,
  ResizeColumnCommand,
  ResizeRowCommand,
  SortRangeCommand,
  SortRangeMultiCommand,
  PasteSpecialCommand,
  MoveRangeCommand,
  SetCommentCommand,
} from './model/commands';
import { CellStyle } from './model/types';
import { ExtToWebviewMessage, WebviewToExtMessage } from './messages';
import { FormulaEngine } from './formula/FormulaEngine';
import { getHtmlForWebview } from './webviewHtml';

export class OdsEditorProvider implements vscode.CustomEditorProvider<OdsDocument> {
  private static readonly viewType = 'odsEditor.spreadsheet';

  private readonly _onDidChangeCustomDocument = new vscode.EventEmitter<
    vscode.CustomDocumentEditEvent<OdsDocument>
  >();
  public readonly onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;

  /** Per-document formula engines, keyed by document URI string */
  private readonly formulaEngines = new Map<string, FormulaEngine>();

  /** Per-document volatile recalc timers (one per document, not per panel) */
  private readonly volatileTimers = new Map<string, ReturnType<typeof setInterval>>();

  /** Active webview panels per document, for volatile recalc broadcast */
  private readonly documentPanels = new Map<string, Set<vscode.Webview>>();

  private getEngine(document: OdsDocument): FormulaEngine {
    const key = document.uri.toString();
    let engine = this.formulaEngines.get(key);
    if (!engine) {
      engine = new FormulaEngine();
      this.formulaEngines.set(key, engine);
    }
    return engine;
  }

  constructor(private readonly context: vscode.ExtensionContext) {}

  // ── Document lifecycle ──────────────────────────────────────────────────

  async openCustomDocument(
    uri: vscode.Uri,
    _openContext: vscode.CustomDocumentOpenContext,
    _token: vscode.CancellationToken,
  ): Promise<OdsDocument> {
    const data = await vscode.workspace.fs.readFile(uri);
    const reader = new OdsReader();
    let model: SpreadsheetModel;
    try {
      model = await reader.read(data);
    } catch (e) {
      // If parsing fails, create an empty model
      model = new SpreadsheetModel();
      vscode.window.showWarningMessage(`Failed to parse ODS file: ${e}`);
    }

    const doc = new OdsDocument(uri, model);

    // Evaluate all formulas using this document's engine
    this.recalcAll(doc);

    return doc;
  }

  async resolveCustomEditor(
    document: OdsDocument,
    webviewPanel: vscode.WebviewPanel,
    _token: vscode.CancellationToken,
  ): Promise<void> {
    webviewPanel.webview.options = {
      enableScripts: true,
      localResourceRoots: [vscode.Uri.joinPath(this.context.extensionUri, 'dist')],
    };

    webviewPanel.webview.html = getHtmlForWebview(webviewPanel.webview, this.context.extensionUri);

    // Listen for messages from webview
    webviewPanel.webview.onDidReceiveMessage((msg: WebviewToExtMessage) => {
      this.handleWebviewMessage(document, webviewPanel.webview, msg);
    });

    // Listen for document changes (edit applied)
    document.onDidChangeDocument((e) => {
      this._onDidChangeCustomDocument.fire({
        document,
        ...e,
      });
      // Send updated model to webview
      this.sendModelSnapshot(webviewPanel.webview, document);
    });

    // Listen for undo/redo — recalculate and refresh webview
    document.onDidUndoRedo(() => {
      this.recalcAll(document);
      this.sendModelSnapshot(webviewPanel.webview, document);
    });

    // Track panels per document for volatile broadcast
    const docKey = document.uri.toString();
    if (!this.documentPanels.has(docKey)) {
      this.documentPanels.set(docKey, new Set());
    }
    this.documentPanels.get(docKey)!.add(webviewPanel.webview);

    // Volatile recalc timer: one per document (not per panel), broadcasts to all open panels
    if (!this.volatileTimers.has(docKey)) {
      const engine = this.getEngine(document);
      const timer = setInterval(() => {
        if (!engine.hasVolatileCells()) return;
        engine.recalcVolatile(document.model);
        const panels = this.documentPanels.get(docKey);
        if (panels) {
          for (const webview of panels) {
            this.postMessage(webview, { type: 'modelSnapshot', data: document.model.serialize() });
          }
        }
      }, 30_000);
      this.volatileTimers.set(docKey, timer);
    }

    webviewPanel.onDidDispose(() => {
      const panels = this.documentPanels.get(docKey);
      if (panels) {
        panels.delete(webviewPanel.webview);
        if (panels.size === 0) {
          // Last panel for this document closed — clean up
          this.documentPanels.delete(docKey);
          const timer = this.volatileTimers.get(docKey);
          if (timer !== undefined) {
            clearInterval(timer);
            this.volatileTimers.delete(docKey);
          }
          this.formulaEngines.delete(docKey);
        }
      }
    });

    // Wait for webview ready, then send initial data
    // The webview sends 'ready' when it has loaded
  }

  async saveCustomDocument(
    document: OdsDocument,
    _cancellation: vscode.CancellationToken,
  ): Promise<void> {
    const writer = new OdsWriter();
    const data = await writer.write(document.model);
    await vscode.workspace.fs.writeFile(document.uri, data);
    document.markSaved();
  }

  async saveCustomDocumentAs(
    document: OdsDocument,
    destination: vscode.Uri,
    _cancellation: vscode.CancellationToken,
  ): Promise<void> {
    const writer = new OdsWriter();
    const data = await writer.write(document.model);
    await vscode.workspace.fs.writeFile(destination, data);
    document.markSaved();
  }

  async revertCustomDocument(
    document: OdsDocument,
    _cancellation: vscode.CancellationToken,
  ): Promise<void> {
    const data = await vscode.workspace.fs.readFile(document.uri);
    const reader = new OdsReader();
    const model = await reader.read(data);
    // Replace model contents
    Object.assign(document.model, model);
    this.recalcAll(document);
  }

  async backupCustomDocument(
    document: OdsDocument,
    context: vscode.CustomDocumentBackupContext,
    _cancellation: vscode.CancellationToken,
  ): Promise<vscode.CustomDocumentBackup> {
    const writer = new OdsWriter();
    const data = await writer.write(document.model);
    await vscode.workspace.fs.writeFile(context.destination, data);
    return {
      id: context.destination.toString(),
      delete: () => {
        vscode.workspace.fs.delete(context.destination).then(undefined, () => {});
      },
    };
  }

  // ── Message handling ──────────────────────────────────────────────────────

  private handleWebviewMessage(
    document: OdsDocument,
    webview: vscode.Webview,
    msg: WebviewToExtMessage,
  ): void {
    const model = document.model;

    switch (msg.type) {
      case 'ready': {
        const initData = model.serialize();
        initData.functionNames = this.getEngine(document).getFunctionNames();
        this.postMessage(webview, { type: 'init', data: initData });
        break;
      }

      case 'editCell': {
        const sheet = model.getSheetByName(msg.sheet);
        if (!sheet) break;

        // Prevent editing cells that are covered by a merge (non-origin cells)
        if (sheet.getCell(msg.col, msg.row).mergedInto !== null) break;

        const oldCell = sheet.getCell(msg.col, msg.row);

        let value: string | number | boolean | null = msg.value;
        let formula: string | null = null;

        if (typeof value === 'string' && value.startsWith('=')) {
          formula = value.substring(1);
          value = null; // Will be computed
        } else if (typeof value === 'string') {
          // Try to parse as number
          const num = Number(value);
          if (value !== '' && !isNaN(num) && isFinite(num)) {
            value = num;
          } else if (value === '') {
            value = null;
          } else if (value.toUpperCase() === 'TRUE') {
            value = true;
          } else if (value.toUpperCase() === 'FALSE') {
            value = false;
          }
        }

        const cmd = new SetCellValueCommand(msg.sheet, msg.col, msg.row, value, formula);
        document.applyEdit(cmd);

        // Recalculate formulas
        this.recalcAll(document);

        // Send updated cells
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'setStyle': {
        const styleId = 'auto-' + Date.now();
        const fullStyle: CellStyle = { id: styleId, ...msg.style };
        model.styles.set(styleId, fullStyle);
        const cmd = new SetCellStyleCommand(msg.sheet, msg.range, styleId);
        document.applyEdit(cmd);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'mergeCells': {
        const cmd = new MergeCellsCommand(msg.sheet, msg.range);
        document.applyEdit(cmd);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'unmergeCells': {
        // Unmerge: reset all cells in range to default merge state
        const uSheet = model.getSheetByName(msg.sheet);
        if (!uSheet) break;
        for (let r = msg.range.startRow; r <= msg.range.endRow; r++) {
          for (let c = msg.range.startCol; c <= msg.range.endCol; c++) {
            const cell = uSheet.getCell(c, r);
            if (cell.mergeColSpan > 1 || cell.mergeRowSpan > 1 || cell.mergedInto) {
              uSheet.setCell(c, r, {
                mergeColSpan: 1,
                mergeRowSpan: 1,
                mergedInto: null,
              });
            }
          }
        }
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'resizeColumn': {
        const cmd = new ResizeColumnCommand(msg.sheet, msg.col, msg.width);
        document.applyEdit(cmd);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'resizeRow': {
        const cmd = new ResizeRowCommand(msg.sheet, msg.row, msg.height);
        document.applyEdit(cmd);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'addSheet': {
        model.addSheet(msg.name);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'removeSheet': {
        model.removeSheet(msg.index);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'renameSheet': {
        model.renameSheet(msg.index, msg.name);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'switchSheet': {
        model.activeSheetIndex = msg.index;
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'sort': {
        const cmd = new SortRangeCommand(msg.sheet, msg.range, msg.column, msg.ascending);
        document.applyEdit(cmd);
        this.recalcAll(document);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'insertRows': {
        const cmd = new InsertRowsCommand(msg.sheet, msg.at, msg.count);
        document.applyEdit(cmd);
        this.recalcAll(document);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'insertColumns': {
        const cmd = new InsertColumnsCommand(msg.sheet, msg.at, msg.count);
        document.applyEdit(cmd);
        this.recalcAll(document);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'deleteRows': {
        const cmd = new DeleteRowsCommand(msg.sheet, msg.at, msg.count);
        document.applyEdit(cmd);
        this.recalcAll(document);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'deleteColumns': {
        const cmd = new DeleteColumnsCommand(msg.sheet, msg.at, msg.count);
        document.applyEdit(cmd);
        this.recalcAll(document);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'setFilter': {
        const sheet = model.getSheetByName(msg.sheet);
        if (!sheet) break;
        if (msg.criteria) {
          // Convert values from array/Set to Set for internal use
          const values =
            msg.criteria.values instanceof Set
              ? msg.criteria.values
              : new Set<string>(
                  Array.isArray(msg.criteria.values)
                    ? msg.criteria.values
                    : Array.from(msg.criteria.values),
                );
          sheet.filters.set(msg.column, { column: msg.column, values });
          // Apply filter: hide rows that don't match
          sheet.hiddenRows.clear();
          const usedRange = sheet.getUsedRange();
          if (usedRange) {
            for (let r = usedRange.startRow; r <= usedRange.endRow; r++) {
              let visible = true;
              for (const [col, criteria] of sheet.filters) {
                const cell = sheet.getCell(col, r);
                const val = String(cell.computedValue ?? cell.rawValue ?? '');
                if (criteria.values.size > 0 && !criteria.values.has(val)) {
                  visible = false;
                  break;
                }
              }
              if (!visible) sheet.hiddenRows.add(r);
            }
          }
        } else {
          sheet.filters.delete(msg.column);
          if (sheet.filters.size === 0) {
            sheet.hiddenRows.clear();
          }
        }
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'freezePanes': {
        const sheet = model.getSheetByName(msg.sheet);
        if (!sheet) break;
        sheet.frozenRows = msg.frozenRows;
        sheet.frozenCols = msg.frozenCols;
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'autoFitColumn': {
        break;
      }

      case 'getColumnValues': {
        const sheet = model.getSheetByName(msg.sheet);
        if (!sheet) break;
        const usedRange = sheet.getUsedRange();
        const valuesSet = new Set<string>();
        if (usedRange) {
          for (let r = usedRange.startRow; r <= usedRange.endRow; r++) {
            const cell = sheet.getCell(msg.column, r);
            valuesSet.add(String(cell.computedValue ?? cell.rawValue ?? ''));
          }
        }
        this.postMessage(webview, {
          type: 'columnValues',
          column: msg.column,
          values: Array.from(valuesSet).sort(),
        });
        break;
      }

      case 'setConditionalFormat': {
        const sheet = model.getSheetByName(msg.sheet);
        if (!sheet) break;
        const existIdx = sheet.conditionalFormats.findIndex((r: any) => r.id === msg.rule.id);
        if (existIdx >= 0) {
          sheet.conditionalFormats[existIdx] = msg.rule;
        } else {
          sheet.conditionalFormats.push(msg.rule);
        }
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'removeConditionalFormat': {
        const sheet = model.getSheetByName(msg.sheet);
        if (!sheet) break;
        sheet.conditionalFormats = sheet.conditionalFormats.filter((r: any) => r.id !== msg.ruleId);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'pasteSpecial': {
        const cmd = new PasteSpecialCommand(
          msg.sheet,
          msg.startCol,
          msg.startRow,
          msg.cells,
          msg.mode,
        );
        document.applyEdit(cmd);
        this.recalcAll(document);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'sortMulti': {
        const cmd = new SortRangeMultiCommand(msg.sheet, msg.range, msg.keys);
        document.applyEdit(cmd);
        this.recalcAll(document);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'reorderSheet': {
        model.reorderSheet(msg.fromIndex, msg.toIndex);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'defineNamedRange': {
        model.defineNamedRange(msg.name, msg.sheet, msg.range);
        this.recalcAll(document);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'deleteNamedRange': {
        model.deleteNamedRange(msg.name);
        this.recalcAll(document);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'setDataValidation': {
        const sheet = model.getSheetByName(msg.sheet);
        if (sheet) {
          const r = msg.validation.range;
          // Remove any validation that overlaps with the given range
          sheet.dataValidations = sheet.dataValidations.filter((v) => {
            const overlap =
              v.range.startCol <= r.endCol &&
              v.range.endCol >= r.startCol &&
              v.range.startRow <= r.endRow &&
              v.range.endRow >= r.startRow;
            return !overlap;
          });
          if (msg.validation.type !== 'none') {
            sheet.dataValidations.push(msg.validation);
          }
          this.sendModelSnapshot(webview, document);
        }
        break;
      }

      case 'removeDataValidation': {
        const sheet = model.getSheetByName(msg.sheet);
        if (sheet) {
          sheet.dataValidations = sheet.dataValidations.filter((v) => v.id !== msg.id);
          this.sendModelSnapshot(webview, document);
        }
        break;
      }

      case 'moveRange': {
        const cmd = new MoveRangeCommand(msg.sheet, msg.fromRange, msg.toCol, msg.toRow);
        document.applyEdit(cmd);
        this.recalcAll(document);
        this.sendModelSnapshot(webview, document);
        break;
      }

      case 'setComment': {
        const cmd = new SetCommentCommand(msg.sheet, msg.col, msg.row, msg.comment);
        document.applyEdit(cmd);
        this.sendModelSnapshot(webview, document);
        break;
      }
    }
  }

  // ── Formula recalculation ─────────────────────────────────────────────────

  private recalcAll(document: OdsDocument): void {
    this.getEngine(document).recalcAll(document.model);
  }

  private postMessage(webview: vscode.Webview, msg: ExtToWebviewMessage): void {
    webview.postMessage(msg);
  }

  private sendModelSnapshot(webview: vscode.Webview, document: OdsDocument): void {
    const data = document.model.serialize();
    data.functionNames = this.getEngine(document).getFunctionNames();
    this.postMessage(webview, { type: 'modelSnapshot', data });
  }
}
