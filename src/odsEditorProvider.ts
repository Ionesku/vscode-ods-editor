import * as vscode from 'vscode';
import { OdsDocument } from './odsDocument';
import { OdsReader } from './ods/OdsReader';
import { OdsWriter } from './ods/OdsWriter';
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
} from './model/commands';
import { CellStyle } from './model/types';
import { ExtToWebviewMessage, WebviewToExtMessage } from './messages';
import { FormulaEngine } from './formula/FormulaEngine';

export class OdsEditorProvider implements vscode.CustomEditorProvider<OdsDocument> {
  private static readonly viewType = 'odsEditor.spreadsheet';

  private readonly _onDidChangeCustomDocument = new vscode.EventEmitter<
    vscode.CustomDocumentEditEvent<OdsDocument>
  >();
  public readonly onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;

  /** Per-document formula engines, keyed by document URI string */
  private readonly formulaEngines = new Map<string, FormulaEngine>();

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

    webviewPanel.webview.html = this.getHtmlForWebview(webviewPanel.webview);

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

    // Volatile recalc timer: re-evaluate RAND/NOW/TODAY every 30 seconds
    const engine = this.getEngine(document);
    const volatileTimer = setInterval(() => {
      if (!engine.hasVolatileCells()) return;
      engine.recalcVolatile(document.model);
      this.sendModelSnapshot(webviewPanel.webview, document);
    }, 30_000);

    webviewPanel.onDidDispose(() => {
      clearInterval(volatileTimer);
      this.formulaEngines.delete(document.uri.toString());
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
      case 'ready':
        this.postMessage(webview, { type: 'init', data: model.serialize() });
        break;

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
            const val = cell.computedValue ?? cell.rawValue;
            valuesSet.add(val !== null ? String(val) : '');
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
    }
  }

  // ── Formula recalculation ─────────────────────────────────────────────────

  private recalcAll(document: OdsDocument): void {
    this.getEngine(document).recalcAll(document.model);
  }

  // ── Webview HTML ──────────────────────────────────────────────────────────

  private getHtmlForWebview(webview: vscode.Webview): string {
    const scriptUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'webview.js'),
    );
    const nonce = getNonce();

    return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'nonce-${nonce}'; style-src 'unsafe-inline';">
  <title>ODS Editor</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    html, body { width: 100%; height: 100%; overflow: hidden; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; font-size: 13px; }
    body { display: flex; flex-direction: column; background: var(--vscode-editor-background, #1e1e1e); color: var(--vscode-editor-foreground, #d4d4d4); }
    #toolbar { display: flex; align-items: center; gap: 4px; padding: 4px 8px; border-bottom: 1px solid var(--vscode-panel-border, #333); background: var(--vscode-editorWidget-background, #252526); flex-shrink: 0; flex-wrap: wrap; min-height: 32px; }
    #toolbar button { background: var(--vscode-button-secondaryBackground, #3a3d41); color: var(--vscode-button-secondaryForeground, #ccc); border: 1px solid transparent; padding: 2px 8px; cursor: pointer; font-size: 12px; border-radius: 2px; min-width: 28px; height: 24px; }
    #toolbar button:hover { background: var(--vscode-button-secondaryHoverBackground, #45494e); }
    #toolbar button.active { background: var(--vscode-button-background, #0e639c); color: var(--vscode-button-foreground, #fff); }
    #toolbar .separator { width: 1px; height: 20px; background: var(--vscode-panel-border, #555); margin: 0 4px; }
    #toolbar input[type="color"] { width: 28px; height: 24px; border: none; padding: 0; cursor: pointer; background: transparent; }
    #formula-bar { display: flex; align-items: center; padding: 2px 8px; border-bottom: 1px solid var(--vscode-panel-border, #333); background: var(--vscode-editor-background, #1e1e1e); flex-shrink: 0; height: 28px; }
    #cell-address { width: 80px; text-align: center; background: var(--vscode-input-background, #3c3c3c); color: var(--vscode-input-foreground, #ccc); border: 1px solid var(--vscode-input-border, #555); padding: 2px 4px; font-size: 12px; margin-right: 8px; }
    #formula-input { flex: 1; background: var(--vscode-input-background, #3c3c3c); color: var(--vscode-input-foreground, #ccc); border: 1px solid var(--vscode-input-border, #555); padding: 2px 4px; font-size: 12px; outline: none; }
    #formula-input:focus { border-color: var(--vscode-focusBorder, #007acc); }
    #canvas-container { flex: 1; position: relative; overflow: hidden; }
    #spreadsheet-canvas { position: absolute; top: 0; left: 0; }
    #cell-editor { position: absolute; display: none; border: 2px solid var(--vscode-focusBorder, #007acc); background: var(--vscode-input-background, #3c3c3c); color: var(--vscode-input-foreground, #ccc); padding: 1px 3px; font-size: 13px; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; outline: none; resize: none; z-index: 10; overflow: hidden; }
    #bottom-bar { display: flex; align-items: center; border-top: 1px solid var(--vscode-panel-border, #333); background: var(--vscode-editorWidget-background, #252526); flex-shrink: 0; height: 28px; }
    #sheet-tabs { display: flex; align-items: center; padding: 0 8px; flex: 1; height: 28px; overflow-x: auto; }
    #status-bar { display: flex; align-items: center; padding: 0 12px; font-size: 12px; color: var(--vscode-descriptionForeground, #999); gap: 12px; white-space: nowrap; }
    #sheet-tabs .tab { padding: 4px 12px; cursor: pointer; font-size: 12px; border: 1px solid transparent; border-bottom: none; white-space: nowrap; }
    #sheet-tabs .tab:hover { background: var(--vscode-list-hoverBackground, #2a2d2e); }
    #sheet-tabs .tab.active { background: var(--vscode-editor-background, #1e1e1e); border-color: var(--vscode-panel-border, #333); border-bottom: 1px solid var(--vscode-editor-background, #1e1e1e); margin-bottom: -1px; }
    #sheet-tabs .add-sheet { padding: 4px 8px; cursor: pointer; font-size: 14px; }
    #sheet-tabs .tab.drag-over { border-left: 2px solid var(--vscode-focusBorder, #007acc); }
    #sheet-tabs .add-sheet:hover { background: var(--vscode-list-hoverBackground, #2a2d2e); }
    .dropdown-wrap { position: relative; display: inline-block; }
    .border-dropdown { display: none; position: absolute; top: 100%; left: 0; background: var(--vscode-menu-background, #252526); border: 1px solid var(--vscode-menu-border, #454545); box-shadow: 0 2px 8px rgba(0,0,0,0.4); z-index: 50; padding: 4px; gap: 2px; flex-wrap: wrap; width: 120px; }
    .border-dropdown.open { display: flex; }
    .border-dropdown button { background: var(--vscode-button-secondaryBackground, #3a3d41); color: var(--vscode-button-secondaryForeground, #ccc); border: 1px solid transparent; padding: 4px; cursor: pointer; border-radius: 2px; width: 32px; height: 28px; display: flex; align-items: center; justify-content: center; }
    .border-dropdown button:hover { background: var(--vscode-button-secondaryHoverBackground, #45494e); }
    #context-menu { position: absolute; display: none; background: var(--vscode-menu-background, #252526); border: 1px solid var(--vscode-menu-border, #454545); box-shadow: 0 2px 8px rgba(0,0,0,0.4); z-index: 100; min-width: 160px; }
    #context-menu .menu-item { padding: 6px 20px; cursor: pointer; font-size: 12px; white-space: nowrap; }
    #context-menu .menu-item:hover { background: var(--vscode-menu-selectionBackground, #094771); color: var(--vscode-menu-selectionForeground, #fff); }
    #context-menu .menu-separator { height: 1px; background: var(--vscode-menu-separatorBackground, #454545); margin: 4px 0; }
    #find-bar { display: none; align-items: center; padding: 4px 8px; border-bottom: 1px solid var(--vscode-panel-border, #333); background: var(--vscode-editorWidget-background, #252526); flex-shrink: 0; gap: 4px; height: 32px; }
    #find-bar.open { display: flex; }
    #find-bar input { background: var(--vscode-input-background, #3c3c3c); color: var(--vscode-input-foreground, #ccc); border: 1px solid var(--vscode-input-border, #555); padding: 2px 6px; font-size: 12px; outline: none; min-width: 180px; }
    #find-bar input:focus { border-color: var(--vscode-focusBorder, #007acc); }
    #find-bar button { background: var(--vscode-button-secondaryBackground, #3a3d41); color: var(--vscode-button-secondaryForeground, #ccc); border: none; padding: 2px 8px; cursor: pointer; font-size: 12px; border-radius: 2px; height: 24px; }
    #find-bar button:hover { background: var(--vscode-button-secondaryHoverBackground, #45494e); }
    #find-bar .find-count { font-size: 11px; color: var(--vscode-descriptionForeground, #999); min-width: 60px; }
    #filter-dropdown { display: none; position: absolute; background: var(--vscode-menu-background, #252526); border: 1px solid var(--vscode-menu-border, #454545); box-shadow: 0 2px 8px rgba(0,0,0,0.5); z-index: 100; min-width: 180px; max-width: 260px; font-size: 12px; }
    #filter-dropdown.open { display: block; }
    .filter-header { padding: 6px 10px; border-bottom: 1px solid var(--vscode-menu-separatorBackground, #454545); }
    .filter-header label { cursor: pointer; }
    #filter-values { max-height: 200px; overflow-y: auto; padding: 4px 10px; }
    #filter-values label { display: block; padding: 2px 0; cursor: pointer; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .filter-actions { display: flex; gap: 4px; padding: 6px 10px; border-top: 1px solid var(--vscode-menu-separatorBackground, #454545); }
    .filter-actions button { flex: 1; background: var(--vscode-button-secondaryBackground, #3a3d41); color: var(--vscode-button-secondaryForeground, #ccc); border: none; padding: 4px; cursor: pointer; border-radius: 2px; font-size: 11px; }
    .filter-actions button:first-child { background: var(--vscode-button-background, #0e639c); color: var(--vscode-button-foreground, #fff); }
    #named-ranges-dialog { display: none; position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); background: var(--vscode-editorWidget-background, #252526); border: 1px solid var(--vscode-menu-border, #454545); box-shadow: 0 4px 16px rgba(0,0,0,0.5); z-index: 200; min-width: 320px; border-radius: 4px; }
    #named-ranges-dialog.open { display: block; }
    #nr-list .nr-item { display: flex; align-items: center; justify-content: space-between; padding: 3px 0; border-bottom: 1px solid var(--vscode-menu-separatorBackground, #353535); }
    #nr-list .nr-item .remove-nr { cursor: pointer; color: #f44; padding: 0 4px; }
    #cond-fmt-dialog { display: none; position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); background: var(--vscode-editorWidget-background, #252526); border: 1px solid var(--vscode-menu-border, #454545); box-shadow: 0 4px 16px rgba(0,0,0,0.5); z-index: 200; min-width: 300px; border-radius: 4px; }
    #cond-fmt-dialog.open { display: block; }
    .dialog-title { padding: 10px 14px; font-weight: bold; border-bottom: 1px solid var(--vscode-menu-separatorBackground, #454545); }
    .dialog-body { padding: 10px 14px; display: flex; flex-direction: column; gap: 8px; }
    .dialog-body label { display: flex; align-items: center; gap: 6px; font-size: 12px; }
    .dialog-body select, .dialog-body input[type="text"] { background: var(--vscode-input-background, #3c3c3c); color: var(--vscode-input-foreground, #ccc); border: 1px solid var(--vscode-input-border, #555); padding: 3px 6px; font-size: 12px; flex: 1; }
    .dialog-body input[type="color"] { width: 28px; height: 22px; border: none; padding: 0; cursor: pointer; }
    .dialog-actions { display: flex; gap: 6px; padding: 10px 14px; border-top: 1px solid var(--vscode-menu-separatorBackground, #454545); justify-content: flex-end; }
    .dialog-actions button { background: var(--vscode-button-secondaryBackground, #3a3d41); color: var(--vscode-button-secondaryForeground, #ccc); border: none; padding: 4px 14px; cursor: pointer; border-radius: 2px; font-size: 12px; }
    .dialog-actions button:first-child { background: var(--vscode-button-background, #0e639c); color: var(--vscode-button-foreground, #fff); }
    #cond-rules-list { padding: 6px 14px; max-height: 150px; overflow-y: auto; }
    #cond-rules-list .cond-rule { display: flex; align-items: center; justify-content: space-between; padding: 3px 0; font-size: 11px; border-bottom: 1px solid var(--vscode-menu-separatorBackground, #353535); }
    #cond-rules-list .cond-rule .remove-rule { cursor: pointer; color: #f44; padding: 0 4px; }
    .fmt-dropdown { display: none; position: absolute; top: 100%; left: 0; background: var(--vscode-menu-background, #252526); border: 1px solid var(--vscode-menu-border, #454545); box-shadow: 0 2px 8px rgba(0,0,0,0.4); z-index: 50; min-width: 160px; }
    .fmt-dropdown.open { display: block; }
    .fmt-dropdown button { display: block; width: 100%; text-align: left; background: transparent; color: var(--vscode-menu-foreground, #ccc); border: none; padding: 5px 12px; cursor: pointer; font-size: 12px; }
    .fmt-dropdown button:hover { background: var(--vscode-menu-selectionBackground, #094771); color: var(--vscode-menu-selectionForeground, #fff); }
    .fmt-separator { height: 1px; background: var(--vscode-menu-separatorBackground, #454545); margin: 4px 0; }
    #custom-fmt-row { display: none; padding: 6px 8px; gap: 4px; }
    #custom-fmt-row.visible { display: flex; }
    #custom-fmt-input { flex: 1; background: var(--vscode-input-background, #3c3c3c); color: var(--vscode-input-foreground, #ccc); border: 1px solid var(--vscode-input-border, #555); padding: 2px 4px; font-size: 11px; min-width: 0; }
    #custom-fmt-apply { background: var(--vscode-button-background, #0e639c); color: var(--vscode-button-foreground, #fff); border: none; padding: 2px 6px; cursor: pointer; font-size: 11px; border-radius: 2px; }
    #custom-fmt-preview { padding: 2px 12px 6px; font-size: 11px; color: var(--vscode-descriptionForeground, #aaa); font-style: italic; }
    #dv-dialog { display: none; position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); background: var(--vscode-editorWidget-background, #252526); border: 1px solid var(--vscode-menu-border, #454545); box-shadow: 0 4px 16px rgba(0,0,0,0.5); z-index: 200; min-width: 300px; border-radius: 4px; }
    #dv-dialog.open { display: block; }
    #dv-list { padding: 6px 14px; max-height: 120px; overflow-y: auto; font-size: 11px; }
    #dv-list .dv-item { display: flex; align-items: center; justify-content: space-between; padding: 3px 0; border-bottom: 1px solid var(--vscode-menu-separatorBackground, #353535); }
    #dv-list .dv-item .remove-dv { cursor: pointer; color: #f44; padding: 0 4px; }
    #dv-dropdown { display: none; position: absolute; background: var(--vscode-menu-background, #252526); border: 1px solid var(--vscode-menu-border, #454545); box-shadow: 0 2px 8px rgba(0,0,0,0.4); z-index: 150; min-width: 120px; max-height: 200px; overflow-y: auto; }
    #dv-dropdown.open { display: block; }
    #dv-dropdown .dv-option { padding: 5px 12px; cursor: pointer; font-size: 12px; }
    #dv-dropdown .dv-option:hover { background: var(--vscode-menu-selectionBackground, #094771); color: var(--vscode-menu-selectionForeground, #fff); }
  </style>
</head>
<body>
  <div id="toolbar">
    <button id="btn-bold" title="Bold (Ctrl+B)"><b>B</b></button>
    <button id="btn-italic" title="Italic (Ctrl+I)"><i>I</i></button>
    <button id="btn-underline" title="Underline (Ctrl+U)"><u>U</u></button>
    <div class="separator"></div>
    <input type="color" id="btn-text-color" title="Text Color" value="#000000">
    <input type="color" id="btn-bg-color" title="Background Color" value="#ffffff">
    <div class="separator"></div>
    <button id="btn-align-left" title="Align Left"><svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><rect x="2" y="3" width="12" height="1.5"/><rect x="2" y="6.5" width="8" height="1.5"/><rect x="2" y="10" width="12" height="1.5"/></svg></button>
    <button id="btn-align-center" title="Align Center"><svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><rect x="2" y="3" width="12" height="1.5"/><rect x="4" y="6.5" width="8" height="1.5"/><rect x="2" y="10" width="12" height="1.5"/></svg></button>
    <button id="btn-align-right" title="Align Right"><svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><rect x="2" y="3" width="12" height="1.5"/><rect x="6" y="6.5" width="8" height="1.5"/><rect x="2" y="10" width="12" height="1.5"/></svg></button>
    <div class="separator"></div>
    <div class="dropdown-wrap">
      <button id="btn-borders" title="Borders"><svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><rect x="1" y="1" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.5"/><line x1="8" y1="1" x2="8" y2="15" stroke="currentColor" stroke-width="1"/><line x1="1" y1="8" x2="15" y2="8" stroke="currentColor" stroke-width="1"/></svg></button>
      <div id="border-menu" class="border-dropdown">
        <button data-border="all" title="All Borders"><svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><rect x="1" y="1" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.5"/><line x1="8" y1="1" x2="8" y2="15" stroke="currentColor" stroke-width="1"/><line x1="1" y1="8" x2="15" y2="8" stroke="currentColor" stroke-width="1"/></svg></button>
        <button data-border="outer" title="Outer Borders"><svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><rect x="1" y="1" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2"/></svg></button>
        <button data-border="none" title="No Borders"><svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><rect x="1" y="1" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1" stroke-dasharray="2,2"/></svg></button>
        <button data-border="bottom" title="Bottom Border"><svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><line x1="1" y1="15" x2="15" y2="15" stroke="currentColor" stroke-width="2"/></svg></button>
        <button data-border="top" title="Top Border"><svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><line x1="1" y1="1" x2="15" y2="1" stroke="currentColor" stroke-width="2"/></svg></button>
        <button data-border="left" title="Left Border"><svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><line x1="1" y1="1" x2="1" y2="15" stroke="currentColor" stroke-width="2"/></svg></button>
        <button data-border="right" title="Right Border"><svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><line x1="15" y1="1" x2="15" y2="15" stroke="currentColor" stroke-width="2"/></svg></button>
      </div>
    </div>
    <div class="separator"></div>
    <div class="dropdown-wrap">
      <button id="btn-numfmt" title="Number Format">123</button>
      <div id="numfmt-menu" class="fmt-dropdown">
        <button data-fmt="general">General</button>
        <button data-fmt="number">Number (1,234.56)</button>
        <button data-fmt="currency">Currency ($1,234.56)</button>
        <button data-fmt="percent">Percent (12.34%)</button>
        <button data-fmt="date">Date (2024-01-15)</button>
        <button data-fmt="int">Integer (1,235)</button>
        <button data-fmt="sci">Scientific (1.23E+4)</button>
        <div class="fmt-separator"></div>
        <button data-fmt="custom">Custom…</button>
        <div id="custom-fmt-row">
          <input type="text" id="custom-fmt-input" placeholder="#,##0.00">
          <button id="custom-fmt-apply">OK</button>
        </div>
        <div id="custom-fmt-preview"></div>
      </div>
    </div>
    <div class="separator"></div>
    <button id="btn-wrap" title="Wrap Text"><svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><rect x="2" y="3" width="12" height="1.2"/><path d="M2 7.5h8a2.5 2.5 0 0 1 0 5H9" fill="none" stroke="currentColor" stroke-width="1.2"/><polyline points="10,11 8.5,12.5 10,14" fill="none" stroke="currentColor" stroke-width="1.2"/></svg></button>
    <button id="btn-freeze" title="Freeze Panes"><svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><rect x="1" y="1" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.2"/><line x1="6" y1="1" x2="6" y2="15" stroke="currentColor" stroke-width="1.5" stroke-dasharray="2,1"/><line x1="1" y1="6" x2="15" y2="6" stroke="currentColor" stroke-width="1.5" stroke-dasharray="2,1"/></svg></button>
    <div class="separator"></div>
    <button id="btn-sort-asc" title="Sort Ascending">A↓</button>
    <button id="btn-sort-desc" title="Sort Descending">Z↓</button>
    <button id="btn-filter" title="Toggle Filter"><svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><polygon points="1,2 15,2 10,8 10,13 6,14 6,8"/></svg></button>
    <button id="btn-cond-fmt" title="Conditional Formatting"><svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><rect x="1" y="1" width="14" height="14" rx="2" fill="none" stroke="currentColor" stroke-width="1.2"/><rect x="3" y="4" width="4" height="3" fill="#4CAF50" opacity="0.7"/><rect x="9" y="4" width="4" height="3" fill="#F44336" opacity="0.7"/><rect x="3" y="9" width="4" height="3" fill="#FF9800" opacity="0.7"/><rect x="9" y="9" width="4" height="3" fill="#2196F3" opacity="0.7"/></svg></button>
    <div class="separator"></div>
    <button id="btn-named-ranges" title="Named Ranges">NR</button>
    <button id="btn-data-validation" title="Data Validation">DV</button>
  </div>
  <div id="formula-bar">
    <input type="text" id="cell-address" readonly>
    <input type="text" id="formula-input" placeholder="Enter value or formula (=SUM(A1:A10))">
  </div>
  <div id="find-bar">
    <input type="text" id="find-input" placeholder="Find...">
    <span class="find-count" id="find-count"></span>
    <button id="find-prev" title="Previous">&#9650;</button>
    <button id="find-next" title="Next">&#9660;</button>
    <input type="text" id="replace-input" placeholder="Replace...">
    <button id="replace-one">Replace</button>
    <button id="replace-all">All</button>
    <button id="find-close" title="Close">&#10005;</button>
  </div>
  <div id="canvas-container">
    <canvas id="spreadsheet-canvas"></canvas>
    <textarea id="cell-editor"></textarea>
  </div>
  <div id="bottom-bar">
    <div id="sheet-tabs"></div>
    <div id="status-bar"></div>
  </div>
  <div id="context-menu"></div>
  <div id="filter-dropdown">
    <div class="filter-header">
      <label><input type="checkbox" id="filter-select-all" checked> Select All</label>
    </div>
    <div id="filter-values"></div>
    <div class="filter-actions">
      <button id="filter-apply">Apply</button>
      <button id="filter-clear">Clear</button>
      <button id="filter-close">Cancel</button>
    </div>
  </div>
  <div id="named-ranges-dialog">
    <div class="dialog-title">Named Ranges</div>
    <div class="dialog-body">
      <label>Name: <input type="text" id="nr-name" placeholder="e.g. SALES_DATA"></label>
      <label>Range: <input type="text" id="nr-range" placeholder="e.g. A1:B10 (uses active sheet)"></label>
    </div>
    <div class="dialog-actions">
      <button id="nr-define">Define</button>
      <button id="nr-close">Close</button>
    </div>
    <div id="nr-list" style="padding:6px 14px;max-height:150px;overflow-y:auto;font-size:11px;"></div>
  </div>
  <div id="dv-dialog">
    <div class="dialog-title">Data Validation</div>
    <div class="dialog-body">
      <label>Range: <input type="text" id="dv-range" placeholder="e.g. A1:A100 (uses active sheet)"></label>
      <label>Type:
        <select id="dv-type">
          <option value="list">List (dropdown)</option>
          <option value="none">None (remove)</option>
        </select>
      </label>
      <label id="dv-source-label">Items: <input type="text" id="dv-source" placeholder="Comma-separated: Yes,No,Maybe"></label>
    </div>
    <div class="dialog-actions">
      <button id="dv-apply">Apply</button>
      <button id="dv-close">Close</button>
    </div>
    <div id="dv-list"></div>
  </div>
  <div id="dv-dropdown"></div>
  <div id="cond-fmt-dialog">
    <div class="dialog-title">Conditional Formatting</div>
    <div class="dialog-body">
      <label>Condition:
        <select id="cond-type">
          <option value="greaterThan">Greater than</option>
          <option value="lessThan">Less than</option>
          <option value="equals">Equals</option>
          <option value="notEquals">Not equals</option>
          <option value="contains">Contains</option>
          <option value="notContains">Not contains</option>
          <option value="between">Between</option>
          <option value="empty">Is empty</option>
          <option value="notEmpty">Is not empty</option>
        </select>
      </label>
      <label>Value: <input type="text" id="cond-value1"></label>
      <label id="cond-value2-label" style="display:none">To: <input type="text" id="cond-value2"></label>
      <label>Background: <input type="color" id="cond-bg" value="#FFEB3B"></label>
      <label>Text color: <input type="color" id="cond-text" value="#000000"></label>
      <div style="display:flex;gap:8px">
        <label><input type="checkbox" id="cond-bold"> Bold</label>
        <label><input type="checkbox" id="cond-italic"> Italic</label>
      </div>
    </div>
    <div class="dialog-actions">
      <button id="cond-apply">Apply</button>
      <button id="cond-cancel">Cancel</button>
    </div>
    <div id="cond-rules-list"></div>
  </div>
  <script nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
  }

  private postMessage(webview: vscode.Webview, msg: ExtToWebviewMessage): void {
    webview.postMessage(msg);
  }

  private sendModelSnapshot(webview: vscode.Webview, document: OdsDocument): void {
    this.postMessage(webview, { type: 'modelSnapshot', data: document.model.serialize() });
  }
}

function getNonce(): string {
  let text = '';
  const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  for (let i = 0; i < 32; i++) {
    text += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return text;
}
