import * as vscode from 'vscode';
import { SpreadsheetModel } from './model/SpreadsheetModel';
import { EditCommand } from './model/commands';

export class OdsDocument implements vscode.CustomDocument {
  private readonly _uri: vscode.Uri;
  private readonly _model: SpreadsheetModel;
  private _savedVersion = 0;
  private _currentVersion = 0;

  private readonly _onDidDispose = new vscode.EventEmitter<void>();
  public readonly onDidDispose = this._onDidDispose.event;

  private readonly _onDidChangeDocument = new vscode.EventEmitter<{
    readonly edit: EditCommand;
    readonly undo: () => void;
    readonly redo: () => void;
  }>();
  public readonly onDidChangeDocument = this._onDidChangeDocument.event;

  /** Fires after an undo or redo operation so the editor can refresh the webview. */
  private readonly _onDidUndoRedo = new vscode.EventEmitter<void>();
  public readonly onDidUndoRedo = this._onDidUndoRedo.event;

  constructor(uri: vscode.Uri, model: SpreadsheetModel) {
    this._uri = uri;
    this._model = model;
  }

  get uri(): vscode.Uri {
    return this._uri;
  }

  get model(): SpreadsheetModel {
    return this._model;
  }

  get isDirty(): boolean {
    return this._currentVersion !== this._savedVersion;
  }

  applyEdit(command: EditCommand): void {
    command.execute(this._model);
    this._currentVersion++;
    this._onDidChangeDocument.fire({
      edit: command,
      undo: () => {
        command.undo(this._model);
        this._currentVersion--;
        this._onDidUndoRedo.fire();
      },
      redo: () => {
        command.execute(this._model);
        this._currentVersion++;
        this._onDidUndoRedo.fire();
      },
    });
  }

  markSaved(): void {
    this._savedVersion = this._currentVersion;
  }

  dispose(): void {
    this._onDidDispose.fire();
    this._onDidDispose.dispose();
    this._onDidChangeDocument.dispose();
    this._onDidUndoRedo.dispose();
  }
}
