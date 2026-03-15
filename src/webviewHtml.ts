import * as vscode from 'vscode';

export function getHtmlForWebview(
  webview: vscode.Webview,
  extensionUri: vscode.Uri,
): string {
  const scriptUri = webview.asWebviewUri(
    vscode.Uri.joinPath(extensionUri, 'dist', 'webview.js'),
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
    #comment-dialog { display: none; position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); background: var(--vscode-editorWidget-background, #252526); border: 1px solid var(--vscode-menu-border, #454545); box-shadow: 0 4px 16px rgba(0,0,0,0.5); z-index: 200; min-width: 280px; border-radius: 4px; }
    #comment-dialog.open { display: block; }
    #comment-text { width: 100%; min-height: 80px; resize: vertical; background: var(--vscode-input-background, #3c3c3c); color: var(--vscode-input-foreground, #ccc); border: 1px solid var(--vscode-input-border, #555); padding: 4px 6px; font-size: 12px; font-family: inherit; box-sizing: border-box; }
    #comment-tooltip { display: none; position: fixed; background: var(--vscode-editorHoverWidget-background, #252526); border: 1px solid var(--vscode-editorHoverWidget-border, #454545); box-shadow: 0 2px 8px rgba(0,0,0,0.4); z-index: 300; padding: 6px 10px; font-size: 12px; max-width: 260px; white-space: pre-wrap; pointer-events: none; }
  </style>
</head>
<body>
  <div id="toolbar">
    <button id="btn-bold" title="Bold (Ctrl+B)"><b>B</b></button>
    <button id="btn-italic" title="Italic (Ctrl+I)"><i>I</i></button>
    <button id="btn-underline" title="Underline (Ctrl+U)"><u>U</u></button>
    <button id="btn-strikethrough" title="Strikethrough"><s>S</s></button>
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
        <button data-fmt="custom">Custom\u2026</button>
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
    <button id="btn-sort-asc" title="Sort Ascending">A\u2193</button>
    <button id="btn-sort-desc" title="Sort Descending">Z\u2193</button>
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
  <div id="comment-dialog">
    <div class="dialog-title">Cell Comment</div>
    <div class="dialog-body">
      <textarea id="comment-text" placeholder="Enter comment..."></textarea>
    </div>
    <div class="dialog-actions">
      <button id="comment-save">Save</button>
      <button id="comment-delete">Delete</button>
      <button id="comment-cancel">Cancel</button>
    </div>
  </div>
  <div id="comment-tooltip"></div>
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
          <option value="colorScale">Color Scale</option>
          <option value="dataBar">Data Bar</option>
        </select>
      </label>
      <label>Value: <input type="text" id="cond-value1"></label>
      <label id="cond-value2-label" style="display:none">To: <input type="text" id="cond-value2"></label>
      <div id="cond-standard-style">
        <label>Background: <input type="color" id="cond-bg" value="#FFEB3B"></label>
        <label>Text color: <input type="color" id="cond-text" value="#000000"></label>
        <div style="display:flex;gap:8px">
          <label><input type="checkbox" id="cond-bold"> Bold</label>
          <label><input type="checkbox" id="cond-italic"> Italic</label>
        </div>
      </div>
      <div id="cond-colorscale-opts" style="display:none">
        <label>Min color: <input type="color" id="cs-min-color" value="#F8696B"></label>
        <label><input type="checkbox" id="cs-use-mid"> Mid color: <input type="color" id="cs-mid-color" value="#FFEB84"></label>
        <label>Max color: <input type="color" id="cs-max-color" value="#63BE7B"></label>
      </div>
      <div id="cond-databar-opts" style="display:none">
        <label>Bar color: <input type="color" id="db-color" value="#4472C4"></label>
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

function getNonce(): string {
  let text = '';
  const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  for (let i = 0; i < 32; i++) {
    text += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return text;
}
