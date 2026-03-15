import * as vscode from 'vscode';
import { OdsEditorProvider } from './odsEditorProvider';
import { OdsWriter } from './ods/OdsWriter';
import { SpreadsheetModel } from './model/SpreadsheetModel';

interface Template {
  label: string;
  description: string;
  build: () => SpreadsheetModel;
}

const templates: Template[] = [
  {
    label: 'Blank',
    description: 'Empty spreadsheet',
    build: () => new SpreadsheetModel(),
  },
  {
    label: 'Budget',
    description: 'Monthly budget tracker with income/expenses',
    build: () => {
      const m = new SpreadsheetModel();
      const s = m.activeSheet;
      s.name = 'Budget';
      const headerStyle = m.getOrCreateStyle({
        id: 'tpl-header',
        bold: true,
        backgroundColor: '#4472C4',
        textColor: '#ffffff',
        horizontalAlign: 'center',
      });
      const catStyle = m.getOrCreateStyle({
        id: 'tpl-cat',
        bold: true,
        backgroundColor: '#D6E4F0',
      });
      const numStyle = m.getOrCreateStyle({
        id: 'tpl-num',
        numberFormat: '$#,##0.00',
        horizontalAlign: 'right',
      });

      const headers = [
        'Category',
        'Jan',
        'Feb',
        'Mar',
        'Apr',
        'May',
        'Jun',
        'Jul',
        'Aug',
        'Sep',
        'Oct',
        'Nov',
        'Dec',
        'Total',
      ];
      headers.forEach((h, i) => {
        s.setCell(i, 0, {
          rawValue: h,
          formula: null,
          computedValue: h,
          styleId: headerStyle.id,
          mergeColSpan: 1,
          mergeRowSpan: 1,
          mergedInto: null,
        });
      });
      s.setColumnWidth(0, 140);

      const income = ['Salary', 'Freelance', 'Other Income'];
      const expenses = [
        'Rent',
        'Utilities',
        'Groceries',
        'Transport',
        'Insurance',
        'Entertainment',
        'Savings',
        'Other',
      ];

      s.setCell(0, 1, {
        rawValue: 'INCOME',
        formula: null,
        computedValue: 'INCOME',
        styleId: catStyle.id,
        mergeColSpan: 1,
        mergeRowSpan: 1,
        mergedInto: null,
      });
      income.forEach((name, i) => {
        s.setCell(0, 2 + i, {
          rawValue: name,
          formula: null,
          computedValue: name,
          styleId: null,
          mergeColSpan: 1,
          mergeRowSpan: 1,
          mergedInto: null,
        });
        for (let m = 1; m <= 12; m++) {
          s.setCell(m, 2 + i, {
            rawValue: 0,
            formula: null,
            computedValue: 0,
            styleId: numStyle.id,
            mergeColSpan: 1,
            mergeRowSpan: 1,
            mergedInto: null,
          });
        }
      });

      const expRow = 2 + income.length + 1;
      s.setCell(0, expRow, {
        rawValue: 'EXPENSES',
        formula: null,
        computedValue: 'EXPENSES',
        styleId: catStyle.id,
        mergeColSpan: 1,
        mergeRowSpan: 1,
        mergedInto: null,
      });
      expenses.forEach((name, i) => {
        s.setCell(0, expRow + 1 + i, {
          rawValue: name,
          formula: null,
          computedValue: name,
          styleId: null,
          mergeColSpan: 1,
          mergeRowSpan: 1,
          mergedInto: null,
        });
        for (let m = 1; m <= 12; m++) {
          s.setCell(m, expRow + 1 + i, {
            rawValue: 0,
            formula: null,
            computedValue: 0,
            styleId: numStyle.id,
            mergeColSpan: 1,
            mergeRowSpan: 1,
            mergedInto: null,
          });
        }
      });

      return m;
    },
  },
  {
    label: 'Invoice',
    description: 'Simple invoice template',
    build: () => {
      const m = new SpreadsheetModel();
      const s = m.activeSheet;
      s.name = 'Invoice';
      const titleStyle = m.getOrCreateStyle({
        id: 'tpl-title',
        bold: true,
        fontSize: 18,
        textColor: '#2E4057',
      });
      const headerStyle = m.getOrCreateStyle({
        id: 'tpl-hdr',
        bold: true,
        backgroundColor: '#2E4057',
        textColor: '#ffffff',
      });
      const moneyStyle = m.getOrCreateStyle({
        id: 'tpl-money',
        numberFormat: '$#,##0.00',
        horizontalAlign: 'right',
      });
      const boldRight = m.getOrCreateStyle({
        id: 'tpl-bold-r',
        bold: true,
        horizontalAlign: 'right',
      });

      s.setCell(0, 0, {
        rawValue: 'INVOICE',
        formula: null,
        computedValue: 'INVOICE',
        styleId: titleStyle.id,
        mergeColSpan: 1,
        mergeRowSpan: 1,
        mergedInto: null,
      });
      s.setCell(0, 2, {
        rawValue: 'Bill To:',
        formula: null,
        computedValue: 'Bill To:',
        styleId: 'tpl-hdr-label',
        mergeColSpan: 1,
        mergeRowSpan: 1,
        mergedInto: null,
      });
      m.getOrCreateStyle({ id: 'tpl-hdr-label', bold: true });
      s.setCell(3, 2, {
        rawValue: 'Invoice #:',
        formula: null,
        computedValue: 'Invoice #:',
        styleId: 'tpl-bold-r',
        mergeColSpan: 1,
        mergeRowSpan: 1,
        mergedInto: null,
      });
      s.setCell(3, 3, {
        rawValue: 'Date:',
        formula: null,
        computedValue: 'Date:',
        styleId: 'tpl-bold-r',
        mergeColSpan: 1,
        mergeRowSpan: 1,
        mergedInto: null,
      });

      s.setColumnWidth(0, 200);
      s.setColumnWidth(1, 80);
      s.setColumnWidth(2, 100);
      s.setColumnWidth(3, 100);
      s.setColumnWidth(4, 100);

      const cols = ['Description', 'Qty', 'Unit Price', 'Amount'];
      const row = 6;
      cols.forEach((c, i) => {
        s.setCell(i, row, {
          rawValue: c,
          formula: null,
          computedValue: c,
          styleId: headerStyle.id,
          mergeColSpan: 1,
          mergeRowSpan: 1,
          mergedInto: null,
        });
      });

      for (let r = 0; r < 5; r++) {
        s.setCell(1, row + 1 + r, {
          rawValue: null,
          formula: null,
          computedValue: null,
          styleId: null,
          mergeColSpan: 1,
          mergeRowSpan: 1,
          mergedInto: null,
        });
        s.setCell(2, row + 1 + r, {
          rawValue: null,
          formula: null,
          computedValue: null,
          styleId: moneyStyle.id,
          mergeColSpan: 1,
          mergeRowSpan: 1,
          mergedInto: null,
        });
        s.setCell(3, row + 1 + r, {
          rawValue: null,
          formula: null,
          computedValue: null,
          styleId: moneyStyle.id,
          mergeColSpan: 1,
          mergeRowSpan: 1,
          mergedInto: null,
        });
      }

      const totalRow = row + 7;
      s.setCell(2, totalRow, {
        rawValue: 'TOTAL:',
        formula: null,
        computedValue: 'TOTAL:',
        styleId: boldRight.id,
        mergeColSpan: 1,
        mergeRowSpan: 1,
        mergedInto: null,
      });
      s.setCell(3, totalRow, {
        rawValue: 0,
        formula: null,
        computedValue: 0,
        styleId: moneyStyle.id,
        mergeColSpan: 1,
        mergeRowSpan: 1,
        mergedInto: null,
      });

      return m;
    },
  },
  {
    label: 'Timesheet',
    description: 'Weekly timesheet tracker',
    build: () => {
      const m = new SpreadsheetModel();
      const s = m.activeSheet;
      s.name = 'Timesheet';
      const headerStyle = m.getOrCreateStyle({
        id: 'tpl-header',
        bold: true,
        backgroundColor: '#548235',
        textColor: '#ffffff',
        horizontalAlign: 'center',
      });
      const totalStyle = m.getOrCreateStyle({
        id: 'tpl-total',
        bold: true,
        backgroundColor: '#E2EFDA',
      });

      const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
      const cols = ['Day', 'Start', 'End', 'Break (h)', 'Hours', 'Notes'];
      cols.forEach((c, i) => {
        s.setCell(i, 0, {
          rawValue: c,
          formula: null,
          computedValue: c,
          styleId: headerStyle.id,
          mergeColSpan: 1,
          mergeRowSpan: 1,
          mergedInto: null,
        });
      });
      s.setColumnWidth(0, 100);
      s.setColumnWidth(5, 200);

      days.forEach((day, i) => {
        s.setCell(0, 1 + i, {
          rawValue: day,
          formula: null,
          computedValue: day,
          styleId: null,
          mergeColSpan: 1,
          mergeRowSpan: 1,
          mergedInto: null,
        });
      });

      s.setCell(0, 8, {
        rawValue: 'TOTAL',
        formula: null,
        computedValue: 'TOTAL',
        styleId: totalStyle.id,
        mergeColSpan: 1,
        mergeRowSpan: 1,
        mergedInto: null,
      });
      s.setCell(4, 8, {
        rawValue: 0,
        formula: null,
        computedValue: 0,
        styleId: totalStyle.id,
        mergeColSpan: 1,
        mergeRowSpan: 1,
        mergedInto: null,
      });

      return m;
    },
  },
  {
    label: 'Contacts',
    description: 'Address book / contact list',
    build: () => {
      const m = new SpreadsheetModel();
      const s = m.activeSheet;
      s.name = 'Contacts';
      const headerStyle = m.getOrCreateStyle({
        id: 'tpl-header',
        bold: true,
        backgroundColor: '#7030A0',
        textColor: '#ffffff',
      });

      const cols = [
        'First Name',
        'Last Name',
        'Email',
        'Phone',
        'Company',
        'Address',
        'City',
        'Notes',
      ];
      const widths = [100, 100, 180, 120, 140, 180, 100, 200];
      cols.forEach((c, i) => {
        s.setCell(i, 0, {
          rawValue: c,
          formula: null,
          computedValue: c,
          styleId: headerStyle.id,
          mergeColSpan: 1,
          mergeRowSpan: 1,
          mergedInto: null,
        });
        s.setColumnWidth(i, widths[i]);
      });

      return m;
    },
  },
  {
    label: 'Project Tracker',
    description: 'Task list with status and priorities',
    build: () => {
      const m = new SpreadsheetModel();
      const s = m.activeSheet;
      s.name = 'Tasks';
      const headerStyle = m.getOrCreateStyle({
        id: 'tpl-header',
        bold: true,
        backgroundColor: '#C55A11',
        textColor: '#ffffff',
      });

      const cols = ['#', 'Task', 'Assignee', 'Priority', 'Status', 'Due Date', 'Notes'];
      const widths = [40, 250, 120, 80, 100, 100, 200];
      cols.forEach((c, i) => {
        s.setCell(i, 0, {
          rawValue: c,
          formula: null,
          computedValue: c,
          styleId: headerStyle.id,
          mergeColSpan: 1,
          mergeRowSpan: 1,
          mergedInto: null,
        });
        s.setColumnWidth(i, widths[i]);
      });

      for (let r = 1; r <= 10; r++) {
        s.setCell(0, r, {
          rawValue: r,
          formula: null,
          computedValue: r,
          styleId: null,
          mergeColSpan: 1,
          mergeRowSpan: 1,
          mergedInto: null,
        });
      }

      return m;
    },
  },
];

// ── Custom template storage ─────────────────────────────────────────────────

interface CustomTemplate {
  label: string;
  description: string;
  odsBase64: string; // base64-encoded ODS file
}

async function loadCustomTemplates(context: vscode.ExtensionContext): Promise<CustomTemplate[]> {
  return context.globalState.get<CustomTemplate[]>('customTemplates', []);
}

async function saveCustomTemplates(
  context: vscode.ExtensionContext,
  list: CustomTemplate[],
): Promise<void> {
  await context.globalState.update('customTemplates', list);
}

export function activate(context: vscode.ExtensionContext): void {
  const provider = new OdsEditorProvider(context);
  context.subscriptions.push(
    vscode.window.registerCustomEditorProvider('odsEditor.spreadsheet', provider, {
      webviewOptions: { retainContextWhenHidden: true },
      supportsMultipleEditorsPerDocument: false,
    }),
  );

  // ── New Table (with template picker) ────────────────────────────────────
  context.subscriptions.push(
    vscode.commands.registerCommand('odsEditor.newTable', async (uri?: vscode.Uri) => {
      let folderUri: vscode.Uri;
      if (uri) {
        const stat = await vscode.workspace.fs.stat(uri);
        folderUri = stat.type === vscode.FileType.Directory ? uri : vscode.Uri.joinPath(uri, '..');
      } else if (vscode.workspace.workspaceFolders?.length) {
        folderUri = vscode.workspace.workspaceFolders[0].uri;
      } else {
        vscode.window.showErrorMessage('No folder open to create a table in.');
        return;
      }

      // Build pick list: built-in + custom
      const customTemplates = await loadCustomTemplates(context);
      const pickLabels: string[] = templates.map((t) => t.label);
      const picks: vscode.QuickPickItem[] = templates.map((t) => ({
        label: t.label,
        description: t.description,
      }));
      if (customTemplates.length > 0) {
        picks.push({ label: '', kind: vscode.QuickPickItemKind.Separator });
        for (const ct of customTemplates) {
          picks.push({ label: ct.label, description: ct.description + ' (custom)' });
        }
      }

      const picked = await vscode.window.showQuickPick(picks, {
        placeHolder: 'Select a template',
        title: 'New Spreadsheet',
      });
      if (!picked) return;

      const isCustom = !pickLabels.includes(picked.label);

      const name = await vscode.window.showInputBox({
        prompt: 'Enter table name',
        value: picked.label === 'Blank' ? 'NewTable' : picked.label,
        validateInput: (v: string) => (v.trim().length === 0 ? 'Name cannot be empty' : undefined),
      });
      if (!name) return;

      const fileName = name.endsWith('.ods') ? name : name + '.ods';
      const fileUri = vscode.Uri.joinPath(folderUri, fileName);

      try {
        await vscode.workspace.fs.stat(fileUri);
        vscode.window.showErrorMessage(`File "${fileName}" already exists.`);
        return;
      } catch {
        // File doesn't exist — good
      }

      let fileData: Uint8Array;
      if (isCustom) {
        const ct = customTemplates.find((t) => t.label === picked.label);
        if (!ct) return;
        fileData = Uint8Array.from(Buffer.from(ct.odsBase64, 'base64'));
      } else {
        const template = templates.find((t) => t.label === picked.label)!;
        const model = template.build();
        const writer = new OdsWriter();
        fileData = await writer.write(model);
      }

      await vscode.workspace.fs.writeFile(fileUri, fileData);
      await vscode.commands.executeCommand('vscode.openWith', fileUri, 'odsEditor.spreadsheet');
    }),
  );

  // ── Save current file as template ───────────────────────────────────────
  context.subscriptions.push(
    vscode.commands.registerCommand('odsEditor.saveAsTemplate', async () => {
      const tabInput = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
      let fileUri: vscode.Uri | undefined;
      if (tabInput && typeof tabInput === 'object' && 'uri' in tabInput) {
        fileUri = (tabInput as { uri: vscode.Uri }).uri;
      }

      if (!fileUri || !fileUri.path.endsWith('.ods')) {
        vscode.window.showErrorMessage('Open an .ods file first to save it as a template.');
        return;
      }

      const label = await vscode.window.showInputBox({
        prompt: 'Template name',
        value: fileUri.path.split('/').pop()?.replace('.ods', '') ?? 'My Template',
      });
      if (!label) return;

      const description =
        (await vscode.window.showInputBox({
          prompt: 'Short description (optional)',
          value: '',
        })) ?? '';

      const fileData = await vscode.workspace.fs.readFile(fileUri);
      const odsBase64 = Buffer.from(fileData).toString('base64');

      const customTemplates = await loadCustomTemplates(context);
      // Replace if same name exists
      const idx = customTemplates.findIndex((t) => t.label === label);
      const entry: CustomTemplate = { label, description, odsBase64 };
      if (idx >= 0) {
        customTemplates[idx] = entry;
      } else {
        customTemplates.push(entry);
      }
      await saveCustomTemplates(context, customTemplates);
      vscode.window.showInformationMessage(
        `Template "${label}" saved. It will appear in the New Table menu.`,
      );
    }),
  );
}

export function deactivate(): void {
  // Nothing to clean up
}
