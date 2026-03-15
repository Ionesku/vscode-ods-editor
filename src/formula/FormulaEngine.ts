import { SpreadsheetModel } from '../model/SpreadsheetModel';
import { CellAddress, FormulaError } from '../model/types';
import { Tokenizer } from './Tokenizer';
import { Parser } from './Parser';
import { Evaluator } from './Evaluator';
import { DependencyGraph } from './DependencyGraph';
import { createDefaultRegistry, FunctionRegistry } from './functions/index';
import { ASTNode, CellValueType } from './types';

export class FormulaEngine {
  private tokenizer = new Tokenizer();
  private parser = new Parser();
  private registry: FunctionRegistry;
  private depGraph = new DependencyGraph();

  /** Addresses of cells that contain at least one volatile function call */
  private volatileCells: CellAddress[] = [];

  constructor() {
    this.registry = createDefaultRegistry();
  }

  /** Returns all registered function names (sorted), useful for autocomplete */
  getFunctionNames(): string[] {
    return this.registry.getAllNames();
  }

  /** True if any formula in the model uses a volatile function (RAND, NOW, TODAY…) */
  hasVolatileCells(): boolean {
    return this.volatileCells.length > 0;
  }

  /** Recalculate all formulas in the entire model */
  recalcAll(model: SpreadsheetModel): void {
    this.depGraph.clear();
    this.volatileCells = [];

    // First pass: parse all formulas and build dependency graph
    const formulaCells: Array<{ addr: CellAddress; ast: ASTNode; volatile: boolean }> = [];

    for (const sheet of model.sheets) {
      for (const entry of sheet.getAllCells()) {
        if (entry.data.formula) {
          const addr: CellAddress = { sheet: sheet.name, col: entry.col, row: entry.row };
          try {
            // Strip array formula braces {=...} before tokenizing
            const rawFormula =
              entry.data.formula.startsWith('{=') && entry.data.formula.endsWith('}')
                ? entry.data.formula.slice(1, -1)
                : entry.data.formula;
            const tokens = this.tokenizer.tokenize(rawFormula);
            const ast = this.parser.parse(tokens);
            const isVolatile = this.containsVolatileCall(ast);
            formulaCells.push({ addr, ast, volatile: isVolatile });
            if (isVolatile) this.volatileCells.push(addr);

            // Extract dependencies
            const deps = this.extractDependencies(ast, sheet.name);
            this.depGraph.setDependencies(addr, deps);
          } catch {
            // Parse error - mark cell with error
            sheet.setCell(entry.col, entry.row, {
              ...entry.data,
              computedValue: FormulaError.NAME,
            });
          }
        }
      }
    }

    // Get evaluation order (and detect circular references)
    const allFormulaAddrs = formulaCells.map((f) => f.addr);
    const { order, cyclic } = this.depGraph.getRecalcOrder(allFormulaAddrs);

    // Mark circular reference cells immediately so they don't evaluate
    for (const addr of cyclic) {
      const sheet = model.getSheetByName(addr.sheet);
      if (!sheet) continue;
      const cell = sheet.getCell(addr.col, addr.row);
      sheet.setCell(addr.col, addr.row, { ...cell, computedValue: FormulaError.CIRC });
    }
    const cyclicKeys = new Set(cyclic.map((a) => `${a.sheet}!${a.col},${a.row}`));

    // Create evaluator
    const getCellValue = (sheetName: string, col: number, row: number): CellValueType => {
      const sheet = model.getSheetByName(sheetName);
      if (!sheet) return FormulaError.REF;
      const cell = sheet.getCell(col, row);
      return (cell.computedValue ?? cell.rawValue ?? null) as CellValueType;
    };

    const evaluator = new Evaluator(getCellValue, this.registry, (name) =>
      model.resolveNamedRange(name),
    );

    // Build a map for quick AST lookup
    const astMap = new Map<string, ASTNode>();
    for (const fc of formulaCells) {
      astMap.set(`${fc.addr.sheet}!${fc.addr.col},${fc.addr.row}`, fc.ast);
    }

    // Evaluate in topological order (skip cyclic cells — already marked #CIRC!)
    for (const addr of order) {
      const key = `${addr.sheet}!${addr.col},${addr.row}`;
      if (cyclicKeys.has(key)) continue;
      const ast = astMap.get(key);
      if (!ast) continue;

      const sheet = model.getSheetByName(addr.sheet);
      if (!sheet) continue;

      try {
        evaluator.currentCell = { sheet: addr.sheet, col: addr.col, row: addr.row };
        const result = evaluator.evaluate(ast, addr.sheet);
        // If result is an array (range), take top-left value
        let value: CellValueType;
        if (Array.isArray(result) && Array.isArray(result[0])) {
          value = (result as CellValueType[][])[0]?.[0] ?? null;
        } else {
          value = result as CellValueType;
        }

        const cell = sheet.getCell(addr.col, addr.row);
        sheet.setCell(addr.col, addr.row, {
          ...cell,
          computedValue: value,
        });
      } catch {
        const cell = sheet.getCell(addr.col, addr.row);
        sheet.setCell(addr.col, addr.row, {
          ...cell,
          computedValue: FormulaError.VALUE,
        });
      }
    }
  }

  /**
   * Re-evaluate only the volatile cells (RAND, NOW, TODAY…) without a full rebuild.
   * Call this on a timer to keep time-based functions fresh.
   */
  recalcVolatile(model: SpreadsheetModel): void {
    if (this.volatileCells.length === 0) return;

    const getCellValue = (sheetName: string, col: number, row: number): CellValueType => {
      const sheet = model.getSheetByName(sheetName);
      if (!sheet) return FormulaError.REF;
      const cell = sheet.getCell(col, row);
      return (cell.computedValue ?? cell.rawValue ?? null) as CellValueType;
    };

    const evaluator = new Evaluator(getCellValue, this.registry, (name) =>
      model.resolveNamedRange(name),
    );

    for (const addr of this.volatileCells) {
      const sheet = model.getSheetByName(addr.sheet);
      if (!sheet) continue;
      const cell = sheet.getCell(addr.col, addr.row);
      if (!cell.formula) continue;
      try {
        evaluator.currentCell = { sheet: addr.sheet, col: addr.col, row: addr.row };
        const rawFormula =
          cell.formula.startsWith('{=') && cell.formula.endsWith('}')
            ? cell.formula.slice(1, -1)
            : cell.formula;
        const tokens = this.tokenizer.tokenize(rawFormula);
        const ast = this.parser.parse(tokens);
        const result = evaluator.evaluate(ast, addr.sheet);
        let value: CellValueType;
        if (Array.isArray(result) && Array.isArray(result[0])) {
          value = (result as CellValueType[][])[0]?.[0] ?? null;
        } else {
          value = result as CellValueType;
        }
        sheet.setCell(addr.col, addr.row, { ...cell, computedValue: value });
      } catch {
        // leave previous value on error
      }
    }
  }

  /** Returns true if the AST contains a call to a volatile function */
  private containsVolatileCall(node: ASTNode): boolean {
    switch (node.type) {
      case 'functionCall':
        if (this.registry.isVolatile(node.name)) return true;
        return node.args.some((a) => this.containsVolatileCall(a));
      case 'binaryOp':
        return this.containsVolatileCall(node.left) || this.containsVolatileCall(node.right);
      case 'unaryOp':
        return this.containsVolatileCall(node.operand);
      default:
        return false;
    }
  }

  /** Extract cell addresses referenced by an AST */
  private extractDependencies(node: ASTNode, contextSheet: string): CellAddress[] {
    const deps: CellAddress[] = [];
    this.walkAST(node, contextSheet, deps);
    return deps;
  }

  private walkAST(node: ASTNode, contextSheet: string, deps: CellAddress[]): void {
    switch (node.type) {
      case 'cellRef':
        deps.push({
          sheet: node.sheet ?? contextSheet,
          col: node.col,
          row: node.row,
        });
        break;
      case 'rangeRef': {
        const sheet = node.sheet ?? contextSheet;
        for (let r = node.startRow; r <= node.endRow; r++) {
          for (let c = node.startCol; c <= node.endCol; c++) {
            deps.push({ sheet, col: c, row: r });
          }
        }
        break;
      }
      case 'namedRangeRef': {
        // Will be resolved at eval time — no static deps to extract here
        break;
      }
      case 'binaryOp':
        this.walkAST(node.left, contextSheet, deps);
        this.walkAST(node.right, contextSheet, deps);
        break;
      case 'unaryOp':
        this.walkAST(node.operand, contextSheet, deps);
        break;
      case 'functionCall':
        for (const arg of node.args) {
          this.walkAST(arg, contextSheet, deps);
        }
        break;
    }
  }
}
