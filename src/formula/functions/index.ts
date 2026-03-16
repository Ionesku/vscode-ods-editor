import { FormulaValue } from '../types';
import { Evaluator } from '../Evaluator';
import { registerMathFunctions } from './math';
import { registerLogicalFunctions } from './logical';
import { registerLookupFunctions } from './lookup';
import { registerTextFunctions } from './text';
import { registerStatisticalFunctions } from './statistical';
import { registerDateTimeFunctions } from './datetime';
import { registerFinancialFunctions } from './financial';

export type FormulaFunction = (args: FormulaValue[], evaluator: Evaluator) => FormulaValue;

export class FunctionRegistry {
  private fns = new Map<string, FormulaFunction>();
  private volatileNames = new Set<string>();

  register(name: string, fn: FormulaFunction): void {
    this.fns.set(name.toUpperCase(), fn);
  }

  /** Mark a function as volatile (re-evaluated on every recalcAll) */
  markVolatile(name: string): void {
    this.volatileNames.add(name.toUpperCase());
  }

  isVolatile(name: string): boolean {
    return this.volatileNames.has(name.toUpperCase());
  }

  get(name: string): FormulaFunction | undefined {
    return this.fns.get(name.toUpperCase());
  }

  has(name: string): boolean {
    return this.fns.has(name.toUpperCase());
  }

  getAllNames(): string[] {
    return Array.from(this.fns.keys()).sort();
  }
}

export function createDefaultRegistry(): FunctionRegistry {
  const registry = new FunctionRegistry();
  registerMathFunctions(registry);
  registerLogicalFunctions(registry);
  registerLookupFunctions(registry);
  registerTextFunctions(registry);
  registerStatisticalFunctions(registry);
  registerDateTimeFunctions(registry);
  registerFinancialFunctions(registry);
  return registry;
}
