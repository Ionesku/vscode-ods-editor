import { FunctionRegistry } from './index';
import { flattenRange, isFormulaError, toNumber } from '../types';
import { FormulaError } from '../../model/types';

export function registerLogicalFunctions(registry: FunctionRegistry): void {
  // IF and IFERROR are handled specially in Evaluator (lazy evaluation)
  // But we still register them so the registry knows they exist
  registry.register('IF', (_args) => {
    // This won't normally be called — Evaluator handles IF specially
    return FormulaError.VALUE;
  });

  registry.register('IFERROR', (_args) => {
    // This won't normally be called — Evaluator handles IFERROR specially
    return FormulaError.VALUE;
  });

  registry.register('AND', (args) => {
    if (args.length === 0) return FormulaError.VALUE;
    for (const arg of args) {
      const values = flattenRange(arg);
      for (const v of values) {
        if (isFormulaError(v)) return v;
        if (typeof v === 'boolean' && !v) return false;
        if (typeof v === 'number' && v === 0) return false;
        // Strings and nulls are ignored
      }
    }
    return true;
  });

  registry.register('OR', (args) => {
    if (args.length === 0) return FormulaError.VALUE;
    for (const arg of args) {
      const values = flattenRange(arg);
      for (const v of values) {
        if (isFormulaError(v)) return v;
        if (typeof v === 'boolean' && v) return true;
        if (typeof v === 'number' && v !== 0) return true;
      }
    }
    return false;
  });

  registry.register('NOT', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const v = flattenRange(args[0])[0];
    if (isFormulaError(v)) return v;
    if (typeof v === 'boolean') return !v;
    if (typeof v === 'number') return v === 0;
    return FormulaError.VALUE;
  });

  registry.register('TRUE', () => true);
  registry.register('FALSE', () => false);

  registry.register('SWITCH', (args) => {
    if (args.length < 3) return FormulaError.VALUE;
    const expr = flattenRange(args[0])[0];
    for (let i = 1; i < args.length - 1; i += 2) {
      const caseVal = flattenRange(args[i])[0];
      if (
        expr === caseVal ||
        (typeof expr === 'string' &&
          typeof caseVal === 'string' &&
          expr.toLowerCase() === caseVal.toLowerCase())
      ) {
        return args[i + 1];
      }
    }
    // Default value (if odd number of args after expression)
    if (args.length % 2 === 0) {
      return args[args.length - 1];
    }
    return FormulaError.NA;
  });

  registry.register('CHOOSE', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const idx = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(idx)) return idx;
    const index = Math.floor(idx);
    if (index < 1 || index >= args.length) return FormulaError.VALUE;
    return args[index];
  });

  registry.register('ISBLANK', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const v = flattenRange(args[0])[0];
    return v === null || v === '';
  });

  registry.register('ISERROR', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const v = flattenRange(args[0])[0];
    return isFormulaError(v);
  });

  registry.register('ISNUMBER', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const v = flattenRange(args[0])[0];
    return typeof v === 'number';
  });

  registry.register('ISTEXT', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const v = flattenRange(args[0])[0];
    return typeof v === 'string';
  });

  registry.register('ISLOGICAL', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const v = flattenRange(args[0])[0];
    return typeof v === 'boolean';
  });
}
