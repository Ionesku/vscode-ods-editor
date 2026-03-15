import { FunctionRegistry } from './index';
import { flattenRange, isFormulaError, toNumber } from '../types';
import { FormulaError } from '../../model/types';

export function registerTextFunctions(registry: FunctionRegistry): void {
  registry.register('CONCATENATE', (args) => {
    let result = '';
    for (const arg of args) {
      const values = flattenRange(arg);
      for (const v of values) {
        if (isFormulaError(v)) return v;
        result += v === null ? '' : String(v);
      }
    }
    return result;
  });

  registry.register('CONCAT', (args) => {
    let result = '';
    for (const arg of args) {
      const values = flattenRange(arg);
      for (const v of values) {
        if (isFormulaError(v)) return v;
        result += v === null ? '' : String(v);
      }
    }
    return result;
  });

  registry.register('TEXTJOIN', (args) => {
    if (args.length < 3) return FormulaError.VALUE;
    const delimiter = String(flattenRange(args[0])[0] ?? '');
    const ignoreEmpty = !!flattenRange(args[1])[0];
    const parts: string[] = [];
    for (let i = 2; i < args.length; i++) {
      const values = flattenRange(args[i]);
      for (const v of values) {
        if (isFormulaError(v)) return v;
        const s = v === null ? '' : String(v);
        if (ignoreEmpty && s === '') continue;
        parts.push(s);
      }
    }
    return parts.join(delimiter);
  });

  registry.register('LEFT', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const s = String(flattenRange(args[0])[0] ?? '');
    const n = args.length >= 2 ? toNumber(flattenRange(args[1])[0]) : 1;
    if (isFormulaError(n)) return n;
    return s.substring(0, Math.max(0, Math.floor(n)));
  });

  registry.register('RIGHT', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const s = String(flattenRange(args[0])[0] ?? '');
    const n = args.length >= 2 ? toNumber(flattenRange(args[1])[0]) : 1;
    if (isFormulaError(n)) return n;
    const count = Math.max(0, Math.floor(n));
    return s.substring(Math.max(0, s.length - count));
  });

  registry.register('MID', (args) => {
    if (args.length < 3) return FormulaError.VALUE;
    const s = String(flattenRange(args[0])[0] ?? '');
    const start = toNumber(flattenRange(args[1])[0]);
    const length = toNumber(flattenRange(args[2])[0]);
    if (isFormulaError(start)) return start;
    if (isFormulaError(length)) return length;
    if (start < 1 || length < 0) return FormulaError.VALUE;
    return s.substring(Math.floor(start) - 1, Math.floor(start) - 1 + Math.floor(length));
  });

  registry.register('LEN', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    return String(flattenRange(args[0])[0] ?? '').length;
  });

  registry.register('TRIM', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    return String(flattenRange(args[0])[0] ?? '')
      .trim()
      .replace(/\s+/g, ' ');
  });

  registry.register('UPPER', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    return String(flattenRange(args[0])[0] ?? '').toUpperCase();
  });

  registry.register('LOWER', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    return String(flattenRange(args[0])[0] ?? '').toLowerCase();
  });

  registry.register('PROPER', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    return String(flattenRange(args[0])[0] ?? '').replace(
      /\w\S*/g,
      (txt) => txt.charAt(0).toUpperCase() + txt.substring(1).toLowerCase(),
    );
  });

  registry.register('SUBSTITUTE', (args) => {
    if (args.length < 3) return FormulaError.VALUE;
    const text = String(flattenRange(args[0])[0] ?? '');
    const oldText = String(flattenRange(args[1])[0] ?? '');
    const newText = String(flattenRange(args[2])[0] ?? '');

    if (args.length >= 4) {
      // Replace nth instance
      const instance = toNumber(flattenRange(args[3])[0]);
      if (isFormulaError(instance)) return instance;
      const n = Math.floor(instance);
      let count = 0;
      let idx = 0;
      while (idx <= text.length) {
        const found = text.indexOf(oldText, idx);
        if (found === -1) break;
        count++;
        if (count === n) {
          return text.substring(0, found) + newText + text.substring(found + oldText.length);
        }
        idx = found + 1;
      }
      return text;
    }

    // Replace all instances
    return text.split(oldText).join(newText);
  });

  registry.register('REPLACE', (args) => {
    if (args.length < 4) return FormulaError.VALUE;
    const text = String(flattenRange(args[0])[0] ?? '');
    const start = toNumber(flattenRange(args[1])[0]);
    const numChars = toNumber(flattenRange(args[2])[0]);
    const newText = String(flattenRange(args[3])[0] ?? '');
    if (isFormulaError(start)) return start;
    if (isFormulaError(numChars)) return numChars;
    const s = Math.floor(start) - 1;
    const n = Math.floor(numChars);
    return text.substring(0, s) + newText + text.substring(s + n);
  });

  registry.register('FIND', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const findText = String(flattenRange(args[0])[0] ?? '');
    const within = String(flattenRange(args[1])[0] ?? '');
    const startPos = args.length >= 3 ? toNumber(flattenRange(args[2])[0]) : 1;
    if (isFormulaError(startPos)) return startPos;
    const idx = within.indexOf(findText, Math.floor(startPos) - 1);
    if (idx === -1) return FormulaError.VALUE;
    return idx + 1;
  });

  registry.register('SEARCH', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const findText = String(flattenRange(args[0])[0] ?? '').toLowerCase();
    const within = String(flattenRange(args[1])[0] ?? '').toLowerCase();
    const startPos = args.length >= 3 ? toNumber(flattenRange(args[2])[0]) : 1;
    if (isFormulaError(startPos)) return startPos;
    const idx = within.indexOf(findText, Math.floor(startPos) - 1);
    if (idx === -1) return FormulaError.VALUE;
    return idx + 1;
  });

  registry.register('REPT', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const text = String(flattenRange(args[0])[0] ?? '');
    const times = toNumber(flattenRange(args[1])[0]);
    if (isFormulaError(times)) return times;
    if (times < 0) return FormulaError.VALUE;
    return text.repeat(Math.floor(times));
  });

  registry.register('EXACT', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const a = String(flattenRange(args[0])[0] ?? '');
    const b = String(flattenRange(args[1])[0] ?? '');
    return a === b;
  });

  registry.register('VALUE', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const v = flattenRange(args[0])[0];
    if (typeof v === 'number') return v;
    const n = Number(v);
    if (isNaN(n)) return FormulaError.VALUE;
    return n;
  });

  registry.register('TEXT', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const v = flattenRange(args[0])[0];
    const fmt = String(flattenRange(args[1])[0] ?? '');
    // Simplified text formatting
    if (typeof v === 'number') {
      if (fmt.includes('0') || fmt.includes('#')) {
        // Count decimal places from format
        const dotIdx = fmt.indexOf('.');
        if (dotIdx >= 0) {
          const decimals = fmt.length - dotIdx - 1;
          return v.toFixed(decimals);
        }
        return Math.round(v).toString();
      }
      return String(v);
    }
    return String(v ?? '');
  });

  registry.register('CHAR', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const n = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(n)) return n;
    return String.fromCharCode(Math.floor(n));
  });

  registry.register('CODE', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const s = String(flattenRange(args[0])[0] ?? '');
    if (s.length === 0) return FormulaError.VALUE;
    return s.charCodeAt(0);
  });
}
