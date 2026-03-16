import { FunctionRegistry } from './index';
import { flattenRange, isFormulaError, toNumber, CellValueType, FormulaValue } from '../types';
import { FormulaError } from '../../model/types';

export function registerMathFunctions(registry: FunctionRegistry): void {
  registry.register('SUM', (args) => {
    let sum = 0;
    for (const arg of args) {
      const values = flattenRange(arg);
      for (const v of values) {
        if (isFormulaError(v)) return v;
        if (typeof v === 'number') sum += v;
        else if (typeof v === 'boolean') sum += v ? 1 : 0;
        // Strings and nulls are ignored in SUM
      }
    }
    return sum;
  });

  registry.register('SUMIF', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const range = flattenRange(args[0]);
    const criteria = args[1];
    const sumRange = args.length >= 3 ? flattenRange(args[2]) : range;

    const matcher = buildMatcher(criteria);
    let sum = 0;
    for (let i = 0; i < range.length; i++) {
      if (matcher(range[i])) {
        const sv = i < sumRange.length ? sumRange[i] : 0;
        if (typeof sv === 'number') sum += sv;
      }
    }
    return sum;
  });

  registry.register('SUMIFS', (args) => {
    if (args.length < 3 || args.length % 2 === 0) return FormulaError.VALUE;
    const sumRange = flattenRange(args[0]);
    const pairs: Array<{ range: CellValueType[]; matcher: (v: CellValueType) => boolean }> = [];
    for (let i = 1; i < args.length; i += 2) {
      pairs.push({
        range: flattenRange(args[i]),
        matcher: buildMatcher(args[i + 1]),
      });
    }
    let sum = 0;
    for (let i = 0; i < sumRange.length; i++) {
      let match = true;
      for (const pair of pairs) {
        if (i >= pair.range.length || !pair.matcher(pair.range[i])) {
          match = false;
          break;
        }
      }
      if (match && typeof sumRange[i] === 'number') {
        sum += sumRange[i] as number;
      }
    }
    return sum;
  });

  registry.register('AVERAGE', (args) => {
    let sum = 0;
    let count = 0;
    for (const arg of args) {
      const values = flattenRange(arg);
      for (const v of values) {
        if (isFormulaError(v)) return v;
        if (typeof v === 'number') {
          sum += v;
          count++;
        } else if (typeof v === 'boolean') {
          sum += v ? 1 : 0;
          count++;
        }
      }
    }
    if (count === 0) return FormulaError.DIV0;
    return sum / count;
  });

  registry.register('AVERAGEIF', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const range = flattenRange(args[0]);
    const criteria = args[1];
    const avgRange = args.length >= 3 ? flattenRange(args[2]) : range;
    const matcher = buildMatcher(criteria);
    let sum = 0,
      count = 0;
    for (let i = 0; i < range.length; i++) {
      if (matcher(range[i])) {
        const sv = i < avgRange.length ? avgRange[i] : null;
        if (typeof sv === 'number') {
          sum += sv;
          count++;
        }
      }
    }
    if (count === 0) return FormulaError.DIV0;
    return sum / count;
  });

  registry.register('MIN', (args) => {
    let min = Infinity;
    let found = false;
    for (const arg of args) {
      const values = flattenRange(arg);
      for (const v of values) {
        if (isFormulaError(v)) return v;
        if (typeof v === 'number') {
          if (v < min) min = v;
          found = true;
        }
      }
    }
    return found ? min : 0;
  });

  registry.register('MAX', (args) => {
    let max = -Infinity;
    let found = false;
    for (const arg of args) {
      const values = flattenRange(arg);
      for (const v of values) {
        if (isFormulaError(v)) return v;
        if (typeof v === 'number') {
          if (v > max) max = v;
          found = true;
        }
      }
    }
    return found ? max : 0;
  });

  registry.register('ABS', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const n = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(n)) return n;
    return Math.abs(n);
  });

  registry.register('ROUND', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const n = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(n)) return n;
    const digits = args.length >= 2 ? toNumber(flattenRange(args[1])[0]) : 0;
    if (isFormulaError(digits)) return digits;
    const factor = Math.pow(10, digits);
    return Math.round(n * factor) / factor;
  });

  registry.register('ROUNDUP', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const n = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(n)) return n;
    const digits = args.length >= 2 ? toNumber(flattenRange(args[1])[0]) : 0;
    if (isFormulaError(digits)) return digits;
    const factor = Math.pow(10, digits);
    return (Math.sign(n) * Math.ceil(Math.abs(n) * factor)) / factor;
  });

  registry.register('ROUNDDOWN', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const n = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(n)) return n;
    const digits = args.length >= 2 ? toNumber(flattenRange(args[1])[0]) : 0;
    if (isFormulaError(digits)) return digits;
    const factor = Math.pow(10, digits);
    return (Math.sign(n) * Math.floor(Math.abs(n) * factor)) / factor;
  });

  registry.register('INT', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const n = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(n)) return n;
    return Math.floor(n);
  });

  registry.register('MOD', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const a = toNumber(flattenRange(args[0])[0]);
    const b = toNumber(flattenRange(args[1])[0]);
    if (isFormulaError(a)) return a;
    if (isFormulaError(b)) return b;
    if (b === 0) return FormulaError.DIV0;
    return a - b * Math.floor(a / b);
  });

  registry.register('POWER', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const base = toNumber(flattenRange(args[0])[0]);
    const exp = toNumber(flattenRange(args[1])[0]);
    if (isFormulaError(base)) return base;
    if (isFormulaError(exp)) return exp;
    return Math.pow(base, exp);
  });

  registry.register('SQRT', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const n = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(n)) return n;
    if (n < 0) return FormulaError.NUM;
    return Math.sqrt(n);
  });

  registry.register('LOG', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const n = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(n)) return n;
    if (n <= 0) return FormulaError.NUM;
    const base = args.length >= 2 ? toNumber(flattenRange(args[1])[0]) : 10;
    if (isFormulaError(base)) return base;
    if (base <= 0 || base === 1) return FormulaError.NUM;
    return Math.log(n) / Math.log(base);
  });

  registry.register('LN', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const n = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(n)) return n;
    if (n <= 0) return FormulaError.NUM;
    return Math.log(n);
  });

  registry.register('LOG10', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const n = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(n)) return n;
    if (n <= 0) return FormulaError.NUM;
    return Math.log10(n);
  });

  registry.register('EXP', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const n = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(n)) return n;
    return Math.exp(n);
  });

  registry.register('PI', () => Math.PI);

  registry.register('RAND', () => Math.random());
  registry.markVolatile('RAND');

  registry.register('RANDBETWEEN', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const lo = toNumber(flattenRange(args[0])[0]);
    const hi = toNumber(flattenRange(args[1])[0]);
    if (isFormulaError(lo)) return lo;
    if (isFormulaError(hi)) return hi;
    const low = Math.ceil(lo);
    const high = Math.floor(hi);
    return Math.floor(Math.random() * (high - low + 1)) + low;
  });
  registry.markVolatile('RANDBETWEEN');

  registry.register('CEILING', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const n = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(n)) return n;
    const sig = args.length >= 2 ? toNumber(flattenRange(args[1])[0]) : 1;
    if (isFormulaError(sig)) return sig;
    if (sig === 0) return 0;
    return Math.ceil(n / sig) * sig;
  });

  registry.register('FLOOR', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const n = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(n)) return n;
    const sig = args.length >= 2 ? toNumber(flattenRange(args[1])[0]) : 1;
    if (isFormulaError(sig)) return sig;
    if (sig === 0) return 0;
    return Math.floor(n / sig) * sig;
  });

  registry.register('SIGN', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const n = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(n)) return n;
    return Math.sign(n);
  });

  registry.register('PRODUCT', (args) => {
    let product = 1;
    let found = false;
    for (const arg of args) {
      const values = flattenRange(arg);
      for (const v of values) {
        if (isFormulaError(v)) return v;
        if (typeof v === 'number') {
          product *= v;
          found = true;
        }
      }
    }
    return found ? product : 0;
  });

  registry.register('COUNT', (args) => {
    let count = 0;
    for (const arg of args) {
      const values = flattenRange(arg);
      for (const v of values) {
        if (typeof v === 'number') count++;
      }
    }
    return count;
  });

  registry.register('COUNTA', (args) => {
    let count = 0;
    for (const arg of args) {
      const values = flattenRange(arg);
      for (const v of values) {
        if (v !== null && v !== undefined && v !== '') count++;
      }
    }
    return count;
  });

  registry.register('COUNTBLANK', (args) => {
    let count = 0;
    for (const arg of args) {
      const values = flattenRange(arg);
      for (const v of values) {
        if (v === null || v === undefined || v === '') count++;
      }
    }
    return count;
  });

  registry.register('COUNTIF', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const range = flattenRange(args[0]);
    const matcher = buildMatcher(args[1]);
    let count = 0;
    for (const v of range) {
      if (matcher(v)) count++;
    }
    return count;
  });

  registry.register('COUNTIFS', (args) => {
    if (args.length < 2 || args.length % 2 !== 0) return FormulaError.VALUE;
    const pairs: Array<{ range: CellValueType[]; matcher: (v: CellValueType) => boolean }> = [];
    for (let i = 0; i < args.length; i += 2) {
      pairs.push({ range: flattenRange(args[i]), matcher: buildMatcher(args[i + 1]) });
    }
    const len = pairs[0].range.length;
    let count = 0;
    for (let i = 0; i < len; i++) {
      let match = true;
      for (const pair of pairs) {
        if (i >= pair.range.length || !pair.matcher(pair.range[i])) { match = false; break; }
      }
      if (match) count++;
    }
    return count;
  });

  registry.register('AVERAGEIFS', (args) => {
    if (args.length < 3 || args.length % 2 === 0) return FormulaError.VALUE;
    const avgRange = flattenRange(args[0]);
    const pairs: Array<{ range: CellValueType[]; matcher: (v: CellValueType) => boolean }> = [];
    for (let i = 1; i < args.length; i += 2) {
      pairs.push({ range: flattenRange(args[i]), matcher: buildMatcher(args[i + 1]) });
    }
    let sum = 0, count = 0;
    for (let i = 0; i < avgRange.length; i++) {
      let match = true;
      for (const pair of pairs) {
        if (i >= pair.range.length || !pair.matcher(pair.range[i])) { match = false; break; }
      }
      if (match && typeof avgRange[i] === 'number') { sum += avgRange[i] as number; count++; }
    }
    if (count === 0) return FormulaError.DIV0;
    return sum / count;
  });

  registry.register('SUMPRODUCT', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    // All ranges must be same size
    const ranges = args.map((a) => flattenRange(a));
    const len = ranges[0].length;
    for (const r of ranges) {
      if (r.length !== len) return FormulaError.VALUE;
    }
    let sum = 0;
    for (let i = 0; i < len; i++) {
      let product = 1;
      for (const r of ranges) {
        const v = r[i];
        if (typeof v === 'number') product *= v;
        else if (typeof v === 'boolean') product *= v ? 1 : 0;
        else product *= 0;
      }
      sum += product;
    }
    return sum;
  });
}

// ── Matcher helper for xIF functions ────────────────────────────────────────

function buildMatcher(criteria: FormulaValue): (v: CellValueType) => boolean {
  // Criteria can be: number, string with operator, string for exact match
  if (typeof criteria === 'number') {
    return (v) => typeof v === 'number' && v === criteria;
  }
  if (typeof criteria === 'boolean') {
    return (v) => v === criteria;
  }

  const s = String(criteria ?? '');

  // Check for comparison operators
  if (s.startsWith('>=')) {
    const n = Number(s.substring(2));
    if (!isNaN(n)) return (v) => typeof v === 'number' && v >= n;
  }
  if (s.startsWith('<=')) {
    const n = Number(s.substring(2));
    if (!isNaN(n)) return (v) => typeof v === 'number' && v <= n;
  }
  if (s.startsWith('<>')) {
    const n = Number(s.substring(2));
    if (!isNaN(n)) return (v) => typeof v === 'number' && v !== n;
    return (v) => String(v ?? '').toLowerCase() !== s.substring(2).toLowerCase();
  }
  if (s.startsWith('>')) {
    const n = Number(s.substring(1));
    if (!isNaN(n)) return (v) => typeof v === 'number' && v > n;
  }
  if (s.startsWith('<')) {
    const n = Number(s.substring(1));
    if (!isNaN(n)) return (v) => typeof v === 'number' && v < n;
  }
  if (s.startsWith('=')) {
    const rest = s.substring(1);
    const n = Number(rest);
    if (!isNaN(n) && rest !== '') return (v) => typeof v === 'number' && v === n;
    return (v) => String(v ?? '').toLowerCase() === rest.toLowerCase();
  }

  // Wildcard support: * and ?
  if (s.includes('*') || s.includes('?')) {
    const pattern = s
      .replace(/[-[\]{}()+.,\\^$|#\s]/g, '\\$&')
      .replace(/\*/g, '.*')
      .replace(/\?/g, '.');
    const re = new RegExp('^' + pattern + '$', 'i');
    return (v) => re.test(String(v ?? ''));
  }

  // Exact match (case-insensitive for strings)
  const num = Number(s);
  if (!isNaN(num) && s !== '') {
    return (v) => typeof v === 'number' && v === num;
  }

  return (v) => String(v ?? '').toLowerCase() === s.toLowerCase();
}
