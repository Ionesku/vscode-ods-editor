import { FunctionRegistry } from './index';
import { FormulaValue, flattenRange, isFormulaError, toNumber, CellValueType } from '../types';
import { FormulaError } from '../../model/types';

export function registerStatisticalFunctions(registry: FunctionRegistry): void {
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
        if (v !== null && v !== '') count++;
      }
    }
    return count;
  });

  registry.register('COUNTBLANK', (args) => {
    let count = 0;
    for (const arg of args) {
      const values = flattenRange(arg);
      for (const v of values) {
        if (v === null || v === '') count++;
      }
    }
    return count;
  });

  registry.register('COUNTIF', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const range = flattenRange(args[0]);
    const criteria = args[1];
    const matcher = buildMatcher(criteria);
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
      pairs.push({
        range: flattenRange(args[i]),
        matcher: buildMatcher(args[i + 1]),
      });
    }
    const len = pairs[0].range.length;
    let count = 0;
    for (let i = 0; i < len; i++) {
      let match = true;
      for (const pair of pairs) {
        if (i >= pair.range.length || !pair.matcher(pair.range[i])) {
          match = false;
          break;
        }
      }
      if (match) count++;
    }
    return count;
  });

  registry.register('MEDIAN', (args) => {
    const nums = collectNumbers(args);
    if (nums.length === 0) return FormulaError.NUM;
    nums.sort((a, b) => a - b);
    const mid = Math.floor(nums.length / 2);
    return nums.length % 2 === 0 ? (nums[mid - 1] + nums[mid]) / 2 : nums[mid];
  });

  registry.register('MODE', (args) => {
    const nums = collectNumbers(args);
    if (nums.length === 0) return FormulaError.NA;
    const freq = new Map<number, number>();
    let maxFreq = 0;
    let mode = nums[0];
    for (const n of nums) {
      const f = (freq.get(n) ?? 0) + 1;
      freq.set(n, f);
      if (f > maxFreq) {
        maxFreq = f;
        mode = n;
      }
    }
    if (maxFreq === 1) return FormulaError.NA;
    return mode;
  });

  registry.register('STDEV', (args) => {
    const nums = collectNumbers(args);
    if (nums.length < 2) return FormulaError.DIV0;
    const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
    const sumSqDiff = nums.reduce((a, b) => a + (b - mean) ** 2, 0);
    return Math.sqrt(sumSqDiff / (nums.length - 1));
  });

  registry.register('STDEVP', (args) => {
    const nums = collectNumbers(args);
    if (nums.length === 0) return FormulaError.DIV0;
    const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
    const sumSqDiff = nums.reduce((a, b) => a + (b - mean) ** 2, 0);
    return Math.sqrt(sumSqDiff / nums.length);
  });

  registry.register('VAR', (args) => {
    const nums = collectNumbers(args);
    if (nums.length < 2) return FormulaError.DIV0;
    const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
    return nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (nums.length - 1);
  });

  registry.register('VARP', (args) => {
    const nums = collectNumbers(args);
    if (nums.length === 0) return FormulaError.DIV0;
    const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
    return nums.reduce((a, b) => a + (b - mean) ** 2, 0) / nums.length;
  });

  registry.register('LARGE', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const nums = collectNumbers([args[0]]);
    const k = toNumber(flattenRange(args[1])[0]);
    if (isFormulaError(k)) return k;
    if (k < 1 || k > nums.length) return FormulaError.NUM;
    nums.sort((a, b) => b - a);
    return nums[Math.floor(k) - 1];
  });

  registry.register('SMALL', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const nums = collectNumbers([args[0]]);
    const k = toNumber(flattenRange(args[1])[0]);
    if (isFormulaError(k)) return k;
    if (k < 1 || k > nums.length) return FormulaError.NUM;
    nums.sort((a, b) => a - b);
    return nums[Math.floor(k) - 1];
  });

  registry.register('PERCENTILE', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const nums = collectNumbers([args[0]]);
    const k = toNumber(flattenRange(args[1])[0]);
    if (isFormulaError(k)) return k;
    if (k < 0 || k > 1 || nums.length === 0) return FormulaError.NUM;
    nums.sort((a, b) => a - b);
    const n = (nums.length - 1) * k;
    const lo = Math.floor(n);
    const hi = Math.ceil(n);
    if (lo === hi) return nums[lo];
    return nums[lo] + (nums[hi] - nums[lo]) * (n - lo);
  });

  registry.register('RANK', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const num = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(num)) return num;
    const nums = collectNumbers([args[1]]);
    const order = args.length >= 3 ? toNumber(flattenRange(args[2])[0]) : 0;
    if (isFormulaError(order)) return order;

    if (order === 0) {
      nums.sort((a, b) => b - a); // Descending
    } else {
      nums.sort((a, b) => a - b); // Ascending
    }
    const idx = nums.indexOf(num);
    if (idx === -1) return FormulaError.NA;
    return idx + 1;
  });

  /**
   * IFS(cond1, val1, cond2, val2, ...) — returns val for the first true condition
   */
  registry.register('IFS', (args) => {
    if (args.length < 2 || args.length % 2 !== 0) return FormulaError.VALUE;
    for (let i = 0; i < args.length; i += 2) {
      const cond = flattenRange(args[i])[0];
      if (isFormulaError(cond)) return cond;
      if (cond) return flattenRange(args[i + 1])[0] ?? null;
    }
    return FormulaError.NA;
  });

  /**
   * MAXIFS(max_range, criteria_range1, criteria1, ...)
   */
  registry.register('MAXIFS', (args) => {
    if (args.length < 3 || (args.length - 1) % 2 !== 0) return FormulaError.VALUE;
    const maxRange = flattenRange(args[0]);
    const pairs: Array<{ range: CellValueType[]; matcher: (v: CellValueType) => boolean }> = [];
    for (let i = 1; i < args.length; i += 2) {
      pairs.push({ range: flattenRange(args[i]), matcher: buildMatcher(args[i + 1]) });
    }
    let result = -Infinity;
    let found = false;
    for (let i = 0; i < maxRange.length; i++) {
      const v = maxRange[i];
      if (typeof v !== 'number') continue;
      let match = true;
      for (const pair of pairs) {
        if (i >= pair.range.length || !pair.matcher(pair.range[i])) {
          match = false;
          break;
        }
      }
      if (match) {
        result = Math.max(result, v);
        found = true;
      }
    }
    return found ? result : 0;
  });

  /**
   * MINIFS(min_range, criteria_range1, criteria1, ...)
   */
  registry.register('MINIFS', (args) => {
    if (args.length < 3 || (args.length - 1) % 2 !== 0) return FormulaError.VALUE;
    const minRange = flattenRange(args[0]);
    const pairs: Array<{ range: CellValueType[]; matcher: (v: CellValueType) => boolean }> = [];
    for (let i = 1; i < args.length; i += 2) {
      pairs.push({ range: flattenRange(args[i]), matcher: buildMatcher(args[i + 1]) });
    }
    let result = Infinity;
    let found = false;
    for (let i = 0; i < minRange.length; i++) {
      const v = minRange[i];
      if (typeof v !== 'number') continue;
      let match = true;
      for (const pair of pairs) {
        if (i >= pair.range.length || !pair.matcher(pair.range[i])) {
          match = false;
          break;
        }
      }
      if (match) {
        result = Math.min(result, v);
        found = true;
      }
    }
    return found ? result : 0;
  });

  registry.register('CORREL', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const xs = collectNumbers([args[0]]);
    const ys = collectNumbers([args[1]]);
    const n = Math.min(xs.length, ys.length);
    if (n < 1) return FormulaError.DIV0;

    const meanX = xs.slice(0, n).reduce((a, b) => a + b, 0) / n;
    const meanY = ys.slice(0, n).reduce((a, b) => a + b, 0) / n;

    let sumXY = 0,
      sumX2 = 0,
      sumY2 = 0;
    for (let i = 0; i < n; i++) {
      const dx = xs[i] - meanX;
      const dy = ys[i] - meanY;
      sumXY += dx * dy;
      sumX2 += dx * dx;
      sumY2 += dy * dy;
    }
    const denom = Math.sqrt(sumX2 * sumY2);
    if (denom === 0) return FormulaError.DIV0;
    return sumXY / denom;
  });
}

function collectNumbers(args: FormulaValue[]): number[] {
  const nums: number[] = [];
  for (const arg of args) {
    const values = flattenRange(arg);
    for (const v of values) {
      if (typeof v === 'number') nums.push(v);
    }
  }
  return nums;
}

function buildMatcher(criteria: FormulaValue): (v: CellValueType) => boolean {
  if (typeof criteria === 'number') return (v) => typeof v === 'number' && v === criteria;
  if (typeof criteria === 'boolean') return (v) => v === criteria;

  const s = String(criteria ?? '');
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

  if (s.includes('*') || s.includes('?')) {
    const pattern = s
      .replace(/[-[\]{}()+.,\\^$|#\s]/g, '\\$&')
      .replace(/\*/g, '.*')
      .replace(/\?/g, '.');
    const re = new RegExp('^' + pattern + '$', 'i');
    return (v) => re.test(String(v ?? ''));
  }

  const num = Number(s);
  if (!isNaN(num) && s !== '') return (v) => typeof v === 'number' && v === num;
  return (v) => String(v ?? '').toLowerCase() === s.toLowerCase();
}
