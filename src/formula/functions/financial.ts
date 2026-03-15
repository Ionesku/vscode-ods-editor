import { FunctionRegistry } from './index';
import { flattenRange, isFormulaError, toNumber } from '../types';
import { FormulaError } from '../../model/types';

export function registerFinancialFunctions(registry: FunctionRegistry): void {
  /**
   * PMT(rate, nper, pv, [fv=0], [type=0])
   * Returns the periodic payment for a loan.
   */
  registry.register('PMT', (args) => {
    if (args.length < 3) return FormulaError.VALUE;
    const rate = toNumber(flattenRange(args[0])[0]);
    const nper = toNumber(flattenRange(args[1])[0]);
    const pv = toNumber(flattenRange(args[2])[0]);
    if (isFormulaError(rate) || isFormulaError(nper) || isFormulaError(pv)) return FormulaError.VALUE;
    const fv = args.length >= 4 ? toNumber(flattenRange(args[3])[0]) : 0;
    const type = args.length >= 5 ? toNumber(flattenRange(args[4])[0]) : 0;
    if (isFormulaError(fv) || isFormulaError(type)) return FormulaError.VALUE;
    return calcPmt(rate, nper, pv, fv, type);
  });

  /**
   * PV(rate, nper, pmt, [fv=0], [type=0])
   * Returns the present value of an investment.
   */
  registry.register('PV', (args) => {
    if (args.length < 3) return FormulaError.VALUE;
    const rate = toNumber(flattenRange(args[0])[0]);
    const nper = toNumber(flattenRange(args[1])[0]);
    const pmt = toNumber(flattenRange(args[2])[0]);
    if (isFormulaError(rate) || isFormulaError(nper) || isFormulaError(pmt)) return FormulaError.VALUE;
    const fv = args.length >= 4 ? toNumber(flattenRange(args[3])[0]) : 0;
    const type = args.length >= 5 ? toNumber(flattenRange(args[4])[0]) : 0;
    if (isFormulaError(fv) || isFormulaError(type)) return FormulaError.VALUE;
    if (rate === 0) return -(pmt * nper + (fv as number));
    const r = rate as number;
    const n = nper as number;
    const f = fv as number;
    const t = type as number;
    return -(f + pmt * (1 + r * t) * ((Math.pow(1 + r, n) - 1) / r)) / Math.pow(1 + r, n);
  });

  /**
   * FV(rate, nper, pmt, [pv=0], [type=0])
   * Returns the future value of an investment.
   */
  registry.register('FV', (args) => {
    if (args.length < 3) return FormulaError.VALUE;
    const rate = toNumber(flattenRange(args[0])[0]);
    const nper = toNumber(flattenRange(args[1])[0]);
    const pmt = toNumber(flattenRange(args[2])[0]);
    if (isFormulaError(rate) || isFormulaError(nper) || isFormulaError(pmt)) return FormulaError.VALUE;
    const pv = args.length >= 4 ? toNumber(flattenRange(args[3])[0]) : 0;
    const type = args.length >= 5 ? toNumber(flattenRange(args[4])[0]) : 0;
    if (isFormulaError(pv) || isFormulaError(type)) return FormulaError.VALUE;
    const r = rate as number;
    const n = nper as number;
    const p = pmt as number;
    const v = pv as number;
    const t = type as number;
    if (r === 0) return -(v + p * n);
    return -(v * Math.pow(1 + r, n) + p * (1 + r * t) * (Math.pow(1 + r, n) - 1) / r);
  });

  /**
   * NPV(rate, value1, value2, ...)
   * Returns the net present value of an investment given a discount rate and future cash flows.
   */
  registry.register('NPV', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const rate = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(rate)) return FormulaError.VALUE;
    const r = rate as number;
    if (r === -1) return FormulaError.DIV0;
    let npv = 0;
    let period = 1;
    for (let i = 1; i < args.length; i++) {
      const values = flattenRange(args[i]);
      for (const v of values) {
        const n = toNumber(v);
        if (isFormulaError(n)) continue;
        npv += (n as number) / Math.pow(1 + r, period);
        period++;
      }
    }
    return npv;
  });

  /**
   * IRR(values, [guess=0.1])
   * Returns the internal rate of return for a series of cash flows.
   * Uses Newton-Raphson iteration.
   */
  registry.register('IRR', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const cashFlows: number[] = [];
    for (const v of flattenRange(args[0])) {
      const n = toNumber(v);
      if (isFormulaError(n)) return FormulaError.VALUE;
      cashFlows.push(n as number);
    }
    if (cashFlows.length === 0) return FormulaError.VALUE;

    const guessArg = args.length >= 2 ? toNumber(flattenRange(args[1])[0]) : 0.1;
    if (isFormulaError(guessArg)) return FormulaError.VALUE;
    let rate = guessArg as number;

    // Newton-Raphson: find rate where NPV = 0
    for (let iter = 0; iter < 100; iter++) {
      let npv = 0;
      let dnpv = 0;
      for (let i = 0; i < cashFlows.length; i++) {
        const denom = Math.pow(1 + rate, i);
        npv += cashFlows[i] / denom;
        if (i > 0) dnpv -= i * cashFlows[i] / (denom * (1 + rate));
      }
      if (Math.abs(dnpv) < 1e-10) return FormulaError.VALUE;
      const newRate = rate - npv / dnpv;
      if (Math.abs(newRate - rate) < 1e-7) return newRate;
      rate = newRate;
    }
    return FormulaError.VALUE; // did not converge
  });

  /**
   * RATE(nper, pmt, pv, [fv=0], [type=0], [guess=0.1])
   * Returns the interest rate per period of an annuity.
   */
  registry.register('RATE', (args) => {
    if (args.length < 3) return FormulaError.VALUE;
    const nper = toNumber(flattenRange(args[0])[0]);
    const pmt = toNumber(flattenRange(args[1])[0]);
    const pv = toNumber(flattenRange(args[2])[0]);
    if (isFormulaError(nper) || isFormulaError(pmt) || isFormulaError(pv)) return FormulaError.VALUE;
    const fv = args.length >= 4 ? toNumber(flattenRange(args[3])[0]) : 0;
    const type = args.length >= 5 ? toNumber(flattenRange(args[4])[0]) : 0;
    const guessArg = args.length >= 6 ? toNumber(flattenRange(args[5])[0]) : 0.1;
    if (isFormulaError(fv) || isFormulaError(type) || isFormulaError(guessArg)) return FormulaError.VALUE;

    const n = nper as number;
    const p = pmt as number;
    const v = pv as number;
    const f = fv as number;
    const t = type as number;
    let rate = guessArg as number;

    // Newton-Raphson on PMT function
    for (let iter = 0; iter < 100; iter++) {
      const computed = calcPmt(rate, n, v, f, t);
      if (isFormulaError(computed)) return FormulaError.VALUE;
      const delta = 1e-6;
      const computed2 = calcPmt(rate + delta, n, v, f, t);
      if (isFormulaError(computed2)) return FormulaError.VALUE;
      const deriv = ((computed2 as number) - (computed as number)) / delta;
      if (Math.abs(deriv) < 1e-12) return FormulaError.VALUE;
      const newRate = rate - ((computed as number) - p) / deriv;
      if (Math.abs(newRate - rate) < 1e-7) return newRate;
      rate = newRate;
    }
    return FormulaError.VALUE;
  });

  /**
   * NPER(rate, pmt, pv, [fv=0], [type=0])
   * Returns the number of periods for an investment.
   */
  registry.register('NPER', (args) => {
    if (args.length < 3) return FormulaError.VALUE;
    const rate = toNumber(flattenRange(args[0])[0]);
    const pmt = toNumber(flattenRange(args[1])[0]);
    const pv = toNumber(flattenRange(args[2])[0]);
    if (isFormulaError(rate) || isFormulaError(pmt) || isFormulaError(pv)) return FormulaError.VALUE;
    const fv = args.length >= 4 ? toNumber(flattenRange(args[3])[0]) : 0;
    const type = args.length >= 5 ? toNumber(flattenRange(args[4])[0]) : 0;
    if (isFormulaError(fv) || isFormulaError(type)) return FormulaError.VALUE;

    const r = rate as number;
    const p = pmt as number;
    const v = pv as number;
    const f = fv as number;
    const t = type as number;

    if (r === 0) {
      if (p === 0) return FormulaError.DIV0;
      return -(v + f) / p;
    }
    const num = p * (1 + r * t) - f * r;
    const den = p * (1 + r * t) + v * r;
    if (num === 0 || den === 0 || num / den <= 0) return FormulaError.NUM;
    return Math.log(num / den) / Math.log(1 + r);
  });

  /**
   * SLN(cost, salvage, life)
   * Straight-line depreciation for one period.
   */
  registry.register('SLN', (args) => {
    if (args.length < 3) return FormulaError.VALUE;
    const cost = toNumber(flattenRange(args[0])[0]);
    const salvage = toNumber(flattenRange(args[1])[0]);
    const life = toNumber(flattenRange(args[2])[0]);
    if (isFormulaError(cost) || isFormulaError(salvage) || isFormulaError(life)) return FormulaError.VALUE;
    if ((life as number) === 0) return FormulaError.DIV0;
    return ((cost as number) - (salvage as number)) / (life as number);
  });
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function calcPmt(
  rate: number,
  nper: number,
  pv: number,
  fv: number,
  type: number,
): number | FormulaError {
  if (nper === 0) return FormulaError.DIV0;
  if (rate === 0) return -(pv + fv) / nper;
  const pvif = Math.pow(1 + rate, nper);
  const pmt = (-pv * pvif - fv) / ((pvif - 1) / rate * (1 + rate * type));
  return pmt;
}
