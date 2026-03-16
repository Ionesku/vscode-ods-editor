import { FunctionRegistry } from './index';
import { flattenRange, isFormulaError, toNumber } from '../types';
import { FormulaError } from '../../model/types';

export function registerDateTimeFunctions(registry: FunctionRegistry): void {
  registry.register('NOW', () => {
    // Return Excel serial date number (days since 1899-12-30) with fractional time
    const now = new Date();
    return dateToSerial(now) + timeToFraction(now);
  });
  registry.markVolatile('NOW');

  registry.register('TODAY', () => {
    return dateToSerial(new Date());
  });
  registry.markVolatile('TODAY');

  registry.register('DATE', (args) => {
    if (args.length < 3) return FormulaError.VALUE;
    const year = toNumber(flattenRange(args[0])[0]);
    const month = toNumber(flattenRange(args[1])[0]);
    const day = toNumber(flattenRange(args[2])[0]);
    if (isFormulaError(year)) return year;
    if (isFormulaError(month)) return month;
    if (isFormulaError(day)) return day;
    const d = new Date(Math.floor(year), Math.floor(month) - 1, Math.floor(day));
    if (isNaN(d.getTime())) return FormulaError.VALUE;
    return dateToSerial(d);
  });

  registry.register('YEAR', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const serial = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(serial)) return serial;
    return serialToDate(serial).getFullYear();
  });

  registry.register('MONTH', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const serial = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(serial)) return serial;
    return serialToDate(serial).getMonth() + 1;
  });

  registry.register('DAY', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const serial = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(serial)) return serial;
    return serialToDate(serial).getDate();
  });

  registry.register('HOUR', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const serial = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(serial)) return serial;
    const frac = serial - Math.floor(serial);
    return Math.floor(frac * 24);
  });

  registry.register('MINUTE', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const serial = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(serial)) return serial;
    const frac = serial - Math.floor(serial);
    return Math.floor((frac * 24 - Math.floor(frac * 24)) * 60);
  });

  registry.register('SECOND', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const serial = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(serial)) return serial;
    const frac = serial - Math.floor(serial);
    const totalSeconds = frac * 86400;
    return Math.floor(totalSeconds % 60);
  });

  registry.register('TIME', (args) => {
    if (args.length < 3) return FormulaError.VALUE;
    const h = toNumber(flattenRange(args[0])[0]);
    const m = toNumber(flattenRange(args[1])[0]);
    const s = toNumber(flattenRange(args[2])[0]);
    if (isFormulaError(h)) return h;
    if (isFormulaError(m)) return m;
    if (isFormulaError(s)) return s;
    return (h * 3600 + m * 60 + s) / 86400;
  });

  registry.register('WEEKDAY', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const serial = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(serial)) return serial;
    const d = serialToDate(serial);
    const returnType = args.length >= 2 ? toNumber(flattenRange(args[1])[0]) : 1;
    if (isFormulaError(returnType)) return returnType;
    const dow = d.getDay(); // 0=Sunday
    if (returnType === 1) return dow + 1; // 1=Sunday...7=Saturday
    if (returnType === 2) return dow === 0 ? 7 : dow; // 1=Monday...7=Sunday
    if (returnType === 3) return dow === 0 ? 6 : dow - 1; // 0=Monday...6=Sunday
    return dow + 1;
  });

  registry.register('WEEKNUM', (args) => {
    if (args.length < 1) return FormulaError.VALUE;
    const serial = toNumber(flattenRange(args[0])[0]);
    if (isFormulaError(serial)) return serial;
    const d = serialToDate(serial);
    const jan1 = new Date(d.getFullYear(), 0, 1);
    const daysDiff = Math.floor((d.getTime() - jan1.getTime()) / 86400000);
    return Math.ceil((daysDiff + jan1.getDay() + 1) / 7);
  });

  registry.register('DATEDIF', (args) => {
    if (args.length < 3) return FormulaError.VALUE;
    const start = toNumber(flattenRange(args[0])[0]);
    const end = toNumber(flattenRange(args[1])[0]);
    const unit = String(flattenRange(args[2])[0] ?? '').toUpperCase();
    if (isFormulaError(start)) return start;
    if (isFormulaError(end)) return end;

    const d1 = serialToDate(start);
    const d2 = serialToDate(end);
    if (d2 < d1) return FormulaError.NUM;

    switch (unit) {
      case 'D':
        return Math.floor((d2.getTime() - d1.getTime()) / 86400000);
      case 'M':
        return (d2.getFullYear() - d1.getFullYear()) * 12 + d2.getMonth() - d1.getMonth();
      case 'Y':
        return d2.getFullYear() - d1.getFullYear();
      default:
        return FormulaError.NUM;
    }
  });

  registry.register('EDATE', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const serial = toNumber(flattenRange(args[0])[0]);
    const months = toNumber(flattenRange(args[1])[0]);
    if (isFormulaError(serial)) return serial;
    if (isFormulaError(months)) return months;
    const d = serialToDate(serial);
    d.setMonth(d.getMonth() + Math.floor(months));
    return dateToSerial(d);
  });

  registry.register('EOMONTH', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const serial = toNumber(flattenRange(args[0])[0]);
    const months = toNumber(flattenRange(args[1])[0]);
    if (isFormulaError(serial)) return serial;
    if (isFormulaError(months)) return months;
    const d = serialToDate(serial);
    d.setMonth(d.getMonth() + Math.floor(months) + 1, 0); // Last day of target month
    return dateToSerial(d);
  });

  registry.register('NETWORKDAYS', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const start = toNumber(flattenRange(args[0])[0]);
    const end = toNumber(flattenRange(args[1])[0]);
    if (isFormulaError(start)) return start;
    if (isFormulaError(end)) return end;

    const d1 = serialToDate(Math.floor(start));
    const d2 = serialToDate(Math.floor(end));
    let count = 0;
    const step = d1 <= d2 ? 1 : -1;
    const current = new Date(d1);
    while (step > 0 ? current <= d2 : current >= d2) {
      const dow = current.getDay();
      if (dow !== 0 && dow !== 6) count++;
      current.setDate(current.getDate() + step);
    }
    return step > 0 ? count : -count;
  });
}

// Excel epoch: 1899-12-30
const EPOCH = new Date(1899, 11, 30);

function dateToSerial(d: Date): number {
  const ms = d.getTime() - EPOCH.getTime();
  return Math.floor(ms / 86400000);
}

function serialToDate(serial: number): Date {
  const d = new Date(EPOCH.getTime() + Math.floor(serial) * 86400000);
  return d;
}

function timeToFraction(d: Date): number {
  return (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400;
}
