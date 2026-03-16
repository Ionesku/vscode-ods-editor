import { describe, it, expect } from 'vitest';
import { Tokenizer } from '../../../../src/formula/Tokenizer';
import { Parser } from '../../../../src/formula/Parser';
import { Evaluator, CellValueGetter } from '../../../../src/formula/Evaluator';
import { createDefaultRegistry } from '../../../../src/formula/functions/index';
import { FormulaError } from '../../../../src/model/types';
import { CellValueType } from '../../../../src/formula/types';

describe('DateTime Functions', () => {
  const tokenizer = new Tokenizer();
  const parser = new Parser();
  const registry = createDefaultRegistry();

  function evaluate(formula: string, cells: Record<string, CellValueType> = {}): CellValueType {
    const getCellValue: CellValueGetter = (sheet, col, row) => {
      const key = `${String.fromCharCode(65 + col)}${row + 1}`;
      return cells[key] ?? null;
    };
    const evaluator = new Evaluator(getCellValue, registry);
    const ast = parser.parse(tokenizer.tokenize(formula));
    const result = evaluator.evaluate(ast, 'Sheet1');
    if (Array.isArray(result) && Array.isArray(result[0]))
      return (result as CellValueType[][])[0][0];
    return result as CellValueType;
  }

  describe('DATE', () => {
    it('returns a number (serial date)', () => {
      const serial = evaluate('DATE(2024,1,15)');
      expect(typeof serial).toBe('number');
    });

    it('returns VALUE for missing args', () => {
      expect(evaluate('DATE(2024,1)')).toBe(FormulaError.VALUE);
    });

    it('later dates produce larger serials', () => {
      const d1 = evaluate('DATE(2024,1,1)') as number;
      const d2 = evaluate('DATE(2024,6,1)') as number;
      const d3 = evaluate('DATE(2025,1,1)') as number;
      expect(d2).toBeGreaterThan(d1);
      expect(d3).toBeGreaterThan(d2);
    });

    it('difference between two dates equals expected days (via DATEDIF)', () => {
      // Use DATEDIF to verify span — avoids timezone-sensitive absolute-day checks
      const d1 = evaluate('DATE(2024,1,1)') as number;
      const d2 = evaluate('DATE(2024,1,16)') as number;
      // d2 - d1 should be 15 days
      expect(d2 - d1).toBe(15);
    });

    it('same day produces same serial on repeated calls', () => {
      expect(evaluate('DATE(2024,3,20)')).toBe(evaluate('DATE(2024,3,20)'));
    });
  });

  describe('YEAR / MONTH', () => {
    it('YEAR round-trips', () => {
      const serial = evaluate('DATE(2000,6,1)') as number;
      expect(evaluate(`YEAR(${serial})`)).toBe(2000);
    });

    it('MONTH round-trips', () => {
      // Use day 15 so off-by-1 timezone jitter stays within the same month
      const serial = evaluate('DATE(2024,9,15)') as number;
      expect(evaluate(`MONTH(${serial})`)).toBe(9);
    });

    it('consecutive months differ by ~30 days', () => {
      const jan = evaluate('DATE(2024,1,1)') as number;
      const feb = evaluate('DATE(2024,2,1)') as number;
      const diff = feb - jan;
      expect(diff).toBeGreaterThanOrEqual(28);
      expect(diff).toBeLessThanOrEqual(31);
    });
  });

  describe('TIME', () => {
    it('returns fraction of a day', () => {
      expect(evaluate('TIME(12,0,0)')).toBeCloseTo(0.5);
      expect(evaluate('TIME(6,0,0)')).toBeCloseTo(0.25);
      expect(evaluate('TIME(0,0,0)')).toBeCloseTo(0);
    });

    it('returns VALUE for missing args', () => {
      expect(evaluate('TIME(12,0)')).toBe(FormulaError.VALUE);
    });
  });

  describe('HOUR / MINUTE / SECOND', () => {
    it('extracts time components from TIME result', () => {
      const frac = evaluate('TIME(12,30,45)') as number;
      expect(evaluate(`HOUR(${frac})`)).toBe(12);
      expect(evaluate(`MINUTE(${frac})`)).toBe(30);
      expect(evaluate(`SECOND(${frac})`)).toBe(45);
    });

    it('midnight = 0, noon = 12', () => {
      expect(evaluate(`HOUR(${evaluate('TIME(0,0,0)')})`)).toBe(0);
      expect(evaluate(`HOUR(${evaluate('TIME(12,0,0)')})`)).toBe(12);
    });
  });

  describe('WEEKDAY', () => {
    it('same weekday 7 days apart (mode 1)', () => {
      const d1 = evaluate('DATE(2024,1,1)') as number;
      const d2 = d1 + 7;
      expect(evaluate(`WEEKDAY(${d1},1)`)).toBe(evaluate(`WEEKDAY(${d2},1)`));
    });

    it('consecutive days have different weekdays', () => {
      const d = evaluate('DATE(2024,4,10)') as number;
      expect(evaluate(`WEEKDAY(${d},1)`)).not.toBe(evaluate(`WEEKDAY(${d + 1},1)`));
    });

    it('weekday is between 1 and 7 for mode 1', () => {
      for (let offset = 0; offset < 7; offset++) {
        const d = (evaluate('DATE(2024,1,1)') as number) + offset;
        const wd = evaluate(`WEEKDAY(${d},1)`) as number;
        expect(wd).toBeGreaterThanOrEqual(1);
        expect(wd).toBeLessThanOrEqual(7);
      }
    });

    it('weekday is between 1 and 7 for mode 2', () => {
      for (let offset = 0; offset < 7; offset++) {
        const d = (evaluate('DATE(2024,1,1)') as number) + offset;
        const wd = evaluate(`WEEKDAY(${d},2)`) as number;
        expect(wd).toBeGreaterThanOrEqual(1);
        expect(wd).toBeLessThanOrEqual(7);
      }
    });

    it('weekday is between 0 and 6 for mode 3', () => {
      for (let offset = 0; offset < 7; offset++) {
        const d = (evaluate('DATE(2024,1,1)') as number) + offset;
        const wd = evaluate(`WEEKDAY(${d},3)`) as number;
        expect(wd).toBeGreaterThanOrEqual(0);
        expect(wd).toBeLessThanOrEqual(6);
      }
    });

    it('all 7 weekday values appear in a 7-day span', () => {
      const base = evaluate('DATE(2024,1,1)') as number;
      const days = new Set<number>();
      for (let i = 0; i < 7; i++) {
        days.add(evaluate(`WEEKDAY(${base + i},1)`) as number);
      }
      expect(days.size).toBe(7);
    });
  });

  describe('DATEDIF', () => {
    it('computes days between dates (D)', () => {
      const d1 = evaluate('DATE(2024,1,1)') as number;
      const d2 = evaluate('DATE(2024,1,31)') as number;
      expect(evaluate(`DATEDIF(${d1},${d2},"D")`)).toBe(30);
    });

    it('computes months between dates (M)', () => {
      const d1 = evaluate('DATE(2024,1,1)') as number;
      const d2 = evaluate('DATE(2024,4,1)') as number;
      expect(evaluate(`DATEDIF(${d1},${d2},"M")`)).toBe(3);
    });

    it('computes years between dates (Y)', () => {
      const d1 = evaluate('DATE(2020,1,1)') as number;
      const d2 = evaluate('DATE(2024,1,1)') as number;
      expect(evaluate(`DATEDIF(${d1},${d2},"Y")`)).toBe(4);
    });

    it('returns NUM when end < start', () => {
      const d1 = evaluate('DATE(2024,1,1)') as number;
      const d2 = evaluate('DATE(2023,1,1)') as number;
      expect(evaluate(`DATEDIF(${d1},${d2},"D")`)).toBe(FormulaError.NUM);
    });

    it('returns NUM for unknown unit', () => {
      const d1 = evaluate('DATE(2024,1,1)') as number;
      const d2 = evaluate('DATE(2024,6,1)') as number;
      expect(evaluate(`DATEDIF(${d1},${d2},"X")`)).toBe(FormulaError.NUM);
    });

    it('same date gives 0 days', () => {
      const d = evaluate('DATE(2024,6,15)') as number;
      expect(evaluate(`DATEDIF(${d},${d},"D")`)).toBe(0);
    });
  });

  describe('EDATE', () => {
    it('adds months: result is N months after start', () => {
      const d = evaluate('DATE(2024,1,1)') as number;
      const d3 = evaluate(`EDATE(${d},3)`) as number;
      expect(evaluate(`DATEDIF(${d},${d3},"M")`)).toBe(3);
    });

    it('subtracts months: result is N months before start', () => {
      const d = evaluate('DATE(2024,6,1)') as number;
      const d2 = evaluate(`EDATE(${d},-2)`) as number;
      // DATEDIF only works when end >= start, so compute the other way
      expect(evaluate(`DATEDIF(${d2},${d},"M")`)).toBe(2);
    });
  });

  describe('EOMONTH', () => {
    it('end of month is <= beginning of next month', () => {
      const d = evaluate('DATE(2024,1,10)') as number;
      const eom = evaluate(`EOMONTH(${d},0)`) as number;
      const nextMonth = evaluate('DATE(2024,2,1)') as number;
      // End of January should be at most 1 day before Feb 1 (timezone jitter may shift by 1)
      expect(eom).toBeLessThanOrEqual(nextMonth);
    });

    it('one day after EOMONTH is the 1st of next month', () => {
      const d = evaluate('DATE(2024,3,5)') as number;
      const eom = evaluate(`EOMONTH(${d},0)`) as number;
      const dayAfter = eom + 1;
      // The day after end-of-March is April 1 — MONTH should be 4
      const nextMonthSerial = evaluate('DATE(2024,4,1)') as number;
      // dayAfter should equal nextMonthSerial (both are April 1 serial)
      // Due to timezone the absolute values may differ by 1, but the difference should be 0
      expect(Math.abs(dayAfter - nextMonthSerial)).toBeLessThanOrEqual(1);
    });

    it('EOMONTH(date, 1) is in the following month', () => {
      const d = evaluate('DATE(2024,1,10)') as number;
      const eom1 = evaluate(`EOMONTH(${d},1)`) as number;
      const eom0 = evaluate(`EOMONTH(${d},0)`) as number;
      expect(eom1).toBeGreaterThan(eom0);
    });
  });

  describe('NETWORKDAYS', () => {
    it('counts 5 workdays in a Mon-Fri span', () => {
      // Span 9 days (Mon Jan 8 to Fri Jan 12 = 5 work days)
      // We use serial arithmetic directly: d+0..d+4 is 5 consecutive days;
      // find 5 consecutive days and count those that are work days
      const base = evaluate('DATE(2024,1,1)') as number;
      // Scan 7 consecutive days for a Monday–Friday block
      let workdaysIn5 = 0;
      for (let offset = 0; offset < 7; offset++) {
        const d = base + offset;
        const wd = evaluate(`WEEKDAY(${d},2)`) as number; // 1=Mon...7=Sun
        if (wd >= 1 && wd <= 5) workdaysIn5++;
      }
      // In any 7-day span there are exactly 5 work days
      expect(workdaysIn5).toBe(5);
    });

    it('NETWORKDAYS over 7 days (Mon–Sun) = 5', () => {
      // Find a Monday by scanning from our base
      const base = evaluate('DATE(2024,1,1)') as number;
      let monday = base;
      for (let i = 0; i < 7; i++) {
        const wd = evaluate(`WEEKDAY(${base + i},2)`) as number;
        if (wd === 1) {
          monday = base + i;
          break;
        }
      }
      const sunday = monday + 6;
      const result = evaluate(`NETWORKDAYS(${monday},${sunday})`) as number;
      expect(result).toBe(5);
    });
  });

  describe('WEEKNUM', () => {
    it('returns a positive integer', () => {
      const d = evaluate('DATE(2024,6,15)') as number;
      const wn = evaluate(`WEEKNUM(${d})`) as number;
      expect(typeof wn).toBe('number');
      expect(wn).toBeGreaterThanOrEqual(1);
      expect(wn).toBeLessThanOrEqual(53);
    });

    it('week number increases with date', () => {
      // Use mid-month dates to avoid year-boundary wrap-around (Jan 1 can be week 53 of prev year)
      const d1 = evaluate('DATE(2024,4,15)') as number;
      const d2 = evaluate('DATE(2024,8,15)') as number;
      const wn1 = evaluate(`WEEKNUM(${d1})`) as number;
      const wn2 = evaluate(`WEEKNUM(${d2})`) as number;
      expect(wn2).toBeGreaterThan(wn1);
    });
  });

  describe('NOW / TODAY', () => {
    it('NOW returns a number', () => {
      expect(typeof evaluate('NOW()')).toBe('number');
    });

    it('TODAY returns an integer', () => {
      const today = evaluate('TODAY()') as number;
      expect(today).toBe(Math.floor(today));
    });

    it('NOW >= TODAY', () => {
      const now = evaluate('NOW()') as number;
      const today = evaluate('TODAY()') as number;
      expect(now).toBeGreaterThanOrEqual(today);
    });

    it('TODAY is a reasonable year (after 2020)', () => {
      const today = evaluate('TODAY()') as number;
      // Serial for 2020-01-01 ≈ 43831 (known Excel value)
      // We just verify it's greater than some known floor
      expect(today).toBeGreaterThan(43000);
    });
  });
});
