import { describe, it, expect } from 'vitest';
import { Tokenizer } from '../../../../src/formula/Tokenizer';
import { Parser } from '../../../../src/formula/Parser';
import { Evaluator, CellValueGetter } from '../../../../src/formula/Evaluator';
import { createDefaultRegistry } from '../../../../src/formula/functions/index';
import { CellValueType } from '../../../../src/formula/types';

describe('Text Functions', () => {
  const tokenizer = new Tokenizer();
  const parser = new Parser();
  const registry = createDefaultRegistry();

  function evaluate(formula: string): CellValueType {
    const getCellValue: CellValueGetter = () => null;
    const evaluator = new Evaluator(getCellValue, registry);
    const ast = parser.parse(tokenizer.tokenize(formula));
    const result = evaluator.evaluate(ast, 'Sheet1');
    if (Array.isArray(result) && Array.isArray(result[0]))
      return (result as CellValueType[][])[0][0];
    return result as CellValueType;
  }

  it('CONCATENATE', () => {
    expect(evaluate('CONCATENATE("hello"," ","world")')).toBe('hello world');
  });

  it('LEFT', () => {
    expect(evaluate('LEFT("hello",3)')).toBe('hel');
    expect(evaluate('LEFT("hello")')).toBe('h');
  });

  it('RIGHT', () => {
    expect(evaluate('RIGHT("hello",3)')).toBe('llo');
    expect(evaluate('RIGHT("hello")')).toBe('o');
  });

  it('MID', () => {
    expect(evaluate('MID("hello world",7,5)')).toBe('world');
  });

  it('LEN', () => {
    expect(evaluate('LEN("hello")')).toBe(5);
    expect(evaluate('LEN("")')).toBe(0);
  });

  it('TRIM', () => {
    expect(evaluate('TRIM("  hello   world  ")')).toBe('hello world');
  });

  it('UPPER', () => {
    expect(evaluate('UPPER("hello")')).toBe('HELLO');
  });

  it('LOWER', () => {
    expect(evaluate('LOWER("HELLO")')).toBe('hello');
  });

  it('PROPER', () => {
    expect(evaluate('PROPER("hello world")')).toBe('Hello World');
  });

  it('SUBSTITUTE', () => {
    expect(evaluate('SUBSTITUTE("hello world","world","earth")')).toBe('hello earth');
  });

  it('SUBSTITUTE nth instance', () => {
    expect(evaluate('SUBSTITUTE("aaa","a","b",2)')).toBe('aba');
  });

  it('REPLACE', () => {
    expect(evaluate('REPLACE("hello",2,3,"a")')).toBe('hao');
  });

  it('FIND', () => {
    expect(evaluate('FIND("lo","hello")')).toBe(4);
  });

  it('REPT', () => {
    expect(evaluate('REPT("ab",3)')).toBe('ababab');
  });

  it('EXACT', () => {
    expect(evaluate('EXACT("hello","hello")')).toBe(true);
    expect(evaluate('EXACT("hello","Hello")')).toBe(false);
  });

  it('TEXT', () => {
    expect(evaluate('TEXT(3.14159,"0.00")')).toBe('3.14');
  });

  it('CHAR and CODE', () => {
    expect(evaluate('CHAR(65)')).toBe('A');
    expect(evaluate('CODE("A")')).toBe(65);
  });
});
