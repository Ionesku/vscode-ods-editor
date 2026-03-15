import { describe, it, expect } from 'vitest';
import { Tokenizer } from '../../../src/formula/Tokenizer';
import { TokenType } from '../../../src/formula/types';

describe('Tokenizer', () => {
  const tokenizer = new Tokenizer();

  function tokenTypes(formula: string): TokenType[] {
    return tokenizer.tokenize(formula).map((t) => t.type);
  }

  function tokenValues(formula: string): string[] {
    return tokenizer
      .tokenize(formula)
      .filter((t) => t.type !== TokenType.EOF)
      .map((t) => t.value);
  }

  it('tokenizes numbers', () => {
    expect(tokenValues('42')).toEqual(['42']);
    expect(tokenValues('3.14')).toEqual(['3.14']);
    expect(tokenValues('1e5')).toEqual(['1e5']);
    expect(tokenValues('.5')).toEqual(['.5']);
  });

  it('tokenizes strings', () => {
    expect(tokenValues('"hello"')).toEqual(['hello']);
    expect(tokenValues('"he""llo"')).toEqual(['he"llo']);
    expect(tokenValues('""')).toEqual(['']);
  });

  it('tokenizes booleans', () => {
    expect(tokenValues('TRUE')).toEqual(['TRUE']);
    expect(tokenValues('FALSE')).toEqual(['FALSE']);
  });

  it('tokenizes cell references', () => {
    expect(tokenValues('A1')).toEqual(['A1']);
    expect(tokenValues('$B$2')).toEqual(['$B$2']);
    expect(tokenValues('AA100')).toEqual(['AA100']);
  });

  it('tokenizes ODS bracket references', () => {
    expect(tokenValues('[.A1]')).toEqual(['A1']);
    expect(tokenValues('[.Sheet1.A1]')).toEqual(['Sheet1.A1']);
  });

  it('tokenizes function names', () => {
    const tokens = tokenizer.tokenize('SUM(A1:A10)');
    expect(tokens[0].type).toBe(TokenType.FunctionName);
    expect(tokens[0].value).toBe('SUM');
  });

  it('tokenizes operators', () => {
    expect(tokenValues('1+2')).toEqual(['1', '+', '2']);
    expect(tokenValues('1>=2')).toEqual(['1', '>=', '2']);
    expect(tokenValues('1<>2')).toEqual(['1', '<>', '2']);
    expect(tokenValues('1<=2')).toEqual(['1', '<=', '2']);
  });

  it('tokenizes complex formulas', () => {
    const values = tokenValues('SUM(A1:B2)+IF(C1>0,C1*2,"no")');
    expect(values).toEqual([
      'SUM',
      '(',
      'A1',
      ':',
      'B2',
      ')',
      '+',
      'IF',
      '(',
      'C1',
      '>',
      '0',
      ',',
      'C1',
      '*',
      '2',
      ',',
      'no',
      ')',
    ]);
  });

  it('strips ODS formula prefixes', () => {
    expect(tokenValues('of:=SUM(A1)')).toEqual(['SUM', '(', 'A1', ')']);
    expect(tokenValues('oooc:=A1+B1')).toEqual(['A1', '+', 'B1']);
  });

  it('handles semicolons as argument separators', () => {
    const tokens = tokenizer.tokenize('SUM(1;2;3)');
    const commas = tokens.filter((t) => t.type === TokenType.Comma);
    expect(commas.length).toBe(2);
  });

  it('handles percentage operator', () => {
    expect(tokenValues('50%')).toEqual(['50', '%']);
  });

  it('handles concatenation operator', () => {
    expect(tokenValues('"a"&"b"')).toEqual(['a', '&', 'b']);
  });
});
