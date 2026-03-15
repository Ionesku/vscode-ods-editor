import { describe, it, expect } from 'vitest';
import { Tokenizer } from '../../../src/formula/Tokenizer';
import { Parser } from '../../../src/formula/Parser';
import { ASTNode } from '../../../src/formula/types';

describe('Parser', () => {
  const tokenizer = new Tokenizer();
  const parser = new Parser();

  function parse(formula: string): ASTNode {
    return parser.parse(tokenizer.tokenize(formula));
  }

  it('parses number literals', () => {
    expect(parse('42')).toEqual({ type: 'number', value: 42 });
  });

  it('parses string literals', () => {
    expect(parse('"hello"')).toEqual({ type: 'string', value: 'hello' });
  });

  it('parses boolean literals', () => {
    expect(parse('TRUE')).toEqual({ type: 'boolean', value: true });
    expect(parse('FALSE')).toEqual({ type: 'boolean', value: false });
  });

  it('parses cell references', () => {
    const node = parse('A1');
    expect(node.type).toBe('cellRef');
    if (node.type === 'cellRef') {
      expect(node.col).toBe(0);
      expect(node.row).toBe(0);
    }
  });

  it('parses absolute cell references', () => {
    const node = parse('$B$2');
    if (node.type === 'cellRef') {
      expect(node.col).toBe(1);
      expect(node.row).toBe(1);
      expect(node.absCol).toBe(true);
      expect(node.absRow).toBe(true);
    }
  });

  it('parses range references', () => {
    const node = parse('A1:B5');
    expect(node.type).toBe('rangeRef');
    if (node.type === 'rangeRef') {
      expect(node.startCol).toBe(0);
      expect(node.startRow).toBe(0);
      expect(node.endCol).toBe(1);
      expect(node.endRow).toBe(4);
    }
  });

  it('parses binary operations with correct precedence', () => {
    const node = parse('1+2*3');
    expect(node.type).toBe('binaryOp');
    if (node.type === 'binaryOp') {
      expect(node.op).toBe('+');
      expect(node.left).toEqual({ type: 'number', value: 1 });
      expect(node.right.type).toBe('binaryOp');
      if (node.right.type === 'binaryOp') {
        expect(node.right.op).toBe('*');
      }
    }
  });

  it('parses unary negation', () => {
    const node = parse('-5');
    expect(node.type).toBe('unaryOp');
    if (node.type === 'unaryOp') {
      expect(node.op).toBe('-');
      expect(node.operand).toEqual({ type: 'number', value: 5 });
    }
  });

  it('parses parentheses', () => {
    const node = parse('(1+2)*3');
    expect(node.type).toBe('binaryOp');
    if (node.type === 'binaryOp') {
      expect(node.op).toBe('*');
      expect(node.left.type).toBe('binaryOp');
    }
  });

  it('parses function calls', () => {
    const node = parse('SUM(A1:A10)');
    expect(node.type).toBe('functionCall');
    if (node.type === 'functionCall') {
      expect(node.name).toBe('SUM');
      expect(node.args.length).toBe(1);
      expect(node.args[0].type).toBe('rangeRef');
    }
  });

  it('parses nested function calls', () => {
    const node = parse('IF(A1>0,SUM(B1:B5),0)');
    expect(node.type).toBe('functionCall');
    if (node.type === 'functionCall') {
      expect(node.name).toBe('IF');
      expect(node.args.length).toBe(3);
    }
  });

  it('parses comparison operators', () => {
    const node = parse('A1>=10');
    expect(node.type).toBe('binaryOp');
    if (node.type === 'binaryOp') {
      expect(node.op).toBe('>=');
    }
  });

  it('parses power operator', () => {
    const node = parse('2^3');
    expect(node.type).toBe('binaryOp');
    if (node.type === 'binaryOp') {
      expect(node.op).toBe('^');
    }
  });

  it('parses concatenation', () => {
    const node = parse('"a"&"b"');
    expect(node.type).toBe('binaryOp');
    if (node.type === 'binaryOp') {
      expect(node.op).toBe('&');
    }
  });

  it('parses percentage as division by 100', () => {
    const node = parse('50%');
    expect(node.type).toBe('binaryOp');
    if (node.type === 'binaryOp') {
      expect(node.op).toBe('/');
      expect(node.right).toEqual({ type: 'number', value: 100 });
    }
  });
});
