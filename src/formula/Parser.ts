import { Token, TokenType, ASTNode } from './types';
import { FormulaError, letterToCol } from '../model/types';

const MAX_PARSE_DEPTH = 100;

export class Parser {
  private tokens: Token[] = [];
  private pos = 0;
  private depth = 0;

  parse(tokens: Token[]): ASTNode {
    this.tokens = tokens;
    this.pos = 0;
    this.depth = 0;
    const result = this.parseExpression();
    return result;
  }

  private current(): Token {
    return this.tokens[this.pos] ?? { type: TokenType.EOF, value: '', position: -1 };
  }

  private peek(): Token {
    return this.tokens[this.pos + 1] ?? { type: TokenType.EOF, value: '', position: -1 };
  }

  private advance(): Token {
    const t = this.current();
    this.pos++;
    return t;
  }

  private expect(type: TokenType): Token {
    const t = this.current();
    if (t.type !== type) {
      throw new Error(
        `Expected ${type} but got ${t.type} ("${t.value}") at position ${t.position}`,
      );
    }
    return this.advance();
  }

  // Expression -> Comparison (('&') Comparison)*
  private parseExpression(): ASTNode {
    if (++this.depth > MAX_PARSE_DEPTH) {
      this.depth--;
      throw new Error('Formula exceeds maximum nesting depth');
    }
    let left = this.parseComparison();
    while (this.current().type === TokenType.Operator && this.current().value === '&') {
      const op = this.advance().value;
      const right = this.parseComparison();
      left = { type: 'binaryOp', op, left, right };
    }
    this.depth--;
    return left;
  }

  // Comparison -> Addition (('=' | '<>' | '<' | '>' | '<=' | '>=') Addition)*
  private parseComparison(): ASTNode {
    let left = this.parseAddition();
    while (
      this.current().type === TokenType.Operator &&
      ['=', '<>', '<', '>', '<=', '>='].includes(this.current().value)
    ) {
      const op = this.advance().value;
      const right = this.parseAddition();
      left = { type: 'binaryOp', op, left, right };
    }
    return left;
  }

  // Addition -> Multiplication (('+' | '-') Multiplication)*
  private parseAddition(): ASTNode {
    let left = this.parseMultiplication();
    while (
      this.current().type === TokenType.Operator &&
      (this.current().value === '+' || this.current().value === '-')
    ) {
      const op = this.advance().value;
      const right = this.parseMultiplication();
      left = { type: 'binaryOp', op, left, right };
    }
    return left;
  }

  // Multiplication -> Unary (('*' | '/') Unary)*
  private parseMultiplication(): ASTNode {
    let left = this.parseUnary();
    while (
      this.current().type === TokenType.Operator &&
      (this.current().value === '*' || this.current().value === '/')
    ) {
      const op = this.advance().value;
      const right = this.parseUnary();
      left = { type: 'binaryOp', op, left, right };
    }
    return left;
  }

  // Unary -> ('-' | '+')? Power
  private parseUnary(): ASTNode {
    if (
      this.current().type === TokenType.Operator &&
      (this.current().value === '-' || this.current().value === '+')
    ) {
      const op = this.advance().value;
      const operand = this.parsePower();
      if (op === '+') return operand;
      return { type: 'unaryOp', op, operand };
    }
    return this.parsePower();
  }

  // Power -> Postfix ('^' Unary)?
  private parsePower(): ASTNode {
    let left = this.parsePostfix();
    if (this.current().type === TokenType.Operator && this.current().value === '^') {
      this.advance();
      const right = this.parseUnary();
      left = { type: 'binaryOp', op: '^', left, right };
    }
    return left;
  }

  // Postfix -> Atom ('%')?
  private parsePostfix(): ASTNode {
    let node = this.parseAtom();
    if (this.current().type === TokenType.Operator && this.current().value === '%') {
      this.advance();
      node = { type: 'binaryOp', op: '/', left: node, right: { type: 'number', value: 100 } };
    }
    return node;
  }

  // Atom -> Number | String | Boolean | CellRef (':' CellRef)? | FunctionCall | '(' Expression ')'
  private parseAtom(): ASTNode {
    const t = this.current();

    switch (t.type) {
      case TokenType.Number:
        this.advance();
        return { type: 'number', value: parseFloat(t.value) };

      case TokenType.String:
        this.advance();
        return { type: 'string', value: t.value };

      case TokenType.Boolean:
        this.advance();
        return { type: 'boolean', value: t.value === 'TRUE' };

      case TokenType.FunctionName:
        return this.parseFunctionCall();

      case TokenType.CellRef:
        return this.parseCellOrRangeRef();

      case TokenType.LParen: {
        this.advance();
        const expr = this.parseExpression();
        this.expect(TokenType.RParen);
        return expr;
      }

      default:
        throw new Error(`Unexpected token: ${t.type} ("${t.value}") at position ${t.position}`);
    }
  }

  private parseFunctionCall(): ASTNode {
    const name = this.advance().value;
    this.expect(TokenType.LParen);

    const args: ASTNode[] = [];
    if (this.current().type !== TokenType.RParen) {
      args.push(this.parseExpression());
      while (this.current().type === TokenType.Comma) {
        this.advance();
        args.push(this.parseExpression());
      }
    }

    this.expect(TokenType.RParen);
    return { type: 'functionCall', name, args };
  }

  private parseCellOrRangeRef(): ASTNode {
    const ref = this.parseCellRef();

    // Check for range operator ':'
    if (this.current().type === TokenType.Colon) {
      this.advance();
      const ref2 = this.parseCellRef();
      if (ref.type === 'cellRef' && ref2.type === 'cellRef') {
        return {
          type: 'rangeRef',
          sheet: ref.sheet,
          startCol: ref.col,
          startRow: ref.row,
          endCol: ref2.col,
          endRow: ref2.row,
        };
      }
    }

    return ref;
  }

  private parseCellRef(): ASTNode {
    const t = this.advance();
    const value = t.value;

    // Parse cell reference: optional_sheet.col_row or col_row
    // ODS format: Sheet1.A1 or just A1, with optional $ for absolute refs
    let sheet: string | undefined;
    let refPart = value;

    // Check for sheet prefix (contains a dot)
    const dotIdx = value.indexOf('.');
    if (dotIdx > 0) {
      sheet = value.substring(0, dotIdx);
      refPart = value.substring(dotIdx + 1);
    }

    // Parse column/row from refPart (e.g., $A$1, B2, AA100)
    let absCol = false;
    let absRow = false;
    let i = 0;

    if (refPart[i] === '$') {
      absCol = true;
      i++;
    }

    let colLetters = '';
    while (
      (i < refPart.length && refPart[i] >= 'A' && refPart[i] <= 'Z') ||
      (refPart[i] >= 'a' && refPart[i] <= 'z')
    ) {
      colLetters += refPart[i].toUpperCase();
      i++;
    }

    if (refPart[i] === '$') {
      absRow = true;
      i++;
    }

    let rowDigits = '';
    while (i < refPart.length && refPart[i] >= '0' && refPart[i] <= '9') {
      rowDigits += refPart[i];
      i++;
    }

    if (colLetters.length === 0 || rowDigits.length === 0) {
      // Looks like a name (e.g. MYRANGE) rather than a cell reference
      if (!sheet && rowDigits.length === 0 && colLetters.length > 0) {
        return {
          type: 'namedRangeRef',
          name: (colLetters + refPart.substring(colLetters.length)).toUpperCase(),
        };
      }
      return { type: 'error', errorType: FormulaError.REF };
    }

    const col = letterToCol(colLetters);
    const row = parseInt(rowDigits, 10) - 1; // 0-based

    return { type: 'cellRef', sheet, col, row, absCol, absRow };
  }
}
