import { Token, TokenType } from './types';

export class Tokenizer {
  private input = '';
  private pos = 0;
  private tokens: Token[] = [];

  tokenize(formula: string): Token[] {
    // Strip ODF prefix
    this.input = formula;
    if (this.input.startsWith('of:=')) this.input = this.input.substring(4);
    else if (this.input.startsWith('oooc:=')) this.input = this.input.substring(6);
    else if (this.input.startsWith('=')) this.input = this.input.substring(1);

    this.pos = 0;
    this.tokens = [];

    while (this.pos < this.input.length) {
      this.skipWhitespace();
      if (this.pos >= this.input.length) break;

      const ch = this.input[this.pos];

      if (ch === '"') {
        this.readString();
      } else if (ch === '(') {
        this.tokens.push({ type: TokenType.LParen, value: '(', position: this.pos++ });
      } else if (ch === ')') {
        this.tokens.push({ type: TokenType.RParen, value: ')', position: this.pos++ });
      } else if (ch === ',') {
        this.tokens.push({ type: TokenType.Comma, value: ',', position: this.pos++ });
      } else if (ch === ';') {
        // Semicolons are used as argument separators in some locales
        this.tokens.push({ type: TokenType.Comma, value: ';', position: this.pos++ });
      } else if (ch === ':') {
        this.tokens.push({ type: TokenType.Colon, value: ':', position: this.pos++ });
      } else if (this.isOperatorStart(ch)) {
        this.readOperator();
      } else if (
        this.isDigit(ch) ||
        (ch === '.' && this.pos + 1 < this.input.length && this.isDigit(this.input[this.pos + 1]))
      ) {
        this.readNumber();
      } else if (ch === '$' || this.isLetter(ch) || ch === '[') {
        this.readIdentifier();
      } else {
        // Skip unknown characters
        this.pos++;
      }
    }

    this.tokens.push({ type: TokenType.EOF, value: '', position: this.pos });
    return this.tokens;
  }

  private skipWhitespace(): void {
    while (
      this.pos < this.input.length &&
      (this.input[this.pos] === ' ' || this.input[this.pos] === '\t')
    ) {
      this.pos++;
    }
  }

  private readString(): void {
    const start = this.pos;
    this.pos++; // skip opening "
    let value = '';
    while (this.pos < this.input.length) {
      if (this.input[this.pos] === '"') {
        if (this.pos + 1 < this.input.length && this.input[this.pos + 1] === '"') {
          value += '"';
          this.pos += 2;
        } else {
          this.pos++; // skip closing "
          break;
        }
      } else {
        value += this.input[this.pos];
        this.pos++;
      }
    }
    this.tokens.push({ type: TokenType.String, value, position: start });
  }

  private readNumber(): void {
    const start = this.pos;
    let value = '';
    while (
      this.pos < this.input.length &&
      (this.isDigit(this.input[this.pos]) || this.input[this.pos] === '.')
    ) {
      value += this.input[this.pos];
      this.pos++;
    }
    // Scientific notation
    if (
      this.pos < this.input.length &&
      (this.input[this.pos] === 'e' || this.input[this.pos] === 'E')
    ) {
      value += this.input[this.pos];
      this.pos++;
      if (
        this.pos < this.input.length &&
        (this.input[this.pos] === '+' || this.input[this.pos] === '-')
      ) {
        value += this.input[this.pos];
        this.pos++;
      }
      while (this.pos < this.input.length && this.isDigit(this.input[this.pos])) {
        value += this.input[this.pos];
        this.pos++;
      }
    }
    this.tokens.push({ type: TokenType.Number, value, position: start });
  }

  private readIdentifier(): void {
    const start = this.pos;

    // Handle ODS-style cell references with brackets: [.A1] or [.Sheet1.A1]
    if (this.input[this.pos] === '[') {
      this.readBracketRef();
      return;
    }

    // Read the full identifier (may include $ for absolute refs, . for sheet separator)
    let value = '';
    while (this.pos < this.input.length) {
      const ch = this.input[this.pos];
      if (this.isLetter(ch) || this.isDigit(ch) || ch === '$' || ch === '_' || ch === '.') {
        value += ch;
        this.pos++;
      } else {
        break;
      }
    }

    const upper = value.toUpperCase();

    // Check if this is a function name (followed by '(')
    const savedPos = this.pos;
    this.skipWhitespace();
    const followedByParen = this.pos < this.input.length && this.input[this.pos] === '(';

    if (followedByParen && (upper === 'TRUE' || upper === 'FALSE' || this.isFunctionLike(value))) {
      this.pos = savedPos; // Restore pos — paren consumed by parser
      this.skipWhitespace();
      this.tokens.push({ type: TokenType.FunctionName, value: upper, position: start });
      return;
    }

    // Restore pos if we skipped whitespace but it's not a function
    this.pos = savedPos;

    // Check for boolean (without parentheses)
    if (upper === 'TRUE' || upper === 'FALSE') {
      this.tokens.push({ type: TokenType.Boolean, value: upper, position: start });
      return;
    }

    // Otherwise treat as cell reference
    this.tokens.push({ type: TokenType.CellRef, value, position: start });
  }

  private readBracketRef(): void {
    const start = this.pos;
    this.pos++; // skip [
    let value = '';
    while (this.pos < this.input.length && this.input[this.pos] !== ']') {
      value += this.input[this.pos];
      this.pos++;
    }
    if (this.pos < this.input.length) this.pos++; // skip ]

    // Strip leading dot: [.A1] -> A1
    if (value.startsWith('.')) value = value.substring(1);

    this.tokens.push({ type: TokenType.CellRef, value, position: start });
  }

  private readOperator(): void {
    const start = this.pos;
    const ch = this.input[this.pos];

    if (ch === '<') {
      if (this.pos + 1 < this.input.length) {
        if (this.input[this.pos + 1] === '=') {
          this.tokens.push({ type: TokenType.Operator, value: '<=', position: start });
          this.pos += 2;
          return;
        }
        if (this.input[this.pos + 1] === '>') {
          this.tokens.push({ type: TokenType.Operator, value: '<>', position: start });
          this.pos += 2;
          return;
        }
      }
    } else if (ch === '>') {
      if (this.pos + 1 < this.input.length && this.input[this.pos + 1] === '=') {
        this.tokens.push({ type: TokenType.Operator, value: '>=', position: start });
        this.pos += 2;
        return;
      }
    }

    this.tokens.push({ type: TokenType.Operator, value: ch, position: start });
    this.pos++;
  }

  private isDigit(ch: string): boolean {
    return ch >= '0' && ch <= '9';
  }

  private isLetter(ch: string): boolean {
    return (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z');
  }

  private isOperatorStart(ch: string): boolean {
    return '+-*/^&=<>%'.includes(ch);
  }

  private isFunctionLike(name: string): boolean {
    // If name contains digits mixed with letters and starts with letters, it might be a cell ref
    // Functions are pure letter names (possibly with dots for sheet refs)
    const upper = name.toUpperCase().replace(/\$/g, '');
    // Cell ref pattern: optional_sheet.col_row or just col_row
    const parts = upper.split('.');
    const lastPart = parts[parts.length - 1];
    // Check if lastPart looks like a cell ref (letters then digits)
    const cellRefMatch = lastPart.match(/^[A-Z]+\d+$/);
    return !cellRefMatch;
  }
}
