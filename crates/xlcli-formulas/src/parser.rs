use logos::Logos;

use crate::ast::*;
use crate::token::Token;

pub fn parse(input: &str) -> Result<Expr, String> {
    let mut parser = Parser::new(input);
    let expr = parser.parse_expr(0)?;
    Ok(expr)
}

struct Parser<'src> {
    tokens: Vec<(Token<'src>, &'src str)>,
    pos: usize,
}

impl<'src> Parser<'src> {
    fn new(input: &'src str) -> Self {
        let mut tokens = Vec::new();
        let mut lex = Token::lexer(input);
        while let Some(tok) = lex.next() {
            if let Ok(t) = tok {
                tokens.push((t, lex.slice()));
            }
        }
        Self { tokens, pos: 0 }
    }

    fn peek(&self) -> Option<&Token<'src>> {
        self.tokens.get(self.pos).map(|(t, _)| t)
    }

    fn advance(&mut self) -> Option<(Token<'src>, &'src str)> {
        let tok = self.tokens.get(self.pos).cloned();
        self.pos += 1;
        tok
    }

    fn expect(&mut self, expected: &Token) -> Result<(), String> {
        match self.peek() {
            Some(t) if t == expected => {
                self.advance();
                Ok(())
            }
            Some(t) => Err(format!("Expected {:?}, got {:?}", expected, t)),
            None => Err(format!("Expected {:?}, got end of input", expected)),
        }
    }

    fn parse_expr(&mut self, min_bp: u8) -> Result<Expr, String> {
        let mut lhs = self.parse_prefix()?;

        loop {
            // Postfix: percent
            if let Some(Token::Percent) = self.peek() {
                self.advance();
                lhs = Expr::Percent(Box::new(lhs));
                continue;
            }

            let Some(op) = self.peek_binop() else {
                break;
            };

            let (l_bp, r_bp) = infix_binding_power(op);
            if l_bp < min_bp {
                break;
            }

            // Special: colon = range operator
            if op == BinOp::Concat && self.peek() == Some(&Token::Colon) {
                // Actually this is handled in parse_prefix for cell refs
                break;
            }

            self.advance(); // consume operator
            let rhs = self.parse_expr(r_bp)?;
            lhs = Expr::BinOp {
                op,
                left: Box::new(lhs),
                right: Box::new(rhs),
            };
        }

        Ok(lhs)
    }

    fn parse_prefix(&mut self) -> Result<Expr, String> {
        match self.peek().cloned() {
            Some(Token::Number(..)) => {
                let (_, slice) = self.advance().unwrap();
                let n: f64 = slice.parse().map_err(|e| format!("Bad number: {}", e))?;
                Ok(Expr::Number(n))
            }
            Some(Token::StringLit(..)) => {
                let (_, slice) = self.advance().unwrap();
                let s = &slice[1..slice.len() - 1];
                let s = s.replace("\\\"", "\"");
                Ok(Expr::String(s))
            }
            Some(Token::Boolean(..)) => {
                let (_, slice) = self.advance().unwrap();
                Ok(Expr::Boolean(slice.eq_ignore_ascii_case("true")))
            }
            Some(Token::SheetRefQuoted(..)) | Some(Token::SheetRef(..)) => {
                let (_, slice) = self.advance().unwrap();
                let cell = parse_sheet_cell_ref(slice)?;
                if self.peek() == Some(&Token::Colon) {
                    self.advance();
                    match self.peek() {
                        Some(Token::CellRef(..)) | Some(Token::SheetRef(..)) | Some(Token::SheetRefQuoted(..)) => {
                            let (_, slice2) = self.advance().unwrap();
                            let mut cell2 = if slice2.contains('!') {
                                parse_sheet_cell_ref(slice2)?
                            } else {
                                parse_cell_ref(slice2)?
                            };
                            if let Expr::CellRef { ref mut sheet, .. } = cell2 {
                                if sheet.is_none() {
                                    if let Expr::CellRef { sheet: ref s, .. } = cell {
                                        *sheet = s.clone();
                                    }
                                }
                            }
                            return Ok(Expr::Range {
                                start: Box::new(cell),
                                end: Box::new(cell2),
                            });
                        }
                        _ => return Err("Expected cell reference after ':'".to_string()),
                    }
                }
                Ok(cell)
            }
            Some(Token::CellRef(..)) => {
                let (_, slice) = self.advance().unwrap();
                let cell = parse_cell_ref(slice)?;

                // Check for range operator ':'
                if self.peek() == Some(&Token::Colon) {
                    self.advance();
                    if let Some(Token::CellRef(..)) = self.peek() {
                        let (_, slice2) = self.advance().unwrap();
                        let cell2 = parse_cell_ref(slice2)?;
                        return Ok(Expr::Range {
                            start: Box::new(cell),
                            end: Box::new(cell2),
                        });
                    } else {
                        return Err("Expected cell reference after ':'".to_string());
                    }
                }

                Ok(cell)
            }
            Some(Token::FuncName(..)) => {
                let (_, slice) = self.advance().unwrap();
                let name = slice.to_uppercase();
                self.expect(&Token::LParen)?;
                let args = self.parse_arg_list()?;
                self.expect(&Token::RParen)?;
                Ok(Expr::FnCall { name, args })
            }
            Some(Token::LParen) => {
                self.advance();
                let expr = self.parse_expr(0)?;
                self.expect(&Token::RParen)?;
                Ok(expr)
            }
            Some(Token::Minus) => {
                self.advance();
                let expr = self.parse_expr(PREFIX_BP)?;
                Ok(Expr::UnaryOp {
                    op: UnaryOp::Neg,
                    expr: Box::new(expr),
                })
            }
            Some(Token::Plus) => {
                self.advance();
                let expr = self.parse_expr(PREFIX_BP)?;
                Ok(Expr::UnaryOp {
                    op: UnaryOp::Pos,
                    expr: Box::new(expr),
                })
            }
            Some(t) => Err(format!("Unexpected token: {:?}", t)),
            None => Err("Unexpected end of formula".to_string()),
        }
    }

    fn parse_arg_list(&mut self) -> Result<Vec<Expr>, String> {
        let mut args = Vec::new();
        if self.peek() == Some(&Token::RParen) {
            return Ok(args);
        }
        args.push(self.parse_expr(0)?);
        while matches!(self.peek(), Some(Token::Comma) | Some(Token::Semicolon)) {
            self.advance();
            args.push(self.parse_expr(0)?);
        }
        Ok(args)
    }

    fn peek_binop(&self) -> Option<BinOp> {
        match self.peek()? {
            Token::Plus => Some(BinOp::Add),
            Token::Minus => Some(BinOp::Sub),
            Token::Star => Some(BinOp::Mul),
            Token::Slash => Some(BinOp::Div),
            Token::Caret => Some(BinOp::Pow),
            Token::Ampersand => Some(BinOp::Concat),
            Token::Eq => Some(BinOp::Eq),
            Token::Neq => Some(BinOp::Neq),
            Token::Lt => Some(BinOp::Lt),
            Token::Gt => Some(BinOp::Gt),
            Token::Lte => Some(BinOp::Lte),
            Token::Gte => Some(BinOp::Gte),
            _ => None,
        }
    }
}

const PREFIX_BP: u8 = 14;

fn infix_binding_power(op: BinOp) -> (u8, u8) {
    match op {
        BinOp::Eq | BinOp::Neq | BinOp::Lt | BinOp::Gt | BinOp::Lte | BinOp::Gte => (2, 3),
        BinOp::Concat => (4, 5),
        BinOp::Add | BinOp::Sub => (6, 7),
        BinOp::Mul | BinOp::Div => (8, 9),
        BinOp::Pow => (11, 10), // right-associative
    }
}

fn parse_sheet_cell_ref(s: &str) -> Result<Expr, String> {
    let bang_pos = s.rfind('!').ok_or("Missing '!' in sheet reference")?;
    let sheet_part = &s[..bang_pos];
    let cell_part = &s[bang_pos + 1..];

    let sheet_name = if sheet_part.starts_with('\'') && sheet_part.ends_with('\'') {
        sheet_part[1..sheet_part.len() - 1].to_string()
    } else {
        sheet_part.to_string()
    };

    let mut cell = parse_cell_ref(cell_part)?;
    if let Expr::CellRef { ref mut sheet, .. } = cell {
        *sheet = Some(sheet_name);
    }
    Ok(cell)
}

fn parse_cell_ref(s: &str) -> Result<Expr, String> {
    let bytes = s.as_bytes();
    let mut i = 0;

    let col_abs = if bytes.get(i) == Some(&b'$') {
        i += 1;
        true
    } else {
        false
    };

    let col_start = i;
    while i < bytes.len() && bytes[i].to_ascii_uppercase().is_ascii_alphabetic() {
        i += 1;
    }
    let col_str = &s[col_start..i].to_uppercase();
    let col_val = xlcli_core::types::CellAddr::parse_col(col_str)
        .ok_or_else(|| format!("Invalid column: {}", col_str))?;

    let row_abs = if bytes.get(i) == Some(&b'$') {
        i += 1;
        true
    } else {
        false
    };

    let row_str = &s[i..];
    let row_val: u32 = row_str
        .parse::<u32>()
        .map_err(|_| format!("Invalid row: {}", row_str))?
        .saturating_sub(1);

    let col = if col_abs {
        ColRef::Absolute(col_val)
    } else {
        ColRef::Relative(col_val)
    };
    let row = if row_abs {
        RowRef::Absolute(row_val)
    } else {
        RowRef::Relative(row_val)
    };

    Ok(Expr::CellRef {
        sheet: None,
        col,
        row,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simple_number() {
        let expr = parse("42").unwrap();
        assert_eq!(expr, Expr::Number(42.0));
    }

    #[test]
    fn test_addition() {
        let expr = parse("1+2").unwrap();
        assert!(matches!(expr, Expr::BinOp { op: BinOp::Add, .. }));
    }

    #[test]
    fn test_cell_ref() {
        let expr = parse("A1").unwrap();
        assert!(matches!(expr, Expr::CellRef { .. }));
    }

    #[test]
    fn test_function_call() {
        let expr = parse("SUM(A1:A10)").unwrap();
        assert!(matches!(expr, Expr::FnCall { .. }));
    }

    #[test]
    fn test_precedence() {
        let expr = parse("1+2*3").unwrap();
        if let Expr::BinOp { op, right, .. } = expr {
            assert_eq!(op, BinOp::Add);
            assert!(matches!(*right, Expr::BinOp { op: BinOp::Mul, .. }));
        } else {
            panic!("Expected BinOp");
        }
    }

    #[test]
    fn test_range() {
        let expr = parse("A1:B10").unwrap();
        assert!(matches!(expr, Expr::Range { .. }));
    }
}
