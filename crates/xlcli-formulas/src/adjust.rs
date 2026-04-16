use xlcli_core::types::CellAddr;

use crate::ast::*;

pub fn adjust_formula(formula: &str, drow: i32, dcol: i32) -> Option<String> {
    let expr = crate::parser::parse(formula).ok()?;
    let shifted = shift_refs(&expr, drow, dcol);
    Some(expr_to_string(&shifted))
}

fn shift_refs(expr: &Expr, drow: i32, dcol: i32) -> Expr {
    match expr {
        Expr::CellRef { sheet, col, row } => Expr::CellRef {
            sheet: sheet.clone(),
            col: shift_col(col, dcol),
            row: shift_row(row, drow),
        },
        Expr::Range { start, end } => Expr::Range {
            start: Box::new(shift_refs(start, drow, dcol)),
            end: Box::new(shift_refs(end, drow, dcol)),
        },
        Expr::BinOp { op, left, right } => Expr::BinOp {
            op: *op,
            left: Box::new(shift_refs(left, drow, dcol)),
            right: Box::new(shift_refs(right, drow, dcol)),
        },
        Expr::UnaryOp { op, expr: inner } => Expr::UnaryOp {
            op: *op,
            expr: Box::new(shift_refs(inner, drow, dcol)),
        },
        Expr::Percent(inner) => Expr::Percent(Box::new(shift_refs(inner, drow, dcol))),
        Expr::FnCall { name, args } => Expr::FnCall {
            name: name.clone(),
            args: args.iter().map(|a| shift_refs(a, drow, dcol)).collect(),
        },
        other => other.clone(),
    }
}

fn shift_col(col: &ColRef, dcol: i32) -> ColRef {
    match col {
        ColRef::Absolute(v) => ColRef::Absolute(*v),
        ColRef::Relative(v) => {
            let new_val = (*v as i32 + dcol).max(0) as u16;
            ColRef::Relative(new_val)
        }
    }
}

fn shift_row(row: &RowRef, drow: i32) -> RowRef {
    match row {
        RowRef::Absolute(v) => RowRef::Absolute(*v),
        RowRef::Relative(v) => {
            let new_val = (*v as i64 + drow as i64).max(0) as u32;
            RowRef::Relative(new_val)
        }
    }
}

fn expr_to_string(expr: &Expr) -> String {
    match expr {
        Expr::Number(n) => {
            if *n == n.floor() && n.abs() < 1e15 {
                format!("{}", *n as i64)
            } else {
                format!("{}", n)
            }
        }
        Expr::String(s) => format!("\"{}\"", s.replace('"', "\\\"")),
        Expr::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        Expr::Error(e) => format!("{:?}", e),
        Expr::CellRef { sheet, col, row } => {
            let mut s = String::new();
            if let Some(sheet_name) = sheet {
                if sheet_name.contains(' ') {
                    s.push_str(&format!("'{}'!", sheet_name));
                } else {
                    s.push_str(&format!("{}!", sheet_name));
                }
            }
            s.push_str(&col_ref_to_string(col));
            s.push_str(&row_ref_to_string(row));
            s
        }
        Expr::Range { start, end } => {
            format!("{}:{}", expr_to_string(start), expr_to_string(end))
        }
        Expr::BinOp { op, left, right } => {
            let op_str = match op {
                BinOp::Add => "+",
                BinOp::Sub => "-",
                BinOp::Mul => "*",
                BinOp::Div => "/",
                BinOp::Pow => "^",
                BinOp::Concat => "&",
                BinOp::Eq => "=",
                BinOp::Neq => "<>",
                BinOp::Lt => "<",
                BinOp::Gt => ">",
                BinOp::Lte => "<=",
                BinOp::Gte => ">=",
            };
            let left_str = maybe_paren(left, *op, true);
            let right_str = maybe_paren(right, *op, false);
            format!("{}{}{}", left_str, op_str, right_str)
        }
        Expr::UnaryOp { op, expr: inner } => {
            let op_str = match op {
                UnaryOp::Neg => "-",
                UnaryOp::Pos => "+",
            };
            format!("{}{}", op_str, expr_to_string(inner))
        }
        Expr::Percent(inner) => format!("{}%", expr_to_string(inner)),
        Expr::FnCall { name, args } => {
            let arg_strs: Vec<String> = args.iter().map(|a| expr_to_string(a)).collect();
            format!("{}({})", name, arg_strs.join(","))
        }
    }
}

fn col_ref_to_string(col: &ColRef) -> String {
    match col {
        ColRef::Absolute(v) => format!("${}", CellAddr::col_name(*v)),
        ColRef::Relative(v) => CellAddr::col_name(*v),
    }
}

fn row_ref_to_string(row: &RowRef) -> String {
    match row {
        RowRef::Absolute(v) => format!("${}", v + 1),
        RowRef::Relative(v) => format!("{}", v + 1),
    }
}

fn precedence(op: BinOp) -> u8 {
    match op {
        BinOp::Eq | BinOp::Neq | BinOp::Lt | BinOp::Gt | BinOp::Lte | BinOp::Gte => 1,
        BinOp::Concat => 2,
        BinOp::Add | BinOp::Sub => 3,
        BinOp::Mul | BinOp::Div => 4,
        BinOp::Pow => 5,
    }
}

fn maybe_paren(expr: &Expr, parent_op: BinOp, is_left: bool) -> String {
    let s = expr_to_string(expr);
    if let Expr::BinOp { op: child_op, .. } = expr {
        let child_prec = precedence(*child_op);
        let parent_prec = precedence(parent_op);
        if child_prec < parent_prec || (child_prec == parent_prec && !is_left) {
            return format!("({})", s);
        }
    }
    s
}
