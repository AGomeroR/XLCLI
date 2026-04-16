use xlcli_core::cell::CellValue;
use xlcli_core::types::{CellAddr, CellError};

use crate::ast::*;
use crate::registry::FunctionRegistry;

pub trait EvalContext {
    fn get_cell_value(&self, addr: CellAddr) -> CellValue;
    fn current_cell(&self) -> CellAddr;
    fn current_sheet(&self) -> u16;
}

pub fn evaluate(expr: &Expr, ctx: &dyn EvalContext, registry: &FunctionRegistry) -> CellValue {
    match expr {
        Expr::Number(n) => CellValue::Number(*n),
        Expr::String(s) => CellValue::String(s.as_str().into()),
        Expr::Boolean(b) => CellValue::Boolean(*b),
        Expr::Error(e) => CellValue::Error(*e),

        Expr::CellRef { sheet: _, col, row } => {
            let addr = CellAddr::new(ctx.current_sheet(), row.value(), col.value());
            ctx.get_cell_value(addr)
        }

        Expr::Range { .. } => CellValue::Error(CellError::Value),

        Expr::UnaryOp { op, expr } => {
            let val = evaluate(expr, ctx, registry);
            match op {
                UnaryOp::Neg => match val.as_f64() {
                    Some(n) => CellValue::Number(-n),
                    None => CellValue::Error(CellError::Value),
                },
                UnaryOp::Pos => match val.as_f64() {
                    Some(n) => CellValue::Number(n),
                    None => CellValue::Error(CellError::Value),
                },
            }
        }

        Expr::Percent(inner) => {
            let val = evaluate(inner, ctx, registry);
            match val.as_f64() {
                Some(n) => CellValue::Number(n / 100.0),
                None => CellValue::Error(CellError::Value),
            }
        }

        Expr::BinOp { op, left, right } => {
            let lval = evaluate(left, ctx, registry);
            let rval = evaluate(right, ctx, registry);
            eval_binop(*op, &lval, &rval)
        }

        Expr::FnCall { name, args } => {
            if let Some(spec) = registry.get(name) {
                let arg_count = args.len();
                if arg_count < spec.min_args {
                    return CellValue::Error(CellError::Value);
                }
                if let Some(max) = spec.max_args {
                    if arg_count > max {
                        return CellValue::Error(CellError::Value);
                    }
                }
                (spec.eval)(args, ctx, registry)
            } else {
                CellValue::Error(CellError::Name)
            }
        }
    }
}

fn eval_binop(op: BinOp, left: &CellValue, right: &CellValue) -> CellValue {
    match op {
        BinOp::Concat => {
            let ls = left.display_value();
            let rs = right.display_value();
            CellValue::String(format!("{}{}", ls, rs).into())
        }
        BinOp::Eq => CellValue::Boolean(cell_values_equal(left, right)),
        BinOp::Neq => CellValue::Boolean(!cell_values_equal(left, right)),
        _ => {
            let ln = match left.as_f64() {
                Some(n) => n,
                None => return CellValue::Error(CellError::Value),
            };
            let rn = match right.as_f64() {
                Some(n) => n,
                None => return CellValue::Error(CellError::Value),
            };
            match op {
                BinOp::Add => CellValue::Number(ln + rn),
                BinOp::Sub => CellValue::Number(ln - rn),
                BinOp::Mul => CellValue::Number(ln * rn),
                BinOp::Div => {
                    if rn == 0.0 {
                        CellValue::Error(CellError::Div0)
                    } else {
                        CellValue::Number(ln / rn)
                    }
                }
                BinOp::Pow => CellValue::Number(ln.powf(rn)),
                BinOp::Lt => CellValue::Boolean(ln < rn),
                BinOp::Gt => CellValue::Boolean(ln > rn),
                BinOp::Lte => CellValue::Boolean(ln <= rn),
                BinOp::Gte => CellValue::Boolean(ln >= rn),
                _ => unreachable!(),
            }
        }
    }
}

fn cell_values_equal(a: &CellValue, b: &CellValue) -> bool {
    match (a, b) {
        (CellValue::Number(a), CellValue::Number(b)) => (a - b).abs() < f64::EPSILON,
        (CellValue::String(a), CellValue::String(b)) => a.eq_ignore_ascii_case(b),
        (CellValue::Boolean(a), CellValue::Boolean(b)) => a == b,
        (CellValue::Empty, CellValue::Empty) => true,
        (CellValue::Empty, CellValue::Number(n)) | (CellValue::Number(n), CellValue::Empty) => {
            *n == 0.0
        }
        (CellValue::Empty, CellValue::String(s)) | (CellValue::String(s), CellValue::Empty) => {
            s.is_empty()
        }
        _ => false,
    }
}

pub fn collect_range_values(
    start: &Expr,
    end: &Expr,
    ctx: &dyn EvalContext,
) -> Vec<CellValue> {
    let (sr, sc) = match start {
        Expr::CellRef { row, col, .. } => (row.value(), col.value()),
        _ => return vec![],
    };
    let (er, ec) = match end {
        Expr::CellRef { row, col, .. } => (row.value(), col.value()),
        _ => return vec![],
    };

    let min_r = sr.min(er);
    let max_r = sr.max(er);
    let min_c = sc.min(ec);
    let max_c = sc.max(ec);
    let sheet = ctx.current_sheet();

    let mut values = Vec::new();
    for r in min_r..=max_r {
        for c in min_c..=max_c {
            values.push(ctx.get_cell_value(CellAddr::new(sheet, r, c)));
        }
    }
    values
}
