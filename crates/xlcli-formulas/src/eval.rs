use xlcli_core::cell::CellValue;
use xlcli_core::types::{CellAddr, CellError};

use crate::ast::*;
use crate::registry::FunctionRegistry;

pub trait EvalContext {
    fn get_cell_value(&self, addr: CellAddr) -> CellValue;
    fn current_cell(&self) -> CellAddr;
    fn current_sheet(&self) -> u16;
    fn resolve_sheet(&self, name: &str) -> Option<u16>;
    fn resolve_named_range(&self, _name: &str) -> Option<(CellAddr, CellAddr)> { None }
}

pub fn evaluate(expr: &Expr, ctx: &dyn EvalContext, registry: &FunctionRegistry) -> CellValue {
    match expr {
        Expr::Number(n) => CellValue::Number(*n),
        Expr::String(s) => CellValue::String(s.as_str().into()),
        Expr::Boolean(b) => CellValue::Boolean(*b),
        Expr::Error(e) => CellValue::Error(*e),

        Expr::CellRef { sheet, col, row } => {
            let sheet_idx = match sheet {
                Some(name) => match ctx.resolve_sheet(name) {
                    Some(idx) => idx,
                    None => return CellValue::Error(CellError::Ref),
                },
                None => ctx.current_sheet(),
            };
            let addr = CellAddr::new(sheet_idx, row.value(), col.value());
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

        Expr::NamedRef(name) => {
            match ctx.resolve_named_range(name) {
                Some((start, _end)) => ctx.get_cell_value(start),
                None => CellValue::Error(CellError::Name),
            }
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

pub fn eval_as_range(arg: &crate::ast::Expr, ctx: &dyn EvalContext) -> Vec<CellValue> {
    match arg {
        crate::ast::Expr::Range { start, end } => collect_range_values(start, end, ctx),
        crate::ast::Expr::NamedRef(name) => collect_named_range_values(name, ctx),
        other => {
            let val = evaluate(other, ctx, &crate::registry::FunctionRegistry::default());
            if val.is_empty() { vec![] } else { vec![val] }
        }
    }
}

pub fn collect_named_range_values(name: &str, ctx: &dyn EvalContext) -> Vec<CellValue> {
    match ctx.resolve_named_range(name) {
        Some((start, end)) => {
            let mut values = Vec::new();
            for r in start.row..=end.row {
                for c in start.col..=end.col {
                    values.push(ctx.get_cell_value(CellAddr::new(start.sheet, r, c)));
                }
            }
            values
        }
        None => vec![],
    }
}

pub fn collect_range_values(
    start: &Expr,
    end: &Expr,
    ctx: &dyn EvalContext,
) -> Vec<CellValue> {
    let (sr, sc, sheet_name) = match start {
        Expr::CellRef { row, col, sheet } => (row.value(), col.value(), sheet.clone()),
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
    let sheet = match sheet_name {
        Some(ref name) => ctx.resolve_sheet(name).unwrap_or(ctx.current_sheet()),
        None => ctx.current_sheet(),
    };

    let mut values = Vec::new();
    for r in min_r..=max_r {
        for c in min_c..=max_c {
            values.push(ctx.get_cell_value(CellAddr::new(sheet, r, c)));
        }
    }
    values
}
