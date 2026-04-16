use xlcli_core::cell::CellValue;
use xlcli_core::types::CellError;

use crate::ast::Expr;
use crate::eval::{evaluate, EvalContext};
use crate::registry::{FnSpec, FunctionRegistry};

pub fn register(reg: &mut FunctionRegistry) {
    reg.register(FnSpec { name: "ISBLANK", min_args: 1, max_args: Some(1), eval: fn_isblank });
    reg.register(FnSpec { name: "ISERROR", min_args: 1, max_args: Some(1), eval: fn_iserror });
    reg.register(FnSpec { name: "ISERR", min_args: 1, max_args: Some(1), eval: fn_iserr });
    reg.register(FnSpec { name: "ISNA", min_args: 1, max_args: Some(1), eval: fn_isna });
    reg.register(FnSpec { name: "ISNUMBER", min_args: 1, max_args: Some(1), eval: fn_isnumber });
    reg.register(FnSpec { name: "ISTEXT", min_args: 1, max_args: Some(1), eval: fn_istext });
    reg.register(FnSpec { name: "ISLOGICAL", min_args: 1, max_args: Some(1), eval: fn_islogical });
    reg.register(FnSpec { name: "ISNONTEXT", min_args: 1, max_args: Some(1), eval: fn_isnontext });
    reg.register(FnSpec { name: "ISEVEN", min_args: 1, max_args: Some(1), eval: fn_iseven });
    reg.register(FnSpec { name: "ISODD", min_args: 1, max_args: Some(1), eval: fn_isodd });
    reg.register(FnSpec { name: "TYPE", min_args: 1, max_args: Some(1), eval: fn_type });
    reg.register(FnSpec { name: "N", min_args: 1, max_args: Some(1), eval: fn_n });
    reg.register(FnSpec { name: "NA", min_args: 0, max_args: Some(0), eval: fn_na });
    reg.register(FnSpec { name: "ERROR.TYPE", min_args: 1, max_args: Some(1), eval: fn_error_type });
    reg.register(FnSpec { name: "SHEET", min_args: 0, max_args: Some(1), eval: fn_sheet });
}

fn fn_isblank(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    CellValue::Boolean(matches!(evaluate(&args[0], ctx, reg), CellValue::Empty))
}

fn fn_iserror(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    CellValue::Boolean(matches!(evaluate(&args[0], ctx, reg), CellValue::Error(_)))
}

fn fn_iserr(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let val = evaluate(&args[0], ctx, reg);
    CellValue::Boolean(matches!(val, CellValue::Error(e) if e != CellError::Na))
}

fn fn_isna(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    CellValue::Boolean(matches!(evaluate(&args[0], ctx, reg), CellValue::Error(CellError::Na)))
}

fn fn_isnumber(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    CellValue::Boolean(matches!(evaluate(&args[0], ctx, reg), CellValue::Number(_)))
}

fn fn_istext(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    CellValue::Boolean(matches!(evaluate(&args[0], ctx, reg), CellValue::String(_)))
}

fn fn_islogical(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    CellValue::Boolean(matches!(evaluate(&args[0], ctx, reg), CellValue::Boolean(_)))
}

fn fn_isnontext(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    CellValue::Boolean(!matches!(evaluate(&args[0], ctx, reg), CellValue::String(_)))
}

fn fn_iseven(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match evaluate(&args[0], ctx, reg).as_f64() {
        Some(n) => CellValue::Boolean((n.trunc() as i64) % 2 == 0),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_isodd(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match evaluate(&args[0], ctx, reg).as_f64() {
        Some(n) => CellValue::Boolean((n.trunc() as i64) % 2 != 0),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_type(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let val = evaluate(&args[0], ctx, reg);
    let type_num = match val {
        CellValue::Number(_) => 1.0,
        CellValue::String(_) => 2.0,
        CellValue::Boolean(_) => 4.0,
        CellValue::Error(_) => 16.0,
        CellValue::Array(_) => 64.0,
        CellValue::Empty => 1.0,
        CellValue::DateTime(_) => 1.0,
    };
    CellValue::Number(type_num)
}

fn fn_n(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let val = evaluate(&args[0], ctx, reg);
    match val {
        CellValue::Number(n) => CellValue::Number(n),
        CellValue::Boolean(true) => CellValue::Number(1.0),
        CellValue::Boolean(false) => CellValue::Number(0.0),
        CellValue::Error(e) => CellValue::Error(e),
        _ => CellValue::Number(0.0),
    }
}

fn fn_na(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Error(CellError::Na)
}

fn fn_error_type(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match evaluate(&args[0], ctx, reg) {
        CellValue::Error(e) => {
            let num = match e {
                CellError::Null => 1.0,
                CellError::Div0 => 2.0,
                CellError::Value => 3.0,
                CellError::Ref => 4.0,
                CellError::Name => 5.0,
                CellError::Num => 6.0,
                CellError::Na => 7.0,
                CellError::GettingData => 8.0,
            };
            CellValue::Number(num)
        }
        _ => CellValue::Error(CellError::Na),
    }
}

fn fn_sheet(_args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Number((ctx.current_sheet() + 1) as f64)
}
