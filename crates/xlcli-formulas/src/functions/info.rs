use xlcli_core::cell::CellValue;
use xlcli_core::types::CellError;

use crate::ast::Expr;
use crate::eval::{evaluate, EvalContext};
use crate::registry::{FnSpec, FunctionRegistry};

pub fn register(reg: &mut FunctionRegistry) {
    reg.register(FnSpec { name: "ISBLANK", description: "TRUE if cell is empty", syntax: "ISBLANK(value)", min_args: 1, max_args: Some(1), eval: fn_isblank });
    reg.register(FnSpec { name: "ISERROR", description: "TRUE if value is any error", syntax: "ISERROR(value)", min_args: 1, max_args: Some(1), eval: fn_iserror });
    reg.register(FnSpec { name: "ISERR", description: "TRUE if value is error except #N/A", syntax: "ISERR(value)", min_args: 1, max_args: Some(1), eval: fn_iserr });
    reg.register(FnSpec { name: "ISNA", description: "TRUE if value is #N/A", syntax: "ISNA(value)", min_args: 1, max_args: Some(1), eval: fn_isna });
    reg.register(FnSpec { name: "ISNUMBER", description: "TRUE if value is a number", syntax: "ISNUMBER(value)", min_args: 1, max_args: Some(1), eval: fn_isnumber });
    reg.register(FnSpec { name: "ISTEXT", description: "TRUE if value is text", syntax: "ISTEXT(value)", min_args: 1, max_args: Some(1), eval: fn_istext });
    reg.register(FnSpec { name: "ISLOGICAL", description: "TRUE if value is boolean", syntax: "ISLOGICAL(value)", min_args: 1, max_args: Some(1), eval: fn_islogical });
    reg.register(FnSpec { name: "ISNONTEXT", description: "TRUE if value is not text", syntax: "ISNONTEXT(value)", min_args: 1, max_args: Some(1), eval: fn_isnontext });
    reg.register(FnSpec { name: "ISEVEN", description: "TRUE if number is even", syntax: "ISEVEN(number)", min_args: 1, max_args: Some(1), eval: fn_iseven });
    reg.register(FnSpec { name: "ISODD", description: "TRUE if number is odd", syntax: "ISODD(number)", min_args: 1, max_args: Some(1), eval: fn_isodd });
    reg.register(FnSpec { name: "TYPE", description: "Returns type of value as number", syntax: "TYPE(value)", min_args: 1, max_args: Some(1), eval: fn_type });
    reg.register(FnSpec { name: "N", description: "Converts value to a number", syntax: "N(value)", min_args: 1, max_args: Some(1), eval: fn_n });
    reg.register(FnSpec { name: "NA", description: "Returns the #N/A error value", syntax: "NA()", min_args: 0, max_args: Some(0), eval: fn_na });
    reg.register(FnSpec { name: "ERROR.TYPE", description: "Returns number for error type", syntax: "ERROR.TYPE(error_val)", min_args: 1, max_args: Some(1), eval: fn_error_type });
    reg.register(FnSpec { name: "SHEET", description: "Returns the sheet number", syntax: "SHEET([value])", min_args: 0, max_args: Some(1), eval: fn_sheet });
    reg.register(FnSpec { name: "ISREF", description: "TRUE if value is a reference", syntax: "ISREF(value)", min_args: 1, max_args: Some(1), eval: fn_isref });
    reg.register(FnSpec { name: "ISFORMULA", description: "TRUE if cell contains a formula", syntax: "ISFORMULA(reference)", min_args: 1, max_args: Some(1), eval: fn_isformula });
    reg.register(FnSpec { name: "SHEETS", description: "Returns number of sheets", syntax: "SHEETS([reference])", min_args: 0, max_args: Some(1), eval: fn_sheets });
    reg.register(FnSpec { name: "CELL", description: "Returns info about a cell", syntax: "CELL(info_type, [reference])", min_args: 1, max_args: Some(2), eval: fn_cell });
    reg.register(FnSpec { name: "INFO", description: "Returns system environment info", syntax: "INFO(type_text)", min_args: 1, max_args: Some(1), eval: fn_info });
    reg.register(FnSpec { name: "NULL", description: "Returns empty/null value", syntax: "NULL()", min_args: 0, max_args: Some(0), eval: fn_null });
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

fn fn_isref(args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Boolean(matches!(&args[0], Expr::CellRef { .. } | Expr::Range { .. }))
}

fn fn_isformula(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Boolean(false) // Would need formula storage access
}

fn fn_sheets(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Number(1.0) // Stub — would need workbook access
}

fn fn_cell(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let info_type = evaluate(&args[0], ctx, reg).display_value().to_lowercase();
    match info_type.as_str() {
        "row" => CellValue::Number((ctx.current_cell().row + 1) as f64),
        "col" => CellValue::Number((ctx.current_cell().col + 1) as f64),
        _ => CellValue::Error(CellError::Value),
    }
}

fn fn_null(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Empty
}

fn fn_info(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let info_type = evaluate(&args[0], ctx, reg).display_value().to_lowercase();
    match info_type.as_str() {
        "osversion" => CellValue::String("xlcli".into()),
        "system" => CellValue::String("xlcli".into()),
        _ => CellValue::Error(CellError::Value),
    }
}
