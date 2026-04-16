use xlcli_core::cell::CellValue;
use xlcli_core::types::CellError;

use crate::ast::Expr;
use crate::eval::{evaluate, EvalContext};
use crate::registry::{FnSpec, FunctionRegistry};

pub fn register(reg: &mut FunctionRegistry) {
    reg.register(FnSpec { name: "IF", min_args: 2, max_args: Some(3), eval: fn_if });
    reg.register(FnSpec { name: "AND", min_args: 1, max_args: None, eval: fn_and });
    reg.register(FnSpec { name: "OR", min_args: 1, max_args: None, eval: fn_or });
    reg.register(FnSpec { name: "NOT", min_args: 1, max_args: Some(1), eval: fn_not });
    reg.register(FnSpec { name: "XOR", min_args: 1, max_args: None, eval: fn_xor });
    reg.register(FnSpec { name: "IFERROR", min_args: 2, max_args: Some(2), eval: fn_iferror });
    reg.register(FnSpec { name: "IFNA", min_args: 2, max_args: Some(2), eval: fn_ifna });
    reg.register(FnSpec { name: "IFS", min_args: 2, max_args: None, eval: fn_ifs });
    reg.register(FnSpec { name: "SWITCH", min_args: 3, max_args: None, eval: fn_switch });
    reg.register(FnSpec { name: "TRUE", min_args: 0, max_args: Some(0), eval: fn_true });
    reg.register(FnSpec { name: "FALSE", min_args: 0, max_args: Some(0), eval: fn_false });
}

fn to_bool(val: &CellValue) -> Option<bool> {
    match val {
        CellValue::Boolean(b) => Some(*b),
        CellValue::Number(n) => Some(*n != 0.0),
        CellValue::Empty => Some(false),
        _ => None,
    }
}

fn fn_if(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let cond = evaluate(&args[0], ctx, reg);
    let is_true = match to_bool(&cond) {
        Some(b) => b,
        None => return CellValue::Error(CellError::Value),
    };
    if is_true {
        evaluate(&args[1], ctx, reg)
    } else if args.len() > 2 {
        evaluate(&args[2], ctx, reg)
    } else {
        CellValue::Boolean(false)
    }
}

fn fn_and(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    for arg in args {
        let val = evaluate(arg, ctx, reg);
        match to_bool(&val) {
            Some(false) => return CellValue::Boolean(false),
            None => return CellValue::Error(CellError::Value),
            _ => {}
        }
    }
    CellValue::Boolean(true)
}

fn fn_or(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    for arg in args {
        let val = evaluate(arg, ctx, reg);
        match to_bool(&val) {
            Some(true) => return CellValue::Boolean(true),
            None => return CellValue::Error(CellError::Value),
            _ => {}
        }
    }
    CellValue::Boolean(false)
}

fn fn_not(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let val = evaluate(&args[0], ctx, reg);
    match to_bool(&val) {
        Some(b) => CellValue::Boolean(!b),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_xor(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let mut count = 0;
    for arg in args {
        let val = evaluate(arg, ctx, reg);
        match to_bool(&val) {
            Some(true) => count += 1,
            None => return CellValue::Error(CellError::Value),
            _ => {}
        }
    }
    CellValue::Boolean(count % 2 == 1)
}

fn fn_iferror(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let val = evaluate(&args[0], ctx, reg);
    if matches!(val, CellValue::Error(_)) {
        evaluate(&args[1], ctx, reg)
    } else {
        val
    }
}

fn fn_ifna(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let val = evaluate(&args[0], ctx, reg);
    if matches!(val, CellValue::Error(CellError::Na)) {
        evaluate(&args[1], ctx, reg)
    } else {
        val
    }
}

fn fn_ifs(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let mut i = 0;
    while i + 1 < args.len() {
        let cond = evaluate(&args[i], ctx, reg);
        match to_bool(&cond) {
            Some(true) => return evaluate(&args[i + 1], ctx, reg),
            Some(false) => {}
            None => return CellValue::Error(CellError::Value),
        }
        i += 2;
    }
    CellValue::Error(CellError::Na)
}

fn fn_switch(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let expr_val = evaluate(&args[0], ctx, reg);
    let mut i = 1;
    while i + 1 < args.len() {
        let case_val = evaluate(&args[i], ctx, reg);
        if expr_val.display_value() == case_val.display_value() {
            return evaluate(&args[i + 1], ctx, reg);
        }
        i += 2;
    }
    if i < args.len() {
        evaluate(&args[i], ctx, reg)
    } else {
        CellValue::Error(CellError::Na)
    }
}

fn fn_true(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Boolean(true)
}

fn fn_false(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Boolean(false)
}
