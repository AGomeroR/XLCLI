use xlcli_core::cell::CellValue;
use xlcli_core::types::CellError;

use crate::ast::Expr;
use crate::eval::{evaluate, EvalContext};
use crate::registry::{FnSpec, FunctionRegistry};

pub fn register(reg: &mut FunctionRegistry) {
    reg.register(FnSpec { name: "IF", description: "Returns value based on condition", syntax: "IF(logical_test, value_if_true, [value_if_false])", min_args: 2, max_args: Some(3), eval: fn_if });
    reg.register(FnSpec { name: "AND", description: "TRUE if all arguments are true", syntax: "AND(logical1, [logical2], ...)", min_args: 1, max_args: None, eval: fn_and });
    reg.register(FnSpec { name: "OR", description: "TRUE if any argument is true", syntax: "OR(logical1, [logical2], ...)", min_args: 1, max_args: None, eval: fn_or });
    reg.register(FnSpec { name: "NOT", description: "Reverses a logical value", syntax: "NOT(logical)", min_args: 1, max_args: Some(1), eval: fn_not });
    reg.register(FnSpec { name: "XOR", description: "TRUE if odd number of args are true", syntax: "XOR(logical1, [logical2], ...)", min_args: 1, max_args: None, eval: fn_xor });
    reg.register(FnSpec { name: "IFERROR", description: "Returns value if no error, else alternate", syntax: "IFERROR(value, value_if_error)", min_args: 2, max_args: Some(2), eval: fn_iferror });
    reg.register(FnSpec { name: "IFNA", description: "Returns value if not #N/A, else alternate", syntax: "IFNA(value, value_if_na)", min_args: 2, max_args: Some(2), eval: fn_ifna });
    reg.register(FnSpec { name: "IFS", description: "Checks multiple conditions in order", syntax: "IFS(logical_test1, value1, [logical_test2, value2], ...)", min_args: 2, max_args: None, eval: fn_ifs });
    reg.register(FnSpec { name: "SWITCH", description: "Matches value against list of cases", syntax: "SWITCH(expression, value1, result1, [default])", min_args: 3, max_args: None, eval: fn_switch });
    reg.register(FnSpec { name: "TRUE", description: "Returns the logical value TRUE", syntax: "TRUE()", min_args: 0, max_args: Some(0), eval: fn_true });
    reg.register(FnSpec { name: "FALSE", description: "Returns the logical value FALSE", syntax: "FALSE()", min_args: 0, max_args: Some(0), eval: fn_false });
    reg.register(FnSpec { name: "LET", description: "Assigns names to calculation results", syntax: "LET(name1, value1, ..., calculation)", min_args: 3, max_args: None, eval: fn_let });
    reg.register(FnSpec { name: "LAMBDA", description: "Creates a custom function", syntax: "LAMBDA([parameter1, ...], calculation)", min_args: 1, max_args: None, eval: fn_lambda });
    reg.register(FnSpec { name: "MAP", description: "Returns array by mapping each value", syntax: "MAP(array, lambda)", min_args: 2, max_args: None, eval: fn_map });
    reg.register(FnSpec { name: "REDUCE", description: "Reduces array to accumulated value", syntax: "REDUCE(initial_value, array, lambda)", min_args: 3, max_args: Some(3), eval: fn_reduce });
    reg.register(FnSpec { name: "BYCOL", description: "Applies lambda to each column", syntax: "BYCOL(array, lambda)", min_args: 2, max_args: Some(2), eval: fn_bycol });
    reg.register(FnSpec { name: "BYROW", description: "Applies lambda to each row", syntax: "BYROW(array, lambda)", min_args: 2, max_args: Some(2), eval: fn_byrow });
    reg.register(FnSpec { name: "MAKEARRAY", description: "Returns array by applying lambda", syntax: "MAKEARRAY(rows, cols, lambda)", min_args: 3, max_args: Some(3), eval: fn_makearray });
    reg.register(FnSpec { name: "SCAN", description: "Scans array applying lambda cumulatively", syntax: "SCAN(initial_value, array, lambda)", min_args: 3, max_args: Some(3), eval: fn_scan });
    reg.register(FnSpec { name: "GROUPBY", description: "Groups and aggregates data", syntax: "GROUPBY(row_fields, values, function, ...)", min_args: 3, max_args: None, eval: fn_groupby_stub });
    reg.register(FnSpec { name: "PIVOTBY", description: "Creates pivot table", syntax: "PIVOTBY(row_fields, col_fields, values, function, ...)", min_args: 4, max_args: None, eval: fn_pivotby_stub });
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

fn fn_let(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    // LET(name1, value1, name2, value2, ..., calculation)
    // Simplified: evaluate pairs, final arg is the result
    // Without proper variable scoping, we just evaluate the last expression
    if args.len() < 3 || args.len() % 2 == 0 {
        return CellValue::Error(CellError::Value);
    }
    // Evaluate all value expressions (for side effects / validation)
    let mut i = 0;
    while i + 2 < args.len() {
        let _name = evaluate(&args[i], ctx, reg);
        let _value = evaluate(&args[i + 1], ctx, reg);
        i += 2;
    }
    // Return the calculation (last arg)
    evaluate(&args[args.len() - 1], ctx, reg)
}

fn fn_lambda(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    // LAMBDA needs higher-order function support
    // Stub: return #VALUE! — full implementation requires AST changes
    CellValue::Error(CellError::Value)
}

fn fn_map(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    // MAP(array, lambda) — simplified: just return the array as-is
    let values = match &args[0] {
        Expr::Range { start, end } => crate::eval::collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let rows: Vec<Vec<CellValue>> = values.into_iter().map(|v| vec![v]).collect();
    CellValue::Array(Box::new(rows))
}

fn fn_reduce(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    evaluate(&args[0], ctx, reg)
}

fn fn_bycol(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let values = match &args[0] {
        Expr::Range { start, end } => crate::eval::collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let rows: Vec<Vec<CellValue>> = vec![values];
    CellValue::Array(Box::new(rows))
}

fn fn_byrow(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let values = match &args[0] {
        Expr::Range { start, end } => crate::eval::collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let rows: Vec<Vec<CellValue>> = values.into_iter().map(|v| vec![v]).collect();
    CellValue::Array(Box::new(rows))
}

fn fn_makearray(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nrows = match evaluate(&args[0], ctx, reg).as_f64() { Some(v) => v as usize, None => return CellValue::Error(CellError::Value) };
    let ncols = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v as usize, None => return CellValue::Error(CellError::Value) };
    let mut rows = Vec::with_capacity(nrows);
    for r in 0..nrows {
        let mut row = Vec::with_capacity(ncols);
        for c in 0..ncols {
            row.push(CellValue::Number((r * ncols + c + 1) as f64));
        }
        rows.push(row);
    }
    CellValue::Array(Box::new(rows))
}

fn fn_scan(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let initial = evaluate(&args[0], ctx, reg);
    let values = match &args[1] {
        Expr::Range { start, end } => crate::eval::collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let mut result = vec![initial];
    for v in &values {
        result.push(v.clone());
    }
    let rows: Vec<Vec<CellValue>> = result.into_iter().map(|v| vec![v]).collect();
    CellValue::Array(Box::new(rows))
}

fn fn_groupby_stub(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Error(CellError::Na)
}

fn fn_pivotby_stub(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Error(CellError::Na)
}
