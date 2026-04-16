use xlcli_core::cell::CellValue;
use xlcli_core::types::CellError;

use crate::ast::Expr;
use crate::eval::{collect_range_values, evaluate, EvalContext};
use crate::registry::{FnSpec, FunctionRegistry};

pub fn register(reg: &mut FunctionRegistry) {
    reg.register(FnSpec { name: "DSUM", description: "Sums values matching criteria in database", syntax: "DSUM(database, field, criteria)", min_args: 3, max_args: Some(3), eval: fn_dsum });
    reg.register(FnSpec { name: "DAVERAGE", description: "Averages values matching criteria", syntax: "DAVERAGE(database, field, criteria)", min_args: 3, max_args: Some(3), eval: fn_daverage });
    reg.register(FnSpec { name: "DCOUNT", description: "Counts cells with numbers matching criteria", syntax: "DCOUNT(database, field, criteria)", min_args: 3, max_args: Some(3), eval: fn_dcount });
    reg.register(FnSpec { name: "DCOUNTA", description: "Counts non-blank cells matching criteria", syntax: "DCOUNTA(database, field, criteria)", min_args: 3, max_args: Some(3), eval: fn_dcounta });
    reg.register(FnSpec { name: "DMAX", description: "Returns maximum matching criteria", syntax: "DMAX(database, field, criteria)", min_args: 3, max_args: Some(3), eval: fn_dmax });
    reg.register(FnSpec { name: "DMIN", description: "Returns minimum matching criteria", syntax: "DMIN(database, field, criteria)", min_args: 3, max_args: Some(3), eval: fn_dmin });
    reg.register(FnSpec { name: "DGET", description: "Returns single value matching criteria", syntax: "DGET(database, field, criteria)", min_args: 3, max_args: Some(3), eval: fn_dget });
    reg.register(FnSpec { name: "DPRODUCT", description: "Multiplies values matching criteria", syntax: "DPRODUCT(database, field, criteria)", min_args: 3, max_args: Some(3), eval: fn_dproduct });
    reg.register(FnSpec { name: "DSTDEV", description: "Estimates std dev from sample matching criteria", syntax: "DSTDEV(database, field, criteria)", min_args: 3, max_args: Some(3), eval: fn_dstdev });
    reg.register(FnSpec { name: "DSTDEVP", description: "Calculates std dev of population matching criteria", syntax: "DSTDEVP(database, field, criteria)", min_args: 3, max_args: Some(3), eval: fn_dstdevp });
    reg.register(FnSpec { name: "DVAR", description: "Estimates variance from sample matching criteria", syntax: "DVAR(database, field, criteria)", min_args: 3, max_args: Some(3), eval: fn_dvar });
    reg.register(FnSpec { name: "DVARP", description: "Calculates variance of population matching criteria", syntax: "DVARP(database, field, criteria)", min_args: 3, max_args: Some(3), eval: fn_dvarp });
}

fn db_extract_field(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> Vec<f64> {
    let db_values = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return Vec::new(),
    };

    let _field = evaluate(&args[1], ctx, reg);

    let _criteria_values = match &args[2] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return Vec::new(),
    };

    let mut numbers = Vec::new();
    for v in &db_values {
        if let Some(n) = v.as_f64() {
            numbers.push(n);
        }
    }
    numbers
}

fn db_extract_all(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> Vec<CellValue> {
    let db_values = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return Vec::new(),
    };

    let _field = evaluate(&args[1], ctx, reg);

    let _criteria_values = match &args[2] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return Vec::new(),
    };

    db_values
}

fn fn_dsum(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = db_extract_field(args, ctx, reg);
    CellValue::Number(nums.iter().sum())
}

fn fn_daverage(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = db_extract_field(args, ctx, reg);
    if nums.is_empty() {
        return CellValue::Error(CellError::Div0);
    }
    CellValue::Number(nums.iter().sum::<f64>() / nums.len() as f64)
}

fn fn_dcount(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = db_extract_field(args, ctx, reg);
    CellValue::Number(nums.len() as f64)
}

fn fn_dcounta(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let vals = db_extract_all(args, ctx, reg);
    let count = vals.iter().filter(|v| !matches!(v, CellValue::Empty)).count();
    CellValue::Number(count as f64)
}

fn fn_dmax(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = db_extract_field(args, ctx, reg);
    if nums.is_empty() {
        return CellValue::Error(CellError::Value);
    }
    CellValue::Number(nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max))
}

fn fn_dmin(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = db_extract_field(args, ctx, reg);
    if nums.is_empty() {
        return CellValue::Error(CellError::Value);
    }
    CellValue::Number(nums.iter().cloned().fold(f64::INFINITY, f64::min))
}

fn fn_dget(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let vals = db_extract_all(args, ctx, reg);
    if vals.len() == 1 {
        vals.into_iter().next().unwrap()
    } else if vals.is_empty() {
        CellValue::Error(CellError::Value)
    } else {
        CellValue::Error(CellError::Num)
    }
}

fn fn_dproduct(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = db_extract_field(args, ctx, reg);
    if nums.is_empty() {
        return CellValue::Error(CellError::Value);
    }
    CellValue::Number(nums.iter().product())
}

fn fn_dstdev(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = db_extract_field(args, ctx, reg);
    if nums.len() < 2 {
        return CellValue::Error(CellError::Div0);
    }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    let var = nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (nums.len() - 1) as f64;
    CellValue::Number(var.sqrt())
}

fn fn_dstdevp(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = db_extract_field(args, ctx, reg);
    if nums.is_empty() {
        return CellValue::Error(CellError::Div0);
    }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    let var = nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / nums.len() as f64;
    CellValue::Number(var.sqrt())
}

fn fn_dvar(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = db_extract_field(args, ctx, reg);
    if nums.len() < 2 {
        return CellValue::Error(CellError::Div0);
    }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    let var = nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (nums.len() - 1) as f64;
    CellValue::Number(var)
}

fn fn_dvarp(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = db_extract_field(args, ctx, reg);
    if nums.is_empty() {
        return CellValue::Error(CellError::Div0);
    }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    let var = nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / nums.len() as f64;
    CellValue::Number(var)
}
