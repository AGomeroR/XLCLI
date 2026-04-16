use xlcli_core::cell::CellValue;
use xlcli_core::types::CellError;

use crate::ast::Expr;
use crate::eval::{collect_range_values, evaluate, EvalContext};
use crate::registry::{FnSpec, FunctionRegistry};

pub fn register(reg: &mut FunctionRegistry) {
    reg.register(FnSpec { name: "SUM", description: "Adds all numbers in a range", syntax: "SUM(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_sum });
    reg.register(FnSpec { name: "AVERAGE", description: "Returns the average of numbers", syntax: "AVERAGE(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_average });
    reg.register(FnSpec { name: "MIN", description: "Returns the smallest value", syntax: "MIN(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_min });
    reg.register(FnSpec { name: "MAX", description: "Returns the largest value", syntax: "MAX(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_max });
    reg.register(FnSpec { name: "ABS", description: "Returns the absolute value", syntax: "ABS(number)", min_args: 1, max_args: Some(1), eval: fn_abs });
    reg.register(FnSpec { name: "ROUND", description: "Rounds to a specified number of digits", syntax: "ROUND(number, num_digits)", min_args: 1, max_args: Some(2), eval: fn_round });
    reg.register(FnSpec { name: "ROUNDUP", description: "Rounds up away from zero", syntax: "ROUNDUP(number, num_digits)", min_args: 1, max_args: Some(2), eval: fn_roundup });
    reg.register(FnSpec { name: "ROUNDDOWN", description: "Rounds down toward zero", syntax: "ROUNDDOWN(number, num_digits)", min_args: 1, max_args: Some(2), eval: fn_rounddown });
    reg.register(FnSpec { name: "INT", description: "Rounds down to nearest integer", syntax: "INT(number)", min_args: 1, max_args: Some(1), eval: fn_int });
    reg.register(FnSpec { name: "MOD", description: "Returns the remainder after division", syntax: "MOD(number, divisor)", min_args: 2, max_args: Some(2), eval: fn_mod });
    reg.register(FnSpec { name: "POWER", description: "Returns a number raised to a power", syntax: "POWER(number, power)", min_args: 2, max_args: Some(2), eval: fn_power });
    reg.register(FnSpec { name: "SQRT", description: "Returns the square root", syntax: "SQRT(number)", min_args: 1, max_args: Some(1), eval: fn_sqrt });
    reg.register(FnSpec { name: "LOG", description: "Returns the logarithm of a number", syntax: "LOG(number, [base])", min_args: 1, max_args: Some(2), eval: fn_log });
    reg.register(FnSpec { name: "LOG10", description: "Returns the base-10 logarithm", syntax: "LOG10(number)", min_args: 1, max_args: Some(1), eval: fn_log10 });
    reg.register(FnSpec { name: "LN", description: "Returns the natural logarithm", syntax: "LN(number)", min_args: 1, max_args: Some(1), eval: fn_ln });
    reg.register(FnSpec { name: "EXP", description: "Returns e raised to a power", syntax: "EXP(number)", min_args: 1, max_args: Some(1), eval: fn_exp });
    reg.register(FnSpec { name: "PI", description: "Returns the value of pi", syntax: "PI()", min_args: 0, max_args: Some(0), eval: fn_pi });
    reg.register(FnSpec { name: "RAND", description: "Returns a random number between 0 and 1", syntax: "RAND()", min_args: 0, max_args: Some(0), eval: fn_rand });
    reg.register(FnSpec { name: "RANDBETWEEN", description: "Returns a random integer between two values", syntax: "RANDBETWEEN(bottom, top)", min_args: 2, max_args: Some(2), eval: fn_randbetween });
    reg.register(FnSpec { name: "CEILING", description: "Rounds up to nearest multiple", syntax: "CEILING(number, significance)", min_args: 2, max_args: Some(2), eval: fn_ceiling });
    reg.register(FnSpec { name: "FLOOR", description: "Rounds down to nearest multiple", syntax: "FLOOR(number, significance)", min_args: 2, max_args: Some(2), eval: fn_floor });
    reg.register(FnSpec { name: "SIGN", description: "Returns the sign of a number", syntax: "SIGN(number)", min_args: 1, max_args: Some(1), eval: fn_sign });
    reg.register(FnSpec { name: "PRODUCT", description: "Multiplies all numbers", syntax: "PRODUCT(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_product });
    reg.register(FnSpec { name: "SUMPRODUCT", description: "Returns sum of products of ranges", syntax: "SUMPRODUCT(array1, [array2], ...)", min_args: 1, max_args: None, eval: fn_sumproduct });
    reg.register(FnSpec { name: "SIN", description: "Returns the sine of an angle", syntax: "SIN(number)", min_args: 1, max_args: Some(1), eval: fn_sin });
    reg.register(FnSpec { name: "COS", description: "Returns the cosine of an angle", syntax: "COS(number)", min_args: 1, max_args: Some(1), eval: fn_cos });
    reg.register(FnSpec { name: "TAN", description: "Returns the tangent of an angle", syntax: "TAN(number)", min_args: 1, max_args: Some(1), eval: fn_tan });
    reg.register(FnSpec { name: "ASIN", description: "Returns the arcsine", syntax: "ASIN(number)", min_args: 1, max_args: Some(1), eval: fn_asin });
    reg.register(FnSpec { name: "ACOS", description: "Returns the arccosine", syntax: "ACOS(number)", min_args: 1, max_args: Some(1), eval: fn_acos });
    reg.register(FnSpec { name: "ATAN", description: "Returns the arctangent", syntax: "ATAN(number)", min_args: 1, max_args: Some(1), eval: fn_atan });
    reg.register(FnSpec { name: "ATAN2", description: "Returns the arctangent from x and y", syntax: "ATAN2(x_num, y_num)", min_args: 2, max_args: Some(2), eval: fn_atan2 });
    reg.register(FnSpec { name: "DEGREES", description: "Converts radians to degrees", syntax: "DEGREES(angle)", min_args: 1, max_args: Some(1), eval: fn_degrees });
    reg.register(FnSpec { name: "RADIANS", description: "Converts degrees to radians", syntax: "RADIANS(angle)", min_args: 1, max_args: Some(1), eval: fn_radians });
    reg.register(FnSpec { name: "EVEN", description: "Rounds up to nearest even integer", syntax: "EVEN(number)", min_args: 1, max_args: Some(1), eval: fn_even });
    reg.register(FnSpec { name: "ODD", description: "Rounds up to nearest odd integer", syntax: "ODD(number)", min_args: 1, max_args: Some(1), eval: fn_odd });
    reg.register(FnSpec { name: "FACT", description: "Returns the factorial", syntax: "FACT(number)", min_args: 1, max_args: Some(1), eval: fn_fact });
    reg.register(FnSpec { name: "GCD", description: "Returns the greatest common divisor", syntax: "GCD(number1, [number2], ...)", min_args: 2, max_args: None, eval: fn_gcd });
    reg.register(FnSpec { name: "LCM", description: "Returns the least common multiple", syntax: "LCM(number1, [number2], ...)", min_args: 2, max_args: None, eval: fn_lcm });
    reg.register(FnSpec { name: "TRUNC", description: "Truncates to an integer", syntax: "TRUNC(number, [num_digits])", min_args: 1, max_args: Some(2), eval: fn_trunc });
    reg.register(FnSpec { name: "QUOTIENT", description: "Returns integer portion of division", syntax: "QUOTIENT(numerator, denominator)", min_args: 2, max_args: Some(2), eval: fn_quotient });
}

fn collect_numbers(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> Vec<f64> {
    let mut nums = Vec::new();
    for arg in args {
        match arg {
            Expr::Range { start, end } => {
                for val in collect_range_values(start, end, ctx) {
                    if let Some(n) = val.as_f64() {
                        nums.push(n);
                    }
                }
            }
            _ => {
                let val = evaluate(arg, ctx, reg);
                if let Some(n) = val.as_f64() {
                    nums.push(n);
                }
            }
        }
    }
    nums
}

fn eval_one_num(args: &[Expr], idx: usize, ctx: &dyn EvalContext, reg: &FunctionRegistry) -> Option<f64> {
    evaluate(&args[idx], ctx, reg).as_f64()
}

fn fn_sum(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    CellValue::Number(nums.iter().sum())
}

fn fn_average(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    if nums.is_empty() {
        CellValue::Error(CellError::Div0)
    } else {
        CellValue::Number(nums.iter().sum::<f64>() / nums.len() as f64)
    }
}

fn fn_min(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    nums.iter().copied().reduce(f64::min).map(CellValue::Number).unwrap_or(CellValue::Number(0.0))
}

fn fn_max(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    nums.iter().copied().reduce(f64::max).map(CellValue::Number).unwrap_or(CellValue::Number(0.0))
}

fn fn_abs(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_one_num(args, 0, ctx, reg) {
        Some(n) => CellValue::Number(n.abs()),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_round(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_one_num(args, 0, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let digits = if args.len() > 1 { eval_one_num(args, 1, ctx, reg).unwrap_or(0.0) as i32 } else { 0 };
    let factor = 10f64.powi(digits);
    CellValue::Number((n * factor).round() / factor)
}

fn fn_roundup(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_one_num(args, 0, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let digits = if args.len() > 1 { eval_one_num(args, 1, ctx, reg).unwrap_or(0.0) as i32 } else { 0 };
    let factor = 10f64.powi(digits);
    CellValue::Number((n * factor).ceil() / factor)
}

fn fn_rounddown(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_one_num(args, 0, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let digits = if args.len() > 1 { eval_one_num(args, 1, ctx, reg).unwrap_or(0.0) as i32 } else { 0 };
    let factor = 10f64.powi(digits);
    CellValue::Number((n * factor).floor() / factor)
}

fn fn_int(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_one_num(args, 0, ctx, reg) {
        Some(n) => CellValue::Number(n.floor()),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_mod(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_one_num(args, 0, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let d = match eval_one_num(args, 1, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if d == 0.0 { CellValue::Error(CellError::Div0) } else { CellValue::Number(n % d) }
}

fn fn_power(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let base = match eval_one_num(args, 0, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let exp = match eval_one_num(args, 1, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    CellValue::Number(base.powf(exp))
}

fn fn_sqrt(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_one_num(args, 0, ctx, reg) {
        Some(n) if n >= 0.0 => CellValue::Number(n.sqrt()),
        Some(_) => CellValue::Error(CellError::Num),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_log(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_one_num(args, 0, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let base = if args.len() > 1 { eval_one_num(args, 1, ctx, reg).unwrap_or(10.0) } else { 10.0 };
    if n <= 0.0 || base <= 0.0 || base == 1.0 { CellValue::Error(CellError::Num) } else { CellValue::Number(n.log(base)) }
}

fn fn_log10(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_one_num(args, 0, ctx, reg) {
        Some(n) if n > 0.0 => CellValue::Number(n.log10()),
        Some(_) => CellValue::Error(CellError::Num),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_ln(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_one_num(args, 0, ctx, reg) {
        Some(n) if n > 0.0 => CellValue::Number(n.ln()),
        Some(_) => CellValue::Error(CellError::Num),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_exp(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_one_num(args, 0, ctx, reg) {
        Some(n) => CellValue::Number(n.exp()),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_pi(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Number(std::f64::consts::PI)
}

fn fn_rand(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Number(rand_f64())
}

fn fn_randbetween(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let lo = match eval_one_num(args, 0, ctx, reg) { Some(n) => n.ceil() as i64, None => return CellValue::Error(CellError::Value) };
    let hi = match eval_one_num(args, 1, ctx, reg) { Some(n) => n.floor() as i64, None => return CellValue::Error(CellError::Value) };
    if lo > hi { return CellValue::Error(CellError::Value); }
    let range = (hi - lo + 1) as f64;
    CellValue::Number(lo as f64 + (rand_f64() * range).floor())
}

fn fn_ceiling(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_one_num(args, 0, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let sig = match eval_one_num(args, 1, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if sig == 0.0 { return CellValue::Number(0.0); }
    CellValue::Number((n / sig).ceil() * sig)
}

fn fn_floor(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_one_num(args, 0, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let sig = match eval_one_num(args, 1, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if sig == 0.0 { return CellValue::Error(CellError::Div0); }
    CellValue::Number((n / sig).floor() * sig)
}

fn fn_sign(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_one_num(args, 0, ctx, reg) {
        Some(n) => CellValue::Number(if n > 0.0 { 1.0 } else if n < 0.0 { -1.0 } else { 0.0 }),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_product(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    CellValue::Number(nums.iter().product())
}

fn fn_sumproduct(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let ranges: Vec<Vec<f64>> = args.iter().map(|a| {
        match a {
            Expr::Range { start, end } => {
                collect_range_values(start, end, ctx).iter().map(|v| v.as_f64().unwrap_or(0.0)).collect()
            }
            _ => vec![evaluate(a, ctx, reg).as_f64().unwrap_or(0.0)],
        }
    }).collect();
    if ranges.is_empty() { return CellValue::Number(0.0); }
    let len = ranges[0].len();
    if ranges.iter().any(|r| r.len() != len) { return CellValue::Error(CellError::Value); }
    let mut sum = 0.0;
    for i in 0..len {
        let mut prod = 1.0;
        for r in &ranges {
            prod *= r[i];
        }
        sum += prod;
    }
    CellValue::Number(sum)
}

fn fn_sin(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, f64::sin) }
fn fn_cos(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, f64::cos) }
fn fn_tan(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, f64::tan) }
fn fn_asin(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, f64::asin) }
fn fn_acos(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, f64::acos) }
fn fn_atan(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, f64::atan) }

fn fn_atan2(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match eval_one_num(args, 0, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let y = match eval_one_num(args, 1, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    CellValue::Number(x.atan2(y))
}

fn fn_degrees(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, f64::to_degrees) }
fn fn_radians(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, f64::to_radians) }

fn fn_even(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_one_num(args, 0, ctx, reg) {
        Some(n) => {
            let c = if n >= 0.0 { n.ceil() } else { n.floor() };
            let c = c as i64;
            CellValue::Number(if c % 2 == 0 { c } else { c + c.signum() } as f64)
        }
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_odd(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_one_num(args, 0, ctx, reg) {
        Some(n) => {
            let c = if n >= 0.0 { n.ceil() } else { n.floor() };
            let c = c as i64;
            CellValue::Number(if c % 2 != 0 { c } else { c + c.signum() } as f64)
        }
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_fact(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_one_num(args, 0, ctx, reg) {
        Some(n) if n >= 0.0 && n <= 170.0 => {
            let n = n.floor() as u64;
            let mut result = 1u64;
            for i in 2..=n {
                result = result.saturating_mul(i);
            }
            CellValue::Number(result as f64)
        }
        Some(_) => CellValue::Error(CellError::Num),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_gcd(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    if nums.is_empty() { return CellValue::Number(0.0); }
    let mut result = nums[0].abs().floor() as u64;
    for &n in &nums[1..] {
        result = gcd(result, n.abs().floor() as u64);
    }
    CellValue::Number(result as f64)
}

fn fn_lcm(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    if nums.is_empty() { return CellValue::Number(0.0); }
    let mut result = nums[0].abs().floor() as u64;
    for &n in &nums[1..] {
        let b = n.abs().floor() as u64;
        if result == 0 && b == 0 { result = 0; } else { result = result / gcd(result, b) * b; }
    }
    CellValue::Number(result as f64)
}

fn fn_trunc(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_one_num(args, 0, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let digits = if args.len() > 1 { eval_one_num(args, 1, ctx, reg).unwrap_or(0.0) as i32 } else { 0 };
    let factor = 10f64.powi(digits);
    CellValue::Number((n * factor).trunc() / factor)
}

fn fn_quotient(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_one_num(args, 0, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let d = match eval_one_num(args, 1, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if d == 0.0 { CellValue::Error(CellError::Div0) } else { CellValue::Number((n / d).trunc()) }
}

fn unary_math(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry, f: fn(f64) -> f64) -> CellValue {
    match eval_one_num(args, 0, ctx, reg) {
        Some(n) => CellValue::Number(f(n)),
        None => CellValue::Error(CellError::Value),
    }
}

fn gcd(mut a: u64, mut b: u64) -> u64 {
    while b != 0 { let t = b; b = a % b; a = t; }
    a
}

fn rand_f64() -> f64 {
    use std::time::SystemTime;
    let seed = SystemTime::now().duration_since(SystemTime::UNIX_EPOCH).unwrap().subsec_nanos();
    seed as f64 / u32::MAX as f64
}
