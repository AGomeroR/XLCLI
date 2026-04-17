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
    reg.register(FnSpec { name: "CEILING.MATH", description: "Rounds up to nearest multiple", syntax: "CEILING.MATH(number, [significance], [mode])", min_args: 1, max_args: Some(3), eval: fn_ceiling_math });
    reg.register(FnSpec { name: "FLOOR.MATH", description: "Rounds down to nearest multiple", syntax: "FLOOR.MATH(number, [significance], [mode])", min_args: 1, max_args: Some(3), eval: fn_floor_math });
    reg.register(FnSpec { name: "SQRTPI", description: "Returns square root of number * pi", syntax: "SQRTPI(number)", min_args: 1, max_args: Some(1), eval: fn_sqrtpi });
    reg.register(FnSpec { name: "SERIESSUM", description: "Returns sum of power series", syntax: "SERIESSUM(x, n, m, coefficients)", min_args: 4, max_args: Some(4), eval: fn_seriessum });
    reg.register(FnSpec { name: "FACTDOUBLE", description: "Returns double factorial", syntax: "FACTDOUBLE(number)", min_args: 1, max_args: Some(1), eval: fn_factdouble });
    reg.register(FnSpec { name: "COMBIN", description: "Returns combinations", syntax: "COMBIN(number, number_chosen)", min_args: 2, max_args: Some(2), eval: fn_combin });
    reg.register(FnSpec { name: "COMBINA", description: "Returns combinations with repetition", syntax: "COMBINA(number, number_chosen)", min_args: 2, max_args: Some(2), eval: fn_combina });
    reg.register(FnSpec { name: "MULTINOMIAL", description: "Returns multinomial coefficient", syntax: "MULTINOMIAL(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_multinomial });
    reg.register(FnSpec { name: "MROUND", description: "Rounds to nearest multiple", syntax: "MROUND(number, multiple)", min_args: 2, max_args: Some(2), eval: fn_mround });
    reg.register(FnSpec { name: "ROMAN", description: "Converts to Roman numeral text", syntax: "ROMAN(number, [form])", min_args: 1, max_args: Some(2), eval: fn_roman });
    reg.register(FnSpec { name: "ARABIC", description: "Converts Roman numeral to number", syntax: "ARABIC(text)", min_args: 1, max_args: Some(1), eval: fn_arabic });
    reg.register(FnSpec { name: "BASE", description: "Converts number to text in given base", syntax: "BASE(number, radix, [min_length])", min_args: 2, max_args: Some(3), eval: fn_base });
    reg.register(FnSpec { name: "DECIMAL", description: "Converts text in given base to number", syntax: "DECIMAL(text, radix)", min_args: 2, max_args: Some(2), eval: fn_decimal });
    reg.register(FnSpec { name: "SUMSQ", description: "Returns sum of squares", syntax: "SUMSQ(number1, [number2], ...)", min_args: 1, max_args: None, eval: fn_sumsq });
    reg.register(FnSpec { name: "SINH", description: "Returns hyperbolic sine", syntax: "SINH(number)", min_args: 1, max_args: Some(1), eval: fn_sinh });
    reg.register(FnSpec { name: "COSH", description: "Returns hyperbolic cosine", syntax: "COSH(number)", min_args: 1, max_args: Some(1), eval: fn_cosh });
    reg.register(FnSpec { name: "TANH", description: "Returns hyperbolic tangent", syntax: "TANH(number)", min_args: 1, max_args: Some(1), eval: fn_tanh });
    reg.register(FnSpec { name: "ASINH", description: "Returns inverse hyperbolic sine", syntax: "ASINH(number)", min_args: 1, max_args: Some(1), eval: fn_asinh });
    reg.register(FnSpec { name: "ACOSH", description: "Returns inverse hyperbolic cosine", syntax: "ACOSH(number)", min_args: 1, max_args: Some(1), eval: fn_acosh });
    reg.register(FnSpec { name: "ATANH", description: "Returns inverse hyperbolic tangent", syntax: "ATANH(number)", min_args: 1, max_args: Some(1), eval: fn_atanh });
    reg.register(FnSpec { name: "SEC", description: "Returns the secant", syntax: "SEC(number)", min_args: 1, max_args: Some(1), eval: fn_sec });
    reg.register(FnSpec { name: "CSC", description: "Returns the cosecant", syntax: "CSC(number)", min_args: 1, max_args: Some(1), eval: fn_csc });
    reg.register(FnSpec { name: "COT", description: "Returns the cotangent", syntax: "COT(number)", min_args: 1, max_args: Some(1), eval: fn_cot });
    reg.register(FnSpec { name: "SECH", description: "Returns hyperbolic secant", syntax: "SECH(number)", min_args: 1, max_args: Some(1), eval: fn_sech });
    reg.register(FnSpec { name: "CSCH", description: "Returns hyperbolic cosecant", syntax: "CSCH(number)", min_args: 1, max_args: Some(1), eval: fn_csch });
    reg.register(FnSpec { name: "COTH", description: "Returns hyperbolic cotangent", syntax: "COTH(number)", min_args: 1, max_args: Some(1), eval: fn_coth });
    reg.register(FnSpec { name: "SUBTOTAL", description: "Returns subtotal with function number", syntax: "SUBTOTAL(function_num, ref1, ...)", min_args: 2, max_args: None, eval: fn_subtotal });
    reg.register(FnSpec { name: "AGGREGATE", description: "Returns aggregate with options", syntax: "AGGREGATE(function_num, options, ref1, ...)", min_args: 3, max_args: None, eval: fn_aggregate });
    reg.register(FnSpec { name: "RANDARRAY", description: "Returns array of random numbers", syntax: "RANDARRAY([rows], [columns], [min], [max], [whole_number])", min_args: 0, max_args: Some(5), eval: fn_randarray });
    reg.register(FnSpec { name: "MMULT", description: "Returns matrix product of two arrays", syntax: "MMULT(array1, array2)", min_args: 2, max_args: None, eval: fn_mmult });
    reg.register(FnSpec { name: "SUMX2MY2", description: "Returns sum of x^2 - y^2", syntax: "SUMX2MY2(array_x, array_y)", min_args: 2, max_args: Some(2), eval: fn_sumx2my2 });
    reg.register(FnSpec { name: "SUMX2PY2", description: "Returns sum of x^2 + y^2", syntax: "SUMX2PY2(array_x, array_y)", min_args: 2, max_args: Some(2), eval: fn_sumx2py2 });
    reg.register(FnSpec { name: "SUMXMY2", description: "Returns sum of (x-y)^2", syntax: "SUMXMY2(array_x, array_y)", min_args: 2, max_args: Some(2), eval: fn_sumxmy2 });
    reg.register(FnSpec { name: "MDETERM", description: "Returns matrix determinant", syntax: "MDETERM(array)", min_args: 1, max_args: Some(1), eval: fn_mdeterm });
    reg.register(FnSpec { name: "MINVERSE", description: "Returns matrix inverse (2x2)", syntax: "MINVERSE(array)", min_args: 1, max_args: Some(1), eval: fn_minverse });
    reg.register(FnSpec { name: "CONVERT", description: "Converts between units", syntax: "CONVERT(number, from_unit, to_unit)", min_args: 3, max_args: Some(3), eval: fn_convert });
    reg.register(FnSpec { name: "DELTA", description: "Tests whether two values are equal", syntax: "DELTA(number1, [number2])", min_args: 1, max_args: Some(2), eval: fn_delta });
    reg.register(FnSpec { name: "GESTEP", description: "Tests whether number >= step", syntax: "GESTEP(number, [step])", min_args: 1, max_args: Some(2), eval: fn_gestep });
    reg.register(FnSpec { name: "ACOT", description: "Returns arccotangent", syntax: "ACOT(number)", min_args: 1, max_args: Some(1), eval: fn_acot });
    reg.register(FnSpec { name: "ACOTH", description: "Returns hyperbolic arccotangent", syntax: "ACOTH(number)", min_args: 1, max_args: Some(1), eval: fn_acoth });
    reg.register(FnSpec { name: "CEILING.PRECISE", description: "Rounds up to nearest significance", syntax: "CEILING.PRECISE(number, [significance])", min_args: 1, max_args: Some(2), eval: fn_ceiling });
    reg.register(FnSpec { name: "FLOOR.PRECISE", description: "Rounds down to nearest significance", syntax: "FLOOR.PRECISE(number, [significance])", min_args: 1, max_args: Some(2), eval: fn_floor });
    reg.register(FnSpec { name: "ISO.CEILING", description: "Rounds up to nearest integer or significance", syntax: "ISO.CEILING(number, [significance])", min_args: 1, max_args: Some(2), eval: fn_ceiling });
    reg.register(FnSpec { name: "MUNIT", description: "Returns the unit matrix", syntax: "MUNIT(dimension)", min_args: 1, max_args: Some(1), eval: fn_munit });
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
            Expr::NamedRef(name) => {
                for val in crate::eval::collect_named_range_values(name, ctx) {
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

fn fn_ceiling_math(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_one_num(args, 0, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let sig = if args.len() > 1 { eval_one_num(args, 1, ctx, reg).unwrap_or(1.0) } else { 1.0 };
    let mode = if args.len() > 2 { eval_one_num(args, 2, ctx, reg).unwrap_or(0.0) } else { 0.0 };
    if sig == 0.0 { return CellValue::Number(0.0); }
    let sig = sig.abs();
    if n >= 0.0 || mode == 0.0 {
        CellValue::Number((n / sig).ceil() * sig)
    } else {
        CellValue::Number(-((-n / sig).floor() * sig))
    }
}

fn fn_floor_math(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_one_num(args, 0, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let sig = if args.len() > 1 { eval_one_num(args, 1, ctx, reg).unwrap_or(1.0) } else { 1.0 };
    let mode = if args.len() > 2 { eval_one_num(args, 2, ctx, reg).unwrap_or(0.0) } else { 0.0 };
    if sig == 0.0 { return CellValue::Number(0.0); }
    let sig = sig.abs();
    if n >= 0.0 || mode == 0.0 {
        CellValue::Number((n / sig).floor() * sig)
    } else {
        CellValue::Number(-((-n / sig).ceil() * sig))
    }
}

fn fn_sqrtpi(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_one_num(args, 0, ctx, reg) {
        Some(n) if n >= 0.0 => CellValue::Number((n * std::f64::consts::PI).sqrt()),
        Some(_) => CellValue::Error(CellError::Num),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_seriessum(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match eval_one_num(args, 0, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let n = match eval_one_num(args, 1, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let m = match eval_one_num(args, 2, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let coeffs = match &args[3] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx)
            .iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>(),
        _ => vec![evaluate(&args[3], ctx, reg).as_f64().unwrap_or(0.0)],
    };
    let mut sum = 0.0;
    for (i, &c) in coeffs.iter().enumerate() {
        sum += c * x.powf(n + m * i as f64);
    }
    CellValue::Number(sum)
}

fn fn_factdouble(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_one_num(args, 0, ctx, reg) {
        Some(n) if n >= -1.0 && n <= 170.0 => {
            let n = n.floor() as i64;
            if n <= 0 { return CellValue::Number(1.0); }
            let mut result = 1.0;
            let mut i = n;
            while i > 0 { result *= i as f64; i -= 2; }
            CellValue::Number(result)
        }
        Some(_) => CellValue::Error(CellError::Num),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_combin(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_one_num(args, 0, ctx, reg) { Some(v) => v.floor() as u64, None => return CellValue::Error(CellError::Value) };
    let k = match eval_one_num(args, 1, ctx, reg) { Some(v) => v.floor() as u64, None => return CellValue::Error(CellError::Value) };
    if k > n { return CellValue::Error(CellError::Num); }
    let k = k.min(n - k);
    let mut result = 1.0;
    for i in 0..k { result *= (n - i) as f64 / (i + 1) as f64; }
    CellValue::Number(result.round())
}

fn fn_combina(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_one_num(args, 0, ctx, reg) { Some(v) => v.floor() as u64, None => return CellValue::Error(CellError::Value) };
    let k = match eval_one_num(args, 1, ctx, reg) { Some(v) => v.floor() as u64, None => return CellValue::Error(CellError::Value) };
    let total = n + k - 1;
    let choose = k;
    if choose > total { return CellValue::Error(CellError::Num); }
    let choose = choose.min(total - choose);
    let mut result = 1.0;
    for i in 0..choose { result *= (total - i) as f64 / (i + 1) as f64; }
    CellValue::Number(result.round())
}

fn fn_multinomial(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    let sum: f64 = nums.iter().sum();
    let mut result = 1.0;
    let mut running = 0.0;
    for &n in &nums {
        let n = n.floor();
        for i in 1..=(n as u64) {
            running += 1.0;
            result *= running / i as f64;
        }
    }
    let _ = sum;
    CellValue::Number(result.round())
}

fn fn_mround(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_one_num(args, 0, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let m = match eval_one_num(args, 1, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    if m == 0.0 { return CellValue::Number(0.0); }
    if n.signum() != m.signum() && n != 0.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number((n / m).round() * m)
}

fn fn_roman(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_one_num(args, 0, ctx, reg) { Some(n) => n.floor() as i64, None => return CellValue::Error(CellError::Value) };
    if n < 0 || n > 3999 { return CellValue::Error(CellError::Value); }
    if n == 0 { return CellValue::String("".into()); }
    let vals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
                (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")];
    let mut result = String::new();
    let mut n = n;
    for &(val, sym) in &vals {
        while n >= val { result.push_str(sym); n -= val; }
    }
    CellValue::String(result.into())
}

fn fn_arabic(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = evaluate(&args[0], ctx, reg).display_value().to_uppercase();
    let mut total = 0i64;
    let mut prev = 0i64;
    for c in s.chars().rev() {
        let val = match c {
            'I' => 1, 'V' => 5, 'X' => 10, 'L' => 50, 'C' => 100, 'D' => 500, 'M' => 1000,
            _ => return CellValue::Error(CellError::Value),
        };
        if val < prev { total -= val; } else { total += val; }
        prev = val;
    }
    CellValue::Number(total as f64)
}

fn fn_base(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_one_num(args, 0, ctx, reg) { Some(n) => n.floor() as u64, None => return CellValue::Error(CellError::Value) };
    let radix = match eval_one_num(args, 1, ctx, reg) { Some(n) => n.floor() as u32, None => return CellValue::Error(CellError::Value) };
    if !(2..=36).contains(&radix) { return CellValue::Error(CellError::Num); }
    let min_len = if args.len() > 2 { eval_one_num(args, 2, ctx, reg).unwrap_or(0.0).floor() as usize } else { 0 };
    let mut result = String::new();
    let mut num = n;
    if num == 0 { result.push('0'); }
    while num > 0 {
        let digit = (num % radix as u64) as u32;
        result.insert(0, char::from_digit(digit, radix).unwrap_or('0').to_ascii_uppercase());
        num /= radix as u64;
    }
    while result.len() < min_len { result.insert(0, '0'); }
    CellValue::String(result.into())
}

fn fn_decimal(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let text = evaluate(&args[0], ctx, reg).display_value();
    let radix = match eval_one_num(args, 1, ctx, reg) { Some(n) => n.floor() as u32, None => return CellValue::Error(CellError::Value) };
    if !(2..=36).contains(&radix) { return CellValue::Error(CellError::Num); }
    match u64::from_str_radix(&text, radix) {
        Ok(n) => CellValue::Number(n as f64),
        Err(_) => CellValue::Error(CellError::Num),
    }
}

fn fn_sumsq(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    CellValue::Number(nums.iter().map(|n| n * n).sum())
}

fn fn_sinh(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, f64::sinh) }
fn fn_cosh(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, f64::cosh) }
fn fn_tanh(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, f64::tanh) }
fn fn_asinh(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, f64::asinh) }
fn fn_acosh(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_one_num(args, 0, ctx, reg) {
        Some(n) if n >= 1.0 => CellValue::Number(n.acosh()),
        Some(_) => CellValue::Error(CellError::Num),
        None => CellValue::Error(CellError::Value),
    }
}
fn fn_atanh(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_one_num(args, 0, ctx, reg) {
        Some(n) if (-1.0..1.0).contains(&n) => CellValue::Number(n.atanh()),
        Some(_) => CellValue::Error(CellError::Num),
        None => CellValue::Error(CellError::Value),
    }
}
fn fn_sec(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, |x| 1.0 / x.cos()) }
fn fn_csc(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, |x| 1.0 / x.sin()) }
fn fn_cot(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, |x| x.cos() / x.sin()) }
fn fn_sech(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, |x| 1.0 / x.cosh()) }
fn fn_csch(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, |x| 1.0 / x.sinh()) }
fn fn_coth(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue { unary_math(args, ctx, reg, |x| x.cosh() / x.sinh()) }

fn fn_subtotal(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let func_num = match eval_one_num(args, 0, ctx, reg) { Some(n) => n as i32, None => return CellValue::Error(CellError::Value) };
    let rest = &args[1..];
    let nums = collect_numbers(rest, ctx, reg);
    match func_num {
        1 | 101 => { if nums.is_empty() { CellValue::Error(CellError::Div0) } else { CellValue::Number(nums.iter().sum::<f64>() / nums.len() as f64) } }
        2 | 102 => CellValue::Number(nums.len() as f64),
        3 | 103 => CellValue::Number(nums.len() as f64),
        4 | 104 => nums.iter().copied().reduce(f64::max).map(CellValue::Number).unwrap_or(CellValue::Number(0.0)),
        5 | 105 => nums.iter().copied().reduce(f64::min).map(CellValue::Number).unwrap_or(CellValue::Number(0.0)),
        6 | 106 => CellValue::Number(nums.iter().product()),
        9 | 109 => CellValue::Number(nums.iter().sum()),
        _ => CellValue::Error(CellError::Value),
    }
}

fn fn_aggregate(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let func_num = match eval_one_num(args, 0, ctx, reg) { Some(n) => n as i32, None => return CellValue::Error(CellError::Value) };
    let _options = match eval_one_num(args, 1, ctx, reg) { Some(n) => n as i32, None => return CellValue::Error(CellError::Value) };
    let rest = &args[2..];
    let nums = collect_numbers(rest, ctx, reg);
    match func_num {
        1 => { if nums.is_empty() { CellValue::Error(CellError::Div0) } else { CellValue::Number(nums.iter().sum::<f64>() / nums.len() as f64) } }
        2 => CellValue::Number(nums.len() as f64),
        4 => nums.iter().copied().reduce(f64::max).map(CellValue::Number).unwrap_or(CellValue::Number(0.0)),
        5 => nums.iter().copied().reduce(f64::min).map(CellValue::Number).unwrap_or(CellValue::Number(0.0)),
        6 => CellValue::Number(nums.iter().product()),
        9 => CellValue::Number(nums.iter().sum()),
        _ => CellValue::Error(CellError::Value),
    }
}

fn fn_randarray(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let rows = if !args.is_empty() { eval_one_num(args, 0, ctx, reg).unwrap_or(1.0) as usize } else { 1 };
    let cols = if args.len() > 1 { eval_one_num(args, 1, ctx, reg).unwrap_or(1.0) as usize } else { 1 };
    let min_val = if args.len() > 2 { eval_one_num(args, 2, ctx, reg).unwrap_or(0.0) } else { 0.0 };
    let max_val = if args.len() > 3 { eval_one_num(args, 3, ctx, reg).unwrap_or(1.0) } else { 1.0 };
    let whole = if args.len() > 4 { eval_one_num(args, 4, ctx, reg).unwrap_or(0.0) != 0.0 } else { false };
    let range = max_val - min_val;
    let mut result = Vec::with_capacity(rows);
    for _ in 0..rows {
        let mut row = Vec::with_capacity(cols);
        for _ in 0..cols {
            let v = min_val + rand_f64() * range;
            row.push(CellValue::Number(if whole { v.floor() } else { v }));
        }
        result.push(row);
    }
    CellValue::Array(Box::new(result))
}

fn fn_mmult(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let (a_vals, a_rows, a_cols) = match &args[0] {
        Expr::Range { start, end } => {
            let (sr, sc, er, ec) = match (start.as_ref(), end.as_ref()) {
                (Expr::CellRef { row: r1, col: c1, .. }, Expr::CellRef { row: r2, col: c2, .. }) =>
                    (r1.value(), c1.value(), r2.value(), c2.value()),
                _ => return CellValue::Error(CellError::Value),
            };
            let rows = (er - sr + 1) as usize;
            let cols = (ec - sc + 1) as usize;
            let vals: Vec<f64> = collect_range_values(start, end, ctx).iter().map(|v| v.as_f64().unwrap_or(0.0)).collect();
            (vals, rows, cols)
        }
        _ => return CellValue::Error(CellError::Value),
    };
    let (b_vals, b_rows, b_cols) = match &args[1] {
        Expr::Range { start, end } => {
            let (sr, sc, er, ec) = match (start.as_ref(), end.as_ref()) {
                (Expr::CellRef { row: r1, col: c1, .. }, Expr::CellRef { row: r2, col: c2, .. }) =>
                    (r1.value(), c1.value(), r2.value(), c2.value()),
                _ => return CellValue::Error(CellError::Value),
            };
            let rows = (er - sr + 1) as usize;
            let cols = (ec - sc + 1) as usize;
            let vals: Vec<f64> = collect_range_values(start, end, ctx).iter().map(|v| v.as_f64().unwrap_or(0.0)).collect();
            (vals, rows, cols)
        }
        _ => return CellValue::Error(CellError::Value),
    };
    if a_cols != b_rows { return CellValue::Error(CellError::Value); }
    let mut result = Vec::with_capacity(a_rows);
    for i in 0..a_rows {
        let mut row = Vec::with_capacity(b_cols);
        for j in 0..b_cols {
            let mut sum = 0.0;
            for k in 0..a_cols {
                sum += a_vals[i * a_cols + k] * b_vals[k * b_cols + j];
            }
            row.push(CellValue::Number(sum));
        }
        result.push(row);
    }
    CellValue::Array(Box::new(result))
}

fn collect_pair_arrays_math(args: &[Expr], ctx: &dyn EvalContext) -> Option<(Vec<f64>, Vec<f64>)> {
    let xs = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx)
            .iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>(),
        _ => return None,
    };
    let ys = match &args[1] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx)
            .iter().filter_map(|v| v.as_f64()).collect::<Vec<_>>(),
        _ => return None,
    };
    let n = xs.len().min(ys.len());
    if n == 0 { return None; }
    Some((xs[..n].to_vec(), ys[..n].to_vec()))
}

fn fn_sumx2my2(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    match collect_pair_arrays_math(args, ctx) {
        Some((xs, ys)) => {
            let sum: f64 = xs.iter().zip(ys.iter()).map(|(x, y)| x * x - y * y).sum();
            CellValue::Number(sum)
        }
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_sumx2py2(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    match collect_pair_arrays_math(args, ctx) {
        Some((xs, ys)) => {
            let sum: f64 = xs.iter().zip(ys.iter()).map(|(x, y)| x * x + y * y).sum();
            CellValue::Number(sum)
        }
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_sumxmy2(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    match collect_pair_arrays_math(args, ctx) {
        Some((xs, ys)) => {
            let sum: f64 = xs.iter().zip(ys.iter()).map(|(x, y)| (x - y).powi(2)).sum();
            CellValue::Number(sum)
        }
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_mdeterm(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let (vals, rows, cols) = match &args[0] {
        Expr::Range { start, end } => {
            let (sr, sc, er, ec) = match (start.as_ref(), end.as_ref()) {
                (Expr::CellRef { row: r1, col: c1, .. }, Expr::CellRef { row: r2, col: c2, .. }) =>
                    (r1.value(), c1.value(), r2.value(), c2.value()),
                _ => return CellValue::Error(CellError::Value),
            };
            let rows = (er - sr + 1) as usize;
            let cols = (ec - sc + 1) as usize;
            let vals: Vec<f64> = collect_range_values(start, end, ctx).iter().map(|v| v.as_f64().unwrap_or(0.0)).collect();
            (vals, rows, cols)
        }
        _ => return CellValue::Error(CellError::Value),
    };
    if rows != cols { return CellValue::Error(CellError::Value); }
    match rows {
        1 => CellValue::Number(vals[0]),
        2 => CellValue::Number(vals[0] * vals[3] - vals[1] * vals[2]),
        3 => {
            let det = vals[0] * (vals[4] * vals[8] - vals[5] * vals[7])
                    - vals[1] * (vals[3] * vals[8] - vals[5] * vals[6])
                    + vals[2] * (vals[3] * vals[7] - vals[4] * vals[6]);
            CellValue::Number(det)
        }
        _ => CellValue::Error(CellError::Value),
    }
}

fn fn_minverse(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let (vals, rows, cols) = match &args[0] {
        Expr::Range { start, end } => {
            let (sr, sc, er, ec) = match (start.as_ref(), end.as_ref()) {
                (Expr::CellRef { row: r1, col: c1, .. }, Expr::CellRef { row: r2, col: c2, .. }) =>
                    (r1.value(), c1.value(), r2.value(), c2.value()),
                _ => return CellValue::Error(CellError::Value),
            };
            let rows = (er - sr + 1) as usize;
            let cols = (ec - sc + 1) as usize;
            let vals: Vec<f64> = collect_range_values(start, end, ctx).iter().map(|v| v.as_f64().unwrap_or(0.0)).collect();
            (vals, rows, cols)
        }
        _ => return CellValue::Error(CellError::Value),
    };
    if rows != cols || rows != 2 { return CellValue::Error(CellError::Value); }
    let det = vals[0] * vals[3] - vals[1] * vals[2];
    if det == 0.0 { return CellValue::Error(CellError::Num); }
    let inv = vec![
        vec![CellValue::Number(vals[3] / det), CellValue::Number(-vals[1] / det)],
        vec![CellValue::Number(-vals[2] / det), CellValue::Number(vals[0] / det)],
    ];
    CellValue::Array(Box::new(inv))
}

fn fn_convert(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Error(CellError::Na)
}

fn fn_delta(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n1 = match eval_one_num(args, 0, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let n2 = if args.len() > 1 { eval_one_num(args, 1, ctx, reg).unwrap_or(0.0) } else { 0.0 };
    CellValue::Number(if (n1 - n2).abs() < f64::EPSILON { 1.0 } else { 0.0 })
}

fn fn_gestep(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_one_num(args, 0, ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let step = if args.len() > 1 { eval_one_num(args, 1, ctx, reg).unwrap_or(0.0) } else { 0.0 };
    CellValue::Number(if n >= step { 1.0 } else { 0.0 })
}

fn fn_acot(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    unary_math(args, ctx, reg, |x| std::f64::consts::FRAC_PI_2 - x.atan())
}

fn fn_acoth(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    unary_math(args, ctx, reg, |x| {
        if x.abs() <= 1.0 { return f64::NAN; }
        0.5 * ((x + 1.0) / (x - 1.0)).ln()
    })
}

fn fn_munit(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let dim = match eval_one_num(args, 0, ctx, reg) { Some(n) => n as usize, None => return CellValue::Error(CellError::Value) };
    if dim == 0 || dim > 100 { return CellValue::Error(CellError::Value); }
    let mut rows = Vec::with_capacity(dim);
    for i in 0..dim {
        let mut row = vec![CellValue::Number(0.0); dim];
        row[i] = CellValue::Number(1.0);
        rows.push(row);
    }
    CellValue::Array(Box::new(rows))
}
