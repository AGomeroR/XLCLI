use xlcli_core::cell::CellValue;
use xlcli_core::types::CellError;

use crate::ast::Expr;
use crate::eval::{evaluate, EvalContext};
use crate::registry::{FnSpec, FunctionRegistry};

pub fn register(reg: &mut FunctionRegistry) {
    reg.register(FnSpec { name: "DEC2BIN", description: "Converts decimal to binary", syntax: "DEC2BIN(number, [places])", min_args: 1, max_args: Some(2), eval: fn_dec2bin });
    reg.register(FnSpec { name: "DEC2OCT", description: "Converts decimal to octal", syntax: "DEC2OCT(number, [places])", min_args: 1, max_args: Some(2), eval: fn_dec2oct });
    reg.register(FnSpec { name: "DEC2HEX", description: "Converts decimal to hexadecimal", syntax: "DEC2HEX(number, [places])", min_args: 1, max_args: Some(2), eval: fn_dec2hex });
    reg.register(FnSpec { name: "BIN2DEC", description: "Converts binary to decimal", syntax: "BIN2DEC(number)", min_args: 1, max_args: Some(1), eval: fn_bin2dec });
    reg.register(FnSpec { name: "BIN2OCT", description: "Converts binary to octal", syntax: "BIN2OCT(number, [places])", min_args: 1, max_args: Some(2), eval: fn_bin2oct });
    reg.register(FnSpec { name: "BIN2HEX", description: "Converts binary to hexadecimal", syntax: "BIN2HEX(number, [places])", min_args: 1, max_args: Some(2), eval: fn_bin2hex });
    reg.register(FnSpec { name: "OCT2DEC", description: "Converts octal to decimal", syntax: "OCT2DEC(number)", min_args: 1, max_args: Some(1), eval: fn_oct2dec });
    reg.register(FnSpec { name: "OCT2BIN", description: "Converts octal to binary", syntax: "OCT2BIN(number, [places])", min_args: 1, max_args: Some(2), eval: fn_oct2bin });
    reg.register(FnSpec { name: "OCT2HEX", description: "Converts octal to hexadecimal", syntax: "OCT2HEX(number, [places])", min_args: 1, max_args: Some(2), eval: fn_oct2hex });
    reg.register(FnSpec { name: "HEX2DEC", description: "Converts hexadecimal to decimal", syntax: "HEX2DEC(number)", min_args: 1, max_args: Some(1), eval: fn_hex2dec });
    reg.register(FnSpec { name: "HEX2BIN", description: "Converts hexadecimal to binary", syntax: "HEX2BIN(number, [places])", min_args: 1, max_args: Some(2), eval: fn_hex2bin });
    reg.register(FnSpec { name: "HEX2OCT", description: "Converts hexadecimal to octal", syntax: "HEX2OCT(number, [places])", min_args: 1, max_args: Some(2), eval: fn_hex2oct });
    reg.register(FnSpec { name: "BITAND", description: "Returns bitwise AND of two numbers", syntax: "BITAND(number1, number2)", min_args: 2, max_args: Some(2), eval: fn_bitand });
    reg.register(FnSpec { name: "BITOR", description: "Returns bitwise OR of two numbers", syntax: "BITOR(number1, number2)", min_args: 2, max_args: Some(2), eval: fn_bitor });
    reg.register(FnSpec { name: "BITXOR", description: "Returns bitwise XOR of two numbers", syntax: "BITXOR(number1, number2)", min_args: 2, max_args: Some(2), eval: fn_bitxor });
    reg.register(FnSpec { name: "BITLSHIFT", description: "Returns number shifted left by bits", syntax: "BITLSHIFT(number, shift_amount)", min_args: 2, max_args: Some(2), eval: fn_bitlshift });
    reg.register(FnSpec { name: "BITRSHIFT", description: "Returns number shifted right by bits", syntax: "BITRSHIFT(number, shift_amount)", min_args: 2, max_args: Some(2), eval: fn_bitrshift });
    reg.register(FnSpec { name: "COMPLEX", description: "Creates complex number from parts", syntax: "COMPLEX(real_num, i_num, [suffix])", min_args: 2, max_args: Some(3), eval: fn_complex });
    reg.register(FnSpec { name: "IMAGINARY", description: "Returns imaginary part of complex number", syntax: "IMAGINARY(inumber)", min_args: 1, max_args: Some(1), eval: fn_imaginary });
    reg.register(FnSpec { name: "IMREAL", description: "Returns real part of complex number", syntax: "IMREAL(inumber)", min_args: 1, max_args: Some(1), eval: fn_imreal });
    reg.register(FnSpec { name: "IMABS", description: "Returns absolute value of complex number", syntax: "IMABS(inumber)", min_args: 1, max_args: Some(1), eval: fn_imabs });
    reg.register(FnSpec { name: "IMSUM", description: "Returns sum of complex numbers", syntax: "IMSUM(inumber1, [inumber2], ...)", min_args: 1, max_args: None, eval: fn_imsum });
    reg.register(FnSpec { name: "DELTA", description: "Tests if two values are equal", syntax: "DELTA(number1, [number2])", min_args: 1, max_args: Some(2), eval: fn_delta });
    reg.register(FnSpec { name: "GESTEP", description: "Tests if number >= step value", syntax: "GESTEP(number, [step])", min_args: 1, max_args: Some(2), eval: fn_gestep });
    reg.register(FnSpec { name: "ERF", description: "Returns the error function", syntax: "ERF(lower_limit, [upper_limit])", min_args: 1, max_args: Some(2), eval: fn_erf });
    reg.register(FnSpec { name: "ERFC", description: "Returns complementary error function", syntax: "ERFC(x)", min_args: 1, max_args: Some(1), eval: fn_erfc });
    reg.register(FnSpec { name: "CONVERT", description: "Converts between measurement units", syntax: "CONVERT(number, from_unit, to_unit)", min_args: 3, max_args: Some(3), eval: fn_convert });
}

fn eval_f64(expr: &Expr, ctx: &dyn EvalContext, reg: &FunctionRegistry) -> Option<f64> {
    evaluate(expr, ctx, reg).as_f64()
}

fn eval_str(expr: &Expr, ctx: &dyn EvalContext, reg: &FunctionRegistry) -> String {
    evaluate(expr, ctx, reg).display_value()
}

fn pad_result(s: &str, places: Option<usize>) -> String {
    match places {
        Some(p) if s.len() < p => format!("{:0>width$}", s, width = p),
        _ => s.to_string(),
    }
}

fn fn_dec2bin(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_f64(&args[0], ctx, reg) { Some(v) => v as i64, None => return CellValue::Error(CellError::Value) };
    let places = if args.len() > 1 { eval_f64(&args[1], ctx, reg).map(|v| v as usize) } else { None };
    if n < -512 || n > 511 { return CellValue::Error(CellError::Num); }
    let s = if n >= 0 { format!("{:b}", n) } else { format!("{:010b}", n as u64 & 0x3FF) };
    CellValue::String(pad_result(&s, places).into())
}

fn fn_dec2oct(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_f64(&args[0], ctx, reg) { Some(v) => v as i64, None => return CellValue::Error(CellError::Value) };
    let places = if args.len() > 1 { eval_f64(&args[1], ctx, reg).map(|v| v as usize) } else { None };
    let s = if n >= 0 { format!("{:o}", n) } else { format!("{:010o}", n as u64 & 0x3FFFFFFF) };
    CellValue::String(pad_result(&s, places).into())
}

fn fn_dec2hex(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_f64(&args[0], ctx, reg) { Some(v) => v as i64, None => return CellValue::Error(CellError::Value) };
    let places = if args.len() > 1 { eval_f64(&args[1], ctx, reg).map(|v| v as usize) } else { None };
    let s = if n >= 0 { format!("{:X}", n) } else { format!("{:010X}", n as u64 & 0xFFFFFFFFFF) };
    CellValue::String(pad_result(&s, places).into())
}

fn fn_bin2dec(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_str(&args[0], ctx, reg);
    match i64::from_str_radix(&s, 2) {
        Ok(n) => {
            let val = if s.len() == 10 && s.starts_with('1') { n - 1024 } else { n };
            CellValue::Number(val as f64)
        }
        Err(_) => CellValue::Error(CellError::Num),
    }
}

fn fn_bin2oct(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_str(&args[0], ctx, reg);
    let places = if args.len() > 1 { eval_f64(&args[1], ctx, reg).map(|v| v as usize) } else { None };
    match i64::from_str_radix(&s, 2) {
        Ok(n) => {
            let val = if s.len() == 10 && s.starts_with('1') { n - 1024 } else { n };
            let r = if val >= 0 { format!("{:o}", val) } else { format!("{:010o}", val as u64 & 0x3FFFFFFF) };
            CellValue::String(pad_result(&r, places).into())
        }
        Err(_) => CellValue::Error(CellError::Num),
    }
}

fn fn_bin2hex(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_str(&args[0], ctx, reg);
    let places = if args.len() > 1 { eval_f64(&args[1], ctx, reg).map(|v| v as usize) } else { None };
    match i64::from_str_radix(&s, 2) {
        Ok(n) => {
            let val = if s.len() == 10 && s.starts_with('1') { n - 1024 } else { n };
            let r = if val >= 0 { format!("{:X}", val) } else { format!("{:010X}", val as u64 & 0xFFFFFFFFFF) };
            CellValue::String(pad_result(&r, places).into())
        }
        Err(_) => CellValue::Error(CellError::Num),
    }
}

fn fn_oct2dec(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_str(&args[0], ctx, reg);
    match i64::from_str_radix(&s, 8) {
        Ok(n) => CellValue::Number(n as f64),
        Err(_) => CellValue::Error(CellError::Num),
    }
}

fn fn_oct2bin(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_str(&args[0], ctx, reg);
    let places = if args.len() > 1 { eval_f64(&args[1], ctx, reg).map(|v| v as usize) } else { None };
    match i64::from_str_radix(&s, 8) {
        Ok(n) => {
            if n < -512 || n > 511 { return CellValue::Error(CellError::Num); }
            let r = if n >= 0 { format!("{:b}", n) } else { format!("{:010b}", n as u64 & 0x3FF) };
            CellValue::String(pad_result(&r, places).into())
        }
        Err(_) => CellValue::Error(CellError::Num),
    }
}

fn fn_oct2hex(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_str(&args[0], ctx, reg);
    let places = if args.len() > 1 { eval_f64(&args[1], ctx, reg).map(|v| v as usize) } else { None };
    match i64::from_str_radix(&s, 8) {
        Ok(n) => {
            let r = if n >= 0 { format!("{:X}", n) } else { format!("{:010X}", n as u64 & 0xFFFFFFFFFF) };
            CellValue::String(pad_result(&r, places).into())
        }
        Err(_) => CellValue::Error(CellError::Num),
    }
}

fn fn_hex2dec(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_str(&args[0], ctx, reg);
    match i64::from_str_radix(&s, 16) {
        Ok(n) => CellValue::Number(n as f64),
        Err(_) => CellValue::Error(CellError::Num),
    }
}

fn fn_hex2bin(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_str(&args[0], ctx, reg);
    let places = if args.len() > 1 { eval_f64(&args[1], ctx, reg).map(|v| v as usize) } else { None };
    match i64::from_str_radix(&s, 16) {
        Ok(n) => {
            if n < -512 || n > 511 { return CellValue::Error(CellError::Num); }
            let r = if n >= 0 { format!("{:b}", n) } else { format!("{:010b}", n as u64 & 0x3FF) };
            CellValue::String(pad_result(&r, places).into())
        }
        Err(_) => CellValue::Error(CellError::Num),
    }
}

fn fn_hex2oct(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_str(&args[0], ctx, reg);
    let places = if args.len() > 1 { eval_f64(&args[1], ctx, reg).map(|v| v as usize) } else { None };
    match i64::from_str_radix(&s, 16) {
        Ok(n) => {
            let r = if n >= 0 { format!("{:o}", n) } else { format!("{:010o}", n as u64 & 0x3FFFFFFF) };
            CellValue::String(pad_result(&r, places).into())
        }
        Err(_) => CellValue::Error(CellError::Num),
    }
}

fn fn_bitand(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let a = match eval_f64(&args[0], ctx, reg) { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    let b = match eval_f64(&args[1], ctx, reg) { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    CellValue::Number((a & b) as f64)
}

fn fn_bitor(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let a = match eval_f64(&args[0], ctx, reg) { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    let b = match eval_f64(&args[1], ctx, reg) { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    CellValue::Number((a | b) as f64)
}

fn fn_bitxor(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let a = match eval_f64(&args[0], ctx, reg) { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    let b = match eval_f64(&args[1], ctx, reg) { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    CellValue::Number((a ^ b) as f64)
}

fn fn_bitlshift(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_f64(&args[0], ctx, reg) { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    let shift = match eval_f64(&args[1], ctx, reg) { Some(v) => v as u32, None => return CellValue::Error(CellError::Value) };
    CellValue::Number((n << shift) as f64)
}

fn fn_bitrshift(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_f64(&args[0], ctx, reg) { Some(v) => v as u64, None => return CellValue::Error(CellError::Value) };
    let shift = match eval_f64(&args[1], ctx, reg) { Some(v) => v as u32, None => return CellValue::Error(CellError::Value) };
    CellValue::Number((n >> shift) as f64)
}

fn fn_complex(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let real = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let imag = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let suffix = if args.len() > 2 { eval_str(&args[2], ctx, reg) } else { "i".to_string() };
    if imag == 0.0 {
        CellValue::String(format!("{}", real).into())
    } else if real == 0.0 {
        CellValue::String(format!("{}{}", imag, suffix).into())
    } else if imag > 0.0 {
        CellValue::String(format!("{}+{}{}", real, imag, suffix).into())
    } else {
        CellValue::String(format!("{}{}{}", real, imag, suffix).into())
    }
}

fn parse_complex(s: &str) -> Option<(f64, f64)> {
    let s = s.trim();
    if s.ends_with('i') || s.ends_with('j') {
        let s = &s[..s.len() - 1];
        if let Some(pos) = s.rfind('+').or_else(|| {
            let p = s.rfind('-')?;
            if p == 0 { None } else { Some(p) }
        }) {
            let real: f64 = s[..pos].parse().ok()?;
            let imag: f64 = s[pos..].parse().ok()?;
            Some((real, imag))
        } else {
            let imag: f64 = if s.is_empty() { 1.0 } else { s.parse().ok()? };
            Some((0.0, imag))
        }
    } else {
        let real: f64 = s.parse().ok()?;
        Some((real, 0.0))
    }
}

fn fn_imaginary(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_str(&args[0], ctx, reg);
    match parse_complex(&s) {
        Some((_, imag)) => CellValue::Number(imag),
        None => CellValue::Error(CellError::Num),
    }
}

fn fn_imreal(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_str(&args[0], ctx, reg);
    match parse_complex(&s) {
        Some((real, _)) => CellValue::Number(real),
        None => CellValue::Error(CellError::Num),
    }
}

fn fn_imabs(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_str(&args[0], ctx, reg);
    match parse_complex(&s) {
        Some((real, imag)) => CellValue::Number((real * real + imag * imag).sqrt()),
        None => CellValue::Error(CellError::Num),
    }
}

fn fn_imsum(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let mut real_sum = 0.0;
    let mut imag_sum = 0.0;
    for arg in args {
        let s = evaluate(arg, ctx, reg).display_value();
        match parse_complex(&s) {
            Some((r, i)) => { real_sum += r; imag_sum += i; }
            None => return CellValue::Error(CellError::Num),
        }
    }
    if imag_sum == 0.0 {
        CellValue::String(format!("{}", real_sum).into())
    } else if imag_sum > 0.0 {
        CellValue::String(format!("{}+{}i", real_sum, imag_sum).into())
    } else {
        CellValue::String(format!("{}{}i", real_sum, imag_sum).into())
    }
}

fn fn_delta(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let a = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let b = if args.len() > 1 { eval_f64(&args[1], ctx, reg).unwrap_or(0.0) } else { 0.0 };
    CellValue::Number(if (a - b).abs() < f64::EPSILON { 1.0 } else { 0.0 })
}

fn fn_gestep(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let number = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let step = if args.len() > 1 { eval_f64(&args[1], ctx, reg).unwrap_or(0.0) } else { 0.0 };
    CellValue::Number(if number >= step { 1.0 } else { 0.0 })
}

fn fn_erf(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    if args.len() > 1 {
        let upper = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
        CellValue::Number(erf_approx(upper) - erf_approx(x))
    } else {
        CellValue::Number(erf_approx(x))
    }
}

fn fn_erfc(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let x = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    CellValue::Number(1.0 - erf_approx(x))
}

fn erf_approx(x: f64) -> f64 {
    let sign = if x >= 0.0 { 1.0 } else { -1.0 };
    let x = x.abs();
    let t = 1.0 / (1.0 + 0.3275911 * x);
    let poly = t * (0.254829592 + t * (-0.284496736 + t * (1.421413741 + t * (-1.453152027 + t * 1.061405429))));
    sign * (1.0 - poly * (-x * x).exp())
}

fn fn_convert(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let number = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let from = eval_str(&args[1], ctx, reg);
    let to = eval_str(&args[2], ctx, reg);

    let from_si = to_si_factor(&from);
    let to_si = to_si_factor(&to);
    match (from_si, to_si) {
        (Some((ff, fc, fg)), Some((tf, tc, tg))) if fg == tg => {
            CellValue::Number((number * ff + fc - tc) / tf)
        }
        _ => CellValue::Error(CellError::Na),
    }
}

fn to_si_factor(unit: &str) -> Option<(f64, f64, &'static str)> {
    // (multiplier, offset, group)
    match unit {
        "m" => Some((1.0, 0.0, "length")),
        "cm" => Some((0.01, 0.0, "length")),
        "mm" => Some((0.001, 0.0, "length")),
        "km" => Some((1000.0, 0.0, "length")),
        "in" => Some((0.0254, 0.0, "length")),
        "ft" => Some((0.3048, 0.0, "length")),
        "yd" => Some((0.9144, 0.0, "length")),
        "mi" => Some((1609.344, 0.0, "length")),
        "Nmi" => Some((1852.0, 0.0, "length")),
        "kg" => Some((1.0, 0.0, "mass")),
        "g" => Some((0.001, 0.0, "mass")),
        "mg" => Some((1e-6, 0.0, "mass")),
        "lbm" => Some((0.45359237, 0.0, "mass")),
        "ozm" => Some((0.028349523125, 0.0, "mass")),
        "C" => Some((1.0, 0.0, "temp")),
        "F" => Some((5.0/9.0, -32.0 * 5.0/9.0, "temp")),
        "K" => Some((1.0, -273.15, "temp")),
        "l" | "lt" => Some((0.001, 0.0, "volume")),
        "ml" => Some((1e-6, 0.0, "volume")),
        "gal" => Some((0.003785411784, 0.0, "volume")),
        "qt" => Some((0.000946352946, 0.0, "volume")),
        "pt" => Some((0.000473176473, 0.0, "volume")),
        "cup" => Some((0.000236588236, 0.0, "volume")),
        "tsp" => Some((4.92892e-6, 0.0, "volume")),
        "tbs" => Some((1.47868e-5, 0.0, "volume")),
        "s" | "sec" => Some((1.0, 0.0, "time")),
        "mn" | "min" => Some((60.0, 0.0, "time")),
        "hr" => Some((3600.0, 0.0, "time")),
        "day" => Some((86400.0, 0.0, "time")),
        "yr" => Some((31557600.0, 0.0, "time")),
        _ => None,
    }
}
