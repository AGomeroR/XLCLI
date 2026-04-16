use xlcli_core::cell::CellValue;
use xlcli_core::types::CellError;

use crate::ast::Expr;
use crate::eval::{evaluate, EvalContext};
use crate::registry::{FnSpec, FunctionRegistry};

pub fn register(reg: &mut FunctionRegistry) {
    reg.register(FnSpec { name: "LEN", min_args: 1, max_args: Some(1), eval: fn_len });
    reg.register(FnSpec { name: "LEFT", min_args: 1, max_args: Some(2), eval: fn_left });
    reg.register(FnSpec { name: "RIGHT", min_args: 1, max_args: Some(2), eval: fn_right });
    reg.register(FnSpec { name: "MID", min_args: 3, max_args: Some(3), eval: fn_mid });
    reg.register(FnSpec { name: "UPPER", min_args: 1, max_args: Some(1), eval: fn_upper });
    reg.register(FnSpec { name: "LOWER", min_args: 1, max_args: Some(1), eval: fn_lower });
    reg.register(FnSpec { name: "PROPER", min_args: 1, max_args: Some(1), eval: fn_proper });
    reg.register(FnSpec { name: "TRIM", min_args: 1, max_args: Some(1), eval: fn_trim });
    reg.register(FnSpec { name: "CLEAN", min_args: 1, max_args: Some(1), eval: fn_clean });
    reg.register(FnSpec { name: "CONCATENATE", min_args: 1, max_args: None, eval: fn_concatenate });
    reg.register(FnSpec { name: "CONCAT", min_args: 1, max_args: None, eval: fn_concatenate });
    reg.register(FnSpec { name: "TEXTJOIN", min_args: 3, max_args: None, eval: fn_textjoin });
    reg.register(FnSpec { name: "REPT", min_args: 2, max_args: Some(2), eval: fn_rept });
    reg.register(FnSpec { name: "SUBSTITUTE", min_args: 3, max_args: Some(4), eval: fn_substitute });
    reg.register(FnSpec { name: "REPLACE", min_args: 4, max_args: Some(4), eval: fn_replace });
    reg.register(FnSpec { name: "FIND", min_args: 2, max_args: Some(3), eval: fn_find });
    reg.register(FnSpec { name: "SEARCH", min_args: 2, max_args: Some(3), eval: fn_search });
    reg.register(FnSpec { name: "TEXT", min_args: 2, max_args: Some(2), eval: fn_text });
    reg.register(FnSpec { name: "VALUE", min_args: 1, max_args: Some(1), eval: fn_value });
    reg.register(FnSpec { name: "EXACT", min_args: 2, max_args: Some(2), eval: fn_exact });
    reg.register(FnSpec { name: "T", min_args: 1, max_args: Some(1), eval: fn_t });
    reg.register(FnSpec { name: "CHAR", min_args: 1, max_args: Some(1), eval: fn_char });
    reg.register(FnSpec { name: "CODE", min_args: 1, max_args: Some(1), eval: fn_code });
    reg.register(FnSpec { name: "NUMBERVALUE", min_args: 1, max_args: Some(3), eval: fn_numbervalue });
}

fn eval_to_string(expr: &Expr, ctx: &dyn EvalContext, reg: &FunctionRegistry) -> String {
    evaluate(expr, ctx, reg).display_value()
}

fn eval_to_f64(expr: &Expr, ctx: &dyn EvalContext, reg: &FunctionRegistry) -> Option<f64> {
    evaluate(expr, ctx, reg).as_f64()
}

fn fn_len(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_to_string(&args[0], ctx, reg);
    CellValue::Number(s.len() as f64)
}

fn fn_left(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_to_string(&args[0], ctx, reg);
    let n = if args.len() > 1 {
        match eval_to_f64(&args[1], ctx, reg) {
            Some(v) => v as usize,
            None => return CellValue::Error(CellError::Value),
        }
    } else {
        1
    };
    let result: String = s.chars().take(n).collect();
    CellValue::String(result.into())
}

fn fn_right(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_to_string(&args[0], ctx, reg);
    let n = if args.len() > 1 {
        match eval_to_f64(&args[1], ctx, reg) {
            Some(v) => v as usize,
            None => return CellValue::Error(CellError::Value),
        }
    } else {
        1
    };
    let chars: Vec<char> = s.chars().collect();
    let start = chars.len().saturating_sub(n);
    let result: String = chars[start..].iter().collect();
    CellValue::String(result.into())
}

fn fn_mid(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_to_string(&args[0], ctx, reg);
    let start = match eval_to_f64(&args[1], ctx, reg) {
        Some(v) if v >= 1.0 => (v as usize) - 1,
        _ => return CellValue::Error(CellError::Value),
    };
    let num = match eval_to_f64(&args[2], ctx, reg) {
        Some(v) if v >= 0.0 => v as usize,
        _ => return CellValue::Error(CellError::Value),
    };
    let result: String = s.chars().skip(start).take(num).collect();
    CellValue::String(result.into())
}

fn fn_upper(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    CellValue::String(eval_to_string(&args[0], ctx, reg).to_uppercase().into())
}

fn fn_lower(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    CellValue::String(eval_to_string(&args[0], ctx, reg).to_lowercase().into())
}

fn fn_proper(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_to_string(&args[0], ctx, reg);
    let mut result = String::with_capacity(s.len());
    let mut capitalize_next = true;
    for ch in s.chars() {
        if ch.is_alphanumeric() {
            if capitalize_next {
                result.extend(ch.to_uppercase());
                capitalize_next = false;
            } else {
                result.extend(ch.to_lowercase());
            }
        } else {
            result.push(ch);
            capitalize_next = true;
        }
    }
    CellValue::String(result.into())
}

fn fn_trim(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_to_string(&args[0], ctx, reg);
    let trimmed: String = s.split_whitespace().collect::<Vec<_>>().join(" ");
    CellValue::String(trimmed.into())
}

fn fn_clean(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_to_string(&args[0], ctx, reg);
    let cleaned: String = s.chars().filter(|c| !c.is_control()).collect();
    CellValue::String(cleaned.into())
}

fn fn_concatenate(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let mut result = String::new();
    for arg in args {
        result.push_str(&eval_to_string(arg, ctx, reg));
    }
    CellValue::String(result.into())
}

fn fn_textjoin(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let delimiter = eval_to_string(&args[0], ctx, reg);
    let ignore_empty = match evaluate(&args[1], ctx, reg) {
        CellValue::Boolean(b) => b,
        v => v.as_f64().map(|n| n != 0.0).unwrap_or(false),
    };
    let mut parts = Vec::new();
    for arg in &args[2..] {
        let s = eval_to_string(arg, ctx, reg);
        if ignore_empty && s.is_empty() {
            continue;
        }
        parts.push(s);
    }
    CellValue::String(parts.join(&delimiter).into())
}

fn fn_rept(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_to_string(&args[0], ctx, reg);
    let n = match eval_to_f64(&args[1], ctx, reg) {
        Some(v) if v >= 0.0 => v as usize,
        _ => return CellValue::Error(CellError::Value),
    };
    CellValue::String(s.repeat(n).into())
}

fn fn_substitute(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let text = eval_to_string(&args[0], ctx, reg);
    let old = eval_to_string(&args[1], ctx, reg);
    let new = eval_to_string(&args[2], ctx, reg);

    if args.len() > 3 {
        let instance = match eval_to_f64(&args[3], ctx, reg) {
            Some(v) if v >= 1.0 => v as usize,
            _ => return CellValue::Error(CellError::Value),
        };
        let mut result = text.clone();
        let mut count = 0;
        let mut search_start = 0;
        while let Some(pos) = result[search_start..].find(&old) {
            count += 1;
            let abs_pos = search_start + pos;
            if count == instance {
                result = format!("{}{}{}", &result[..abs_pos], new, &result[abs_pos + old.len()..]);
                break;
            }
            search_start = abs_pos + old.len();
        }
        CellValue::String(result.into())
    } else {
        CellValue::String(text.replace(&old, &new).into())
    }
}

fn fn_replace(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let text = eval_to_string(&args[0], ctx, reg);
    let start = match eval_to_f64(&args[1], ctx, reg) {
        Some(v) if v >= 1.0 => (v as usize) - 1,
        _ => return CellValue::Error(CellError::Value),
    };
    let num = match eval_to_f64(&args[2], ctx, reg) {
        Some(v) if v >= 0.0 => v as usize,
        _ => return CellValue::Error(CellError::Value),
    };
    let new_text = eval_to_string(&args[3], ctx, reg);
    let chars: Vec<char> = text.chars().collect();
    let before: String = chars.iter().take(start).collect();
    let after: String = chars.iter().skip(start + num).collect();
    CellValue::String(format!("{}{}{}", before, new_text, after).into())
}

fn fn_find(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let find_text = eval_to_string(&args[0], ctx, reg);
    let within = eval_to_string(&args[1], ctx, reg);
    let start = if args.len() > 2 {
        match eval_to_f64(&args[2], ctx, reg) {
            Some(v) if v >= 1.0 => (v as usize) - 1,
            _ => return CellValue::Error(CellError::Value),
        }
    } else {
        0
    };
    match within[start..].find(&find_text) {
        Some(pos) => CellValue::Number((start + pos + 1) as f64),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_search(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let find_text = eval_to_string(&args[0], ctx, reg).to_lowercase();
    let within = eval_to_string(&args[1], ctx, reg).to_lowercase();
    let start = if args.len() > 2 {
        match eval_to_f64(&args[2], ctx, reg) {
            Some(v) if v >= 1.0 => (v as usize) - 1,
            _ => return CellValue::Error(CellError::Value),
        }
    } else {
        0
    };
    match within[start..].find(&find_text) {
        Some(pos) => CellValue::Number((start + pos + 1) as f64),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_text(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let val = evaluate(&args[0], ctx, reg);
    let _format = eval_to_string(&args[1], ctx, reg);
    // Simplified: just convert to string representation
    CellValue::String(val.display_value().into())
}

fn fn_value(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_to_string(&args[0], ctx, reg);
    match s.trim().parse::<f64>() {
        Ok(n) => CellValue::Number(n),
        Err(_) => CellValue::Error(CellError::Value),
    }
}

fn fn_exact(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let a = eval_to_string(&args[0], ctx, reg);
    let b = eval_to_string(&args[1], ctx, reg);
    CellValue::Boolean(a == b)
}

fn fn_t(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let val = evaluate(&args[0], ctx, reg);
    match val {
        CellValue::String(_) => val,
        _ => CellValue::String("".into()),
    }
}

fn fn_char(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_to_f64(&args[0], ctx, reg) {
        Some(n) if (1.0..=255.0).contains(&n) => {
            let c = n as u8 as char;
            CellValue::String(c.to_string().into())
        }
        _ => CellValue::Error(CellError::Value),
    }
}

fn fn_code(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_to_string(&args[0], ctx, reg);
    match s.chars().next() {
        Some(c) => CellValue::Number(c as u32 as f64),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_numbervalue(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let text = eval_to_string(&args[0], ctx, reg);
    let decimal_sep = if args.len() > 1 {
        eval_to_string(&args[1], ctx, reg)
    } else {
        ".".to_string()
    };
    let group_sep = if args.len() > 2 {
        eval_to_string(&args[2], ctx, reg)
    } else {
        ",".to_string()
    };
    let cleaned = text.replace(&group_sep, "").replace(&decimal_sep, ".");
    match cleaned.trim().parse::<f64>() {
        Ok(n) => CellValue::Number(n),
        Err(_) => CellValue::Error(CellError::Value),
    }
}
