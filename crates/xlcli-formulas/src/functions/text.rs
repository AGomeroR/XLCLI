use xlcli_core::cell::CellValue;
use xlcli_core::types::CellError;

use crate::ast::Expr;
use crate::eval::{evaluate, EvalContext};
use crate::registry::{FnSpec, FunctionRegistry};

pub fn register(reg: &mut FunctionRegistry) {
    reg.register(FnSpec { name: "LEN", description: "Returns the number of characters", syntax: "LEN(text)", min_args: 1, max_args: Some(1), eval: fn_len });
    reg.register(FnSpec { name: "LEFT", description: "Returns leftmost characters", syntax: "LEFT(text, [num_chars])", min_args: 1, max_args: Some(2), eval: fn_left });
    reg.register(FnSpec { name: "RIGHT", description: "Returns rightmost characters", syntax: "RIGHT(text, [num_chars])", min_args: 1, max_args: Some(2), eval: fn_right });
    reg.register(FnSpec { name: "MID", description: "Returns characters from the middle", syntax: "MID(text, start_num, num_chars)", min_args: 3, max_args: Some(3), eval: fn_mid });
    reg.register(FnSpec { name: "UPPER", description: "Converts text to uppercase", syntax: "UPPER(text)", min_args: 1, max_args: Some(1), eval: fn_upper });
    reg.register(FnSpec { name: "LOWER", description: "Converts text to lowercase", syntax: "LOWER(text)", min_args: 1, max_args: Some(1), eval: fn_lower });
    reg.register(FnSpec { name: "PROPER", description: "Capitalizes each word", syntax: "PROPER(text)", min_args: 1, max_args: Some(1), eval: fn_proper });
    reg.register(FnSpec { name: "TRIM", description: "Removes extra spaces", syntax: "TRIM(text)", min_args: 1, max_args: Some(1), eval: fn_trim });
    reg.register(FnSpec { name: "CLEAN", description: "Removes non-printable characters", syntax: "CLEAN(text)", min_args: 1, max_args: Some(1), eval: fn_clean });
    reg.register(FnSpec { name: "CONCATENATE", description: "Joins text strings together", syntax: "CONCATENATE(text1, [text2], ...)", min_args: 1, max_args: None, eval: fn_concatenate });
    reg.register(FnSpec { name: "CONCAT", description: "Joins text strings together", syntax: "CONCAT(text1, [text2], ...)", min_args: 1, max_args: None, eval: fn_concatenate });
    reg.register(FnSpec { name: "TEXTJOIN", description: "Joins text with a delimiter", syntax: "TEXTJOIN(delimiter, ignore_empty, text1, ...)", min_args: 3, max_args: None, eval: fn_textjoin });
    reg.register(FnSpec { name: "REPT", description: "Repeats text a given number of times", syntax: "REPT(text, number_times)", min_args: 2, max_args: Some(2), eval: fn_rept });
    reg.register(FnSpec { name: "SUBSTITUTE", description: "Replaces occurrences of text", syntax: "SUBSTITUTE(text, old_text, new_text, [instance])", min_args: 3, max_args: Some(4), eval: fn_substitute });
    reg.register(FnSpec { name: "REPLACE", description: "Replaces characters by position", syntax: "REPLACE(old_text, start_num, num_chars, new_text)", min_args: 4, max_args: Some(4), eval: fn_replace });
    reg.register(FnSpec { name: "FIND", description: "Finds text within text (case-sensitive)", syntax: "FIND(find_text, within_text, [start_num])", min_args: 2, max_args: Some(3), eval: fn_find });
    reg.register(FnSpec { name: "SEARCH", description: "Finds text within text (case-insensitive)", syntax: "SEARCH(find_text, within_text, [start_num])", min_args: 2, max_args: Some(3), eval: fn_search });
    reg.register(FnSpec { name: "TEXT", description: "Formats a number as text", syntax: "TEXT(value, format_text)", min_args: 2, max_args: Some(2), eval: fn_text });
    reg.register(FnSpec { name: "VALUE", description: "Converts text to a number", syntax: "VALUE(text)", min_args: 1, max_args: Some(1), eval: fn_value });
    reg.register(FnSpec { name: "EXACT", description: "Checks if two strings are identical", syntax: "EXACT(text1, text2)", min_args: 2, max_args: Some(2), eval: fn_exact });
    reg.register(FnSpec { name: "T", description: "Returns text or empty string", syntax: "T(value)", min_args: 1, max_args: Some(1), eval: fn_t });
    reg.register(FnSpec { name: "CHAR", description: "Returns character from code number", syntax: "CHAR(number)", min_args: 1, max_args: Some(1), eval: fn_char });
    reg.register(FnSpec { name: "CODE", description: "Returns numeric code for a character", syntax: "CODE(text)", min_args: 1, max_args: Some(1), eval: fn_code });
    reg.register(FnSpec { name: "NUMBERVALUE", description: "Converts text to number with locale", syntax: "NUMBERVALUE(text, [decimal_sep], [group_sep])", min_args: 1, max_args: Some(3), eval: fn_numbervalue });
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
