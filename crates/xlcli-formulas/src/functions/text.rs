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
    reg.register(FnSpec { name: "UNICHAR", description: "Returns Unicode character", syntax: "UNICHAR(number)", min_args: 1, max_args: Some(1), eval: fn_unichar });
    reg.register(FnSpec { name: "UNICODE", description: "Returns Unicode code of first character", syntax: "UNICODE(text)", min_args: 1, max_args: Some(1), eval: fn_unicode });
    reg.register(FnSpec { name: "FIXED", description: "Formats number with fixed decimals", syntax: "FIXED(number, [decimals], [no_commas])", min_args: 1, max_args: Some(3), eval: fn_fixed });
    reg.register(FnSpec { name: "DOLLAR", description: "Formats number as currency text", syntax: "DOLLAR(number, [decimals])", min_args: 1, max_args: Some(2), eval: fn_dollar });
    reg.register(FnSpec { name: "ENCODEURL", description: "URL-encodes a text string", syntax: "ENCODEURL(text)", min_args: 1, max_args: Some(1), eval: fn_encodeurl });
    reg.register(FnSpec { name: "TEXTBEFORE", description: "Returns text before delimiter", syntax: "TEXTBEFORE(text, delimiter, [instance_num])", min_args: 2, max_args: Some(3), eval: fn_textbefore });
    reg.register(FnSpec { name: "TEXTAFTER", description: "Returns text after delimiter", syntax: "TEXTAFTER(text, delimiter, [instance_num])", min_args: 2, max_args: Some(3), eval: fn_textafter });
    reg.register(FnSpec { name: "TEXTSPLIT", description: "Splits text by delimiters", syntax: "TEXTSPLIT(text, col_delimiter, [row_delimiter])", min_args: 2, max_args: Some(3), eval: fn_textsplit });
    reg.register(FnSpec { name: "VALUETOTEXT", description: "Returns text from any value", syntax: "VALUETOTEXT(value, [format])", min_args: 1, max_args: Some(2), eval: fn_valuetotext });
    reg.register(FnSpec { name: "ARRAYTOTEXT", description: "Returns text from array", syntax: "ARRAYTOTEXT(array, [format])", min_args: 1, max_args: Some(2), eval: fn_arraytotext });
    reg.register(FnSpec { name: "ASC", description: "Converts full-width to half-width", syntax: "ASC(text)", min_args: 1, max_args: Some(1), eval: fn_asc });
    reg.register(FnSpec { name: "LEFTB", description: "Returns leftmost bytes of text", syntax: "LEFTB(text, [num_bytes])", min_args: 1, max_args: Some(2), eval: fn_left });
    reg.register(FnSpec { name: "RIGHTB", description: "Returns rightmost bytes of text", syntax: "RIGHTB(text, [num_bytes])", min_args: 1, max_args: Some(2), eval: fn_right });
    reg.register(FnSpec { name: "MIDB", description: "Returns bytes from middle of text", syntax: "MIDB(text, start_num, num_bytes)", min_args: 3, max_args: Some(3), eval: fn_mid });
    reg.register(FnSpec { name: "LENB", description: "Returns number of bytes in text", syntax: "LENB(text)", min_args: 1, max_args: Some(1), eval: fn_len });
    reg.register(FnSpec { name: "BAHTTEXT", description: "Converts number to Thai Baht text", syntax: "BAHTTEXT(number)", min_args: 1, max_args: Some(1), eval: fn_bahttext });
    reg.register(FnSpec { name: "PHONETIC", description: "Returns phonetic (furigana) text", syntax: "PHONETIC(reference)", min_args: 1, max_args: Some(1), eval: fn_phonetic });
    reg.register(FnSpec { name: "DBCS", description: "Converts half-width to full-width characters", syntax: "DBCS(text)", min_args: 1, max_args: Some(1), eval: fn_dbcs });
    reg.register(FnSpec { name: "FORMULATEXT", description: "Returns formula as text", syntax: "FORMULATEXT(reference)", min_args: 1, max_args: Some(1), eval: fn_formulatext });
    reg.register(FnSpec { name: "ISOMITTED", description: "Tests if argument is omitted", syntax: "ISOMITTED(argument)", min_args: 1, max_args: Some(1), eval: fn_isomitted });
    reg.register(FnSpec { name: "REGEXTEST", description: "Tests if text matches regex pattern", syntax: "REGEXTEST(text, pattern)", min_args: 2, max_args: Some(2), eval: fn_regextest });
    reg.register(FnSpec { name: "REGEXEXTRACT", description: "Extracts regex match from text", syntax: "REGEXEXTRACT(text, pattern)", min_args: 2, max_args: Some(2), eval: fn_regexextract });
    reg.register(FnSpec { name: "REGEXREPLACE", description: "Replaces regex match in text", syntax: "REGEXREPLACE(text, pattern, replacement)", min_args: 3, max_args: Some(3), eval: fn_regexreplace });
    reg.register(FnSpec { name: "FINDB", description: "Finds text position (byte-based)", syntax: "FINDB(find_text, within_text, [start_num])", min_args: 2, max_args: Some(3), eval: fn_find });
    reg.register(FnSpec { name: "SEARCHB", description: "Searches text position (byte-based)", syntax: "SEARCHB(find_text, within_text, [start_num])", min_args: 2, max_args: Some(3), eval: fn_search });
    reg.register(FnSpec { name: "REPLACEB", description: "Replaces text by position (byte-based)", syntax: "REPLACEB(old_text, start_num, num_bytes, new_text)", min_args: 4, max_args: Some(4), eval: fn_replace });
    reg.register(FnSpec { name: "JIS", description: "Converts half-width to full-width characters", syntax: "JIS(text)", min_args: 1, max_args: Some(1), eval: fn_dbcs });
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

fn fn_unichar(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_to_f64(&args[0], ctx, reg) {
        Some(n) if n >= 1.0 => {
            match char::from_u32(n as u32) {
                Some(c) => CellValue::String(c.to_string().into()),
                None => CellValue::Error(CellError::Value),
            }
        }
        _ => CellValue::Error(CellError::Value),
    }
}

fn fn_unicode(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_to_string(&args[0], ctx, reg);
    match s.chars().next() {
        Some(c) => CellValue::Number(c as u32 as f64),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_fixed(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_to_f64(&args[0], ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let decimals = if args.len() > 1 { eval_to_f64(&args[1], ctx, reg).unwrap_or(2.0) as usize } else { 2 };
    let no_commas = if args.len() > 2 {
        match evaluate(&args[2], ctx, reg) {
            CellValue::Boolean(b) => b,
            CellValue::Number(v) => v != 0.0,
            _ => false,
        }
    } else { false };
    let formatted = format!("{:.prec$}", n, prec = decimals);
    if no_commas {
        CellValue::String(formatted.into())
    } else {
        let parts: Vec<&str> = formatted.split('.').collect();
        let int_part = parts[0];
        let negative = int_part.starts_with('-');
        let digits: String = int_part.chars().filter(|c| c.is_ascii_digit()).collect();
        let mut with_commas = String::new();
        for (i, c) in digits.chars().rev().enumerate() {
            if i > 0 && i % 3 == 0 { with_commas.insert(0, ','); }
            with_commas.insert(0, c);
        }
        if negative { with_commas.insert(0, '-'); }
        if parts.len() > 1 {
            with_commas.push('.');
            with_commas.push_str(parts[1]);
        }
        CellValue::String(with_commas.into())
    }
}

fn fn_dollar(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let n = match eval_to_f64(&args[0], ctx, reg) { Some(n) => n, None => return CellValue::Error(CellError::Value) };
    let decimals = if args.len() > 1 { eval_to_f64(&args[1], ctx, reg).unwrap_or(2.0) as usize } else { 2 };
    let formatted = format!("{:.prec$}", n.abs(), prec = decimals);
    let parts: Vec<&str> = formatted.split('.').collect();
    let digits: String = parts[0].chars().filter(|c| c.is_ascii_digit()).collect();
    let mut with_commas = String::new();
    for (i, c) in digits.chars().rev().enumerate() {
        if i > 0 && i % 3 == 0 { with_commas.insert(0, ','); }
        with_commas.insert(0, c);
    }
    let mut result = String::from("$");
    if n < 0.0 { result.insert(0, '-'); }
    result.push_str(&with_commas);
    if parts.len() > 1 {
        result.push('.');
        result.push_str(parts[1]);
    }
    CellValue::String(result.into())
}

fn fn_encodeurl(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_to_string(&args[0], ctx, reg);
    let encoded: String = s.chars().map(|c| {
        if c.is_ascii_alphanumeric() || "-._~".contains(c) {
            c.to_string()
        } else {
            format!("%{:02X}", c as u32)
        }
    }).collect();
    CellValue::String(encoded.into())
}

fn fn_textbefore(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let text = eval_to_string(&args[0], ctx, reg);
    let delim = eval_to_string(&args[1], ctx, reg);
    let instance = if args.len() > 2 { eval_to_f64(&args[2], ctx, reg).unwrap_or(1.0) as usize } else { 1 };
    let mut start = 0;
    for i in 0..instance {
        match text[start..].find(&delim) {
            Some(pos) => {
                if i + 1 == instance { return CellValue::String(text[..start + pos].to_string().into()); }
                start += pos + delim.len();
            }
            None => return CellValue::Error(CellError::Na),
        }
    }
    CellValue::Error(CellError::Na)
}

fn fn_textafter(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let text = eval_to_string(&args[0], ctx, reg);
    let delim = eval_to_string(&args[1], ctx, reg);
    let instance = if args.len() > 2 { eval_to_f64(&args[2], ctx, reg).unwrap_or(1.0) as usize } else { 1 };
    let mut start = 0;
    for i in 0..instance {
        match text[start..].find(&delim) {
            Some(pos) => {
                if i + 1 == instance { return CellValue::String(text[start + pos + delim.len()..].to_string().into()); }
                start += pos + delim.len();
            }
            None => return CellValue::Error(CellError::Na),
        }
    }
    CellValue::Error(CellError::Na)
}

fn fn_textsplit(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let text = eval_to_string(&args[0], ctx, reg);
    let col_delim = eval_to_string(&args[1], ctx, reg);
    let parts: Vec<CellValue> = text.split(&col_delim).map(|s| CellValue::String(s.to_string().into())).collect();
    let rows: Vec<Vec<CellValue>> = parts.into_iter().map(|v| vec![v]).collect();
    CellValue::Array(Box::new(rows))
}

fn fn_valuetotext(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let val = evaluate(&args[0], ctx, reg);
    CellValue::String(val.display_value().into())
}

fn fn_arraytotext(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let val = evaluate(&args[0], ctx, reg);
    CellValue::String(val.display_value().into())
}

fn fn_bahttext(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_to_string(&args[0], ctx, reg);
    CellValue::String(s.into())
}

fn fn_phonetic(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_to_string(&args[0], ctx, reg);
    CellValue::String(s.into())
}

fn fn_dbcs(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_to_string(&args[0], ctx, reg);
    let converted: String = s.chars().map(|c| {
        let code = c as u32;
        if (0x21..=0x7E).contains(&code) {
            char::from_u32(code - 0x21 + 0xFF01).unwrap_or(c)
        } else if c == ' ' {
            '\u{3000}'
        } else {
            c
        }
    }).collect();
    CellValue::String(converted.into())
}

fn fn_formulatext(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Error(CellError::Na)
}

fn fn_isomitted(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Boolean(false)
}

fn fn_regextest(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Boolean(false)
}

fn fn_regexextract(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Error(CellError::Na)
}

fn fn_regexreplace(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_to_string(&args[0], ctx, reg);
    CellValue::String(s.into())
}

fn fn_asc(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = eval_to_string(&args[0], ctx, reg);
    let converted: String = s.chars().map(|c| {
        let code = c as u32;
        if (0xFF01..=0xFF5E).contains(&code) {
            char::from_u32(code - 0xFF01 + 0x21).unwrap_or(c)
        } else if code == 0x3000 {
            ' '
        } else {
            c
        }
    }).collect();
    CellValue::String(converted.into())
}
