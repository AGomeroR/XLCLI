use xlcli_core::cell::CellValue;
use xlcli_core::types::{CellAddr, CellError};

use crate::ast::Expr;
use crate::eval::{collect_range_values, evaluate, EvalContext};
use crate::registry::{FnSpec, FunctionRegistry};

pub fn register(reg: &mut FunctionRegistry) {
    reg.register(FnSpec { name: "VLOOKUP", description: "Looks up value in first column of range", syntax: "VLOOKUP(lookup_value, table_array, col_index, [range_lookup])", min_args: 3, max_args: Some(4), eval: fn_vlookup });
    reg.register(FnSpec { name: "HLOOKUP", description: "Looks up value in first row of range", syntax: "HLOOKUP(lookup_value, table_array, row_index, [range_lookup])", min_args: 3, max_args: Some(4), eval: fn_hlookup });
    reg.register(FnSpec { name: "INDEX", description: "Returns value at row/col in range", syntax: "INDEX(array, row_num, [col_num])", min_args: 2, max_args: Some(3), eval: fn_index });
    reg.register(FnSpec { name: "MATCH", description: "Returns position of value in range", syntax: "MATCH(lookup_value, lookup_array, [match_type])", min_args: 2, max_args: Some(3), eval: fn_match });
    reg.register(FnSpec { name: "XLOOKUP", description: "Searches range and returns matching item", syntax: "XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])", min_args: 3, max_args: Some(6), eval: fn_xlookup });
    reg.register(FnSpec { name: "CHOOSE", description: "Returns value from list by index", syntax: "CHOOSE(index_num, value1, [value2], ...)", min_args: 2, max_args: None, eval: fn_choose });
    reg.register(FnSpec { name: "ROW", description: "Returns the row number", syntax: "ROW([reference])", min_args: 0, max_args: Some(1), eval: fn_row });
    reg.register(FnSpec { name: "COLUMN", description: "Returns the column number", syntax: "COLUMN([reference])", min_args: 0, max_args: Some(1), eval: fn_column });
    reg.register(FnSpec { name: "ROWS", description: "Returns number of rows in a range", syntax: "ROWS(array)", min_args: 1, max_args: Some(1), eval: fn_rows });
    reg.register(FnSpec { name: "COLUMNS", description: "Returns number of columns in a range", syntax: "COLUMNS(array)", min_args: 1, max_args: Some(1), eval: fn_columns });
    reg.register(FnSpec { name: "ADDRESS", description: "Creates a cell address as text", syntax: "ADDRESS(row_num, col_num, [abs_num], [a1], [sheet])", min_args: 2, max_args: Some(5), eval: fn_address });
    reg.register(FnSpec { name: "INDIRECT", description: "Returns reference from text string", syntax: "INDIRECT(ref_text, [a1])", min_args: 1, max_args: Some(2), eval: fn_indirect });
    reg.register(FnSpec { name: "OFFSET", description: "Returns a range offset from a reference", syntax: "OFFSET(reference, rows, cols, [height], [width])", min_args: 3, max_args: Some(5), eval: fn_offset });
}

fn range_dims(start: &Expr, end: &Expr) -> Option<(u32, u16, u32, u16)> {
    let (sr, sc) = match start {
        Expr::CellRef { row, col, .. } => (row.value(), col.value()),
        _ => return None,
    };
    let (er, ec) = match end {
        Expr::CellRef { row, col, .. } => (row.value(), col.value()),
        _ => return None,
    };
    Some((sr.min(er), sc.min(ec), sr.max(er), sc.max(ec)))
}

fn fn_vlookup(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let lookup_val = evaluate(&args[0], ctx, reg);
    let col_idx = match evaluate(&args[2], ctx, reg).as_f64() {
        Some(n) if n >= 1.0 => (n as usize) - 1,
        _ => return CellValue::Error(CellError::Value),
    };
    let exact = if args.len() > 3 {
        match evaluate(&args[3], ctx, reg) {
            CellValue::Boolean(b) => !b,
            CellValue::Number(n) => n == 0.0,
            _ => false,
        }
    } else {
        false
    };

    let (min_r, min_c, max_r, max_c) = match &args[1] {
        Expr::Range { start, end } => match range_dims(start, end) {
            Some(d) => d,
            None => return CellValue::Error(CellError::Ref),
        },
        _ => return CellValue::Error(CellError::Value),
    };

    let num_cols = (max_c - min_c + 1) as usize;
    if col_idx >= num_cols {
        return CellValue::Error(CellError::Ref);
    }

    let sheet = ctx.current_sheet();
    let lookup_str = lookup_val.display_value();

    for r in min_r..=max_r {
        let cell_val = ctx.get_cell_value(CellAddr::new(sheet, r, min_c));
        if exact {
            if cell_val.display_value().eq_ignore_ascii_case(&lookup_str) {
                return ctx.get_cell_value(CellAddr::new(sheet, r, min_c + col_idx as u16));
            }
        } else {
            if let (Some(lv), Some(cv)) = (lookup_val.as_f64(), cell_val.as_f64()) {
                if cv > lv {
                    if r > min_r {
                        return ctx.get_cell_value(CellAddr::new(sheet, r - 1, min_c + col_idx as u16));
                    } else {
                        return CellValue::Error(CellError::Na);
                    }
                }
            }
        }
    }

    if exact {
        CellValue::Error(CellError::Na)
    } else {
        ctx.get_cell_value(CellAddr::new(sheet, max_r, min_c + col_idx as u16))
    }
}

fn fn_hlookup(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let lookup_val = evaluate(&args[0], ctx, reg);
    let row_idx = match evaluate(&args[2], ctx, reg).as_f64() {
        Some(n) if n >= 1.0 => (n as usize) - 1,
        _ => return CellValue::Error(CellError::Value),
    };
    let exact = if args.len() > 3 {
        match evaluate(&args[3], ctx, reg) {
            CellValue::Boolean(b) => !b,
            CellValue::Number(n) => n == 0.0,
            _ => false,
        }
    } else {
        false
    };

    let (min_r, min_c, max_r, max_c) = match &args[1] {
        Expr::Range { start, end } => match range_dims(start, end) {
            Some(d) => d,
            None => return CellValue::Error(CellError::Ref),
        },
        _ => return CellValue::Error(CellError::Value),
    };

    let num_rows = (max_r - min_r + 1) as usize;
    if row_idx >= num_rows {
        return CellValue::Error(CellError::Ref);
    }

    let sheet = ctx.current_sheet();
    let lookup_str = lookup_val.display_value();

    for c in min_c..=max_c {
        let cell_val = ctx.get_cell_value(CellAddr::new(sheet, min_r, c));
        if exact {
            if cell_val.display_value().eq_ignore_ascii_case(&lookup_str) {
                return ctx.get_cell_value(CellAddr::new(sheet, min_r + row_idx as u32, c));
            }
        }
    }
    CellValue::Error(CellError::Na)
}

fn fn_index(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let row_num = match evaluate(&args[1], ctx, reg).as_f64() {
        Some(n) => n as u32,
        None => return CellValue::Error(CellError::Value),
    };
    let col_num = if args.len() > 2 {
        match evaluate(&args[2], ctx, reg).as_f64() {
            Some(n) => n as u16,
            None => return CellValue::Error(CellError::Value),
        }
    } else {
        1
    };

    let (min_r, min_c, max_r, max_c) = match &args[0] {
        Expr::Range { start, end } => match range_dims(start, end) {
            Some(d) => d,
            None => return CellValue::Error(CellError::Ref),
        },
        _ => return CellValue::Error(CellError::Value),
    };

    if row_num < 1 || col_num < 1 {
        return CellValue::Error(CellError::Value);
    }
    let target_r = min_r + row_num - 1;
    let target_c = min_c + col_num - 1;
    if target_r > max_r || target_c > max_c {
        return CellValue::Error(CellError::Ref);
    }

    ctx.get_cell_value(CellAddr::new(ctx.current_sheet(), target_r, target_c))
}

fn fn_match(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let lookup = evaluate(&args[0], ctx, reg);
    let match_type = if args.len() > 2 {
        match evaluate(&args[2], ctx, reg).as_f64() {
            Some(n) => n as i32,
            None => return CellValue::Error(CellError::Value),
        }
    } else {
        1
    };

    let values = match &args[1] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };

    let lookup_str = lookup.display_value();

    match match_type {
        0 => {
            for (i, v) in values.iter().enumerate() {
                if v.display_value().eq_ignore_ascii_case(&lookup_str) {
                    return CellValue::Number((i + 1) as f64);
                }
            }
            CellValue::Error(CellError::Na)
        }
        1 => {
            let lookup_n = lookup.as_f64();
            let mut last_match = None;
            for (i, v) in values.iter().enumerate() {
                if let (Some(ln), Some(vn)) = (lookup_n, v.as_f64()) {
                    if vn <= ln {
                        last_match = Some(i);
                    }
                }
            }
            match last_match {
                Some(i) => CellValue::Number((i + 1) as f64),
                None => CellValue::Error(CellError::Na),
            }
        }
        -1 => {
            let lookup_n = lookup.as_f64();
            let mut last_match = None;
            for (i, v) in values.iter().enumerate() {
                if let (Some(ln), Some(vn)) = (lookup_n, v.as_f64()) {
                    if vn >= ln {
                        last_match = Some(i);
                    }
                }
            }
            match last_match {
                Some(i) => CellValue::Number((i + 1) as f64),
                None => CellValue::Error(CellError::Na),
            }
        }
        _ => CellValue::Error(CellError::Value),
    }
}

fn fn_xlookup(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let lookup = evaluate(&args[0], ctx, reg);
    let lookup_arr = match &args[1] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let return_arr = match &args[2] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let if_not_found = if args.len() > 3 {
        Some(evaluate(&args[3], ctx, reg))
    } else {
        None
    };
    let _match_mode = if args.len() > 4 {
        evaluate(&args[4], ctx, reg).as_f64().unwrap_or(0.0) as i32
    } else {
        0
    };

    let lookup_str = lookup.display_value();
    for (i, v) in lookup_arr.iter().enumerate() {
        if v.display_value().eq_ignore_ascii_case(&lookup_str) {
            return return_arr.get(i).cloned().unwrap_or(CellValue::Error(CellError::Na));
        }
    }
    if_not_found.unwrap_or(CellValue::Error(CellError::Na))
}

fn fn_choose(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let idx = match evaluate(&args[0], ctx, reg).as_f64() {
        Some(n) if n >= 1.0 => n as usize,
        _ => return CellValue::Error(CellError::Value),
    };
    if idx >= args.len() {
        return CellValue::Error(CellError::Value);
    }
    evaluate(&args[idx], ctx, reg)
}

fn fn_row(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    if args.is_empty() {
        CellValue::Number((ctx.current_cell().row + 1) as f64)
    } else {
        match &args[0] {
            Expr::CellRef { row, .. } => CellValue::Number((row.value() + 1) as f64),
            Expr::Range { start, .. } => match start.as_ref() {
                Expr::CellRef { row, .. } => CellValue::Number((row.value() + 1) as f64),
                _ => CellValue::Error(CellError::Value),
            },
            _ => CellValue::Error(CellError::Value),
        }
    }
}

fn fn_column(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    if args.is_empty() {
        CellValue::Number((ctx.current_cell().col + 1) as f64)
    } else {
        match &args[0] {
            Expr::CellRef { col, .. } => CellValue::Number((col.value() + 1) as f64),
            Expr::Range { start, .. } => match start.as_ref() {
                Expr::CellRef { col, .. } => CellValue::Number((col.value() + 1) as f64),
                _ => CellValue::Error(CellError::Value),
            },
            _ => CellValue::Error(CellError::Value),
        }
    }
}

fn fn_rows(args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    match &args[0] {
        Expr::Range { start, end } => {
            match (start.as_ref(), end.as_ref()) {
                (Expr::CellRef { row: r1, .. }, Expr::CellRef { row: r2, .. }) => {
                    let diff = (r2.value() as i64 - r1.value() as i64).unsigned_abs() + 1;
                    CellValue::Number(diff as f64)
                }
                _ => CellValue::Error(CellError::Value),
            }
        }
        _ => CellValue::Number(1.0),
    }
}

fn fn_columns(args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    match &args[0] {
        Expr::Range { start, end } => {
            match (start.as_ref(), end.as_ref()) {
                (Expr::CellRef { col: c1, .. }, Expr::CellRef { col: c2, .. }) => {
                    let diff = (c2.value() as i64 - c1.value() as i64).unsigned_abs() + 1;
                    CellValue::Number(diff as f64)
                }
                _ => CellValue::Error(CellError::Value),
            }
        }
        _ => CellValue::Number(1.0),
    }
}

fn fn_address(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let row = match evaluate(&args[0], ctx, reg).as_f64() {
        Some(n) if n >= 1.0 => n as u32,
        _ => return CellValue::Error(CellError::Value),
    };
    let col = match evaluate(&args[1], ctx, reg).as_f64() {
        Some(n) if n >= 1.0 => n as u16,
        _ => return CellValue::Error(CellError::Value),
    };
    let abs_type = if args.len() > 2 {
        evaluate(&args[2], ctx, reg).as_f64().unwrap_or(1.0) as u8
    } else {
        1
    };

    let col_name = xlcli_core::types::CellAddr::col_name(col - 1);
    let result = match abs_type {
        1 => format!("${}${}", col_name, row),
        2 => format!("{}${}", col_name, row),
        3 => format!("${}{}", col_name, row),
        4 => format!("{}{}", col_name, row),
        _ => return CellValue::Error(CellError::Value),
    };
    CellValue::String(result.into())
}

fn fn_indirect(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    // INDIRECT requires runtime cell address resolution — stub for now
    CellValue::Error(CellError::Ref)
}

fn fn_offset(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    // OFFSET requires dynamic range creation — stub for now
    CellValue::Error(CellError::Ref)
}
