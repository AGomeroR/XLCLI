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
    reg.register(FnSpec { name: "TRANSPOSE", description: "Transposes rows and columns", syntax: "TRANSPOSE(array)", min_args: 1, max_args: Some(1), eval: fn_transpose });
    reg.register(FnSpec { name: "XMATCH", description: "Returns position using match mode", syntax: "XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])", min_args: 2, max_args: Some(4), eval: fn_xmatch });
    reg.register(FnSpec { name: "LOOKUP", description: "Looks up value in one range", syntax: "LOOKUP(lookup_value, lookup_vector, [result_vector])", min_args: 2, max_args: Some(3), eval: fn_lookup });
    reg.register(FnSpec { name: "SORT", description: "Sorts a range or array", syntax: "SORT(array, [sort_index], [sort_order], [by_col])", min_args: 1, max_args: Some(4), eval: fn_sort });
    reg.register(FnSpec { name: "UNIQUE", description: "Returns unique values", syntax: "UNIQUE(array, [by_col], [exactly_once])", min_args: 1, max_args: Some(3), eval: fn_unique });
    reg.register(FnSpec { name: "SEQUENCE", description: "Generates a sequence of numbers", syntax: "SEQUENCE(rows, [columns], [start], [step])", min_args: 1, max_args: Some(4), eval: fn_sequence });
    reg.register(FnSpec { name: "FILTER", description: "Filters array by conditions", syntax: "FILTER(array, include, [if_empty])", min_args: 2, max_args: Some(3), eval: fn_filter });
    reg.register(FnSpec { name: "SORTBY", description: "Sorts array by another array", syntax: "SORTBY(array, by_array1, [sort_order1], ...)", min_args: 2, max_args: None, eval: fn_sortby });
    reg.register(FnSpec { name: "TAKE", description: "Returns rows/cols from start or end", syntax: "TAKE(array, rows, [columns])", min_args: 2, max_args: Some(3), eval: fn_take });
    reg.register(FnSpec { name: "DROP", description: "Drops rows/cols from start or end", syntax: "DROP(array, rows, [columns])", min_args: 2, max_args: Some(3), eval: fn_drop });
    reg.register(FnSpec { name: "TOCOL", description: "Returns array as single column", syntax: "TOCOL(array, [ignore], [scan_by_column])", min_args: 1, max_args: Some(3), eval: fn_tocol });
    reg.register(FnSpec { name: "TOROW", description: "Returns array as single row", syntax: "TOROW(array, [ignore], [scan_by_column])", min_args: 1, max_args: Some(3), eval: fn_torow });
    reg.register(FnSpec { name: "WRAPROWS", description: "Wraps row of values into rows", syntax: "WRAPROWS(vector, wrap_count, [pad_with])", min_args: 2, max_args: Some(3), eval: fn_wraprows });
    reg.register(FnSpec { name: "WRAPCOLS", description: "Wraps row of values into columns", syntax: "WRAPCOLS(vector, wrap_count, [pad_with])", min_args: 2, max_args: Some(3), eval: fn_wrapcols });
    reg.register(FnSpec { name: "CHOOSECOLS", description: "Returns specified columns from array", syntax: "CHOOSECOLS(array, col_num1, [col_num2], ...)", min_args: 2, max_args: None, eval: fn_choosecols });
    reg.register(FnSpec { name: "CHOOSEROWS", description: "Returns specified rows from array", syntax: "CHOOSEROWS(array, row_num1, [row_num2], ...)", min_args: 2, max_args: None, eval: fn_chooserows });
    reg.register(FnSpec { name: "AREAS", description: "Returns number of areas in reference", syntax: "AREAS(reference)", min_args: 1, max_args: Some(1), eval: fn_areas });
    reg.register(FnSpec { name: "HYPERLINK", description: "Creates a hyperlink", syntax: "HYPERLINK(link_location, [friendly_name])", min_args: 1, max_args: Some(2), eval: fn_hyperlink });
    reg.register(FnSpec { name: "GETPIVOTDATA", description: "Returns data from a PivotTable", syntax: "GETPIVOTDATA(data_field, pivot_table, ...)", min_args: 2, max_args: None, eval: fn_getpivotdata });
    reg.register(FnSpec { name: "RTD", description: "Returns real-time data", syntax: "RTD(progID, server, topic1, ...)", min_args: 3, max_args: None, eval: fn_rtd });
    reg.register(FnSpec { name: "EXPAND", description: "Expands array to specified dimensions", syntax: "EXPAND(array, rows, [columns], [pad_with])", min_args: 2, max_args: Some(4), eval: fn_expand });
    reg.register(FnSpec { name: "HSTACK", description: "Stacks arrays horizontally", syntax: "HSTACK(array1, [array2], ...)", min_args: 1, max_args: None, eval: fn_hstack });
    reg.register(FnSpec { name: "VSTACK", description: "Stacks arrays vertically", syntax: "VSTACK(array1, [array2], ...)", min_args: 1, max_args: None, eval: fn_vstack });
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
    CellValue::Error(CellError::Ref)
}

fn fn_transpose(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let (sr, sc, er, ec) = match &args[0] {
        Expr::Range { start, end } => match range_dims(start, end) {
            Some(d) => d,
            None => return CellValue::Error(CellError::Value),
        },
        _ => return CellValue::Error(CellError::Value),
    };
    let rows = (er - sr + 1) as usize;
    let cols = (ec - sc + 1) as usize;
    let sheet = ctx.current_sheet();
    let mut result = vec![vec![CellValue::Empty; rows]; cols];
    for r in 0..rows {
        for c in 0..cols {
            result[c][r] = ctx.get_cell_value(CellAddr::new(sheet, sr + r as u32, sc + c as u16));
        }
    }
    CellValue::Array(Box::new(result))
}

fn fn_xmatch(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let lookup = evaluate(&args[0], ctx, reg);
    let values = match &args[1] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let match_mode = if args.len() > 2 {
        evaluate(&args[2], ctx, reg).as_f64().unwrap_or(0.0) as i32
    } else { 0 };

    let lookup_str = lookup.display_value();
    match match_mode {
        0 => {
            for (i, v) in values.iter().enumerate() {
                if v.display_value().eq_ignore_ascii_case(&lookup_str) {
                    return CellValue::Number((i + 1) as f64);
                }
            }
            CellValue::Error(CellError::Na)
        }
        -1 => {
            let lookup_n = lookup.as_f64();
            let mut best: Option<(usize, f64)> = None;
            for (i, v) in values.iter().enumerate() {
                if let (Some(ln), Some(vn)) = (lookup_n, v.as_f64()) {
                    if vn <= ln {
                        if best.is_none() || vn > best.unwrap().1 { best = Some((i, vn)); }
                    }
                }
            }
            match best {
                Some((i, _)) => CellValue::Number((i + 1) as f64),
                None => CellValue::Error(CellError::Na),
            }
        }
        1 => {
            let lookup_n = lookup.as_f64();
            let mut best: Option<(usize, f64)> = None;
            for (i, v) in values.iter().enumerate() {
                if let (Some(ln), Some(vn)) = (lookup_n, v.as_f64()) {
                    if vn >= ln {
                        if best.is_none() || vn < best.unwrap().1 { best = Some((i, vn)); }
                    }
                }
            }
            match best {
                Some((i, _)) => CellValue::Number((i + 1) as f64),
                None => CellValue::Error(CellError::Na),
            }
        }
        _ => CellValue::Error(CellError::Value),
    }
}

fn fn_lookup(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let lookup = evaluate(&args[0], ctx, reg);
    let lookup_vec = match &args[1] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let result_vec = if args.len() > 2 {
        match &args[2] {
            Expr::Range { start, end } => collect_range_values(start, end, ctx),
            _ => return CellValue::Error(CellError::Value),
        }
    } else {
        lookup_vec.clone()
    };

    let lookup_n = lookup.as_f64();
    let mut last_match = None;
    for (i, v) in lookup_vec.iter().enumerate() {
        if let (Some(ln), Some(vn)) = (lookup_n, v.as_f64()) {
            if vn <= ln { last_match = Some(i); }
        } else if v.display_value().eq_ignore_ascii_case(&lookup.display_value()) {
            last_match = Some(i);
        }
    }
    match last_match {
        Some(i) => result_vec.get(i).cloned().unwrap_or(CellValue::Error(CellError::Na)),
        None => CellValue::Error(CellError::Na),
    }
}

fn fn_sort(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let values = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let order = if args.len() > 2 {
        evaluate(&args[2], ctx, reg).as_f64().unwrap_or(1.0) as i32
    } else { 1 };

    let mut sorted = values;
    sorted.sort_by(|a, b| {
        let cmp = match (a.as_f64(), b.as_f64()) {
            (Some(an), Some(bn)) => an.partial_cmp(&bn).unwrap_or(std::cmp::Ordering::Equal),
            _ => a.display_value().cmp(&b.display_value()),
        };
        if order == -1 { cmp.reverse() } else { cmp }
    });
    let rows: Vec<Vec<CellValue>> = sorted.into_iter().map(|v| vec![v]).collect();
    CellValue::Array(Box::new(rows))
}

fn fn_unique(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let values = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let mut seen = std::collections::HashSet::new();
    let mut unique = Vec::new();
    for v in values {
        let key = v.display_value();
        if seen.insert(key) {
            unique.push(v);
        }
    }
    let rows: Vec<Vec<CellValue>> = unique.into_iter().map(|v| vec![v]).collect();
    CellValue::Array(Box::new(rows))
}

fn fn_sequence(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let rows = match evaluate(&args[0], ctx, reg).as_f64() {
        Some(n) if n >= 1.0 => n as usize,
        _ => return CellValue::Error(CellError::Value),
    };
    let cols = if args.len() > 1 {
        match evaluate(&args[1], ctx, reg).as_f64() {
            Some(n) if n >= 1.0 => n as usize,
            _ => return CellValue::Error(CellError::Value),
        }
    } else { 1 };
    let start = if args.len() > 2 { evaluate(&args[2], ctx, reg).as_f64().unwrap_or(1.0) } else { 1.0 };
    let step = if args.len() > 3 { evaluate(&args[3], ctx, reg).as_f64().unwrap_or(1.0) } else { 1.0 };

    let mut result = Vec::with_capacity(rows);
    let mut val = start;
    for _ in 0..rows {
        let mut row = Vec::with_capacity(cols);
        for _ in 0..cols {
            row.push(CellValue::Number(val));
            val += step;
        }
        result.push(row);
    }
    CellValue::Array(Box::new(result))
}

fn fn_filter(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let values = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let include = match &args[1] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let filtered: Vec<CellValue> = values.iter().enumerate()
        .filter(|(i, _)| include.get(*i).and_then(|v| v.as_f64()).unwrap_or(0.0) != 0.0
                        || matches!(include.get(*i), Some(CellValue::Boolean(true))))
        .map(|(_, v)| v.clone())
        .collect();
    if filtered.is_empty() {
        if args.len() > 2 { return evaluate(&args[2], ctx, reg); }
        return CellValue::Error(CellError::Na);
    }
    let rows: Vec<Vec<CellValue>> = filtered.into_iter().map(|v| vec![v]).collect();
    CellValue::Array(Box::new(rows))
}

fn fn_sortby(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let values = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let keys = match &args[1] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let mut indexed: Vec<(usize, &CellValue)> = keys.iter().enumerate().collect();
    indexed.sort_by(|a, b| {
        match (a.1.as_f64(), b.1.as_f64()) {
            (Some(an), Some(bn)) => an.partial_cmp(&bn).unwrap_or(std::cmp::Ordering::Equal),
            _ => a.1.display_value().cmp(&b.1.display_value()),
        }
    });
    let rows: Vec<Vec<CellValue>> = indexed.iter()
        .filter_map(|(i, _)| values.get(*i).map(|v| vec![v.clone()]))
        .collect();
    CellValue::Array(Box::new(rows))
}

fn fn_take(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let values = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let n = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v as i32, None => return CellValue::Error(CellError::Value) };
    let taken: Vec<CellValue> = if n >= 0 {
        values.into_iter().take(n as usize).collect()
    } else {
        let len = values.len();
        let skip = len.saturating_sub((-n) as usize);
        values.into_iter().skip(skip).collect()
    };
    let rows: Vec<Vec<CellValue>> = taken.into_iter().map(|v| vec![v]).collect();
    CellValue::Array(Box::new(rows))
}

fn fn_drop(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let values = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let n = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v as i32, None => return CellValue::Error(CellError::Value) };
    let dropped: Vec<CellValue> = if n >= 0 {
        values.into_iter().skip(n as usize).collect()
    } else {
        let len = values.len();
        let take_count = len.saturating_sub((-n) as usize);
        values.into_iter().take(take_count).collect()
    };
    let rows: Vec<Vec<CellValue>> = dropped.into_iter().map(|v| vec![v]).collect();
    CellValue::Array(Box::new(rows))
}

fn fn_tocol(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let values = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let rows: Vec<Vec<CellValue>> = values.into_iter().map(|v| vec![v]).collect();
    CellValue::Array(Box::new(rows))
}

fn fn_torow(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let values = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    CellValue::Array(Box::new(vec![values]))
}

fn fn_wraprows(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let values = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let wrap = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v as usize, None => return CellValue::Error(CellError::Value) };
    if wrap == 0 { return CellValue::Error(CellError::Value); }
    let pad = if args.len() > 2 { evaluate(&args[2], ctx, reg) } else { CellValue::Error(CellError::Na) };
    let mut rows = Vec::new();
    for chunk in values.chunks(wrap) {
        let mut row: Vec<CellValue> = chunk.to_vec();
        while row.len() < wrap { row.push(pad.clone()); }
        rows.push(row);
    }
    CellValue::Array(Box::new(rows))
}

fn fn_wrapcols(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let values = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let wrap = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v as usize, None => return CellValue::Error(CellError::Value) };
    if wrap == 0 { return CellValue::Error(CellError::Value); }
    let pad = if args.len() > 2 { evaluate(&args[2], ctx, reg) } else { CellValue::Error(CellError::Na) };
    let num_cols = (values.len() + wrap - 1) / wrap;
    let mut rows = vec![vec![pad.clone(); num_cols]; wrap];
    for (i, v) in values.into_iter().enumerate() {
        let col = i / wrap;
        let row = i % wrap;
        rows[row][col] = v;
    }
    CellValue::Array(Box::new(rows))
}

fn fn_choosecols(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    // Simplified: works with flat range, returns selected elements
    let values = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let mut selected = Vec::new();
    for arg in &args[1..] {
        let idx = match evaluate(arg, ctx, reg).as_f64() { Some(v) => (v as usize).saturating_sub(1), None => return CellValue::Error(CellError::Value) };
        selected.push(values.get(idx).cloned().unwrap_or(CellValue::Error(CellError::Na)));
    }
    let rows: Vec<Vec<CellValue>> = selected.into_iter().map(|v| vec![v]).collect();
    CellValue::Array(Box::new(rows))
}

fn fn_areas(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Number(1.0)
}

fn fn_hyperlink(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    if args.len() > 1 {
        evaluate(&args[1], ctx, reg)
    } else {
        let url = evaluate(&args[0], ctx, reg).display_value();
        CellValue::String(url.into())
    }
}

fn fn_getpivotdata(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Error(CellError::Na)
}

fn fn_rtd(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    CellValue::Error(CellError::Na)
}

fn fn_chooserows(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let values = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let mut selected = Vec::new();
    for arg in &args[1..] {
        let idx = match evaluate(arg, ctx, reg).as_f64() { Some(v) => (v as usize).saturating_sub(1), None => return CellValue::Error(CellError::Value) };
        selected.push(values.get(idx).cloned().unwrap_or(CellValue::Error(CellError::Na)));
    }
    let rows: Vec<Vec<CellValue>> = selected.into_iter().map(|v| vec![v]).collect();
    CellValue::Array(Box::new(rows))
}

fn fn_expand(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let values = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let target_rows = match evaluate(&args[1], ctx, reg).as_f64() { Some(v) => v as usize, None => return CellValue::Error(CellError::Value) };
    let target_cols = if args.len() > 2 { evaluate(&args[2], ctx, reg).as_f64().unwrap_or(1.0) as usize } else { 1 };
    let pad = if args.len() > 3 { evaluate(&args[3], ctx, reg) } else { CellValue::Error(CellError::Na) };
    let mut rows = Vec::with_capacity(target_rows);
    for r in 0..target_rows {
        let mut row = Vec::with_capacity(target_cols);
        for c in 0..target_cols {
            let idx = r * target_cols + c;
            row.push(values.get(idx).cloned().unwrap_or_else(|| pad.clone()));
        }
        rows.push(row);
    }
    CellValue::Array(Box::new(rows))
}

fn fn_hstack(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let mut all_cols: Vec<Vec<CellValue>> = Vec::new();
    for arg in args {
        let values = match arg {
            Expr::Range { start, end } => collect_range_values(start, end, ctx),
            _ => continue,
        };
        all_cols.push(values);
    }
    if all_cols.is_empty() { return CellValue::Error(CellError::Value); }
    let max_len = all_cols.iter().map(|c| c.len()).max().unwrap_or(0);
    let mut rows = Vec::with_capacity(max_len);
    for i in 0..max_len {
        let mut row = Vec::new();
        for col in &all_cols {
            row.push(col.get(i).cloned().unwrap_or(CellValue::Error(CellError::Na)));
        }
        rows.push(row);
    }
    CellValue::Array(Box::new(rows))
}

fn fn_vstack(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let mut all_values: Vec<Vec<CellValue>> = Vec::new();
    for arg in args {
        let values = match arg {
            Expr::Range { start, end } => collect_range_values(start, end, ctx),
            _ => continue,
        };
        for v in values {
            all_values.push(vec![v]);
        }
    }
    if all_values.is_empty() { return CellValue::Error(CellError::Value); }
    CellValue::Array(Box::new(all_values))
}
