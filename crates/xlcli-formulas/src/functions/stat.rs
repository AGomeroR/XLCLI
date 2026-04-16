use xlcli_core::cell::CellValue;
use xlcli_core::types::CellError;

use crate::ast::Expr;
use crate::eval::{collect_range_values, evaluate, EvalContext};
use crate::registry::{FnSpec, FunctionRegistry};

pub fn register(reg: &mut FunctionRegistry) {
    reg.register(FnSpec { name: "COUNT", min_args: 1, max_args: None, eval: fn_count });
    reg.register(FnSpec { name: "COUNTA", min_args: 1, max_args: None, eval: fn_counta });
    reg.register(FnSpec { name: "COUNTBLANK", min_args: 1, max_args: Some(1), eval: fn_countblank });
    reg.register(FnSpec { name: "COUNTIF", min_args: 2, max_args: Some(2), eval: fn_countif });
    reg.register(FnSpec { name: "SUMIF", min_args: 2, max_args: Some(3), eval: fn_sumif });
    reg.register(FnSpec { name: "AVERAGEIF", min_args: 2, max_args: Some(3), eval: fn_averageif });
    reg.register(FnSpec { name: "MEDIAN", min_args: 1, max_args: None, eval: fn_median });
    reg.register(FnSpec { name: "MODE", min_args: 1, max_args: None, eval: fn_mode });
    reg.register(FnSpec { name: "STDEV", min_args: 1, max_args: None, eval: fn_stdev });
    reg.register(FnSpec { name: "STDEVP", min_args: 1, max_args: None, eval: fn_stdevp });
    reg.register(FnSpec { name: "VAR", min_args: 1, max_args: None, eval: fn_var });
    reg.register(FnSpec { name: "VARP", min_args: 1, max_args: None, eval: fn_varp });
    reg.register(FnSpec { name: "LARGE", min_args: 2, max_args: Some(2), eval: fn_large });
    reg.register(FnSpec { name: "SMALL", min_args: 2, max_args: Some(2), eval: fn_small });
    reg.register(FnSpec { name: "RANK", min_args: 2, max_args: Some(3), eval: fn_rank });
    reg.register(FnSpec { name: "PERCENTILE", min_args: 2, max_args: Some(2), eval: fn_percentile });
    reg.register(FnSpec { name: "CORREL", min_args: 2, max_args: Some(2), eval: fn_correl });
    reg.register(FnSpec { name: "MINIFS", min_args: 3, max_args: None, eval: fn_minifs });
    reg.register(FnSpec { name: "MAXIFS", min_args: 3, max_args: None, eval: fn_maxifs });
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
                if let Some(n) = evaluate(arg, ctx, reg).as_f64() {
                    nums.push(n);
                }
            }
        }
    }
    nums
}

fn collect_all_values(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> Vec<CellValue> {
    let mut vals = Vec::new();
    for arg in args {
        match arg {
            Expr::Range { start, end } => {
                vals.extend(collect_range_values(start, end, ctx));
            }
            _ => {
                vals.push(evaluate(arg, ctx, reg));
            }
        }
    }
    vals
}

fn fn_count(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let vals = collect_all_values(args, ctx, reg);
    let count = vals.iter().filter(|v| v.as_f64().is_some()).count();
    CellValue::Number(count as f64)
}

fn fn_counta(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let vals = collect_all_values(args, ctx, reg);
    let count = vals.iter().filter(|v| !matches!(v, CellValue::Empty)).count();
    CellValue::Number(count as f64)
}

fn fn_countblank(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let vals = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let count = vals.iter().filter(|v| matches!(v, CellValue::Empty)).count();
    CellValue::Number(count as f64)
}

fn matches_criteria(val: &CellValue, criteria: &str) -> bool {
    if let Some(rest) = criteria.strip_prefix(">=") {
        if let (Some(vn), Ok(cn)) = (val.as_f64(), rest.parse::<f64>()) {
            return vn >= cn;
        }
    } else if let Some(rest) = criteria.strip_prefix("<=") {
        if let (Some(vn), Ok(cn)) = (val.as_f64(), rest.parse::<f64>()) {
            return vn <= cn;
        }
    } else if let Some(rest) = criteria.strip_prefix("<>") {
        return val.display_value() != rest;
    } else if let Some(rest) = criteria.strip_prefix('>') {
        if let (Some(vn), Ok(cn)) = (val.as_f64(), rest.parse::<f64>()) {
            return vn > cn;
        }
    } else if let Some(rest) = criteria.strip_prefix('<') {
        if let (Some(vn), Ok(cn)) = (val.as_f64(), rest.parse::<f64>()) {
            return vn < cn;
        }
    } else if let Some(rest) = criteria.strip_prefix('=') {
        return val.display_value().eq_ignore_ascii_case(rest);
    }

    if let Ok(cn) = criteria.parse::<f64>() {
        if let Some(vn) = val.as_f64() {
            return (vn - cn).abs() < f64::EPSILON;
        }
    }
    val.display_value().eq_ignore_ascii_case(criteria)
}

fn fn_countif(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let vals = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let criteria = evaluate(&args[1], ctx, reg).display_value();
    let count = vals.iter().filter(|v| matches_criteria(v, &criteria)).count();
    CellValue::Number(count as f64)
}

fn fn_sumif(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let range_vals = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let criteria = evaluate(&args[1], ctx, reg).display_value();
    let sum_vals = if args.len() > 2 {
        match &args[2] {
            Expr::Range { start, end } => collect_range_values(start, end, ctx),
            _ => return CellValue::Error(CellError::Value),
        }
    } else {
        range_vals.clone()
    };

    let mut sum = 0.0;
    for (i, v) in range_vals.iter().enumerate() {
        if matches_criteria(v, &criteria) {
            if let Some(n) = sum_vals.get(i).and_then(|v| v.as_f64()) {
                sum += n;
            }
        }
    }
    CellValue::Number(sum)
}

fn fn_averageif(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let range_vals = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };
    let criteria = evaluate(&args[1], ctx, reg).display_value();
    let sum_vals = if args.len() > 2 {
        match &args[2] {
            Expr::Range { start, end } => collect_range_values(start, end, ctx),
            _ => return CellValue::Error(CellError::Value),
        }
    } else {
        range_vals.clone()
    };

    let mut sum = 0.0;
    let mut count = 0;
    for (i, v) in range_vals.iter().enumerate() {
        if matches_criteria(v, &criteria) {
            if let Some(n) = sum_vals.get(i).and_then(|v| v.as_f64()) {
                sum += n;
                count += 1;
            }
        }
    }
    if count == 0 {
        CellValue::Error(CellError::Div0)
    } else {
        CellValue::Number(sum / count as f64)
    }
}

fn fn_median(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let mut nums = collect_numbers(args, ctx, reg);
    if nums.is_empty() {
        return CellValue::Error(CellError::Num);
    }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap());
    let mid = nums.len() / 2;
    if nums.len() % 2 == 0 {
        CellValue::Number((nums[mid - 1] + nums[mid]) / 2.0)
    } else {
        CellValue::Number(nums[mid])
    }
}

fn fn_mode(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    if nums.is_empty() {
        return CellValue::Error(CellError::Na);
    }
    let mut counts = std::collections::HashMap::new();
    for n in &nums {
        let key = n.to_bits();
        *counts.entry(key).or_insert(0) += 1;
    }
    let max_count = *counts.values().max().unwrap();
    if max_count == 1 {
        return CellValue::Error(CellError::Na);
    }
    for n in &nums {
        if counts[&n.to_bits()] == max_count {
            return CellValue::Number(*n);
        }
    }
    CellValue::Error(CellError::Na)
}

fn variance(nums: &[f64], sample: bool) -> Option<f64> {
    let n = nums.len();
    if n == 0 || (sample && n == 1) {
        return None;
    }
    let mean = nums.iter().sum::<f64>() / n as f64;
    let sum_sq: f64 = nums.iter().map(|x| (x - mean).powi(2)).sum();
    let divisor = if sample { (n - 1) as f64 } else { n as f64 };
    Some(sum_sq / divisor)
}

fn fn_stdev(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    match variance(&nums, true) {
        Some(v) => CellValue::Number(v.sqrt()),
        None => CellValue::Error(CellError::Div0),
    }
}

fn fn_stdevp(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    match variance(&nums, false) {
        Some(v) => CellValue::Number(v.sqrt()),
        None => CellValue::Error(CellError::Div0),
    }
}

fn fn_var(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    match variance(&nums, true) {
        Some(v) => CellValue::Number(v),
        None => CellValue::Error(CellError::Div0),
    }
}

fn fn_varp(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nums = collect_numbers(args, ctx, reg);
    match variance(&nums, false) {
        Some(v) => CellValue::Number(v),
        None => CellValue::Error(CellError::Div0),
    }
}

fn fn_large(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let mut nums = match &args[0] {
        Expr::Range { start, end } => {
            collect_range_values(start, end, ctx)
                .iter()
                .filter_map(|v| v.as_f64())
                .collect::<Vec<_>>()
        }
        _ => return CellValue::Error(CellError::Value),
    };
    let k = match evaluate(&args[1], ctx, reg).as_f64() {
        Some(n) if n >= 1.0 => n as usize,
        _ => return CellValue::Error(CellError::Value),
    };
    if k > nums.len() {
        return CellValue::Error(CellError::Num);
    }
    nums.sort_by(|a, b| b.partial_cmp(a).unwrap());
    CellValue::Number(nums[k - 1])
}

fn fn_small(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let mut nums = match &args[0] {
        Expr::Range { start, end } => {
            collect_range_values(start, end, ctx)
                .iter()
                .filter_map(|v| v.as_f64())
                .collect::<Vec<_>>()
        }
        _ => return CellValue::Error(CellError::Value),
    };
    let k = match evaluate(&args[1], ctx, reg).as_f64() {
        Some(n) if n >= 1.0 => n as usize,
        _ => return CellValue::Error(CellError::Value),
    };
    if k > nums.len() {
        return CellValue::Error(CellError::Num);
    }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap());
    CellValue::Number(nums[k - 1])
}

fn fn_rank(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let number = match evaluate(&args[0], ctx, reg).as_f64() {
        Some(n) => n,
        None => return CellValue::Error(CellError::Value),
    };
    let nums = match &args[1] {
        Expr::Range { start, end } => {
            collect_range_values(start, end, ctx)
                .iter()
                .filter_map(|v| v.as_f64())
                .collect::<Vec<_>>()
        }
        _ => return CellValue::Error(CellError::Value),
    };
    let descending = if args.len() > 2 {
        evaluate(&args[2], ctx, reg).as_f64().unwrap_or(0.0) == 0.0
    } else {
        true
    };

    let rank = if descending {
        nums.iter().filter(|&&n| n > number).count() + 1
    } else {
        nums.iter().filter(|&&n| n < number).count() + 1
    };
    CellValue::Number(rank as f64)
}

fn fn_percentile(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let mut nums = match &args[0] {
        Expr::Range { start, end } => {
            collect_range_values(start, end, ctx)
                .iter()
                .filter_map(|v| v.as_f64())
                .collect::<Vec<_>>()
        }
        _ => return CellValue::Error(CellError::Value),
    };
    let k = match evaluate(&args[1], ctx, reg).as_f64() {
        Some(n) if (0.0..=1.0).contains(&n) => n,
        _ => return CellValue::Error(CellError::Num),
    };
    if nums.is_empty() {
        return CellValue::Error(CellError::Num);
    }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap());
    let n = nums.len() - 1;
    let idx = k * n as f64;
    let lower = idx.floor() as usize;
    let upper = idx.ceil() as usize;
    let frac = idx - lower as f64;
    let result = nums[lower] * (1.0 - frac) + nums[upper] * frac;
    CellValue::Number(result)
}

fn fn_correl(args: &[Expr], ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let xs = match &args[0] {
        Expr::Range { start, end } => {
            collect_range_values(start, end, ctx)
                .iter()
                .filter_map(|v| v.as_f64())
                .collect::<Vec<_>>()
        }
        _ => return CellValue::Error(CellError::Value),
    };
    let ys = match &args[1] {
        Expr::Range { start, end } => {
            collect_range_values(start, end, ctx)
                .iter()
                .filter_map(|v| v.as_f64())
                .collect::<Vec<_>>()
        }
        _ => return CellValue::Error(CellError::Value),
    };
    let n = xs.len().min(ys.len());
    if n < 2 {
        return CellValue::Error(CellError::Na);
    }
    let mean_x = xs[..n].iter().sum::<f64>() / n as f64;
    let mean_y = ys[..n].iter().sum::<f64>() / n as f64;
    let mut sum_xy = 0.0;
    let mut sum_x2 = 0.0;
    let mut sum_y2 = 0.0;
    for i in 0..n {
        let dx = xs[i] - mean_x;
        let dy = ys[i] - mean_y;
        sum_xy += dx * dy;
        sum_x2 += dx * dx;
        sum_y2 += dy * dy;
    }
    let denom = (sum_x2 * sum_y2).sqrt();
    if denom == 0.0 {
        CellValue::Error(CellError::Div0)
    } else {
        CellValue::Number(sum_xy / denom)
    }
}

fn fn_minifs(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    conditional_aggregate(args, ctx, reg, |vals| {
        vals.iter().cloned().fold(f64::INFINITY, f64::min)
    })
}

fn fn_maxifs(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    conditional_aggregate(args, ctx, reg, |vals| {
        vals.iter().cloned().fold(f64::NEG_INFINITY, f64::max)
    })
}

fn conditional_aggregate(
    args: &[Expr],
    ctx: &dyn EvalContext,
    reg: &FunctionRegistry,
    agg: fn(&[f64]) -> f64,
) -> CellValue {
    let target_vals = match &args[0] {
        Expr::Range { start, end } => collect_range_values(start, end, ctx),
        _ => return CellValue::Error(CellError::Value),
    };

    let mut mask = vec![true; target_vals.len()];

    let mut i = 1;
    while i + 1 < args.len() {
        let criteria_range = match &args[i] {
            Expr::Range { start, end } => collect_range_values(start, end, ctx),
            _ => return CellValue::Error(CellError::Value),
        };
        let criteria = evaluate(&args[i + 1], ctx, reg).display_value();
        for (j, v) in criteria_range.iter().enumerate() {
            if j < mask.len() && !matches_criteria(v, &criteria) {
                mask[j] = false;
            }
        }
        i += 2;
    }

    let filtered: Vec<f64> = target_vals
        .iter()
        .enumerate()
        .filter(|(j, _)| mask.get(*j).copied().unwrap_or(false))
        .filter_map(|(_, v)| v.as_f64())
        .collect();

    if filtered.is_empty() {
        CellValue::Number(0.0)
    } else {
        CellValue::Number(agg(&filtered))
    }
}
