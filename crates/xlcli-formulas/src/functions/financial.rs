use xlcli_core::cell::CellValue;
use xlcli_core::types::CellError;

use crate::ast::Expr;
use crate::eval::{evaluate, EvalContext};
use crate::registry::{FnSpec, FunctionRegistry};

pub fn register(reg: &mut FunctionRegistry) {
    reg.register(FnSpec { name: "PMT", description: "Returns periodic loan payment", syntax: "PMT(rate, nper, pv, [fv], [type])", min_args: 3, max_args: Some(5), eval: fn_pmt });
    reg.register(FnSpec { name: "FV", description: "Returns future value of investment", syntax: "FV(rate, nper, pmt, [pv], [type])", min_args: 3, max_args: Some(5), eval: fn_fv });
    reg.register(FnSpec { name: "PV", description: "Returns present value of investment", syntax: "PV(rate, nper, pmt, [fv], [type])", min_args: 3, max_args: Some(5), eval: fn_pv });
    reg.register(FnSpec { name: "NPV", description: "Returns net present value of cash flows", syntax: "NPV(rate, value1, [value2], ...)", min_args: 2, max_args: None, eval: fn_npv });
    reg.register(FnSpec { name: "IRR", description: "Returns internal rate of return", syntax: "IRR(values, [guess])", min_args: 1, max_args: Some(2), eval: fn_irr });
    reg.register(FnSpec { name: "RATE", description: "Returns interest rate per period", syntax: "RATE(nper, pmt, pv, [fv], [type], [guess])", min_args: 3, max_args: Some(6), eval: fn_rate });
    reg.register(FnSpec { name: "NPER", description: "Returns number of payment periods", syntax: "NPER(rate, pmt, pv, [fv], [type])", min_args: 3, max_args: Some(5), eval: fn_nper });
    reg.register(FnSpec { name: "IPMT", description: "Returns interest portion of payment", syntax: "IPMT(rate, per, nper, pv, [fv], [type])", min_args: 4, max_args: Some(6), eval: fn_ipmt });
    reg.register(FnSpec { name: "PPMT", description: "Returns principal portion of payment", syntax: "PPMT(rate, per, nper, pv, [fv], [type])", min_args: 4, max_args: Some(6), eval: fn_ppmt });
    reg.register(FnSpec { name: "CUMIPMT", description: "Returns cumulative interest paid", syntax: "CUMIPMT(rate, nper, pv, start, end, type)", min_args: 6, max_args: Some(6), eval: fn_cumipmt });
    reg.register(FnSpec { name: "CUMPRINC", description: "Returns cumulative principal paid", syntax: "CUMPRINC(rate, nper, pv, start, end, type)", min_args: 6, max_args: Some(6), eval: fn_cumprinc });
    reg.register(FnSpec { name: "SLN", description: "Returns straight-line depreciation", syntax: "SLN(cost, salvage, life)", min_args: 3, max_args: Some(3), eval: fn_sln });
    reg.register(FnSpec { name: "SYD", description: "Returns sum-of-years depreciation", syntax: "SYD(cost, salvage, life, per)", min_args: 4, max_args: Some(4), eval: fn_syd });
    reg.register(FnSpec { name: "DB", description: "Returns fixed-declining balance depreciation", syntax: "DB(cost, salvage, life, period, [month])", min_args: 4, max_args: Some(5), eval: fn_db });
    reg.register(FnSpec { name: "DDB", description: "Returns double-declining balance depreciation", syntax: "DDB(cost, salvage, life, period, [factor])", min_args: 4, max_args: Some(5), eval: fn_ddb });
    reg.register(FnSpec { name: "EFFECT", description: "Returns effective annual interest rate", syntax: "EFFECT(nominal_rate, npery)", min_args: 2, max_args: Some(2), eval: fn_effect });
    reg.register(FnSpec { name: "NOMINAL", description: "Returns nominal annual interest rate", syntax: "NOMINAL(effect_rate, npery)", min_args: 2, max_args: Some(2), eval: fn_nominal });
    reg.register(FnSpec { name: "DOLLARDE", description: "Converts fractional dollar to decimal", syntax: "DOLLARDE(fractional_dollar, fraction)", min_args: 2, max_args: Some(2), eval: fn_dollarde });
    reg.register(FnSpec { name: "DOLLARFR", description: "Converts decimal dollar to fractional", syntax: "DOLLARFR(decimal_dollar, fraction)", min_args: 2, max_args: Some(2), eval: fn_dollarfr });
    reg.register(FnSpec { name: "PDURATION", description: "Returns periods for investment to reach value", syntax: "PDURATION(rate, pv, fv)", min_args: 3, max_args: Some(3), eval: fn_pduration });
    reg.register(FnSpec { name: "RRI", description: "Returns equivalent interest rate for growth", syntax: "RRI(nper, pv, fv)", min_args: 3, max_args: Some(3), eval: fn_rri });
    reg.register(FnSpec { name: "DISC", description: "Returns discount rate for a security", syntax: "DISC(settlement, maturity, pr, redemption, [basis])", min_args: 4, max_args: Some(4), eval: fn_disc });
    reg.register(FnSpec { name: "TBILLEQ", description: "Returns bond-equivalent yield for T-bill", syntax: "TBILLEQ(settlement, maturity, discount)", min_args: 3, max_args: Some(3), eval: fn_tbilleq });
    reg.register(FnSpec { name: "TBILLPRICE", description: "Returns price per 100 face for T-bill", syntax: "TBILLPRICE(settlement, maturity, discount)", min_args: 3, max_args: Some(3), eval: fn_tbillprice });
    reg.register(FnSpec { name: "TBILLYIELD", description: "Returns yield for a T-bill", syntax: "TBILLYIELD(settlement, maturity, pr)", min_args: 3, max_args: Some(3), eval: fn_tbillyield });
    reg.register(FnSpec { name: "XNPV", description: "Returns NPV for irregular cash flows", syntax: "XNPV(rate, values, dates)", min_args: 3, max_args: Some(3), eval: fn_xnpv });
    reg.register(FnSpec { name: "MIRR", description: "Returns modified internal rate of return", syntax: "MIRR(values, finance_rate, reinvest_rate)", min_args: 3, max_args: Some(3), eval: fn_mirr });
    reg.register(FnSpec { name: "ISPMT", description: "Returns interest for even-principal loan", syntax: "ISPMT(rate, per, nper, pv)", min_args: 4, max_args: Some(4), eval: fn_ispmt });
}

fn eval_f64(expr: &Expr, ctx: &dyn EvalContext, reg: &FunctionRegistry) -> Option<f64> {
    evaluate(expr, ctx, reg).as_f64()
}

fn fn_pmt(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let rate = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let nper = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pv = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let fv = if args.len() > 3 { eval_f64(&args[3], ctx, reg).unwrap_or(0.0) } else { 0.0 };
    let pmt_type = if args.len() > 4 { eval_f64(&args[4], ctx, reg).unwrap_or(0.0) as i32 } else { 0 };

    if rate == 0.0 {
        return CellValue::Number(-(pv + fv) / nper);
    }
    let pmt = (-rate * (pv * (1.0 + rate).powf(nper) + fv)) /
              ((1.0 + rate * pmt_type as f64) * ((1.0 + rate).powf(nper) - 1.0));
    CellValue::Number(pmt)
}

fn fn_fv(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let rate = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let nper = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pmt = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pv = if args.len() > 3 { eval_f64(&args[3], ctx, reg).unwrap_or(0.0) } else { 0.0 };
    let pmt_type = if args.len() > 4 { eval_f64(&args[4], ctx, reg).unwrap_or(0.0) as i32 } else { 0 };

    if rate == 0.0 {
        return CellValue::Number(-(pv + pmt * nper));
    }
    let pow = (1.0 + rate).powf(nper);
    let fv = -(pv * pow + pmt * (1.0 + rate * pmt_type as f64) * (pow - 1.0) / rate);
    CellValue::Number(fv)
}

fn fn_pv(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let rate = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let nper = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pmt = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let fv = if args.len() > 3 { eval_f64(&args[3], ctx, reg).unwrap_or(0.0) } else { 0.0 };
    let pmt_type = if args.len() > 4 { eval_f64(&args[4], ctx, reg).unwrap_or(0.0) as i32 } else { 0 };

    if rate == 0.0 {
        return CellValue::Number(-(fv + pmt * nper));
    }
    let pow = (1.0 + rate).powf(nper);
    let pv = -(fv + pmt * (1.0 + rate * pmt_type as f64) * (pow - 1.0) / rate) / pow;
    CellValue::Number(pv)
}

fn fn_npv(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let rate = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let mut npv = 0.0;
    for (i, arg) in args[1..].iter().enumerate() {
        let cf = match eval_f64(arg, ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
        npv += cf / (1.0 + rate).powi(i as i32 + 1);
    }
    CellValue::Number(npv)
}

fn fn_irr(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let values: Vec<f64> = match &args[0] {
        crate::ast::Expr::Range { start, end } => {
            crate::eval::collect_range_values(start, end, ctx)
                .iter()
                .filter_map(|v| v.as_f64())
                .collect()
        }
        _ => return CellValue::Error(CellError::Value),
    };
    let mut guess = if args.len() > 1 { eval_f64(&args[1], ctx, reg).unwrap_or(0.1) } else { 0.1 };

    for _ in 0..100 {
        let mut npv = 0.0;
        let mut dnpv = 0.0;
        for (i, &cf) in values.iter().enumerate() {
            let pow = (1.0 + guess).powi(i as i32);
            npv += cf / pow;
            if i > 0 {
                dnpv -= (i as f64) * cf / ((1.0 + guess).powi(i as i32 + 1));
            }
        }
        if dnpv.abs() < 1e-12 { break; }
        let new_guess = guess - npv / dnpv;
        if (new_guess - guess).abs() < 1e-10 {
            return CellValue::Number(new_guess);
        }
        guess = new_guess;
    }
    CellValue::Error(CellError::Num)
}

fn fn_rate(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nper = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pmt = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pv = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let fv = if args.len() > 3 { eval_f64(&args[3], ctx, reg).unwrap_or(0.0) } else { 0.0 };
    let pmt_type = if args.len() > 4 { eval_f64(&args[4], ctx, reg).unwrap_or(0.0) as i32 } else { 0 };
    let mut guess = if args.len() > 5 { eval_f64(&args[5], ctx, reg).unwrap_or(0.1) } else { 0.1 };

    for _ in 0..100 {
        let pow = (1.0 + guess).powf(nper);
        let f = pv * pow + pmt * (1.0 + guess * pmt_type as f64) * (pow - 1.0) / guess + fv;
        let df = pv * nper * (1.0 + guess).powf(nper - 1.0)
            + pmt * (1.0 + guess * pmt_type as f64) * nper * (1.0 + guess).powf(nper - 1.0) / guess
            - pmt * (1.0 + guess * pmt_type as f64) * (pow - 1.0) / (guess * guess);
        if df.abs() < 1e-12 { break; }
        let new_guess = guess - f / df;
        if (new_guess - guess).abs() < 1e-10 {
            return CellValue::Number(new_guess);
        }
        guess = new_guess;
    }
    CellValue::Error(CellError::Num)
}

fn fn_nper(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let rate = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pmt = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pv = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let fv = if args.len() > 3 { eval_f64(&args[3], ctx, reg).unwrap_or(0.0) } else { 0.0 };
    let pmt_type = if args.len() > 4 { eval_f64(&args[4], ctx, reg).unwrap_or(0.0) as i32 } else { 0 };

    if rate == 0.0 {
        if pmt == 0.0 { return CellValue::Error(CellError::Num); }
        return CellValue::Number(-(pv + fv) / pmt);
    }
    let z = pmt * (1.0 + rate * pmt_type as f64) / rate;
    let numer = -fv + z;
    let denom = pv + z;
    if numer <= 0.0 || denom <= 0.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number((numer / denom).ln() / (1.0 + rate).ln())
}

fn fn_ipmt(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let rate = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let per = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let nper = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pv = match eval_f64(&args[3], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let fv = if args.len() > 4 { eval_f64(&args[4], ctx, reg).unwrap_or(0.0) } else { 0.0 };
    let pmt_type = if args.len() > 5 { eval_f64(&args[5], ctx, reg).unwrap_or(0.0) as i32 } else { 0 };

    let pmt = calc_pmt(rate, nper, pv, fv, pmt_type);
    let fv_at_per = pv * (1.0 + rate).powf(per - 1.0)
        + pmt * (1.0 + rate * pmt_type as f64) * ((1.0 + rate).powf(per - 1.0) - 1.0) / rate;
    CellValue::Number(fv_at_per * rate)
}

fn fn_ppmt(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let rate = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let per = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let nper = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pv = match eval_f64(&args[3], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let fv = if args.len() > 4 { eval_f64(&args[4], ctx, reg).unwrap_or(0.0) } else { 0.0 };
    let pmt_type = if args.len() > 5 { eval_f64(&args[5], ctx, reg).unwrap_or(0.0) as i32 } else { 0 };

    let pmt = calc_pmt(rate, nper, pv, fv, pmt_type);
    let fv_at_per = pv * (1.0 + rate).powf(per - 1.0)
        + pmt * (1.0 + rate * pmt_type as f64) * ((1.0 + rate).powf(per - 1.0) - 1.0) / rate;
    let ipmt = fv_at_per * rate;
    CellValue::Number(pmt - ipmt)
}

fn calc_pmt(rate: f64, nper: f64, pv: f64, fv: f64, pmt_type: i32) -> f64 {
    if rate == 0.0 {
        return -(pv + fv) / nper;
    }
    (-rate * (pv * (1.0 + rate).powf(nper) + fv)) /
        ((1.0 + rate * pmt_type as f64) * ((1.0 + rate).powf(nper) - 1.0))
}

fn fn_cumipmt(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let rate = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let nper = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pv = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let start = match eval_f64(&args[3], ctx, reg) { Some(v) => v as i32, None => return CellValue::Error(CellError::Value) };
    let end = match eval_f64(&args[4], ctx, reg) { Some(v) => v as i32, None => return CellValue::Error(CellError::Value) };
    let pmt_type = match eval_f64(&args[5], ctx, reg) { Some(v) => v as i32, None => return CellValue::Error(CellError::Value) };

    let pmt = calc_pmt(rate, nper, pv, 0.0, pmt_type);
    let mut total = 0.0;
    for per in start..=end {
        let fv_at = pv * (1.0 + rate).powi(per - 1)
            + pmt * (1.0 + rate * pmt_type as f64) * ((1.0 + rate).powi(per - 1) - 1.0) / rate;
        total += fv_at * rate;
    }
    CellValue::Number(total)
}

fn fn_cumprinc(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let rate = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let nper = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pv = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let start = match eval_f64(&args[3], ctx, reg) { Some(v) => v as i32, None => return CellValue::Error(CellError::Value) };
    let end = match eval_f64(&args[4], ctx, reg) { Some(v) => v as i32, None => return CellValue::Error(CellError::Value) };
    let pmt_type = match eval_f64(&args[5], ctx, reg) { Some(v) => v as i32, None => return CellValue::Error(CellError::Value) };

    let pmt = calc_pmt(rate, nper, pv, 0.0, pmt_type);
    let mut total = 0.0;
    for per in start..=end {
        let fv_at = pv * (1.0 + rate).powi(per - 1)
            + pmt * (1.0 + rate * pmt_type as f64) * ((1.0 + rate).powi(per - 1) - 1.0) / rate;
        let ipmt = fv_at * rate;
        total += pmt - ipmt;
    }
    CellValue::Number(total)
}

fn fn_sln(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let cost = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let salvage = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let life = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    if life == 0.0 { return CellValue::Error(CellError::Div0); }
    CellValue::Number((cost - salvage) / life)
}

fn fn_syd(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let cost = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let salvage = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let life = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let per = match eval_f64(&args[3], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let sum = life * (life + 1.0) / 2.0;
    CellValue::Number((cost - salvage) * (life - per + 1.0) / sum)
}

fn fn_db(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let cost = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let salvage = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let life = match eval_f64(&args[2], ctx, reg) { Some(v) => v as i32, None => return CellValue::Error(CellError::Value) };
    let period = match eval_f64(&args[3], ctx, reg) { Some(v) => v as i32, None => return CellValue::Error(CellError::Value) };
    let month = if args.len() > 4 { eval_f64(&args[4], ctx, reg).unwrap_or(12.0) as i32 } else { 12 };

    if cost <= 0.0 || life <= 0 { return CellValue::Error(CellError::Num); }
    let rate = (1.0 - (salvage / cost).powf(1.0 / life as f64) * 1000.0).round() / 1000.0;
    let mut total_dep = 0.0;
    let mut value = cost;

    for p in 1..=period {
        let dep = if p == 1 {
            cost * rate * month as f64 / 12.0
        } else if p == life + 1 {
            value * rate * (12.0 - month as f64) / 12.0
        } else {
            value * rate
        };
        total_dep = dep;
        value -= dep;
    }
    CellValue::Number(total_dep)
}

fn fn_ddb(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let cost = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let salvage = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let life = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let period = match eval_f64(&args[3], ctx, reg) { Some(v) => v as i32, None => return CellValue::Error(CellError::Value) };
    let factor = if args.len() > 4 { eval_f64(&args[4], ctx, reg).unwrap_or(2.0) } else { 2.0 };

    let mut value = cost;
    let mut dep = 0.0;
    for p in 1..=period {
        dep = (value * factor / life).min(value - salvage).max(0.0);
        value -= dep;
        if p < period { dep = 0.0; }
    }
    CellValue::Number(dep)
}

fn fn_effect(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nominal = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let npery = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    if npery < 1.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number((1.0 + nominal / npery).powf(npery) - 1.0)
}

fn fn_nominal(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let effect = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let npery = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    if npery < 1.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number(npery * ((1.0 + effect).powf(1.0 / npery) - 1.0))
}

fn fn_dollarde(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let dollar = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let fraction = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    if fraction < 1.0 { return CellValue::Error(CellError::Div0); }
    let int_part = dollar.trunc();
    let frac_part = dollar.fract();
    CellValue::Number(int_part + frac_part * 10.0_f64.powf((fraction as i32 as f64).log10().ceil()) / fraction)
}

fn fn_dollarfr(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let dollar = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let fraction = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    if fraction < 1.0 { return CellValue::Error(CellError::Div0); }
    let int_part = dollar.trunc();
    let frac_part = dollar.fract();
    CellValue::Number(int_part + frac_part * fraction / 10.0_f64.powf((fraction as i32 as f64).log10().ceil()))
}

fn fn_pduration(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let rate = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pv = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let fv = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    if rate <= 0.0 || pv <= 0.0 || fv <= 0.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number((fv.ln() - pv.ln()) / (1.0 + rate).ln())
}

fn fn_rri(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let nper = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pv = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let fv = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    if nper <= 0.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number((fv / pv).powf(1.0 / nper) - 1.0)
}

fn fn_disc(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let _settlement = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let _maturity = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pr = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let redemption = match eval_f64(&args[3], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let days = (_maturity - _settlement).abs();
    if days == 0.0 || redemption == 0.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number((redemption - pr) / redemption * (360.0 / days))
}

fn fn_tbilleq(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let _settlement = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let _maturity = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let discount = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let dsm = (_maturity - _settlement).abs();
    if dsm == 0.0 { return CellValue::Error(CellError::Num); }
    let price = 100.0 * (1.0 - discount * dsm / 360.0);
    CellValue::Number((100.0 - price) / price * (365.0 / dsm))
}

fn fn_tbillprice(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let _settlement = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let _maturity = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let discount = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let dsm = (_maturity - _settlement).abs();
    CellValue::Number(100.0 * (1.0 - discount * dsm / 360.0))
}

fn fn_tbillyield(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let _settlement = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let _maturity = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let price = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let dsm = (_maturity - _settlement).abs();
    if dsm == 0.0 || price == 0.0 { return CellValue::Error(CellError::Num); }
    CellValue::Number((100.0 - price) / price * (360.0 / dsm))
}

fn fn_xnpv(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let rate = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let values: Vec<f64> = match &args[1] {
        crate::ast::Expr::Range { start, end } => {
            crate::eval::collect_range_values(start, end, ctx)
                .iter().filter_map(|v| v.as_f64()).collect()
        }
        _ => return CellValue::Error(CellError::Value),
    };
    let dates: Vec<f64> = match &args[2] {
        crate::ast::Expr::Range { start, end } => {
            crate::eval::collect_range_values(start, end, ctx)
                .iter().filter_map(|v| v.as_f64()).collect()
        }
        _ => return CellValue::Error(CellError::Value),
    };
    if values.len() != dates.len() || values.is_empty() { return CellValue::Error(CellError::Num); }
    let d0 = dates[0];
    let mut xnpv = 0.0;
    for (i, &cf) in values.iter().enumerate() {
        xnpv += cf / (1.0 + rate).powf((dates[i] - d0) / 365.0);
    }
    CellValue::Number(xnpv)
}

fn fn_mirr(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let values: Vec<f64> = match &args[0] {
        crate::ast::Expr::Range { start, end } => {
            crate::eval::collect_range_values(start, end, ctx)
                .iter().filter_map(|v| v.as_f64()).collect()
        }
        _ => return CellValue::Error(CellError::Value),
    };
    let finance_rate = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let reinvest_rate = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };

    let n = values.len() as f64;
    let mut pv_neg = 0.0;
    let mut fv_pos = 0.0;
    for (i, &cf) in values.iter().enumerate() {
        if cf < 0.0 {
            pv_neg += cf / (1.0 + finance_rate).powi(i as i32);
        } else {
            fv_pos += cf * (1.0 + reinvest_rate).powi((values.len() - 1 - i) as i32);
        }
    }
    if pv_neg == 0.0 { return CellValue::Error(CellError::Div0); }
    CellValue::Number((-fv_pos / pv_neg).powf(1.0 / (n - 1.0)) - 1.0)
}

fn fn_ispmt(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let rate = match eval_f64(&args[0], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let per = match eval_f64(&args[1], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let nper = match eval_f64(&args[2], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    let pv = match eval_f64(&args[3], ctx, reg) { Some(v) => v, None => return CellValue::Error(CellError::Value) };
    CellValue::Number(pv * rate * (per / nper - 1.0))
}
