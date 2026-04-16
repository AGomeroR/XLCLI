use chrono::{Datelike, Duration, NaiveDate, NaiveDateTime, NaiveTime, Timelike};

use xlcli_core::cell::CellValue;
use xlcli_core::types::CellError;

use crate::ast::Expr;
use crate::eval::{evaluate, EvalContext};
use crate::registry::{FnSpec, FunctionRegistry};

pub fn register(reg: &mut FunctionRegistry) {
    reg.register(FnSpec { name: "NOW", min_args: 0, max_args: Some(0), eval: fn_now });
    reg.register(FnSpec { name: "TODAY", min_args: 0, max_args: Some(0), eval: fn_today });
    reg.register(FnSpec { name: "DATE", min_args: 3, max_args: Some(3), eval: fn_date });
    reg.register(FnSpec { name: "TIME", min_args: 3, max_args: Some(3), eval: fn_time });
    reg.register(FnSpec { name: "YEAR", min_args: 1, max_args: Some(1), eval: fn_year });
    reg.register(FnSpec { name: "MONTH", min_args: 1, max_args: Some(1), eval: fn_month });
    reg.register(FnSpec { name: "DAY", min_args: 1, max_args: Some(1), eval: fn_day });
    reg.register(FnSpec { name: "HOUR", min_args: 1, max_args: Some(1), eval: fn_hour });
    reg.register(FnSpec { name: "MINUTE", min_args: 1, max_args: Some(1), eval: fn_minute });
    reg.register(FnSpec { name: "SECOND", min_args: 1, max_args: Some(1), eval: fn_second });
    reg.register(FnSpec { name: "WEEKDAY", min_args: 1, max_args: Some(2), eval: fn_weekday });
    reg.register(FnSpec { name: "WEEKNUM", min_args: 1, max_args: Some(2), eval: fn_weeknum });
    reg.register(FnSpec { name: "DATEVALUE", min_args: 1, max_args: Some(1), eval: fn_datevalue });
    reg.register(FnSpec { name: "DAYS", min_args: 2, max_args: Some(2), eval: fn_days });
    reg.register(FnSpec { name: "EDATE", min_args: 2, max_args: Some(2), eval: fn_edate });
    reg.register(FnSpec { name: "EOMONTH", min_args: 2, max_args: Some(2), eval: fn_eomonth });
    reg.register(FnSpec { name: "DATEDIF", min_args: 3, max_args: Some(3), eval: fn_datedif });
    reg.register(FnSpec { name: "ISOWEEKNUM", min_args: 1, max_args: Some(1), eval: fn_isoweeknum });
}

const EXCEL_EPOCH: i64 = 25569; // days between 1899-12-30 and 1970-01-01

fn serial_to_date(serial: f64) -> Option<NaiveDate> {
    let days = serial as i64 - EXCEL_EPOCH;
    NaiveDate::from_ymd_opt(1970, 1, 1).and_then(|epoch| epoch.checked_add_signed(Duration::days(days)))
}

fn date_to_serial(date: NaiveDate) -> f64 {
    let epoch = NaiveDate::from_ymd_opt(1970, 1, 1).unwrap();
    let days = (date - epoch).num_days();
    (days + EXCEL_EPOCH) as f64
}

fn serial_to_datetime(serial: f64) -> Option<NaiveDateTime> {
    let date = serial_to_date(serial)?;
    let frac = serial.fract();
    let secs = (frac * 86400.0).round() as u32;
    let time = NaiveTime::from_num_seconds_from_midnight_opt(secs, 0)?;
    Some(NaiveDateTime::new(date, time))
}

fn eval_as_serial(expr: &Expr, ctx: &dyn EvalContext, reg: &FunctionRegistry) -> Option<f64> {
    let val = evaluate(expr, ctx, reg);
    match val {
        CellValue::Number(n) => Some(n),
        CellValue::DateTime(dt) => Some(date_to_serial(dt.date())),
        _ => val.as_f64(),
    }
}

fn fn_now(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let now = chrono::Local::now().naive_local();
    CellValue::DateTime(now)
}

fn fn_today(_args: &[Expr], _ctx: &dyn EvalContext, _reg: &FunctionRegistry) -> CellValue {
    let today = chrono::Local::now().date_naive();
    CellValue::Number(date_to_serial(today))
}

fn fn_date(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let year = match evaluate(&args[0], ctx, reg).as_f64() {
        Some(n) => n as i32,
        None => return CellValue::Error(CellError::Value),
    };
    let month = match evaluate(&args[1], ctx, reg).as_f64() {
        Some(n) => n as u32,
        None => return CellValue::Error(CellError::Value),
    };
    let day = match evaluate(&args[2], ctx, reg).as_f64() {
        Some(n) => n as u32,
        None => return CellValue::Error(CellError::Value),
    };

    match NaiveDate::from_ymd_opt(year, month, day) {
        Some(date) => CellValue::Number(date_to_serial(date)),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_time(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let hour = match evaluate(&args[0], ctx, reg).as_f64() {
        Some(n) => n as u32,
        None => return CellValue::Error(CellError::Value),
    };
    let min = match evaluate(&args[1], ctx, reg).as_f64() {
        Some(n) => n as u32,
        None => return CellValue::Error(CellError::Value),
    };
    let sec = match evaluate(&args[2], ctx, reg).as_f64() {
        Some(n) => n as u32,
        None => return CellValue::Error(CellError::Value),
    };

    let total_secs = hour * 3600 + min * 60 + sec;
    CellValue::Number(total_secs as f64 / 86400.0)
}

fn fn_year(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_as_serial(&args[0], ctx, reg).and_then(serial_to_date) {
        Some(date) => CellValue::Number(date.year() as f64),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_month(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_as_serial(&args[0], ctx, reg).and_then(serial_to_date) {
        Some(date) => CellValue::Number(date.month() as f64),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_day(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_as_serial(&args[0], ctx, reg).and_then(serial_to_date) {
        Some(date) => CellValue::Number(date.day() as f64),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_hour(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_as_serial(&args[0], ctx, reg).and_then(serial_to_datetime) {
        Some(dt) => CellValue::Number(dt.hour() as f64),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_minute(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_as_serial(&args[0], ctx, reg).and_then(serial_to_datetime) {
        Some(dt) => CellValue::Number(dt.minute() as f64),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_second(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    match eval_as_serial(&args[0], ctx, reg).and_then(serial_to_datetime) {
        Some(dt) => CellValue::Number(dt.second() as f64),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_weekday(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let date = match eval_as_serial(&args[0], ctx, reg).and_then(serial_to_date) {
        Some(d) => d,
        None => return CellValue::Error(CellError::Value),
    };
    let return_type = if args.len() > 1 {
        evaluate(&args[1], ctx, reg).as_f64().unwrap_or(1.0) as u8
    } else {
        1
    };
    let weekday = date.weekday().num_days_from_sunday(); // 0=Sun, 6=Sat
    let result = match return_type {
        1 => weekday + 1,       // 1=Sun..7=Sat
        2 => (weekday + 6) % 7 + 1, // 1=Mon..7=Sun
        3 => (weekday + 6) % 7,     // 0=Mon..6=Sun
        _ => return CellValue::Error(CellError::Num),
    };
    CellValue::Number(result as f64)
}

fn fn_weeknum(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let date = match eval_as_serial(&args[0], ctx, reg).and_then(serial_to_date) {
        Some(d) => d,
        None => return CellValue::Error(CellError::Value),
    };
    let jan1 = NaiveDate::from_ymd_opt(date.year(), 1, 1).unwrap();
    let day_of_year = (date - jan1).num_days();
    let jan1_weekday = jan1.weekday().num_days_from_sunday();
    let week = (day_of_year as u32 + jan1_weekday) / 7 + 1;
    CellValue::Number(week as f64)
}

fn fn_datevalue(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let s = evaluate(&args[0], ctx, reg).display_value();
    let formats = ["%Y-%m-%d", "%m/%d/%Y", "%d-%b-%Y", "%B %d, %Y"];
    for fmt in &formats {
        if let Ok(date) = NaiveDate::parse_from_str(&s, fmt) {
            return CellValue::Number(date_to_serial(date));
        }
    }
    CellValue::Error(CellError::Value)
}

fn fn_days(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let end = match eval_as_serial(&args[0], ctx, reg) {
        Some(n) => n,
        None => return CellValue::Error(CellError::Value),
    };
    let start = match eval_as_serial(&args[1], ctx, reg) {
        Some(n) => n,
        None => return CellValue::Error(CellError::Value),
    };
    CellValue::Number((end - start).trunc())
}

fn fn_edate(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let date = match eval_as_serial(&args[0], ctx, reg).and_then(serial_to_date) {
        Some(d) => d,
        None => return CellValue::Error(CellError::Value),
    };
    let months = match evaluate(&args[1], ctx, reg).as_f64() {
        Some(n) => n as i32,
        None => return CellValue::Error(CellError::Value),
    };

    let total_months = date.year() * 12 + date.month() as i32 - 1 + months;
    let new_year = total_months.div_euclid(12);
    let new_month = (total_months.rem_euclid(12) + 1) as u32;
    let new_day = date.day().min(days_in_month(new_year, new_month));

    match NaiveDate::from_ymd_opt(new_year, new_month, new_day) {
        Some(d) => CellValue::Number(date_to_serial(d)),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_eomonth(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let date = match eval_as_serial(&args[0], ctx, reg).and_then(serial_to_date) {
        Some(d) => d,
        None => return CellValue::Error(CellError::Value),
    };
    let months = match evaluate(&args[1], ctx, reg).as_f64() {
        Some(n) => n as i32,
        None => return CellValue::Error(CellError::Value),
    };

    let total_months = date.year() * 12 + date.month() as i32 - 1 + months;
    let new_year = total_months.div_euclid(12);
    let new_month = (total_months.rem_euclid(12) + 1) as u32;
    let last_day = days_in_month(new_year, new_month);

    match NaiveDate::from_ymd_opt(new_year, new_month, last_day) {
        Some(d) => CellValue::Number(date_to_serial(d)),
        None => CellValue::Error(CellError::Value),
    }
}

fn fn_datedif(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let start = match eval_as_serial(&args[0], ctx, reg).and_then(serial_to_date) {
        Some(d) => d,
        None => return CellValue::Error(CellError::Value),
    };
    let end = match eval_as_serial(&args[1], ctx, reg).and_then(serial_to_date) {
        Some(d) => d,
        None => return CellValue::Error(CellError::Value),
    };
    let unit = evaluate(&args[2], ctx, reg).display_value().to_uppercase();

    if start > end {
        return CellValue::Error(CellError::Num);
    }

    match unit.as_str() {
        "Y" => CellValue::Number((end.year() - start.year()) as f64 - if end.ordinal() < start.ordinal() { 1.0 } else { 0.0 }),
        "M" => {
            let months = (end.year() - start.year()) * 12 + end.month() as i32 - start.month() as i32;
            let adj = if end.day() < start.day() { 1 } else { 0 };
            CellValue::Number((months - adj) as f64)
        }
        "D" => CellValue::Number((end - start).num_days() as f64),
        _ => CellValue::Error(CellError::Num),
    }
}

fn fn_isoweeknum(args: &[Expr], ctx: &dyn EvalContext, reg: &FunctionRegistry) -> CellValue {
    let date = match eval_as_serial(&args[0], ctx, reg).and_then(serial_to_date) {
        Some(d) => d,
        None => return CellValue::Error(CellError::Value),
    };
    CellValue::Number(date.iso_week().week() as f64)
}

fn days_in_month(year: i32, month: u32) -> u32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 => {
            if (year % 4 == 0 && year % 100 != 0) || year % 400 == 0 {
                29
            } else {
                28
            }
        }
        _ => 30,
    }
}
