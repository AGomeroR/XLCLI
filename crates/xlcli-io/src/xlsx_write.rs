use std::path::Path;

use rust_xlsxwriter::{
    Color as XlsxColor, ConditionalFormat2ColorScale, ConditionalFormat3ColorScale,
    ConditionalFormatBlank, ConditionalFormatCell, ConditionalFormatCellRule,
    ConditionalFormatCustomIcon, ConditionalFormatDataBar, ConditionalFormatFormula,
    ConditionalFormatIconSet, ConditionalFormatIconType, ConditionalFormatText,
    ConditionalFormatTextRule, ConditionalFormatType, Format, FormatUnderline, Formula,
    Workbook as XlsxWorkbook,
};
use xlcli_core::cell::CellValue;
use xlcli_core::condfmt::{
    CfValueKind, ColorStop, CondRule, Condition, IconSetKind, IconThreshold, StyleOverlay,
    StyleSpec,
};
use xlcli_core::range::CellRange;
use xlcli_core::style::{CellStyle, Color};
use xlcli_core::workbook::Workbook;

use crate::writer::FileWriter;

pub struct XlsxWriter;

impl FileWriter for XlsxWriter {
    fn write(&self, workbook: &Workbook, path: &Path) -> anyhow::Result<()> {
        let mut xlsx = XlsxWorkbook::new();

        for sheet in &workbook.sheets {
            let ws = xlsx.add_worksheet();
            ws.set_name(&sheet.name)?;

            for (&(row, col), cell) in sheet.cells_iter() {
                let cs = workbook.style_pool.get(cell.style_id).clone();
                let fmt = cellstyle_to_format(&cs);
                let has_style = cell.style_id != 0;

                match &cell.value {
                    CellValue::Empty => {
                        if has_style {
                            ws.write_blank(row, col, &fmt)?;
                        }
                    }
                    CellValue::Number(n) => {
                        if has_style { ws.write_number_with_format(row, col, *n, &fmt)?; }
                        else { ws.write_number(row, col, *n)?; }
                    }
                    CellValue::String(s) => {
                        if has_style { ws.write_string_with_format(row, col, s.as_str(), &fmt)?; }
                        else { ws.write_string(row, col, s.as_str())?; }
                    }
                    CellValue::Boolean(b) => {
                        if has_style { ws.write_boolean_with_format(row, col, *b, &fmt)?; }
                        else { ws.write_boolean(row, col, *b)?; }
                    }
                    CellValue::DateTime(dt) => {
                        let serial = datetime_to_excel_serial(dt);
                        let f = if has_style {
                            fmt.clone().set_num_format("yyyy-mm-dd hh:mm:ss")
                        } else {
                            Format::new().set_num_format("yyyy-mm-dd hh:mm:ss")
                        };
                        ws.write_number_with_format(row, col, serial, &f)?;
                    }
                    CellValue::Error(_) => {
                        ws.write_string(row, col, &cell.value.display_value())?;
                    }
                    CellValue::Array(_) => {
                        ws.write_string(row, col, "{...}")?;
                    }
                }

                if let Some(formula) = &cell.formula {
                    ws.write_formula(row, col, formula.as_str())?;
                }
            }

            for rule in &sheet.cond_rules {
                emit_cond_rule(ws, rule)?;
            }
        }

        xlsx.save(path)?;
        Ok(())
    }

    fn extensions(&self) -> &[&str] {
        &["xlsx"]
    }
}

fn datetime_to_excel_serial(dt: &chrono::NaiveDateTime) -> f64 {
    use chrono::NaiveDate;
    let base = NaiveDate::from_ymd_opt(1899, 12, 30)
        .unwrap()
        .and_hms_opt(0, 0, 0)
        .unwrap();
    let diff = *dt - base;
    diff.num_milliseconds() as f64 / 86_400_000.0
}

fn cellstyle_to_format(cs: &CellStyle) -> Format {
    let mut f = Format::new();
    if cs.bold { f = f.set_bold(); }
    if cs.italic { f = f.set_italic(); }
    if cs.strikethrough { f = f.set_font_strikethrough(); }
    if cs.double_underline { f = f.set_underline(FormatUnderline::Double); }
    else if cs.underline { f = f.set_underline(FormatUnderline::Single); }
    if let Some(c) = cs.fg_color { f = f.set_font_color(color_to_xlsx(&c)); }
    if let Some(c) = cs.bg_color { f = f.set_background_color(color_to_xlsx(&c)); }
    f
}

fn color_to_xlsx(c: &Color) -> XlsxColor {
    let rgb: u32 = ((c.r as u32) << 16) | ((c.g as u32) << 8) | (c.b as u32);
    XlsxColor::RGB(rgb)
}

fn overlay_to_format(o: &StyleOverlay) -> Format {
    let mut f = Format::new();
    if o.bold == Some(true) { f = f.set_bold(); }
    if o.italic == Some(true) { f = f.set_italic(); }
    if o.strikethrough == Some(true) { f = f.set_font_strikethrough(); }
    if o.double_underline == Some(true) {
        f = f.set_underline(FormatUnderline::Double);
    } else if o.underline == Some(true) {
        f = f.set_underline(FormatUnderline::Single);
    }
    if let Some(Some(c)) = o.fg_color { f = f.set_font_color(color_to_xlsx(&c)); }
    if let Some(Some(c)) = o.bg_color { f = f.set_background_color(color_to_xlsx(&c)); }
    f
}

fn emit_cond_rule(
    ws: &mut rust_xlsxwriter::Worksheet,
    rule: &CondRule,
) -> anyhow::Result<()> {
    match &rule.style {
        StyleSpec::Overlay(o) => emit_overlay_rule(ws, &rule.range, &rule.cond, o),
        StyleSpec::ColorScale(stops) => emit_color_scale(ws, &rule.range, stops),
        StyleSpec::DataBar { min, max, color } => emit_data_bar(ws, &rule.range, min, max, color),
        StyleSpec::IconSet { kind, thresholds, reverse, show_value } => {
            emit_icon_set(ws, &rule.range, *kind, thresholds, *reverse, *show_value)
        }
    }
}

fn emit_overlay_rule(
    ws: &mut rust_xlsxwriter::Worksheet,
    range: &CellRange,
    cond: &Condition,
    overlay: &StyleOverlay,
) -> anyhow::Result<()> {
    let fmt = overlay_to_format(overlay);
    let (r0, c0, r1, c1) = (range.start.row, range.start.col, range.end.row, range.end.col);

    match cond {
        Condition::Always => {
            let cf = ConditionalFormatFormula::new()
                .set_rule("TRUE")
                .set_format(fmt);
            ws.add_conditional_format(r0, c0, r1, c1, &cf)?;
        }
        Condition::Gt(n) => emit_cell(ws, r0, c0, r1, c1, ConditionalFormatCellRule::GreaterThan(*n), fmt)?,
        Condition::Lt(n) => emit_cell(ws, r0, c0, r1, c1, ConditionalFormatCellRule::LessThan(*n), fmt)?,
        Condition::Gte(n) => emit_cell(ws, r0, c0, r1, c1, ConditionalFormatCellRule::GreaterThanOrEqualTo(*n), fmt)?,
        Condition::Lte(n) => emit_cell(ws, r0, c0, r1, c1, ConditionalFormatCellRule::LessThanOrEqualTo(*n), fmt)?,
        Condition::Eq(n) => emit_cell(ws, r0, c0, r1, c1, ConditionalFormatCellRule::EqualTo(*n), fmt)?,
        Condition::Neq(n) => emit_cell(ws, r0, c0, r1, c1, ConditionalFormatCellRule::NotEqualTo(*n), fmt)?,
        Condition::Between(a, b) => emit_cell(ws, r0, c0, r1, c1, ConditionalFormatCellRule::Between(*a, *b), fmt)?,
        Condition::NotBetween(a, b) => emit_cell(ws, r0, c0, r1, c1, ConditionalFormatCellRule::NotBetween(*a, *b), fmt)?,
        Condition::Contains(s) => emit_text(ws, r0, c0, r1, c1, ConditionalFormatTextRule::Contains(s.clone()), fmt)?,
        Condition::NotContains(s) => emit_text(ws, r0, c0, r1, c1, ConditionalFormatTextRule::DoesNotContain(s.clone()), fmt)?,
        Condition::BeginsWith(s) => emit_text(ws, r0, c0, r1, c1, ConditionalFormatTextRule::BeginsWith(s.clone()), fmt)?,
        Condition::EndsWith(s) => emit_text(ws, r0, c0, r1, c1, ConditionalFormatTextRule::EndsWith(s.clone()), fmt)?,
        Condition::Blanks => {
            let cf = ConditionalFormatBlank::new().set_format(fmt);
            ws.add_conditional_format(r0, c0, r1, c1, &cf)?;
        }
        Condition::NonBlanks => {
            let cf = ConditionalFormatBlank::new().invert().set_format(fmt);
            ws.add_conditional_format(r0, c0, r1, c1, &cf)?;
        }
        Condition::Expression(expr) => {
            let cf = ConditionalFormatFormula::new()
                .set_rule(expr.as_str())
                .set_format(fmt);
            ws.add_conditional_format(r0, c0, r1, c1, &cf)?;
        }
        // Unsupported variants in overlay path: silently skip
        Condition::ContainsErrors | Condition::NotContainsErrors
        | Condition::DuplicateValues | Condition::UniqueValues
        | Condition::Top { .. } | Condition::Average { .. }
        | Condition::TimePeriod(_) => {}
    }
    Ok(())
}

fn cf_value(v: &CfValueKind) -> (ConditionalFormatType, f64, Option<String>) {
    match v {
        CfValueKind::Number(n) => (ConditionalFormatType::Number, *n, None),
        CfValueKind::Percent(n) => (ConditionalFormatType::Percent, *n, None),
        CfValueKind::Percentile(n) => (ConditionalFormatType::Percentile, *n, None),
        CfValueKind::Min => (ConditionalFormatType::Lowest, 0.0, None),
        CfValueKind::Max => (ConditionalFormatType::Highest, 0.0, None),
        CfValueKind::Formula(s) => (ConditionalFormatType::Formula, 0.0, Some(s.clone())),
    }
}

fn emit_color_scale(
    ws: &mut rust_xlsxwriter::Worksheet,
    range: &CellRange,
    stops: &[ColorStop],
) -> anyhow::Result<()> {
    let (r0, c0, r1, c1) = (range.start.row, range.start.col, range.end.row, range.end.col);
    if stops.len() <= 2 {
        let mut cf = ConditionalFormat2ColorScale::new();
        if let Some(s) = stops.first() {
            cf = apply_min(cf, &s.value, &s.color, false);
        }
        if let Some(s) = stops.get(1) {
            cf = apply_max(cf, &s.value, &s.color, false);
        }
        ws.add_conditional_format(r0, c0, r1, c1, &cf)?;
    } else {
        let mut cf = ConditionalFormat3ColorScale::new();
        cf = apply_min3(cf, &stops[0].value, &stops[0].color);
        cf = apply_mid3(cf, &stops[1].value, &stops[1].color);
        let last = stops.last().unwrap();
        cf = apply_max3(cf, &last.value, &last.color);
        ws.add_conditional_format(r0, c0, r1, c1, &cf)?;
    }
    Ok(())
}

fn apply_min(
    cf: ConditionalFormat2ColorScale,
    v: &CfValueKind,
    col: &Color,
    _midpoint: bool,
) -> ConditionalFormat2ColorScale {
    let (t, n, fx) = cf_value(v);
    let cf = match fx {
        Some(f) => cf.set_minimum(t, Formula::new(f.as_str())),
        None => cf.set_minimum(t, n),
    };
    cf.set_minimum_color(color_to_xlsx(col))
}

fn apply_max(
    cf: ConditionalFormat2ColorScale,
    v: &CfValueKind,
    col: &Color,
    _midpoint: bool,
) -> ConditionalFormat2ColorScale {
    let (t, n, fx) = cf_value(v);
    let cf = match fx {
        Some(f) => cf.set_maximum(t, Formula::new(f.as_str())),
        None => cf.set_maximum(t, n),
    };
    cf.set_maximum_color(color_to_xlsx(col))
}

fn apply_min3(
    cf: ConditionalFormat3ColorScale,
    v: &CfValueKind,
    col: &Color,
) -> ConditionalFormat3ColorScale {
    let (t, n, fx) = cf_value(v);
    let cf = match fx {
        Some(f) => cf.set_minimum(t, Formula::new(f.as_str())),
        None => cf.set_minimum(t, n),
    };
    cf.set_minimum_color(color_to_xlsx(col))
}

fn apply_mid3(
    cf: ConditionalFormat3ColorScale,
    v: &CfValueKind,
    col: &Color,
) -> ConditionalFormat3ColorScale {
    let (t, n, fx) = cf_value(v);
    let cf = match fx {
        Some(f) => cf.set_midpoint(t, Formula::new(f.as_str())),
        None => cf.set_midpoint(t, n),
    };
    cf.set_midpoint_color(color_to_xlsx(col))
}

fn apply_max3(
    cf: ConditionalFormat3ColorScale,
    v: &CfValueKind,
    col: &Color,
) -> ConditionalFormat3ColorScale {
    let (t, n, fx) = cf_value(v);
    let cf = match fx {
        Some(f) => cf.set_maximum(t, Formula::new(f.as_str())),
        None => cf.set_maximum(t, n),
    };
    cf.set_maximum_color(color_to_xlsx(col))
}

fn emit_data_bar(
    ws: &mut rust_xlsxwriter::Worksheet,
    range: &CellRange,
    min: &CfValueKind,
    max: &CfValueKind,
    color: &Color,
) -> anyhow::Result<()> {
    let (r0, c0, r1, c1) = (range.start.row, range.start.col, range.end.row, range.end.col);
    let mut cf = ConditionalFormatDataBar::new().set_fill_color(color_to_xlsx(color));
    let (mt, mn, mfx) = cf_value(min);
    cf = match mfx {
        Some(f) => cf.set_minimum(mt, Formula::new(f.as_str())),
        None => cf.set_minimum(mt, mn),
    };
    let (xt, xn, xfx) = cf_value(max);
    cf = match xfx {
        Some(f) => cf.set_maximum(xt, Formula::new(f.as_str())),
        None => cf.set_maximum(xt, xn),
    };
    ws.add_conditional_format(r0, c0, r1, c1, &cf)?;
    Ok(())
}

fn icon_kind_to_xlsx(k: IconSetKind) -> ConditionalFormatIconType {
    use IconSetKind::*;
    match k {
        ThreeArrows => ConditionalFormatIconType::ThreeArrows,
        ThreeArrowsGray => ConditionalFormatIconType::ThreeArrowsGray,
        ThreeFlags => ConditionalFormatIconType::ThreeFlags,
        ThreeSigns => ConditionalFormatIconType::ThreeSigns,
        ThreeSymbols => ConditionalFormatIconType::ThreeSymbols,
        ThreeSymbols2 => ConditionalFormatIconType::ThreeSymbolsCircled,
        ThreeTrafficLights1 => ConditionalFormatIconType::ThreeTrafficLights,
        ThreeTrafficLights2 => ConditionalFormatIconType::ThreeTrafficLightsWithRim,
        FourArrows => ConditionalFormatIconType::FourArrows,
        FourArrowsGray => ConditionalFormatIconType::FourArrowsGray,
        FourRating => ConditionalFormatIconType::FourHistograms,
        FourRedToBlack => ConditionalFormatIconType::FourRedToBlack,
        FourTrafficLights => ConditionalFormatIconType::FourTrafficLights,
        FiveArrows => ConditionalFormatIconType::FiveArrows,
        FiveArrowsGray => ConditionalFormatIconType::FiveArrowsGray,
        FiveRating => ConditionalFormatIconType::FiveHistograms,
        FiveQuarters => ConditionalFormatIconType::FiveQuadrants,
    }
}

fn emit_icon_set(
    ws: &mut rust_xlsxwriter::Worksheet,
    range: &CellRange,
    kind: IconSetKind,
    thresholds: &[IconThreshold],
    reverse: bool,
    show_value: bool,
) -> anyhow::Result<()> {
    let (r0, c0, r1, c1) = (range.start.row, range.start.col, range.end.row, range.end.col);
    let mut cf = ConditionalFormatIconSet::new()
        .set_icon_type(icon_kind_to_xlsx(kind))
        .reverse_icons(reverse)
        .show_icons_only(!show_value);

    if !thresholds.is_empty() {
        let icons: Vec<ConditionalFormatCustomIcon> = thresholds
            .iter()
            .map(|t| {
                let (ty, n, fx) = cf_value(&t.value);
                let icon = ConditionalFormatCustomIcon::new();
                let icon = match fx {
                    Some(f) => icon.set_rule(ty, Formula::new(f.as_str())),
                    None => icon.set_rule(ty, n),
                };
                icon.set_greater_than(!t.gte)
            })
            .collect();
        cf = cf.set_icons(&icons);
    }

    ws.add_conditional_format(r0, c0, r1, c1, &cf)?;
    Ok(())
}

fn emit_text(
    ws: &mut rust_xlsxwriter::Worksheet,
    r0: u32, c0: u16, r1: u32, c1: u16,
    rule: ConditionalFormatTextRule,
    fmt: Format,
) -> anyhow::Result<()> {
    let cf = ConditionalFormatText::new().set_rule(rule).set_format(fmt);
    ws.add_conditional_format(r0, c0, r1, c1, &cf)?;
    Ok(())
}

fn emit_cell(
    ws: &mut rust_xlsxwriter::Worksheet,
    r0: u32, c0: u16, r1: u32, c1: u16,
    rule: ConditionalFormatCellRule<f64>,
    fmt: Format,
) -> anyhow::Result<()> {
    let cf = ConditionalFormatCell::new().set_rule(rule).set_format(fmt);
    ws.add_conditional_format(r0, c0, r1, c1, &cf)?;
    Ok(())
}
