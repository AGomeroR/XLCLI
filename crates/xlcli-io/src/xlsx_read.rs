use std::collections::HashMap;
use std::io::Read;
use std::path::Path;

use calamine::{open_workbook_auto, Data, Reader};
use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;
use xlcli_core::cell::{Cell, CellValue};
use xlcli_core::condfmt::{CondRule, Condition, StyleOverlay};
use xlcli_core::range::CellRange;
use xlcli_core::sheet::Sheet;
use xlcli_core::style::Color;
use xlcli_core::types::CellAddr;
use xlcli_core::workbook::Workbook;

use crate::reader::FileReader;

pub struct XlsxReader;

impl FileReader for XlsxReader {
    fn read(&self, path: &Path) -> anyhow::Result<Workbook> {
        let mut xl = open_workbook_auto(path)?;
        let sheet_names: Vec<String> = xl.sheet_names().to_vec();

        let mut wb = Workbook::new();
        wb.sheets.clear();
        wb.file_path = Some(path.display().to_string());

        for name in &sheet_names {
            let mut sheet = Sheet::new(name.clone());

            if let Ok(range) = xl.worksheet_range(name) {
                for (row_idx, row) in range.rows().enumerate() {
                    for (col_idx, cell_data) in row.iter().enumerate() {
                        let value = convert_calamine_data(cell_data);
                        if !value.is_empty() {
                            sheet.set_cell(
                                row_idx as u32,
                                col_idx as u16,
                                Cell::new(value),
                            );
                        }
                    }
                }
            }

            wb.sheets.push(sheet);
        }

        if wb.sheets.is_empty() {
            wb.sheets.push(Sheet::new("Sheet1"));
        }

        if path.extension().and_then(|s| s.to_str()) == Some("xlsx") {
            let _ = load_cond_rules(path, &mut wb);
            let _ = load_cell_styles(path, &mut wb);
        }

        Ok(wb)
    }

    fn extensions(&self) -> &[&str] {
        &["xlsx", "xls", "ods"]
    }
}

fn convert_calamine_data(data: &Data) -> CellValue {
    match data {
        Data::Empty => CellValue::Empty,
        Data::String(s) => CellValue::String(s.as_str().into()),
        Data::Float(f) => CellValue::Number(*f),
        Data::Int(i) => CellValue::Number(*i as f64),
        Data::Bool(b) => CellValue::Boolean(*b),
        Data::DateTime(dt) => {
            if let Some(naive) = excel_datetime_to_naive(dt) {
                CellValue::DateTime(naive)
            } else {
                CellValue::String(format!("{:?}", dt).into())
            }
        }
        Data::Error(e) => CellValue::Error(convert_calamine_error(e)),
        Data::DateTimeIso(s) => {
            if let Ok(dt) = chrono::NaiveDateTime::parse_from_str(s, "%Y-%m-%dT%H:%M:%S") {
                CellValue::DateTime(dt)
            } else {
                CellValue::String(s.as_str().into())
            }
        }
        Data::DurationIso(s) => CellValue::String(s.as_str().into()),
    }
}

fn excel_datetime_to_naive(dt: &calamine::ExcelDateTime) -> Option<chrono::NaiveDateTime> {
    dt.as_datetime()
}

fn convert_calamine_error(err: &calamine::CellErrorType) -> xlcli_core::types::CellError {
    use calamine::CellErrorType;
    use xlcli_core::types::CellError;
    match err {
        CellErrorType::Div0 => CellError::Div0,
        CellErrorType::NA => CellError::Na,
        CellErrorType::Name => CellError::Name,
        CellErrorType::Null => CellError::Null,
        CellErrorType::Num => CellError::Num,
        CellErrorType::Ref => CellError::Ref,
        CellErrorType::Value => CellError::Value,
        CellErrorType::GettingData => CellError::GettingData,
    }
}

// ---------- Conditional format round-trip load ----------

fn load_cond_rules(path: &Path, wb: &mut Workbook) -> anyhow::Result<usize> {
    let file = std::fs::File::open(path)?;
    let mut zip = zip::ZipArchive::new(file)?;

    let dxfs = read_dxfs(&mut zip).unwrap_or_default();
    let names_in_order = read_workbook_sheets(&mut zip).unwrap_or_default();
    let rels = read_workbook_rels(&mut zip).unwrap_or_default();

    let mut total = 0usize;
    let mut per_sheet: Vec<String> = Vec::new();
    let mut seen_raw = 0usize;
    let mut skipped = 0usize;

    for (sheet_idx, sheet) in wb.sheets.iter_mut().enumerate() {
        let Some((_, rid)) = names_in_order.iter().find(|(n, _)| n == &sheet.name) else {
            per_sheet.push(format!("{}:?no-rid", sheet.name));
            continue;
        };
        let Some(target) = rels.get(rid) else {
            per_sheet.push(format!("{}:?no-target", sheet.name));
            continue;
        };
        let full = if target.starts_with('/') {
            target.trim_start_matches('/').to_string()
        } else {
            format!("xl/{}", target)
        };
        match read_sheet_cond_formats(&mut zip, &full, &dxfs, sheet_idx as u16) {
            Ok((rules, raw, skip)) => {
                seen_raw += raw;
                skipped += skip;
                total += rules.len();
                per_sheet.push(format!("{}:{}", sheet.name, rules.len()));
                sheet.cond_rules.extend(rules);
            }
            Err(_) => {
                per_sheet.push(format!("{}:err", sheet.name));
            }
        }
    }
    // Stash diagnostic into workbook via lightweight hack: use load_diagnostic in caller.
    // Here we return total; caller composes message. But we need skipped too.
    // Encode into load_diagnostic directly.
    wb.load_diagnostic = Some(format!(
        "CF load: dxfs={} sheets=[{}] rules_ok={} raw={} skipped={}",
        dxfs.len(), per_sheet.join(","), total, seen_raw, skipped
    ));
    Ok(total)
}

fn read_zip_to_string(zip: &mut zip::ZipArchive<std::fs::File>, name: &str) -> anyhow::Result<String> {
    let mut f = zip.by_name(name)?;
    let mut s = String::new();
    f.read_to_string(&mut s)?;
    Ok(s)
}

fn read_workbook_sheets(
    zip: &mut zip::ZipArchive<std::fs::File>,
) -> anyhow::Result<Vec<(String, String)>> {
    let xml = read_zip_to_string(zip, "xl/workbook.xml")?;
    let mut r = XmlReader::from_str(&xml);
    r.config_mut().trim_text(true);
    let mut out = Vec::new();
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf)? {
            Event::Empty(e) | Event::Start(e) if e.local_name().as_ref() == b"sheet" => {
                let mut name = String::new();
                let mut rid = String::new();
                for a in e.attributes().flatten() {
                    let key = a.key.local_name();
                    let val = a.unescape_value().unwrap_or_default().to_string();
                    match key.as_ref() {
                        b"name" => name = val,
                        b"id" => rid = val,
                        _ => {}
                    }
                }
                if !name.is_empty() && !rid.is_empty() {
                    out.push((name, rid));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

fn read_workbook_rels(
    zip: &mut zip::ZipArchive<std::fs::File>,
) -> anyhow::Result<HashMap<String, String>> {
    let xml = read_zip_to_string(zip, "xl/_rels/workbook.xml.rels")?;
    let mut r = XmlReader::from_str(&xml);
    r.config_mut().trim_text(true);
    let mut out = HashMap::new();
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf)? {
            Event::Empty(e) | Event::Start(e) if e.local_name().as_ref() == b"Relationship" => {
                let mut id = String::new();
                let mut target = String::new();
                for a in e.attributes().flatten() {
                    let key = a.key.local_name();
                    let val = a.unescape_value().unwrap_or_default().to_string();
                    match key.as_ref() {
                        b"Id" => id = val,
                        b"Target" => target = val,
                        _ => {}
                    }
                }
                if !id.is_empty() && !target.is_empty() {
                    out.insert(id, target);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

#[derive(Debug, Default, Clone)]
struct Dxf {
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    double_underline: Option<bool>,
    strikethrough: Option<bool>,
    fg_color: Option<Color>,
    bg_color: Option<Color>,
}

impl Dxf {
    fn to_overlay(&self) -> StyleOverlay {
        StyleOverlay {
            bold: self.bold,
            italic: self.italic,
            underline: self.underline,
            double_underline: self.double_underline,
            strikethrough: self.strikethrough,
            fg_color: self.fg_color.map(Some),
            bg_color: self.bg_color.map(Some),
        }
    }
}

fn read_dxfs(zip: &mut zip::ZipArchive<std::fs::File>) -> anyhow::Result<Vec<Dxf>> {
    let xml = read_zip_to_string(zip, "xl/styles.xml")?;
    let mut r = XmlReader::from_str(&xml);
    r.config_mut().trim_text(true);
    let mut out = Vec::new();
    let mut buf = Vec::new();
    let mut in_dxfs = false;
    let mut cur: Option<Dxf> = None;
    let mut in_font = false;
    let mut in_fill = false;

    loop {
        match r.read_event_into(&mut buf)? {
            Event::Start(e) => {
                let n = e.local_name();
                match n.as_ref() {
                    b"dxfs" => in_dxfs = true,
                    b"dxf" if in_dxfs => cur = Some(Dxf::default()),
                    b"font" if cur.is_some() => in_font = true,
                    b"fill" if cur.is_some() => in_fill = true,
                    _ => {}
                }
            }
            Event::Empty(e) => {
                let n = e.local_name();
                let Some(d) = cur.as_mut() else { buf.clear(); continue };
                match n.as_ref() {
                    b"b" if in_font => d.bold = Some(true),
                    b"i" if in_font => d.italic = Some(true),
                    b"strike" if in_font => d.strikethrough = Some(true),
                    b"u" if in_font => {
                        let mut is_double = false;
                        for a in e.attributes().flatten() {
                            if a.key.local_name().as_ref() == b"val" {
                                let v = a.unescape_value().unwrap_or_default().to_string();
                                if v == "double" || v == "doubleAccounting" {
                                    is_double = true;
                                }
                            }
                        }
                        if is_double { d.double_underline = Some(true); }
                        else { d.underline = Some(true); }
                    }
                    b"color" if in_font => {
                        if let Some(c) = parse_color_attr(&e) {
                            d.fg_color = Some(c);
                        }
                    }
                    b"bgColor" if in_fill => {
                        if let Some(c) = parse_color_attr(&e) {
                            d.bg_color = Some(c);
                        }
                    }
                    b"fgColor" if in_fill => {
                        if let Some(c) = parse_color_attr(&e) {
                            // Excel stores pattern fill solid color in fgColor
                            d.bg_color = Some(c);
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let n = e.local_name();
                match n.as_ref() {
                    b"dxfs" => in_dxfs = false,
                    b"dxf" => {
                        if let Some(d) = cur.take() { out.push(d); }
                    }
                    b"font" => in_font = false,
                    b"fill" => in_fill = false,
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

fn parse_color_attr(e: &quick_xml::events::BytesStart) -> Option<Color> {
    let mut rgb: Option<String> = None;
    let mut theme: Option<u32> = None;
    let mut indexed: Option<u32> = None;
    let mut tint: f64 = 0.0;
    let mut auto = false;
    for a in e.attributes().flatten() {
        let v = a.unescape_value().ok()?.to_string();
        match a.key.local_name().as_ref() {
            b"rgb" => rgb = Some(v),
            b"theme" => theme = v.parse().ok(),
            b"indexed" => indexed = v.parse().ok(),
            b"tint" => tint = v.parse().unwrap_or(0.0),
            b"auto" => auto = v == "1" || v.eq_ignore_ascii_case("true"),
            _ => {}
        }
    }
    if auto { return None; }
    // indexed 64/65 = System Foreground / Background (automatic)
    if matches!(indexed, Some(64) | Some(65)) { return None; }
    // theme 0 (lt1) and 1 (dk1) are system-mapped window fg/bg — treat as automatic
    if matches!(theme, Some(0) | Some(1)) && tint == 0.0 { return None; }
    let base = if let Some(s) = rgb {
        parse_argb(&s)?
    } else if let Some(t) = theme {
        theme_color(t)?
    } else if let Some(i) = indexed {
        indexed_color(i)?
    } else {
        return None;
    };
    Some(apply_tint(base, tint))
}

fn apply_tint(c: Color, tint: f64) -> Color {
    if tint == 0.0 { return c; }
    let f = |ch: u8| -> u8 {
        let v = ch as f64;
        let nv = if tint < 0.0 { v * (1.0 + tint) } else { v + (255.0 - v) * tint };
        nv.clamp(0.0, 255.0) as u8
    };
    Color::new(f(c.r), f(c.g), f(c.b))
}

fn theme_color(idx: u32) -> Option<Color> {
    // Default Office theme palette (0=lt1, 1=dk1, 2=lt2, 3=dk2, 4=accent1..9=accent6, 10=hlink, 11=fhlink)
    let rgb = match idx {
        0 => 0xFFFFFF, 1 => 0x000000, 2 => 0xE7E6E6, 3 => 0x44546A,
        4 => 0x4472C4, 5 => 0xED7D31, 6 => 0xA5A5A5, 7 => 0xFFC000,
        8 => 0x5B9BD5, 9 => 0x70AD47, 10 => 0x0563C1, 11 => 0x954F72,
        _ => return None,
    };
    Some(Color::new(((rgb >> 16) & 0xFF) as u8, ((rgb >> 8) & 0xFF) as u8, (rgb & 0xFF) as u8))
}

fn indexed_color(idx: u32) -> Option<Color> {
    // Legacy Excel 64-color indexed palette (abbreviated to common ones; 64/65 reserved for auto)
    let rgb: u32 = match idx {
        0 => 0x000000, 1 => 0xFFFFFF, 2 => 0xFF0000, 3 => 0x00FF00, 4 => 0x0000FF,
        5 => 0xFFFF00, 6 => 0xFF00FF, 7 => 0x00FFFF, 8 => 0x000000, 9 => 0xFFFFFF,
        10 => 0xFF0000, 11 => 0x00FF00, 12 => 0x0000FF, 13 => 0xFFFF00, 14 => 0xFF00FF,
        15 => 0x00FFFF, 16 => 0x800000, 17 => 0x008000, 18 => 0x000080, 19 => 0x808000,
        20 => 0x800080, 21 => 0x008080, 22 => 0xC0C0C0, 23 => 0x808080, 24 => 0x9999FF,
        25 => 0x993366, 26 => 0xFFFFCC, 27 => 0xCCFFFF, 28 => 0x660066, 29 => 0xFF8080,
        30 => 0x0066CC, 31 => 0xCCCCFF, 32 => 0x000080, 33 => 0xFF00FF, 34 => 0xFFFF00,
        35 => 0x00FFFF, 36 => 0x800080, 37 => 0x800000, 38 => 0x008080, 39 => 0x0000FF,
        40 => 0x00CCFF, 41 => 0xCCFFFF, 42 => 0xCCFFCC, 43 => 0xFFFF99, 44 => 0x99CCFF,
        45 => 0xFF99CC, 46 => 0xCC99FF, 47 => 0xFFCC99, 48 => 0x3366FF, 49 => 0x33CCCC,
        50 => 0x99CC00, 51 => 0xFFCC00, 52 => 0xFF9900, 53 => 0xFF6600, 54 => 0x666699,
        55 => 0x969696, 56 => 0x003366, 57 => 0x339966, 58 => 0x003300, 59 => 0x333300,
        60 => 0x993300, 61 => 0x993366, 62 => 0x333399, 63 => 0x333333,
        _ => return None,
    };
    Some(Color::new(((rgb >> 16) & 0xFF) as u8, ((rgb >> 8) & 0xFF) as u8, (rgb & 0xFF) as u8))
}

fn parse_argb(s: &str) -> Option<Color> {
    let s = s.trim();
    let hex = if s.len() == 8 { &s[2..] } else if s.len() == 6 { s } else { return None };
    let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
    Some(Color::new(r, g, b))
}

fn read_sheet_cond_formats(
    zip: &mut zip::ZipArchive<std::fs::File>,
    path: &str,
    dxfs: &[Dxf],
    sheet_idx: u16,
) -> anyhow::Result<(Vec<CondRule>, usize, usize)> {
    let xml = read_zip_to_string(zip, path)?;
    let mut r = XmlReader::from_str(&xml);
    r.config_mut().trim_text(true);
    let mut out = Vec::new();
    let mut raw_total = 0usize;
    let mut skipped = 0usize;
    let mut buf = Vec::new();

    let mut cur_sqref: Option<String> = None;
    let mut cur_rule: Option<PendingRule> = None;
    let mut in_formula = false;
    let mut formulas: Vec<String> = Vec::new();

    loop {
        match r.read_event_into(&mut buf)? {
            Event::Start(e) => {
                match e.local_name().as_ref() {
                    b"conditionalFormatting" => {
                        cur_sqref = attr(&e, b"sqref");
                    }
                    b"cfRule" => {
                        cur_rule = Some(PendingRule {
                            ty: attr(&e, b"type").unwrap_or_default(),
                            operator: attr(&e, b"operator").unwrap_or_default(),
                            dxf_id: attr(&e, b"dxfId").and_then(|s| s.parse().ok()),
                            text: attr(&e, b"text").unwrap_or_default(),
                        });
                        formulas.clear();
                    }
                    b"formula" => in_formula = true,
                    _ => {}
                }
            }
            Event::Empty(e) => {
                if e.local_name().as_ref() == b"cfRule" {
                    raw_total += 1;
                    let pending = PendingRule {
                        ty: attr(&e, b"type").unwrap_or_default(),
                        operator: attr(&e, b"operator").unwrap_or_default(),
                        dxf_id: attr(&e, b"dxfId").and_then(|s| s.parse().ok()),
                        text: attr(&e, b"text").unwrap_or_default(),
                    };
                    if let Some(sqref) = &cur_sqref {
                        match build_rule(&pending, &[], sqref, dxfs, sheet_idx) {
                            Some(rule) => out.push(rule),
                            None => skipped += 1,
                        }
                    } else {
                        skipped += 1;
                    }
                }
            }
            Event::Text(t) if in_formula => {
                formulas.push(t.unescape().unwrap_or_default().to_string());
            }
            Event::End(e) => {
                match e.local_name().as_ref() {
                    b"formula" => in_formula = false,
                    b"cfRule" => {
                        raw_total += 1;
                        if let (Some(pending), Some(sqref)) = (cur_rule.take(), cur_sqref.as_ref()) {
                            match build_rule(&pending, &formulas, sqref, dxfs, sheet_idx) {
                                Some(rule) => out.push(rule),
                                None => skipped += 1,
                            }
                        } else {
                            skipped += 1;
                        }
                        formulas.clear();
                    }
                    b"conditionalFormatting" => {
                        cur_sqref = None;
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok((out, raw_total, skipped))
}

// ---------- Per-cell style round-trip load ----------

#[derive(Debug, Clone, Default)]
struct XlsxFont {
    bold: bool,
    italic: bool,
    underline: bool,
    double_underline: bool,
    strikethrough: bool,
    color: Option<Color>,
}

#[derive(Debug, Clone, Default)]
struct XlsxFill {
    bg: Option<Color>,
}

#[derive(Debug, Clone, Default)]
struct XlsxXf {
    font_id: usize,
    fill_id: usize,
    apply_font: bool,
    apply_fill: bool,
}

fn load_cell_styles(path: &Path, wb: &mut Workbook) -> anyhow::Result<()> {
    let file = std::fs::File::open(path)?;
    let mut zip = zip::ZipArchive::new(file)?;
    let styles_xml = match read_zip_to_string(&mut zip, "xl/styles.xml") {
        Ok(s) => s,
        Err(_) => return Ok(()),
    };
    let (fonts, fills, xfs) = parse_styles_xml(&styles_xml).unwrap_or_default();

    // Map xlsx xf index -> xlcli style pool id
    let mut xf_to_pool: Vec<u32> = Vec::with_capacity(xfs.len());
    for xf in &xfs {
        let mut cs = xlcli_core::style::CellStyle::default();
        if xf.apply_font || xf.font_id != 0 {
            if let Some(f) = fonts.get(xf.font_id) {
                cs.bold = f.bold;
                cs.italic = f.italic;
                cs.underline = f.underline;
                cs.double_underline = f.double_underline;
                cs.strikethrough = f.strikethrough;
                cs.fg_color = f.color;
            }
        }
        if xf.apply_fill || xf.fill_id > 1 {
            if let Some(fl) = fills.get(xf.fill_id) {
                cs.bg_color = fl.bg;
            }
        }
        let id = wb.style_pool.get_or_insert(cs);
        xf_to_pool.push(id);
    }

    let names_in_order = read_workbook_sheets(&mut zip).unwrap_or_default();
    let rels = read_workbook_rels(&mut zip).unwrap_or_default();

    let mut styled_cells = 0usize;
    for sheet in wb.sheets.iter_mut() {
        let Some((_, rid)) = names_in_order.iter().find(|(n, _)| n == &sheet.name) else { continue };
        let Some(target) = rels.get(rid) else { continue };
        let full = if target.starts_with('/') {
            target.trim_start_matches('/').to_string()
        } else {
            format!("xl/{}", target)
        };
        let Ok(xml) = read_zip_to_string(&mut zip, &full) else { continue };
        styled_cells += apply_cell_styles(sheet, &xml, &xf_to_pool);
    }

    let prev = wb.load_diagnostic.clone().unwrap_or_default();
    wb.load_diagnostic = Some(format!(
        "{}{}fonts={} fills={} xfs={} styled_cells={}",
        prev, if prev.is_empty() { "" } else { " | " },
        fonts.len(), fills.len(), xfs.len(), styled_cells));
    Ok(())
}

fn parse_styles_xml(xml: &str) -> Option<(Vec<XlsxFont>, Vec<XlsxFill>, Vec<XlsxXf>)> {
    let mut r = XmlReader::from_str(xml);
    r.config_mut().trim_text(true);
    let mut fonts: Vec<XlsxFont> = Vec::new();
    let mut fills: Vec<XlsxFill> = Vec::new();
    let mut xfs: Vec<XlsxXf> = Vec::new();
    let mut buf = Vec::new();

    let mut in_fonts = false;
    let mut in_fills = false;
    let mut in_cellxfs = false;
    let mut cur_font: Option<XlsxFont> = None;
    let mut cur_fill: Option<XlsxFill> = None;
    let mut in_pattern_fill = false;

    loop {
        let ev = r.read_event_into(&mut buf).ok()?;
        match ev {
            Event::Start(e) => {
                match e.local_name().as_ref() {
                    b"fonts" => in_fonts = true,
                    b"fills" => in_fills = true,
                    b"cellXfs" => in_cellxfs = true,
                    b"font" if in_fonts => cur_font = Some(XlsxFont::default()),
                    b"fill" if in_fills => cur_fill = Some(XlsxFill::default()),
                    b"patternFill" if cur_fill.is_some() => in_pattern_fill = true,
                    b"xf" if in_cellxfs => {
                        xfs.push(parse_xf_attrs(&e));
                    }
                    _ => {}
                }
            }
            Event::Empty(e) => {
                let n = e.local_name();
                match n.as_ref() {
                    b"font" if in_fonts => fonts.push(XlsxFont::default()),
                    b"fill" if in_fills => fills.push(XlsxFill::default()),
                    b"xf" if in_cellxfs => xfs.push(parse_xf_attrs(&e)),
                    b"b" if cur_font.is_some() => { cur_font.as_mut().unwrap().bold = true; }
                    b"i" if cur_font.is_some() => { cur_font.as_mut().unwrap().italic = true; }
                    b"strike" if cur_font.is_some() => { cur_font.as_mut().unwrap().strikethrough = true; }
                    b"u" if cur_font.is_some() => {
                        let mut dbl = false;
                        for a in e.attributes().flatten() {
                            if a.key.local_name().as_ref() == b"val" {
                                let v = a.unescape_value().unwrap_or_default().to_string();
                                if v == "double" || v == "doubleAccounting" { dbl = true; }
                            }
                        }
                        let f = cur_font.as_mut().unwrap();
                        if dbl { f.double_underline = true; } else { f.underline = true; }
                    }
                    b"color" if cur_font.is_some() => {
                        if let Some(c) = parse_color_attr(&e) {
                            cur_font.as_mut().unwrap().color = Some(c);
                        }
                    }
                    b"fgColor" if in_pattern_fill => {
                        if let Some(c) = parse_color_attr(&e) {
                            cur_fill.as_mut().unwrap().bg = Some(c);
                        }
                    }
                    b"bgColor" if in_pattern_fill => {
                        if cur_fill.as_ref().map_or(true, |f| f.bg.is_none()) {
                            if let Some(c) = parse_color_attr(&e) {
                                cur_fill.as_mut().unwrap().bg = Some(c);
                            }
                        }
                    }
                    b"patternFill" if cur_fill.is_some() => {
                        // inline patternFill with attrs only; check patternType
                        let mut solid = false;
                        for a in e.attributes().flatten() {
                            if a.key.local_name().as_ref() == b"patternType" {
                                let v = a.unescape_value().unwrap_or_default().to_string();
                                if v == "solid" { solid = true; }
                            }
                        }
                        if !solid {
                            cur_fill.as_mut().unwrap().bg = None;
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                match e.local_name().as_ref() {
                    b"fonts" => in_fonts = false,
                    b"fills" => in_fills = false,
                    b"cellXfs" => in_cellxfs = false,
                    b"font" => { if let Some(f) = cur_font.take() { fonts.push(f); } }
                    b"fill" => { if let Some(f) = cur_fill.take() { fills.push(f); } }
                    b"patternFill" => in_pattern_fill = false,
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Some((fonts, fills, xfs))
}

fn parse_xf_attrs(e: &quick_xml::events::BytesStart) -> XlsxXf {
    let mut xf = XlsxXf::default();
    for a in e.attributes().flatten() {
        let val = a.unescape_value().unwrap_or_default().to_string();
        match a.key.local_name().as_ref() {
            b"fontId" => xf.font_id = val.parse().unwrap_or(0),
            b"fillId" => xf.fill_id = val.parse().unwrap_or(0),
            b"applyFont" => xf.apply_font = val == "1" || val.eq_ignore_ascii_case("true"),
            b"applyFill" => xf.apply_fill = val == "1" || val.eq_ignore_ascii_case("true"),
            _ => {}
        }
    }
    xf
}

fn apply_cell_styles(sheet: &mut Sheet, xml: &str, xf_to_pool: &[u32]) -> usize {
    let mut r = XmlReader::from_str(xml);
    r.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut count = 0usize;
    loop {
        let ev = match r.read_event_into(&mut buf) { Ok(e) => e, Err(_) => break };
        match ev {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"c" => {
                let mut ref_addr: Option<String> = None;
                let mut s_idx: Option<usize> = None;
                for a in e.attributes().flatten() {
                    let val = a.unescape_value().unwrap_or_default().to_string();
                    match a.key.local_name().as_ref() {
                        b"r" => ref_addr = Some(val),
                        b"s" => s_idx = val.parse().ok(),
                        _ => {}
                    }
                }
                if let (Some(addr), Some(s)) = (ref_addr, s_idx) {
                    if s == 0 { buf.clear(); continue; }
                    if let Some(pool_id) = xf_to_pool.get(s).copied() {
                        if let Some((row, col)) = parse_a1(&addr) {
                            // Ensure cell exists before tagging style (may be empty row).
                            if sheet.get_cell(row, col).is_none() {
                                sheet.set_cell(row, col, xlcli_core::cell::Cell::new(xlcli_core::cell::CellValue::Empty));
                            }
                            // Mutate via direct hashmap access: use helper
                            if set_cell_style(sheet, row, col, pool_id) {
                                count += 1;
                            }
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    count
}

fn set_cell_style(sheet: &mut Sheet, row: u32, col: u16, style_id: u32) -> bool {
    if let Some(cell) = sheet.get_cell(row, col).cloned() {
        let mut c = cell;
        c.style_id = style_id;
        sheet.set_cell(row, col, c);
        return true;
    }
    false
}

struct PendingRule {
    ty: String,
    operator: String,
    dxf_id: Option<usize>,
    text: String,
}

fn attr(e: &quick_xml::events::BytesStart, key: &[u8]) -> Option<String> {
    for a in e.attributes().flatten() {
        if a.key.local_name().as_ref() == key {
            return a.unescape_value().ok().map(|v| v.to_string());
        }
    }
    None
}

fn build_rule(
    p: &PendingRule,
    formulas: &[String],
    sqref: &str,
    dxfs: &[Dxf],
    sheet_idx: u16,
) -> Option<CondRule> {
    let range = parse_sqref(sqref, sheet_idx)?;
    let parse_num = |s: &str| -> Option<f64> {
        s.trim().trim_start_matches('=').trim().parse().ok()
    };
    let cond = match p.ty.as_str() {
        "cellIs" => {
            let f0 = formulas.get(0).and_then(|s| parse_num(s))?;
            match p.operator.as_str() {
                "greaterThan" => Condition::Gt(f0),
                "lessThan" => Condition::Lt(f0),
                "greaterThanOrEqual" => Condition::Gte(f0),
                "lessThanOrEqual" => Condition::Lte(f0),
                "equal" => Condition::Eq(f0),
                "notEqual" => Condition::Neq(f0),
                "between" => {
                    let f1 = formulas.get(1).and_then(|s| parse_num(s))?;
                    Condition::Between(f0, f1)
                }
                "notBetween" => {
                    // Map notBetween to no-op equivalent; skip
                    return None;
                }
                _ => return None,
            }
        }
        "containsText" => {
            let s = if !p.text.is_empty() { p.text.clone() }
                    else { extract_contains_text(formulas.get(0)?)? };
            Condition::Contains(s)
        }
        "containsBlanks" => Condition::Blanks,
        "notContainsBlanks" => Condition::NonBlanks,
        "expression" => {
            let f = formulas.get(0).map(|s| s.trim().to_ascii_uppercase());
            if f.as_deref() == Some("TRUE") {
                Condition::Always
            } else {
                return None;
            }
        }
        _ => return None,
    };
    let style = p.dxf_id
        .and_then(|id| dxfs.get(id))
        .map(|d| d.to_overlay())
        .unwrap_or_default();
    Some(CondRule { range, cond, style: xlcli_core::condfmt::StyleSpec::Overlay(style) })
}

fn extract_contains_text(formula: &str) -> Option<String> {
    // e.g. NOT(ISERROR(SEARCH("foo",A1)))
    let start = formula.find('"')?;
    let rest = &formula[start + 1..];
    let end = rest.find('"')?;
    Some(rest[..end].to_string())
}

fn parse_sqref(sqref: &str, sheet: u16) -> Option<CellRange> {
    let first = sqref.split_whitespace().next()?;
    let (s, e) = if let Some((a, b)) = first.split_once(':') {
        (a, b)
    } else {
        (first, first)
    };
    let (r0, c0) = parse_a1(s)?;
    let (r1, c1) = parse_a1(e)?;
    Some(CellRange::new(
        CellAddr::new(sheet, r0, c0),
        CellAddr::new(sheet, r1, c1),
    ))
}

fn parse_a1(s: &str) -> Option<(u32, u16)> {
    let bytes = s.as_bytes();
    let mut i = 0;
    let mut col: u32 = 0;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        col = col * 26 + (bytes[i].to_ascii_uppercase() - b'A' + 1) as u32;
        i += 1;
    }
    if col == 0 || i == 0 || i == bytes.len() {
        return None;
    }
    let row: u32 = std::str::from_utf8(&bytes[i..]).ok()?.parse().ok()?;
    if row == 0 { return None; }
    Some((row - 1, (col - 1) as u16))
}
