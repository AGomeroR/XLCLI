use std::collections::{HashMap, HashSet};

use crate::cell::{Cell, CellValue};
use crate::condfmt::{
    CfValueKind, ColorStop, CondRule, Condition, IconSetKind, IconThreshold, StyleOverlay, StyleSpec,
    TimePeriod,
};
use crate::style::Color;

#[derive(Debug, Clone, PartialEq)]
pub enum FilterCondition {
    Eq(String),
    NotEq(String),
    Gt(f64),
    Lt(f64),
    Gte(f64),
    Lte(f64),
    Contains(String),
    Blanks,
    NonBlanks,
    TopN(usize),
    BottomN(usize),
    ValueSet(HashSet<String>),
}

#[derive(Debug)]
pub struct Sheet {
    pub name: String,
    cells: HashMap<(u32, u16), Cell>,
    col_widths: HashMap<u16, f64>,
    row_heights: HashMap<u32, f64>,
    extent: (u32, u16),
    pub freeze: Option<(u32, u16)>,
    pub hidden_rows: HashSet<u32>,
    pub hidden_cols: HashSet<u16>,
    pub header_row: Option<u32>,
    pub filters: HashMap<u16, FilterCondition>,
    pub filter_range: Option<(u32, u32, u16, u16)>,
    pub base_style: StyleOverlay,
    pub cond_rules: Vec<CondRule>,
}

impl Sheet {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            cells: HashMap::new(),
            col_widths: HashMap::new(),
            row_heights: HashMap::new(),
            extent: (0, 0),
            freeze: None,
            hidden_rows: HashSet::new(),
            hidden_cols: HashSet::new(),
            header_row: None,
            filters: HashMap::new(),
            filter_range: None,
            base_style: StyleOverlay::default(),
            cond_rules: Vec::new(),
        }
    }

    pub fn effective_style(&self, row: u32, col: u16, base: &crate::style::CellStyle) -> crate::style::CellStyle {
        let mut s = base.clone();
        self.base_style.apply(&mut s);
        let val = self.get_cell_value(row, col);
        for rule in &self.cond_rules {
            let r = &rule.range;
            let in_range = row >= r.start.row && row <= r.end.row
                && col >= r.start.col && col <= r.end.col;
            if !in_range { continue; }

            let matches = if rule.cond.needs_range_context() {
                self.eval_contextual(&rule.cond, &rule.range, row, col, val)
            } else {
                rule.cond.matches(val)
            };

            match &rule.style {
                StyleSpec::Overlay(o) => {
                    if matches { o.apply(&mut s); }
                }
                StyleSpec::ColorScale(stops) => {
                    if let Some(n) = cell_num_value(val) {
                        if let Some(c) = color_scale_at(stops, &rule.range, self, n) {
                            s.bg_color = c;
                        }
                    }
                }
                StyleSpec::DataBar { min, max, color } => {
                    if let Some(n) = cell_num_value(val) {
                        if let Some(c) = data_bar_tint(*color, min, max, &rule.range, self, n) {
                            s.bg_color = c;
                        }
                    }
                }
                StyleSpec::IconSet { .. } => {
                    // Handled by icon_for_cell for render text-prefix, not style.
                }
            }
        }
        s
    }

    /// Returns the unicode icon char to prefix the display value for any IconSet
    /// rule covering (row,col), or None.
    pub fn icon_for_cell(&self, row: u32, col: u16) -> Option<&'static str> {
        let val = self.get_cell_value(row, col);
        let n = cell_num_value(val)?;
        for rule in &self.cond_rules {
            let r = &rule.range;
            if row < r.start.row || row > r.end.row || col < r.start.col || col > r.end.col { continue; }
            if let StyleSpec::IconSet { kind, thresholds, reverse, .. } = &rule.style {
                return Some(icon_for(*kind, thresholds, *reverse, &rule.range, self, n));
            }
        }
        None
    }

    fn eval_contextual(&self, cond: &Condition, range: &crate::range::CellRange, row: u32, col: u16, val: &CellValue) -> bool {
        match cond {
            Condition::DuplicateValues | Condition::UniqueValues => {
                let disp = val.display_value();
                if disp.is_empty() { return false; }
                let mut count = 0usize;
                for r in range.start.row..=range.end.row {
                    for c in range.start.col..=range.end.col {
                        let v = self.get_cell_value(r, c);
                        if !v.is_empty() && v.display_value() == disp {
                            count += 1;
                            if count > 1 { break; }
                        }
                    }
                    if count > 1 { break; }
                }
                match cond {
                    Condition::DuplicateValues => count > 1,
                    Condition::UniqueValues => count == 1,
                    _ => false,
                }
            }
            Condition::Top { count, percent, bottom } => {
                let n = match cell_num_value(val) { Some(v) => v, None => return false };
                let mut vals: Vec<f64> = Vec::new();
                for r in range.start.row..=range.end.row {
                    for c in range.start.col..=range.end.col {
                        if let Some(v) = cell_num_value(self.get_cell_value(r, c)) {
                            vals.push(v);
                        }
                    }
                }
                if vals.is_empty() { return false; }
                if *bottom {
                    vals.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
                } else {
                    vals.sort_by(|a, b| b.partial_cmp(a).unwrap_or(std::cmp::Ordering::Equal));
                }
                let k = if *percent {
                    (((*count as f64) / 100.0) * (vals.len() as f64)).ceil() as usize
                } else {
                    *count as usize
                };
                let k = k.min(vals.len()).max(1);
                let cutoff = vals[k - 1];
                if *bottom { n <= cutoff } else { n >= cutoff }
            }
            Condition::Average { above, equal, std_dev } => {
                let n = match cell_num_value(val) { Some(v) => v, None => return false };
                let mut sum = 0.0;
                let mut count = 0.0;
                for r in range.start.row..=range.end.row {
                    for c in range.start.col..=range.end.col {
                        if let Some(v) = cell_num_value(self.get_cell_value(r, c)) {
                            sum += v;
                            count += 1.0;
                        }
                    }
                }
                if count == 0.0 { return false; }
                let avg = sum / count;
                let threshold = if *std_dev != 0 {
                    // compute stddev
                    let mut sq = 0.0;
                    for r in range.start.row..=range.end.row {
                        for c in range.start.col..=range.end.col {
                            if let Some(v) = cell_num_value(self.get_cell_value(r, c)) {
                                sq += (v - avg).powi(2);
                            }
                        }
                    }
                    let sd = (sq / count).sqrt();
                    avg + (*std_dev as f64) * sd
                } else {
                    avg
                };
                if *above {
                    if *equal { n >= threshold } else { n > threshold }
                } else {
                    if *equal { n <= threshold } else { n < threshold }
                }
            }
            Condition::TimePeriod(tp) => eval_time_period(*tp, val),
            Condition::Expression(_) => false, // deferred
            _ => false,
        }
    }
}  // <-- was incorrectly closing impl Sheet; impl continues after helpers

fn cell_num_value(v: &CellValue) -> Option<f64> {
    match v {
        CellValue::Number(n) => Some(*n),
        CellValue::Boolean(b) => Some(if *b { 1.0 } else { 0.0 }),
        CellValue::DateTime(dt) => {
            let base = chrono::NaiveDate::from_ymd_opt(1899, 12, 30).unwrap().and_hms_opt(0, 0, 0).unwrap();
            Some((*dt - base).num_milliseconds() as f64 / 86_400_000.0)
        }
        _ => None,
    }
}

fn eval_time_period(tp: TimePeriod, v: &CellValue) -> bool {
    let dt = match v {
        CellValue::DateTime(d) => *d,
        _ => return false,
    };
    let today = chrono::Local::now().date_naive();
    let d = dt.date();
    use chrono::Datelike;
    use TimePeriod::*;
    let weekday = today.weekday().num_days_from_monday() as i64;
    match tp {
        Today => d == today,
        Yesterday => d == today - chrono::Duration::days(1),
        Tomorrow => d == today + chrono::Duration::days(1),
        Last7Days => {
            let diff = (today - d).num_days();
            (0..=6).contains(&diff)
        }
        ThisWeek => {
            let start = today - chrono::Duration::days(weekday);
            let end = start + chrono::Duration::days(6);
            d >= start && d <= end
        }
        LastWeek => {
            let start = today - chrono::Duration::days(weekday + 7);
            let end = start + chrono::Duration::days(6);
            d >= start && d <= end
        }
        NextWeek => {
            let start = today - chrono::Duration::days(weekday) + chrono::Duration::days(7);
            let end = start + chrono::Duration::days(6);
            d >= start && d <= end
        }
        ThisMonth => d.year() == today.year() && d.month() == today.month(),
        LastMonth => {
            let (y, m) = if today.month() == 1 { (today.year() - 1, 12) } else { (today.year(), today.month() - 1) };
            d.year() == y && d.month() == m
        }
        NextMonth => {
            let (y, m) = if today.month() == 12 { (today.year() + 1, 1) } else { (today.year(), today.month() + 1) };
            d.year() == y && d.month() == m
        }
    }
}

fn resolve_cf_value(v: &CfValueKind, range: &crate::range::CellRange, sheet: &Sheet) -> f64 {
    match v {
        CfValueKind::Number(n) => *n,
        CfValueKind::Formula(_) => 0.0,
        CfValueKind::Min => range_numeric_min(range, sheet).unwrap_or(0.0),
        CfValueKind::Max => range_numeric_max(range, sheet).unwrap_or(0.0),
        CfValueKind::Percent(p) => {
            let lo = range_numeric_min(range, sheet).unwrap_or(0.0);
            let hi = range_numeric_max(range, sheet).unwrap_or(1.0);
            lo + (hi - lo) * (p / 100.0)
        }
        CfValueKind::Percentile(p) => percentile(range, sheet, *p),
    }
}

fn range_numeric_min(range: &crate::range::CellRange, sheet: &Sheet) -> Option<f64> {
    let mut m: Option<f64> = None;
    for r in range.start.row..=range.end.row {
        for c in range.start.col..=range.end.col {
            if let Some(n) = cell_num_value(sheet.get_cell_value(r, c)) {
                m = Some(m.map_or(n, |cur| cur.min(n)));
            }
        }
    }
    m
}

fn range_numeric_max(range: &crate::range::CellRange, sheet: &Sheet) -> Option<f64> {
    let mut m: Option<f64> = None;
    for r in range.start.row..=range.end.row {
        for c in range.start.col..=range.end.col {
            if let Some(n) = cell_num_value(sheet.get_cell_value(r, c)) {
                m = Some(m.map_or(n, |cur| cur.max(n)));
            }
        }
    }
    m
}

fn percentile(range: &crate::range::CellRange, sheet: &Sheet, p: f64) -> f64 {
    let mut vals: Vec<f64> = Vec::new();
    for r in range.start.row..=range.end.row {
        for c in range.start.col..=range.end.col {
            if let Some(n) = cell_num_value(sheet.get_cell_value(r, c)) {
                vals.push(n);
            }
        }
    }
    if vals.is_empty() { return 0.0; }
    vals.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let rank = (p / 100.0) * (vals.len() - 1) as f64;
    let lo = rank.floor() as usize;
    let hi = rank.ceil() as usize;
    if lo == hi { return vals[lo]; }
    let frac = rank - lo as f64;
    vals[lo] + (vals[hi] - vals[lo]) * frac
}

fn color_scale_at(stops: &[ColorStop], range: &crate::range::CellRange, sheet: &Sheet, n: f64) -> Option<Option<Color>> {
    if stops.len() < 2 { return None; }
    let mut resolved: Vec<(f64, Color)> = stops.iter()
        .map(|s| (resolve_cf_value(&s.value, range, sheet), s.color))
        .collect();
    resolved.sort_by(|a, b| a.0.partial_cmp(&b.0).unwrap_or(std::cmp::Ordering::Equal));
    if n <= resolved[0].0 { return Some(Some(resolved[0].1)); }
    if n >= resolved.last().unwrap().0 { return Some(Some(resolved.last().unwrap().1)); }
    for w in resolved.windows(2) {
        let (lo_v, lo_c) = w[0];
        let (hi_v, hi_c) = w[1];
        if n >= lo_v && n <= hi_v {
            let t = if hi_v == lo_v { 0.0 } else { (n - lo_v) / (hi_v - lo_v) };
            return Some(Some(lerp_color(lo_c, hi_c, t)));
        }
    }
    None
}

fn lerp_color(a: Color, b: Color, t: f64) -> Color {
    let t = t.clamp(0.0, 1.0);
    Color::new(
        ((a.r as f64) + ((b.r as f64) - (a.r as f64)) * t).round() as u8,
        ((a.g as f64) + ((b.g as f64) - (a.g as f64)) * t).round() as u8,
        ((a.b as f64) + ((b.b as f64) - (a.b as f64)) * t).round() as u8,
    )
}

fn data_bar_tint(color: Color, min: &CfValueKind, max: &CfValueKind, range: &crate::range::CellRange, sheet: &Sheet, n: f64) -> Option<Option<Color>> {
    let lo = resolve_cf_value(min, range, sheet);
    let hi = resolve_cf_value(max, range, sheet);
    if hi <= lo { return None; }
    let t = ((n - lo) / (hi - lo)).clamp(0.0, 1.0);
    // Lerp white -> color by t (so bigger = stronger color)
    let white = Color::new(255, 255, 255);
    Some(Some(lerp_color(white, color, t)))
}

fn icon_for(kind: IconSetKind, thresholds: &[IconThreshold], reverse: bool, range: &crate::range::CellRange, sheet: &Sheet, n: f64) -> &'static str {
    let icons = kind.icons();
    // thresholds.len() == icons.len() - 1 typically. Bucket index selects icon.
    let mut bucket = 0usize;
    for (i, t) in thresholds.iter().enumerate() {
        let tv = resolve_cf_value(&t.value, range, sheet);
        let passes = if t.gte { n >= tv } else { n > tv };
        if passes { bucket = i + 1; }
    }
    let idx = if reverse { icons.len() - 1 - bucket } else { bucket };
    icons.get(idx).copied().unwrap_or("")
}

impl Sheet {
    pub fn get_cell(&self, row: u32, col: u16) -> Option<&Cell> {
        self.cells.get(&(row, col))
    }

    pub fn get_cell_value(&self, row: u32, col: u16) -> &CellValue {
        self.cells
            .get(&(row, col))
            .map(|c| &c.value)
            .unwrap_or(&CellValue::Empty)
    }

    pub fn set_cell(&mut self, row: u32, col: u16, cell: Cell) {
        if row >= self.extent.0 {
            self.extent.0 = row + 1;
        }
        if col >= self.extent.1 {
            self.extent.1 = col + 1;
        }
        self.cells.insert((row, col), cell);
    }

    pub fn set_value(&mut self, row: u32, col: u16, value: CellValue) {
        if value.is_empty() {
            self.cells.remove(&(row, col));
        } else {
            self.set_cell(row, col, Cell::new(value));
        }
    }

    pub fn remove_cell(&mut self, row: u32, col: u16) -> Option<Cell> {
        self.cells.remove(&(row, col))
    }

    pub fn extent(&self) -> (u32, u16) {
        self.extent
    }

    pub fn row_count(&self) -> u32 {
        self.extent.0
    }

    pub fn col_count(&self) -> u16 {
        self.extent.1
    }

    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    pub fn col_width(&self, col: u16) -> f64 {
        self.col_widths.get(&col).copied().unwrap_or(10.0)
    }

    pub fn set_col_width(&mut self, col: u16, width: f64) {
        self.col_widths.insert(col, width);
    }

    pub fn row_height(&self, row: u32) -> f64 {
        self.row_heights.get(&row).copied().unwrap_or(1.0)
    }

    pub fn set_row_height(&mut self, row: u32, height: f64) {
        self.row_heights.insert(row, height);
    }

    pub fn cells_iter(&self) -> impl Iterator<Item = (&(u32, u16), &Cell)> {
        self.cells.iter()
    }

    pub fn insert_row(&mut self, at_row: u32) {
        let mut new_cells = HashMap::new();
        for (&(r, c), cell) in &self.cells {
            if r >= at_row {
                new_cells.insert((r + 1, c), cell.clone());
            } else {
                new_cells.insert((r, c), cell.clone());
            }
        }
        self.cells = new_cells;
        self.extent.0 += 1;
    }

    pub fn delete_row(&mut self, at_row: u32) -> Vec<(u16, Cell)> {
        let mut removed = Vec::new();
        let mut new_cells = HashMap::new();
        for (&(r, c), cell) in &self.cells {
            if r == at_row {
                removed.push((c, cell.clone()));
            } else if r > at_row {
                new_cells.insert((r - 1, c), cell.clone());
            } else {
                new_cells.insert((r, c), cell.clone());
            }
        }
        self.cells = new_cells;
        if self.extent.0 > 0 {
            self.extent.0 -= 1;
        }
        removed
    }

    pub fn insert_col(&mut self, at_col: u16) {
        let mut new_cells = HashMap::new();
        for (&(r, c), cell) in &self.cells {
            if c >= at_col {
                new_cells.insert((r, c + 1), cell.clone());
            } else {
                new_cells.insert((r, c), cell.clone());
            }
        }
        self.cells = new_cells;
        self.extent.1 += 1;
    }

    pub fn apply_filters(&mut self) {
        self.hidden_rows.clear();
        if self.filters.is_empty() {
            return;
        }
        let (min_row, max_row, _min_col, _max_col) = match self.filter_range {
            Some(r) => r,
            None => return,
        };
        let header = self.header_row;
        let start_row = if let Some(h) = header {
            if h >= min_row && h <= max_row { h + 1 } else { min_row }
        } else {
            min_row
        };

        // For TopN/BottomN, precompute sorted values per column
        let mut top_bottom_sets: HashMap<u16, HashSet<u32>> = HashMap::new();
        for (col, cond) in &self.filters {
            match cond {
                FilterCondition::TopN(n) | FilterCondition::BottomN(n) => {
                    let mut vals: Vec<(u32, f64)> = (start_row..=max_row)
                        .filter_map(|r| {
                            self.get_cell_value(r, *col).as_f64().map(|v| (r, v))
                        })
                        .collect();
                    match cond {
                        FilterCondition::TopN(_) => vals.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal)),
                        FilterCondition::BottomN(_) => vals.sort_by(|a, b| a.1.partial_cmp(&b.1).unwrap_or(std::cmp::Ordering::Equal)),
                        _ => unreachable!(),
                    }
                    let keep: HashSet<u32> = vals.iter().take(*n).map(|(r, _)| *r).collect();
                    top_bottom_sets.insert(*col, keep);
                }
                _ => {}
            }
        }

        for row in start_row..=max_row {
            let mut visible = true;
            for (col, cond) in &self.filters {
                let val = self.get_cell_value(row, *col);
                let pass = match cond {
                    FilterCondition::Eq(s) => val.display_value().eq_ignore_ascii_case(s),
                    FilterCondition::NotEq(s) => !val.display_value().eq_ignore_ascii_case(s),
                    FilterCondition::Gt(n) => val.as_f64().map_or(false, |v| v > *n),
                    FilterCondition::Lt(n) => val.as_f64().map_or(false, |v| v < *n),
                    FilterCondition::Gte(n) => val.as_f64().map_or(false, |v| v >= *n),
                    FilterCondition::Lte(n) => val.as_f64().map_or(false, |v| v <= *n),
                    FilterCondition::Contains(s) => val.display_value().to_lowercase().contains(&s.to_lowercase()),
                    FilterCondition::Blanks => val.is_empty(),
                    FilterCondition::NonBlanks => !val.is_empty(),
                    FilterCondition::TopN(_) | FilterCondition::BottomN(_) => {
                        top_bottom_sets.get(col).map_or(false, |set| set.contains(&row))
                    }
                    FilterCondition::ValueSet(set) => {
                        let dv = val.display_value();
                        set.contains(&dv) || (val.is_empty() && set.contains(""))
                    }
                };
                if !pass {
                    visible = false;
                    break;
                }
            }
            if !visible {
                self.hidden_rows.insert(row);
            }
        }
    }

    pub fn delete_col(&mut self, at_col: u16) -> Vec<(u32, Cell)> {
        let mut removed = Vec::new();
        let mut new_cells = HashMap::new();
        for (&(r, c), cell) in &self.cells {
            if c == at_col {
                removed.push((r, cell.clone()));
            } else if c > at_col {
                new_cells.insert((r, c - 1), cell.clone());
            } else {
                new_cells.insert((r, c), cell.clone());
            }
        }
        self.cells = new_cells;
        if self.extent.1 > 0 {
            self.extent.1 -= 1;
        }
        removed
    }
}
