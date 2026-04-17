use std::collections::{HashMap, HashSet};

use crate::cell::{Cell, CellValue};
use crate::condfmt::{CondRule, StyleOverlay};

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
        let addr = crate::types::CellAddr::new(0, row, col);
        let val = self.get_cell_value(row, col);
        for rule in &self.cond_rules {
            if rule.applies_to(&addr) && rule.cond.matches(val) {
                rule.style.apply(&mut s);
            }
        }
        s
    }

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
