use std::collections::{HashMap, HashSet};

use crate::cell::{Cell, CellValue};

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
        }
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
