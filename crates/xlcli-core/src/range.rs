use crate::types::CellAddr;

#[derive(Debug, Clone, PartialEq)]
pub struct CellRange {
    pub start: CellAddr,
    pub end: CellAddr,
}

impl CellRange {
    pub fn new(start: CellAddr, end: CellAddr) -> Self {
        Self { start, end }
    }

    pub fn rows(&self) -> u32 {
        self.end.row.saturating_sub(self.start.row) + 1
    }

    pub fn cols(&self) -> u16 {
        self.end.col.saturating_sub(self.start.col) + 1
    }

    pub fn contains(&self, addr: &CellAddr) -> bool {
        addr.sheet == self.start.sheet
            && addr.row >= self.start.row
            && addr.row <= self.end.row
            && addr.col >= self.start.col
            && addr.col <= self.end.col
    }

    pub fn iter(&self) -> CellRangeIter {
        CellRangeIter {
            range: self.clone(),
            current_row: self.start.row,
            current_col: self.start.col,
        }
    }
}

pub struct CellRangeIter {
    range: CellRange,
    current_row: u32,
    current_col: u16,
}

impl Iterator for CellRangeIter {
    type Item = CellAddr;

    fn next(&mut self) -> Option<Self::Item> {
        if self.current_row > self.range.end.row {
            return None;
        }
        let addr = CellAddr::new(self.range.start.sheet, self.current_row, self.current_col);
        self.current_col += 1;
        if self.current_col > self.range.end.col {
            self.current_col = self.range.start.col;
            self.current_row += 1;
        }
        Some(addr)
    }
}

#[derive(Debug, Clone)]
pub struct NamedRange {
    pub name: String,
    pub range: CellRange,
}
