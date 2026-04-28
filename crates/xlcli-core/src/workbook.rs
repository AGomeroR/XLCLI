use crate::range::NamedRange;
use crate::sheet::Sheet;
use crate::style::StylePool;

#[derive(Debug)]
pub struct Workbook {
    pub sheets: Vec<Sheet>,
    pub active_sheet: usize,
    pub style_pool: StylePool,
    pub named_ranges: Vec<NamedRange>,
    pub file_path: Option<String>,
    pub load_diagnostic: Option<String>,
}

impl Workbook {
    pub fn new() -> Self {
        let mut wb = Self {
            sheets: Vec::new(),
            active_sheet: 0,
            style_pool: StylePool::new(),
            named_ranges: Vec::new(),
            file_path: None,
            load_diagnostic: None,
        };
        wb.sheets.push(Sheet::new("Sheet1"));
        wb
    }

    pub fn active_sheet(&self) -> &Sheet {
        &self.sheets[self.active_sheet]
    }

    pub fn active_sheet_mut(&mut self) -> &mut Sheet {
        &mut self.sheets[self.active_sheet]
    }

    pub fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    pub fn add_sheet(&mut self, name: impl Into<String>) -> usize {
        let idx = self.sheets.len();
        self.sheets.push(Sheet::new(name));
        idx
    }

    pub fn sheet_by_name(&self, name: &str) -> Option<(usize, &Sheet)> {
        self.sheets
            .iter()
            .enumerate()
            .find(|(_, s)| s.name == name)
    }

    pub fn sheet_by_name_mut(&mut self, name: &str) -> Option<(usize, &mut Sheet)> {
        self.sheets
            .iter_mut()
            .enumerate()
            .find(|(_, s)| s.name == name)
    }
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}
