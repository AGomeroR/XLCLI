use std::path::Path;

use xlcli_core::cell::{Cell, CellValue};
use xlcli_core::sheet::Sheet;
use xlcli_core::workbook::Workbook;

use crate::reader::FileReader;

pub struct CsvReader {
    delimiter: u8,
}

impl CsvReader {
    pub fn new(delimiter: u8) -> Self {
        Self { delimiter }
    }
}

impl FileReader for CsvReader {
    fn read(&self, path: &Path) -> anyhow::Result<Workbook> {
        let mut rdr = csv::ReaderBuilder::new()
            .delimiter(self.delimiter)
            .has_headers(false)
            .flexible(true)
            .from_path(path)?;

        let mut wb = Workbook::new();
        wb.sheets.clear();
        wb.file_path = Some(path.display().to_string());

        let stem = path
            .file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("Sheet1");
        let mut sheet = Sheet::new(stem);

        for (row_idx, result) in rdr.records().enumerate() {
            let record = result?;
            for (col_idx, field) in record.iter().enumerate() {
                let value = parse_csv_value(field);
                if !value.is_empty() {
                    sheet.set_cell(row_idx as u32, col_idx as u16, Cell::new(value));
                }
            }
        }

        wb.sheets.push(sheet);
        Ok(wb)
    }

    fn extensions(&self) -> &[&str] {
        &["csv", "tsv"]
    }
}

fn parse_csv_value(s: &str) -> CellValue {
    if s.is_empty() {
        return CellValue::Empty;
    }

    if let Ok(n) = s.parse::<f64>() {
        return CellValue::Number(n);
    }

    match s.to_uppercase().as_str() {
        "TRUE" => CellValue::Boolean(true),
        "FALSE" => CellValue::Boolean(false),
        _ => CellValue::String(s.into()),
    }
}
