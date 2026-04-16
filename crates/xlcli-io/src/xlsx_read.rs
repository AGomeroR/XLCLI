use std::path::Path;

use calamine::{open_workbook_auto, Data, Reader};
use xlcli_core::cell::{Cell, CellValue};
use xlcli_core::sheet::Sheet;
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
