use std::path::Path;

use rust_xlsxwriter::{Format, Workbook as XlsxWorkbook};
use xlcli_core::cell::CellValue;
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
                match &cell.value {
                    CellValue::Empty => {}
                    CellValue::Number(n) => {
                        ws.write_number(row, col, *n)?;
                    }
                    CellValue::String(s) => {
                        ws.write_string(row, col, s.as_str())?;
                    }
                    CellValue::Boolean(b) => {
                        ws.write_boolean(row, col, *b)?;
                    }
                    CellValue::DateTime(dt) => {
                        let serial = datetime_to_excel_serial(dt);
                        let fmt = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");
                        ws.write_number_with_format(row, col, serial, &fmt)?;
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
