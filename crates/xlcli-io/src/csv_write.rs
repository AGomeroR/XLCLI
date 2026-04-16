use std::path::Path;

use xlcli_core::workbook::Workbook;

use crate::writer::FileWriter;

pub struct CsvWriter {
    delimiter: u8,
}

impl CsvWriter {
    pub fn new(delimiter: u8) -> Self {
        Self { delimiter }
    }
}

impl FileWriter for CsvWriter {
    fn write(&self, workbook: &Workbook, path: &Path) -> anyhow::Result<()> {
        let sheet = workbook.active_sheet();
        let (rows, cols) = sheet.extent();

        let mut wtr = csv::WriterBuilder::new()
            .delimiter(self.delimiter)
            .from_path(path)?;

        for row in 0..rows {
            let mut record = Vec::new();
            for col in 0..cols {
                record.push(sheet.get_cell_value(row, col).display_value());
            }
            wtr.write_record(&record)?;
        }

        wtr.flush()?;
        Ok(())
    }

    fn extensions(&self) -> &[&str] {
        &["csv", "tsv"]
    }
}
