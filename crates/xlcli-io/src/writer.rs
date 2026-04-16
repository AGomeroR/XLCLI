use std::path::Path;
use xlcli_core::workbook::Workbook;

pub trait FileWriter {
    fn write(&self, workbook: &Workbook, path: &Path) -> anyhow::Result<()>;
    fn extensions(&self) -> &[&str];
}

pub fn write_file(workbook: &Workbook, path: &Path) -> anyhow::Result<()> {
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_lowercase();

    match ext.as_str() {
        "xlsx" => {
            let writer = crate::xlsx_write::XlsxWriter;
            writer.write(workbook, path)
        }
        "csv" => {
            let writer = crate::csv_write::CsvWriter::new(b',');
            writer.write(workbook, path)
        }
        "tsv" => {
            let writer = crate::csv_write::CsvWriter::new(b'\t');
            writer.write(workbook, path)
        }
        _ => anyhow::bail!("unsupported write format: {}", ext),
    }
}
