use std::path::Path;
use xlcli_core::workbook::Workbook;

pub trait FileReader {
    fn read(&self, path: &Path) -> anyhow::Result<Workbook>;
    fn extensions(&self) -> &[&str];
}

pub fn read_file(path: &Path) -> anyhow::Result<Workbook> {
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_lowercase();

    match ext.as_str() {
        "xlsx" | "xls" | "ods" => {
            let reader = crate::xlsx_read::XlsxReader;
            reader.read(path)
        }
        "csv" => {
            let reader = crate::csv_read::CsvReader::new(b',');
            reader.read(path)
        }
        "tsv" => {
            let reader = crate::csv_read::CsvReader::new(b'\t');
            reader.read(path)
        }
        _ => anyhow::bail!("unsupported file format: {}", ext),
    }
}
