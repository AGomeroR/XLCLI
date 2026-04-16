use thiserror::Error;

#[derive(Error, Debug)]
pub enum XlcliError {
    #[error("cell reference out of bounds: ({row}, {col})")]
    OutOfBounds { row: u32, col: u16 },

    #[error("sheet not found: {0}")]
    SheetNotFound(String),

    #[error("invalid cell address: {0}")]
    InvalidCellAddress(String),

    #[error("circular reference detected")]
    CircularReference,

    #[error(transparent)]
    Other(#[from] Box<dyn std::error::Error + Send + Sync>),
}

pub type Result<T> = std::result::Result<T, XlcliError>;
