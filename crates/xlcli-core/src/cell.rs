use chrono::NaiveDateTime;
use compact_str::CompactString;
use std::fmt;

use crate::types::CellError;

#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    Empty,
    Number(f64),
    String(CompactString),
    Boolean(bool),
    DateTime(NaiveDateTime),
    Error(CellError),
    Array(Box<Vec<Vec<CellValue>>>),
}

impl Default for CellValue {
    fn default() -> Self {
        CellValue::Empty
    }
}

impl CellValue {
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }

    pub fn as_f64(&self) -> Option<f64> {
        match self {
            CellValue::Number(n) => Some(*n),
            CellValue::Boolean(b) => Some(if *b { 1.0 } else { 0.0 }),
            _ => None,
        }
    }

    pub fn as_str(&self) -> Option<&str> {
        match self {
            CellValue::String(s) => Some(s.as_str()),
            _ => None,
        }
    }

    pub fn display_value(&self) -> String {
        match self {
            CellValue::Empty => String::new(),
            CellValue::Number(n) => {
                if *n == n.floor() && n.abs() < 1e15 {
                    format!("{}", *n as i64)
                } else {
                    format!("{}", n)
                }
            }
            CellValue::String(s) => s.to_string(),
            CellValue::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            CellValue::DateTime(dt) => dt.format("%Y-%m-%d %H:%M:%S").to_string(),
            CellValue::Error(e) => e.to_string(),
            CellValue::Array(_) => "{...}".to_string(),
        }
    }
}

impl fmt::Display for CellValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.display_value())
    }
}

#[derive(Debug, Clone)]
pub struct Cell {
    pub value: CellValue,
    pub formula: Option<String>,
    pub style_id: u32,
}

impl Cell {
    pub fn new(value: CellValue) -> Self {
        Self {
            value,
            formula: None,
            style_id: 0,
        }
    }

    pub fn with_formula(value: CellValue, formula: String) -> Self {
        Self {
            value,
            formula: Some(formula),
            style_id: 0,
        }
    }
}

impl Default for Cell {
    fn default() -> Self {
        Self::new(CellValue::Empty)
    }
}
