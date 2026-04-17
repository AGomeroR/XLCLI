use crate::cell::CellValue;
use crate::range::CellRange;
use crate::style::Color;
use crate::types::CellAddr;

#[derive(Debug, Clone, PartialEq)]
pub enum Condition {
    Always,
    Gt(f64),
    Lt(f64),
    Gte(f64),
    Lte(f64),
    Eq(f64),
    Neq(f64),
    Between(f64, f64),
    Contains(String),
    Blanks,
    NonBlanks,
}

impl Condition {
    pub fn matches(&self, v: &CellValue) -> bool {
        use Condition::*;
        match self {
            Always => true,
            Blanks => matches!(v, CellValue::Empty),
            NonBlanks => !matches!(v, CellValue::Empty),
            Contains(s) => {
                let text = match v {
                    CellValue::String(t) => t.to_string(),
                    CellValue::Number(n) => n.to_string(),
                    CellValue::Boolean(b) => b.to_string(),
                    _ => return false,
                };
                text.to_lowercase().contains(&s.to_lowercase())
            }
            _ => {
                let n = match v {
                    CellValue::Number(n) => *n,
                    CellValue::Boolean(b) => if *b { 1.0 } else { 0.0 },
                    _ => return false,
                };
                match self {
                    Gt(x) => n > *x,
                    Lt(x) => n < *x,
                    Gte(x) => n >= *x,
                    Lte(x) => n <= *x,
                    Eq(x) => n == *x,
                    Neq(x) => n != *x,
                    Between(a, b) => n >= *a && n <= *b,
                    _ => false,
                }
            }
        }
    }
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct StyleOverlay {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub double_underline: Option<bool>,
    pub strikethrough: Option<bool>,
    pub overline: Option<bool>,
    pub fg_color: Option<Option<Color>>,
    pub bg_color: Option<Option<Color>>,
}

impl StyleOverlay {
    pub fn is_empty(&self) -> bool {
        *self == StyleOverlay::default()
    }

    pub fn apply(&self, s: &mut crate::style::CellStyle) {
        if let Some(b) = self.bold { s.bold = b; }
        if let Some(b) = self.italic { s.italic = b; }
        if let Some(b) = self.underline { s.underline = b; }
        if let Some(b) = self.double_underline { s.double_underline = b; }
        if let Some(b) = self.strikethrough { s.strikethrough = b; }
        if let Some(b) = self.overline { s.overline = b; }
        if let Some(c) = self.fg_color { s.fg_color = c; }
        if let Some(c) = self.bg_color { s.bg_color = c; }
    }
}

#[derive(Debug, Clone)]
pub struct CondRule {
    pub range: CellRange,
    pub cond: Condition,
    pub style: StyleOverlay,
}

impl CondRule {
    pub fn applies_to(&self, addr: &CellAddr) -> bool {
        self.range.contains(addr)
    }
}
