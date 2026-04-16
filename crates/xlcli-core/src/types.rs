use std::fmt;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellAddr {
    pub sheet: u16,
    pub row: u32,
    pub col: u16,
}

impl CellAddr {
    pub fn new(sheet: u16, row: u32, col: u16) -> Self {
        Self { sheet, row, col }
    }

    pub fn col_name(col: u16) -> String {
        let mut name = String::new();
        let mut c = col as u32;
        loop {
            name.insert(0, (b'A' + (c % 26) as u8) as char);
            if c < 26 {
                break;
            }
            c = c / 26 - 1;
        }
        name
    }

    pub fn parse_col(s: &str) -> Option<u16> {
        let mut col: u32 = 0;
        for b in s.bytes() {
            if !b.is_ascii_uppercase() {
                return None;
            }
            col = col * 26 + (b - b'A') as u32 + 1;
        }
        if col == 0 {
            return None;
        }
        Some((col - 1) as u16)
    }

    pub fn display_name(&self) -> String {
        format!("{}{}", Self::col_name(self.col), self.row + 1)
    }
}

impl fmt::Display for CellAddr {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}{}", Self::col_name(self.col), self.row + 1)
    }
}

impl PartialOrd for CellAddr {
    fn partial_cmp(&self, other: &Self) -> Option<std::cmp::Ordering> {
        Some(self.cmp(other))
    }
}

impl Ord for CellAddr {
    fn cmp(&self, other: &Self) -> std::cmp::Ordering {
        self.sheet
            .cmp(&other.sheet)
            .then(self.row.cmp(&other.row))
            .then(self.col.cmp(&other.col))
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum CellError {
    Div0,
    Na,
    Name,
    Null,
    Num,
    Ref,
    Value,
    GettingData,
}

impl fmt::Display for CellError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            CellError::Div0 => write!(f, "#DIV/0!"),
            CellError::Na => write!(f, "#N/A"),
            CellError::Name => write!(f, "#NAME?"),
            CellError::Null => write!(f, "#NULL!"),
            CellError::Num => write!(f, "#NUM!"),
            CellError::Ref => write!(f, "#REF!"),
            CellError::Value => write!(f, "#VALUE!"),
            CellError::GettingData => write!(f, "#GETTING_DATA"),
        }
    }
}
