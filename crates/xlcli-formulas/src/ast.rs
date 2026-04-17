use xlcli_core::types::CellError;

#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    Number(f64),
    String(String),
    Boolean(bool),
    Error(CellError),

    CellRef {
        sheet: Option<String>,
        col: ColRef,
        row: RowRef,
    },
    Range {
        start: Box<Expr>,
        end: Box<Expr>,
    },
    BinOp {
        op: BinOp,
        left: Box<Expr>,
        right: Box<Expr>,
    },
    UnaryOp {
        op: UnaryOp,
        expr: Box<Expr>,
    },
    Percent(Box<Expr>),
    FnCall {
        name: String,
        args: Vec<Expr>,
    },
    NamedRef(String),
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ColRef {
    Absolute(u16),
    Relative(u16),
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum RowRef {
    Absolute(u32),
    Relative(u32),
}

impl ColRef {
    pub fn value(&self) -> u16 {
        match self {
            ColRef::Absolute(v) | ColRef::Relative(v) => *v,
        }
    }
}

impl RowRef {
    pub fn value(&self) -> u32 {
        match self {
            RowRef::Absolute(v) | RowRef::Relative(v) => *v,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum BinOp {
    Add,
    Sub,
    Mul,
    Div,
    Pow,
    Concat,
    Eq,
    Neq,
    Lt,
    Gt,
    Lte,
    Gte,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum UnaryOp {
    Neg,
    Pos,
}
