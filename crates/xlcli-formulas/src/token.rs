use logos::Logos;


#[derive(Logos, Debug, Clone, PartialEq)]
#[logos(skip r"[ \t]+")]
pub enum Token<'src> {
    #[regex(r"[0-9]+(\.[0-9]+)?([eE][+-]?[0-9]+)?")]
    Number(&'src str),

    #[regex(r#""([^"\\]|\\.)*""#)]
    StringLit(&'src str),

    #[regex(r"(?i)(TRUE|FALSE)")]
    Boolean(&'src str),

    #[regex(r"[A-Za-z_][A-Za-z0-9_.]*")]
    Ident(&'src str),

    #[regex(r"'[^']+'\!\$?[A-Za-z]{1,3}\$?[0-9]+")]
    SheetRefQuoted(&'src str),

    #[regex(r"[A-Za-z_][A-Za-z0-9_]*\!\$?[A-Za-z]{1,3}\$?[0-9]+", priority = 5)]
    SheetRef(&'src str),

    #[regex(r"\$?[A-Za-z]{1,3}\$?[0-9]+")]
    CellRef(&'src str),

    #[token("+")]
    Plus,
    #[token("-")]
    Minus,
    #[token("*")]
    Star,
    #[token("/")]
    Slash,
    #[token("^")]
    Caret,
    #[token("&")]
    Ampersand,
    #[token("=")]
    Eq,
    #[token("<>")]
    Neq,
    #[token("<=")]
    Lte,
    #[token(">=")]
    Gte,
    #[token("<")]
    Lt,
    #[token(">")]
    Gt,
    #[token("%")]
    Percent,
    #[token("(")]
    LParen,
    #[token(")")]
    RParen,
    #[token(",")]
    Comma,
    #[token(":")]
    Colon,
    #[token(";")]
    Semicolon,
    #[token("!")]
    Bang,
}
