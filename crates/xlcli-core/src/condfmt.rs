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
    NotBetween(f64, f64),
    Contains(String),
    NotContains(String),
    BeginsWith(String),
    EndsWith(String),
    Blanks,
    NonBlanks,
    ContainsErrors,
    NotContainsErrors,
    DuplicateValues,
    UniqueValues,
    Top { count: u32, percent: bool, bottom: bool },
    Average { above: bool, equal: bool, std_dev: i32 },
    TimePeriod(TimePeriod),
    Expression(String),
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum TimePeriod {
    Today, Yesterday, Tomorrow,
    Last7Days, ThisWeek, LastWeek, NextWeek,
    ThisMonth, LastMonth, NextMonth,
}

impl Condition {
    pub fn matches(&self, v: &CellValue) -> bool {
        use Condition::*;
        match self {
            Always => true,
            Blanks => matches!(v, CellValue::Empty),
            NonBlanks => !matches!(v, CellValue::Empty),
            ContainsErrors => matches!(v, CellValue::Error(_)),
            NotContainsErrors => !matches!(v, CellValue::Error(_)),
            Contains(s) | NotContains(s) | BeginsWith(s) | EndsWith(s) => {
                let text = match v {
                    CellValue::String(t) => t.to_string(),
                    CellValue::Number(n) => n.to_string(),
                    CellValue::Boolean(b) => b.to_string(),
                    _ => return matches!(self, NotContains(_)),
                };
                let hay = text.to_lowercase();
                let needle = s.to_lowercase();
                match self {
                    Contains(_) => hay.contains(&needle),
                    NotContains(_) => !hay.contains(&needle),
                    BeginsWith(_) => hay.starts_with(&needle),
                    EndsWith(_) => hay.ends_with(&needle),
                    _ => false,
                }
            }
            // Context-dependent rules: handled at sheet level (need range stats)
            DuplicateValues | UniqueValues | Top { .. } | Average { .. }
            | TimePeriod(_) | Expression(_) => false,
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
                    NotBetween(a, b) => n < *a || n > *b,
                    _ => false,
                }
            }
        }
    }

    /// Whether this variant needs per-range aggregate stats for evaluation.
    pub fn needs_range_context(&self) -> bool {
        use Condition::*;
        matches!(self,
            DuplicateValues | UniqueValues | Top { .. } | Average { .. }
            | TimePeriod(_) | Expression(_))
    }
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct StyleOverlay {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub double_underline: Option<bool>,
    pub strikethrough: Option<bool>,
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
        if let Some(c) = self.fg_color { s.fg_color = c; }
        if let Some(c) = self.bg_color { s.bg_color = c; }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum CfValueKind {
    Number(f64),
    Percent(f64),
    Percentile(f64),
    Min,
    Max,
    Formula(String),
}

#[derive(Debug, Clone)]
pub struct ColorStop {
    pub value: CfValueKind,
    pub color: Color,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum IconSetKind {
    ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeSigns, ThreeSymbols, ThreeSymbols2,
    ThreeTrafficLights1, ThreeTrafficLights2,
    FourArrows, FourArrowsGray, FourRating, FourRedToBlack, FourTrafficLights,
    FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters,
}

impl IconSetKind {
    pub fn icon_count(&self) -> usize {
        use IconSetKind::*;
        match self {
            ThreeArrows | ThreeArrowsGray | ThreeFlags | ThreeSigns | ThreeSymbols
            | ThreeSymbols2 | ThreeTrafficLights1 | ThreeTrafficLights2 => 3,
            FourArrows | FourArrowsGray | FourRating | FourRedToBlack | FourTrafficLights => 4,
            FiveArrows | FiveArrowsGray | FiveRating | FiveQuarters => 5,
        }
    }

    /// Unicode approximation for terminal display, index 0 = lowest bucket.
    pub fn icons(&self) -> &'static [&'static str] {
        use IconSetKind::*;
        match self {
            ThreeArrows | ThreeArrowsGray         => &["\u{25BC}", "\u{25B6}", "\u{25B2}"],
            ThreeFlags                            => &["\u{2691}", "\u{2691}", "\u{2691}"],
            ThreeSigns                            => &["\u{25C6}", "\u{25B2}", "\u{25CF}"],
            ThreeSymbols | ThreeSymbols2          => &["\u{2717}", "!", "\u{2713}"],
            ThreeTrafficLights1 | ThreeTrafficLights2 => &["\u{25CF}", "\u{25CF}", "\u{25CF}"],
            FourArrows | FourArrowsGray           => &["\u{25BC}", "\u{2198}", "\u{2197}", "\u{25B2}"],
            FourRating                            => &["\u{2581}", "\u{2584}", "\u{2586}", "\u{2588}"],
            FourRedToBlack                        => &["\u{25CF}", "\u{25CF}", "\u{25CF}", "\u{25CF}"],
            FourTrafficLights                     => &["\u{25CF}", "\u{25CF}", "\u{25CF}", "\u{25CF}"],
            FiveArrows | FiveArrowsGray           => &["\u{25BC}", "\u{2198}", "\u{25B6}", "\u{2197}", "\u{25B2}"],
            FiveRating                            => &["\u{2581}", "\u{2583}", "\u{2585}", "\u{2587}", "\u{2588}"],
            FiveQuarters                          => &["\u{25CB}", "\u{25D4}", "\u{25D1}", "\u{25D5}", "\u{25CF}"],
        }
    }
}

#[derive(Debug, Clone)]
pub struct IconThreshold {
    pub value: CfValueKind,
    pub gte: bool, // true = >=, false = >
}

/// The style action of a rule. Classic rules apply an overlay (dxf). Gradient/Bar/Icon
/// rules compute per-cell style dynamically.
#[derive(Debug, Clone)]
pub enum StyleSpec {
    Overlay(StyleOverlay),
    ColorScale(Vec<ColorStop>),
    DataBar { min: CfValueKind, max: CfValueKind, color: Color },
    IconSet { kind: IconSetKind, thresholds: Vec<IconThreshold>, reverse: bool, show_value: bool },
}

impl Default for StyleSpec {
    fn default() -> Self {
        StyleSpec::Overlay(StyleOverlay::default())
    }
}

#[derive(Debug, Clone)]
pub struct CondRule {
    pub range: CellRange,
    pub cond: Condition,
    pub style: StyleSpec,
}

impl CondRule {
    pub fn applies_to(&self, addr: &CellAddr) -> bool {
        self.range.contains(addr)
    }

    /// Backwards-compat helper: construct a classic overlay rule.
    pub fn classic(range: CellRange, cond: Condition, overlay: StyleOverlay) -> Self {
        Self { range, cond, style: StyleSpec::Overlay(overlay) }
    }
}
