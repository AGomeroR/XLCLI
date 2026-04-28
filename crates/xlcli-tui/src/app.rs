use std::collections::HashMap;

use xlcli_core::cell::{Cell, CellValue};
use xlcli_core::dep_graph::DepGraph;
use xlcli_core::sheet::FilterCondition;
use xlcli_core::types::CellAddr;
use xlcli_core::workbook::Workbook;
use xlcli_formulas::{FunctionRegistry, EvalContext, extract_refs_with_resolver};

use crate::clipboard::Clipboard;
use crate::config::Config;
use crate::mode::Mode;
use crate::undo::{UndoEntry, UndoStack};
use crate::viewport::Viewport;

pub const MAX_ROW: u32 = 1_048_575;
pub const MAX_COL: u16 = 16_383;

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum SortType {
    AZ,
    ZA,
    Numeric,
    NumericDesc,
    CaseSensitive,
    Natural,
}

impl SortType {
    pub fn label(&self) -> &str {
        match self {
            SortType::AZ => "A-Z",
            SortType::ZA => "Z-A",
            SortType::Numeric => "Numeric",
            SortType::NumericDesc => "Numeric Desc",
            SortType::CaseSensitive => "Case Sensitive",
            SortType::Natural => "Natural",
        }
    }

    pub fn all() -> &'static [SortType] {
        &[SortType::AZ, SortType::ZA, SortType::Numeric, SortType::NumericDesc, SortType::CaseSensitive, SortType::Natural]
    }

    pub fn next(&self) -> SortType {
        let all = Self::all();
        let idx = all.iter().position(|s| s == self).unwrap_or(0);
        all[(idx + 1) % all.len()]
    }

    pub fn prev(&self) -> SortType {
        let all = Self::all();
        let idx = all.iter().position(|s| s == self).unwrap_or(0);
        all[(idx + all.len() - 1) % all.len()]
    }
}

#[derive(Debug, Clone)]
pub struct SortKey {
    pub col: u16,
    pub sort_type: SortType,
}

#[derive(Debug, Clone, PartialEq)]
pub enum SortDialogFocus {
    Range,
    SortByCol,
    SortByType,
    ThenByCol,
    ThenByType,
    HasHeaders,
    BtnSort,
    BtnCancel,
}

impl SortDialogFocus {
    pub fn next(&self) -> Self {
        match self {
            Self::Range => Self::SortByCol,
            Self::SortByCol => Self::SortByType,
            Self::SortByType => Self::ThenByCol,
            Self::ThenByCol => Self::ThenByType,
            Self::ThenByType => Self::HasHeaders,
            Self::HasHeaders => Self::BtnSort,
            Self::BtnSort => Self::BtnCancel,
            Self::BtnCancel => Self::Range,
        }
    }

    pub fn prev(&self) -> Self {
        match self {
            Self::Range => Self::BtnCancel,
            Self::SortByCol => Self::Range,
            Self::SortByType => Self::SortByCol,
            Self::ThenByCol => Self::SortByType,
            Self::ThenByType => Self::ThenByCol,
            Self::HasHeaders => Self::ThenByType,
            Self::BtnSort => Self::HasHeaders,
            Self::BtnCancel => Self::BtnSort,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum SortDropdown {
    None,
    SortByCol,
    SortByType,
    ThenByCol,
    ThenByType,
}

#[derive(Debug, Clone)]
pub struct SortDialog {
    pub visible: bool,
    pub focus: SortDialogFocus,
    pub range_text: String,
    pub min_row: u32,
    pub max_row: u32,
    pub min_col: u16,
    pub max_col: u16,
    pub sort_by_col: u16,
    pub sort_by_type: SortType,
    pub then_by_col: Option<u16>,
    pub then_by_type: SortType,
    pub has_headers: bool,
    pub open_dropdown: SortDropdown,
    pub dropdown_scroll: usize,
    pub screen_x: u16,
    pub screen_y: u16,
    pub screen_w: u16,
    pub screen_h: u16,
    pub dd_item_x: u16,
    pub dd_item_y: u16,
    pub dd_item_w: u16,
    pub dd_item_count: u16,
}

pub struct SheetState {
    pub cursor: CellAddr,
    pub viewport: Viewport,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum FilterType {
    Eq,
    NotEq,
    Gt,
    Lt,
    Gte,
    Lte,
    Contains,
    Blanks,
    NonBlanks,
    TopN,
    BottomN,
}

impl FilterType {
    pub fn label(&self) -> &str {
        match self {
            FilterType::Eq => "= Equal",
            FilterType::NotEq => "!= Not Equal",
            FilterType::Gt => "> Greater",
            FilterType::Lt => "< Less",
            FilterType::Gte => ">= Gr/Equal",
            FilterType::Lte => "<= Le/Equal",
            FilterType::Contains => "Contains",
            FilterType::Blanks => "Blanks",
            FilterType::NonBlanks => "Non-Blanks",
            FilterType::TopN => "Top N",
            FilterType::BottomN => "Bottom N",
        }
    }

    pub fn all() -> &'static [FilterType] {
        &[
            FilterType::Eq, FilterType::NotEq,
            FilterType::Gt, FilterType::Lt, FilterType::Gte, FilterType::Lte,
            FilterType::Contains, FilterType::Blanks, FilterType::NonBlanks,
            FilterType::TopN, FilterType::BottomN,
        ]
    }

    pub fn needs_value(&self) -> bool {
        !matches!(self, FilterType::Blanks | FilterType::NonBlanks)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum FilterDialogFocus {
    Column,
    FilterType,
    Value,
    BtnApply,
    BtnCancel,
}

impl FilterDialogFocus {
    pub fn next(&self) -> Self {
        match self {
            Self::Column => Self::FilterType,
            Self::FilterType => Self::Value,
            Self::Value => Self::BtnApply,
            Self::BtnApply => Self::BtnCancel,
            Self::BtnCancel => Self::Column,
        }
    }

    pub fn prev(&self) -> Self {
        match self {
            Self::Column => Self::BtnCancel,
            Self::FilterType => Self::Column,
            Self::Value => Self::FilterType,
            Self::BtnApply => Self::Value,
            Self::BtnCancel => Self::BtnApply,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum FilterDropdown {
    None,
    Column,
    FilterType,
}

#[derive(Debug, Clone)]
pub struct FilterDialog {
    pub visible: bool,
    pub focus: FilterDialogFocus,
    pub col: u16,
    pub filter_type: FilterType,
    pub value_buf: String,
    pub min_col: u16,
    pub max_col: u16,
    pub open_dropdown: FilterDropdown,
    pub dropdown_scroll: usize,
    pub screen_x: u16,
    pub screen_y: u16,
    pub screen_w: u16,
    pub screen_h: u16,
    pub dd_item_x: u16,
    pub dd_item_y: u16,
    pub dd_item_w: u16,
    pub dd_item_count: u16,
    pub all_values: Vec<(String, bool)>,
    pub filtered_values: Vec<usize>,
    pub value_selected: usize,
    pub value_scroll: usize,
    pub value_search: String,
    pub all_checked: bool,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum CfCond {
    Gt, Lt, Gte, Lte, Eq, Neq, Between, Contains, Blanks, NonBlanks,
}

impl CfCond {
    pub fn label(&self) -> &str {
        match self {
            CfCond::Gt => "> Greater",
            CfCond::Lt => "< Less",
            CfCond::Gte => ">= Gr/Equal",
            CfCond::Lte => "<= Le/Equal",
            CfCond::Eq => "= Equal",
            CfCond::Neq => "!= Not Equal",
            CfCond::Between => "Between A B",
            CfCond::Contains => "Contains",
            CfCond::Blanks => "Blanks",
            CfCond::NonBlanks => "Non-Blanks",
        }
    }
    pub fn all() -> &'static [CfCond] {
        &[CfCond::Gt, CfCond::Lt, CfCond::Gte, CfCond::Lte, CfCond::Eq, CfCond::Neq,
          CfCond::Between, CfCond::Contains, CfCond::Blanks, CfCond::NonBlanks]
    }
}

pub const CF_COLORS: &[&str] = &[
    "none", "red", "green", "blue", "yellow", "cyan", "magenta", "orange", "white", "black", "gray",
];

#[derive(Debug, Clone, PartialEq)]
pub enum CfDialogFocus {
    Conditional,
    Range,
    Cond,
    Val1,
    Val2,
    Bold, Italic, Under, DUnder, Strike,
    Bg, BgHex, Fg, FgHex,
    RulesList,
    BtnBase,
    BtnDismiss,
    BtnApply,
    BtnClose,
    BtnDelete,
    BtnCleanAll,
}

impl CfDialogFocus {
    pub fn next(&self, conditional: bool) -> Self {
        use CfDialogFocus::*;
        match self {
            Range => Bold,
            Bold => Italic,
            Italic => Under,
            Under => DUnder,
            DUnder => Strike,
            Strike => Bg,
            Bg => BgHex,
            BgHex => Fg,
            Fg => FgHex,
            FgHex => Conditional,
            Conditional => if conditional { Cond } else { BtnBase },
            Cond => Val1,
            Val1 => Val2,
            Val2 => BtnBase,
            RulesList => BtnBase,
            BtnBase => BtnDismiss,
            BtnDismiss => BtnApply,
            BtnApply => BtnClose,
            BtnClose => Range,
            BtnDelete => BtnClose,
            BtnCleanAll => BtnClose,
        }
    }
    pub fn prev(&self, conditional: bool) -> Self {
        use CfDialogFocus::*;
        match self {
            Range => BtnClose,
            Bold => Range,
            Italic => Bold,
            Under => Italic,
            DUnder => Under,
            Strike => DUnder,
            Bg => Strike,
            BgHex => Bg,
            Fg => BgHex,
            FgHex => Fg,
            Conditional => FgHex,
            Cond => Conditional,
            Val1 => Cond,
            Val2 => Val1,
            RulesList => if conditional { Val2 } else { Conditional },
            BtnBase => if conditional { Val2 } else { Conditional },
            BtnDismiss => BtnBase,
            BtnApply => BtnDismiss,
            BtnClose => BtnApply,
            BtnDelete => BtnClose,
            BtnCleanAll => BtnClose,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum CfDropdown { None, Cond, Bg, Fg }

#[derive(Debug, Clone, Default)]
pub struct CfListDialog {
    pub visible: bool,
    pub selected: usize,
    pub scroll: usize,
    pub pending_d: bool,
    pub rect: (u16, u16, u16, u16),
}

#[derive(Debug, Clone)]
pub struct CfDialog {
    pub visible: bool,
    pub focus: CfDialogFocus,
    pub editing_text: bool,
    pub conditional: bool,
    pub range_text: String,
    pub cond: CfCond,
    pub val1: String,
    pub val2: String,
    pub bold: bool,
    pub italic: bool,
    pub under: bool,
    pub dunder: bool,
    pub strike: bool,
    pub bg_idx: usize, // into CF_COLORS (0 = none)
    pub fg_idx: usize,
    pub bg_hex: String, // optional "#RRGGBB", overrides preset when valid
    pub fg_hex: String,
    pub open_dropdown: CfDropdown,
    pub dropdown_scroll: usize,
    pub rules_selected: usize,
    pub rules_scroll: usize,
    pub screen_x: u16,
    pub screen_y: u16,
    pub screen_w: u16,
    pub screen_h: u16,
    pub dd_item_x: u16,
    pub dd_item_y: u16,
    pub dd_item_w: u16,
    pub dd_item_count: u16,
    // Click-area rects filled during render (x, y, w)
    pub rect_conditional: (u16, u16, u16),
    pub rect_range: (u16, u16, u16),
    pub rect_cond: (u16, u16, u16),
    pub rect_val1: (u16, u16, u16),
    pub rect_val2: (u16, u16, u16),
    pub rect_bold: (u16, u16, u16),
    pub rect_italic: (u16, u16, u16),
    pub rect_under: (u16, u16, u16),
    pub rect_dunder: (u16, u16, u16),
    pub rect_strike: (u16, u16, u16),
    pub rect_bg: (u16, u16, u16),
    pub rect_bg_hex: (u16, u16, u16),
    pub rect_fg: (u16, u16, u16),
    pub rect_fg_hex: (u16, u16, u16),
    pub rect_rules: (u16, u16, u16, u16), // x,y,w,h
    pub rect_apply: (u16, u16, u16),
    pub rect_base: (u16, u16, u16),
    pub rect_dismiss: (u16, u16, u16),
    pub rect_delete: (u16, u16, u16),
    pub rect_cleanall: (u16, u16, u16),
    pub rect_close: (u16, u16, u16),
}

pub struct App {
    pub workbook: Workbook,
    pub cursor: CellAddr,
    pub viewport: Viewport,
    pub mode: Mode,
    pub command_buffer: String,
    pub cmd_palette_selected: usize,
    pub should_quit: bool,
    pub status_message: Option<String>,
    pub pending_key: Option<char>,
    pub edit_buffer: String,
    pub edit_cursor: usize,
    pub undo_stack: UndoStack,
    pub clipboard: Clipboard,
    pub visual_anchor: Option<CellAddr>,
    pub modified: bool,
    pub config: Config,
    pub formula_registry: FunctionRegistry,
    pub autocomplete: AutocompleteState,
    pub dep_graph: DepGraph,
    pub sheet_states: HashMap<usize, SheetState>,
    pub formula_origin_sheet: Option<u16>,
    pub sort_dialog: SortDialog,
    pub filter_dialog: FilterDialog,
    pub cf_dialog: CfDialog,
    pub cf_list: CfListDialog,
    pub formula_error: Option<xlcli_formulas::ParseError>,
    pub search_buffer: String,
    pub search_results: Vec<(u32, u16)>,
    pub search_idx: usize,
    pub search_active: bool,
}

impl App {
    pub fn recompute_formula_error(&mut self) {
        self.formula_error = if self.mode == Mode::Insert {
            if let Some(body) = self.edit_buffer.strip_prefix('=') {
                if body.trim().is_empty() {
                    None
                } else {
                    xlcli_formulas::parse(body).err()
                }
            } else {
                None
            }
        } else {
            None
        };
    }
}

pub struct AutocompleteState {
    pub visible: bool,
    pub matches: Vec<&'static str>,
    pub selected: usize,
    pub active_function: Option<String>,
}

impl App {
    pub fn new(workbook: Workbook, config: Config) -> Self {
        let diag = workbook.load_diagnostic.clone();
        let mut app = Self {
            workbook,
            cursor: CellAddr::new(0, 0, 0),
            viewport: Viewport::new(),
            mode: Mode::Normal,
            command_buffer: String::new(),
            cmd_palette_selected: 0,
            should_quit: false,
            status_message: diag,
            pending_key: None,
            edit_buffer: String::new(),
            edit_cursor: 0,
            undo_stack: UndoStack::new(),
            clipboard: Clipboard::new(),
            visual_anchor: None,
            modified: false,
            config,
            formula_registry: FunctionRegistry::default(),
            autocomplete: AutocompleteState { visible: false, matches: Vec::new(), selected: 0, active_function: None },
            dep_graph: DepGraph::new(),
            sheet_states: HashMap::new(),
            formula_origin_sheet: None,
            sort_dialog: SortDialog {
                visible: false,
                focus: SortDialogFocus::SortByCol,
                range_text: String::new(),
                min_row: 0, max_row: 0, min_col: 0, max_col: 0,
                sort_by_col: 0,
                sort_by_type: SortType::AZ,
                then_by_col: None,
                then_by_type: SortType::AZ,
                has_headers: false,
                open_dropdown: SortDropdown::None,
                dropdown_scroll: 0,
                screen_x: 0, screen_y: 0, screen_w: 42, screen_h: 11,
                dd_item_x: 0, dd_item_y: 0, dd_item_w: 0, dd_item_count: 0,
            },
            filter_dialog: FilterDialog {
                visible: false,
                focus: FilterDialogFocus::Column,
                col: 0,
                filter_type: FilterType::Eq,
                value_buf: String::new(),
                min_col: 0, max_col: 0,
                open_dropdown: FilterDropdown::None,
                dropdown_scroll: 0,
                screen_x: 0, screen_y: 0, screen_w: 44, screen_h: 11,
                dd_item_x: 0, dd_item_y: 0, dd_item_w: 0, dd_item_count: 0,
                all_values: Vec::new(),
                filtered_values: Vec::new(),
                value_selected: 0,
                value_scroll: 0,
                value_search: String::new(),
                all_checked: true,
            },
            cf_dialog: CfDialog {
                visible: false,
                focus: CfDialogFocus::Range,
                editing_text: false,
                conditional: false,
                range_text: String::new(),
                cond: CfCond::Gt,
                val1: String::new(),
                val2: String::new(),
                bold: false, italic: false, under: false, dunder: false, strike: false,
                bg_idx: 0, fg_idx: 0,
                bg_hex: String::new(), fg_hex: String::new(),
                open_dropdown: CfDropdown::None,
                dropdown_scroll: 0,
                rules_selected: 0,
                rules_scroll: 0,
                screen_x: 0, screen_y: 0, screen_w: 56, screen_h: 24,
                dd_item_x: 0, dd_item_y: 0, dd_item_w: 0, dd_item_count: 0,
                rect_conditional: (0,0,0),
                rect_range: (0,0,0), rect_cond: (0,0,0), rect_val1: (0,0,0), rect_val2: (0,0,0),
                rect_bold: (0,0,0), rect_italic: (0,0,0), rect_under: (0,0,0),
                rect_dunder: (0,0,0), rect_strike: (0,0,0),
                rect_bg: (0,0,0), rect_bg_hex: (0,0,0), rect_fg: (0,0,0), rect_fg_hex: (0,0,0), rect_rules: (0,0,0,0),
                rect_apply: (0,0,0), rect_base: (0,0,0), rect_dismiss: (0,0,0), rect_delete: (0,0,0),
                rect_cleanall: (0,0,0), rect_close: (0,0,0),
            },
            cf_list: CfListDialog::default(),
            formula_error: None,
            search_buffer: String::new(),
            search_results: Vec::new(),
            search_idx: 0,
            search_active: false,
        };
        app.build_dep_graph();
        app.recalc_all();
        app
    }

    pub fn active_sheet(&self) -> &xlcli_core::sheet::Sheet {
        self.workbook.active_sheet()
    }

    pub fn move_cursor(&mut self, drow: i32, dcol: i32) {
        let new_row = (self.cursor.row as i64 + drow as i64).clamp(0, MAX_ROW as i64) as u32;
        let new_col = (self.cursor.col as i64 + dcol as i64).clamp(0, MAX_COL as i64) as u16;

        self.cursor.row = new_row;
        self.cursor.col = new_col;
    }

    pub fn jump_to_cell(&mut self, row: u32, col: u16) {
        self.cursor.row = row;
        self.cursor.col = col;
    }

    pub fn jump_to_start(&mut self) {
        self.cursor.row = 0;
        self.cursor.col = 0;
    }

    pub fn jump_to_last_row(&mut self) {
        let max_row = self.workbook.active_sheet().row_count().saturating_sub(1);
        self.cursor.row = max_row;
    }

    pub fn jump_to_col_start(&mut self) {
        self.cursor.col = 0;
    }

    pub fn jump_to_col_end(&mut self) {
        let max_col = self.workbook.active_sheet().col_count().saturating_sub(1);
        self.cursor.col = max_col;
    }

    // --- Search ---

    pub fn enter_search(&mut self) {
        self.mode = Mode::Search;
        self.search_buffer.clear();
    }

    pub fn execute_search(&mut self) {
        let q = self.search_buffer.to_lowercase();
        self.search_results.clear();
        self.search_idx = 0;
        if q.is_empty() {
            self.search_active = false;
            self.mode = Mode::Normal;
            return;
        }
        let sheet = self.workbook.active_sheet();
        let mut hits: Vec<(u32, u16)> = sheet
            .cells_iter()
            .filter_map(|(&(r, c), cell)| {
                let s = cell.value.display_value();
                if !s.is_empty() && s.to_lowercase().contains(&q) {
                    Some((r, c))
                } else {
                    None
                }
            })
            .collect();
        hits.sort_by(|a, b| a.0.cmp(&b.0).then(a.1.cmp(&b.1)));
        self.search_results = hits;
        self.mode = Mode::Normal;
        if self.search_results.is_empty() {
            self.search_active = false;
            self.status_message = Some(format!("Pattern not found: {}", self.search_buffer));
            return;
        }
        self.search_active = true;
        let cur = (self.cursor.row, self.cursor.col);
        self.search_idx = self
            .search_results
            .iter()
            .position(|p| *p >= cur)
            .unwrap_or(0);
        self.jump_to_search_match();
    }

    pub fn search_next(&mut self) {
        if !self.search_active || self.search_results.is_empty() {
            return;
        }
        self.search_idx = (self.search_idx + 1) % self.search_results.len();
        self.jump_to_search_match();
    }

    pub fn search_prev(&mut self) {
        if !self.search_active || self.search_results.is_empty() {
            return;
        }
        self.search_idx = if self.search_idx == 0 {
            self.search_results.len() - 1
        } else {
            self.search_idx - 1
        };
        self.jump_to_search_match();
    }

    fn jump_to_search_match(&mut self) {
        if let Some((r, c)) = self.search_results.get(self.search_idx).copied() {
            self.cursor.row = r;
            self.cursor.col = c;
            self.status_message = Some(format!(
                "/{}  [{}/{}]",
                self.search_buffer,
                self.search_idx + 1,
                self.search_results.len()
            ));
        }
    }

    pub fn clear_search(&mut self) {
        self.search_active = false;
        self.search_results.clear();
        self.search_buffer.clear();
    }

    // --- Editing ---

    pub fn enter_insert(&mut self) {
        let current = self
            .workbook
            .active_sheet()
            .get_cell(self.cursor.row, self.cursor.col)
            .map(|c| {
                if let Some(ref f) = c.formula {
                    format!("={}", f)
                } else {
                    c.value.display_value()
                }
            })
            .unwrap_or_default();
        self.edit_buffer = current;
        self.edit_cursor = self.edit_buffer.len();
        self.mode = Mode::Insert;
    }

    pub fn enter_insert_clear(&mut self) {
        self.edit_buffer.clear();
        self.edit_cursor = 0;
        self.mode = Mode::Insert;
    }

    pub fn enter_insert_with(&mut self, prefix: &str) {
        self.edit_buffer = prefix.to_string();
        self.edit_cursor = self.edit_buffer.len();
        self.mode = Mode::Insert;
        self.update_autocomplete();
    }

    pub fn update_autocomplete(&mut self) {
        self.autocomplete.active_function = None;

        if !self.edit_buffer.starts_with('=') || self.edit_buffer.len() < 2 {
            self.autocomplete.visible = false;
            self.autocomplete.matches.clear();
            self.autocomplete.selected = 0;
            return;
        }

        let after_eq = &self.edit_buffer[1..];

        if self.config.formula_autocomplete.show_signature {
            self.autocomplete.active_function = self.find_active_function(after_eq);
        }

        if !self.config.formula_autocomplete.enabled {
            self.autocomplete.visible = false;
            self.autocomplete.matches.clear();
            self.autocomplete.selected = 0;
            return;
        }

        let func_start = after_eq.rfind(|c: char| "+-*/^&=<>,(;".contains(c))
            .map(|i| i + 1)
            .unwrap_or(0);
        let partial = &after_eq[func_start..];

        if partial.is_empty() || partial.contains('(') || partial.chars().any(|c| c.is_ascii_digit() && func_start == 0 && partial.starts_with(c)) {
            self.autocomplete.visible = false;
            self.autocomplete.matches.clear();
            self.autocomplete.selected = 0;
            return;
        }

        let upper = partial.to_uppercase();
        let names = self.formula_registry.names();
        let matches: Vec<&'static str> = names.into_iter()
            .filter(|n| n.starts_with(&upper))
            .collect();

        if matches.is_empty() || (matches.len() == 1 && matches[0] == upper) {
            self.autocomplete.visible = false;
            self.autocomplete.matches.clear();
            self.autocomplete.selected = 0;
        } else {
            self.autocomplete.visible = true;
            self.autocomplete.matches = matches;
            if self.autocomplete.selected >= self.autocomplete.matches.len() {
                self.autocomplete.selected = 0;
            }
        }
    }

    fn find_active_function(&self, expr: &str) -> Option<String> {
        let mut depth = 0i32;
        for (i, c) in expr.char_indices().rev() {
            match c {
                ')' => depth += 1,
                '(' => {
                    if depth > 0 {
                        depth -= 1;
                    } else {
                        let before = &expr[..i];
                        let name_start = before.rfind(|c: char| !c.is_ascii_alphanumeric() && c != '.' && c != '_')
                            .map(|p| p + 1)
                            .unwrap_or(0);
                        let name = before[name_start..].to_uppercase();
                        if !name.is_empty() && self.formula_registry.get(&name).is_some() {
                            return Some(name);
                        }
                    }
                }
                _ => {}
            }
        }
        None
    }

    pub fn accept_autocomplete(&mut self) {
        if !self.autocomplete.visible || self.autocomplete.matches.is_empty() {
            return;
        }
        let chosen = self.autocomplete.matches[self.autocomplete.selected];
        let after_eq = &self.edit_buffer[1..];
        let func_start = after_eq.rfind(|c: char| "+-*/^&=<>,(;".contains(c))
            .map(|i| i + 1)
            .unwrap_or(0);

        let prefix = format!("={}", &after_eq[..func_start]);
        self.edit_buffer = format!("{}{}(", prefix, chosen);
        self.edit_cursor = self.edit_buffer.len();
        self.autocomplete.visible = false;
        self.autocomplete.matches.clear();
        self.autocomplete.selected = 0;
        self.update_autocomplete();
    }

    pub fn confirm_edit(&mut self) {
        // Return to origin sheet if we were browsing for cross-sheet refs
        if let Some(origin) = self.formula_origin_sheet.take() {
            self.save_sheet_state();
            self.workbook.active_sheet = origin as usize;
            self.restore_sheet_state(origin as usize);
        }
        let addr = self.cursor;
        let sheet = self.workbook.active_sheet_mut();
        let old = sheet.get_cell(addr.row, addr.col).cloned();

        let input = self.edit_buffer.trim().to_string();
        let new_cell = if input.is_empty() {
            None
        } else if let Some(formula_text) = input.strip_prefix('=') {
            let formula_text = formula_text.to_string();
            if let Ok(expr) = xlcli_formulas::parse(&formula_text) {
                let sheets = &self.workbook.sheets;
                let deps = extract_refs_with_resolver(&expr, addr.sheet, |name: &str| {
                    sheets.iter().enumerate()
                        .find(|(_, s)| s.name.eq_ignore_ascii_case(name))
                        .map(|(i, _)| i as u16)
                });
                self.dep_graph.set_dependencies(addr, deps);
            }
            let value = self.eval_formula(&formula_text, addr);
            Some(Cell::with_formula(value, formula_text))
        } else {
            self.dep_graph.remove_cell(addr);
            let value = parse_input_value(&input);
            if value.is_empty() { None } else { Some(Cell::new(value)) }
        };

        let sheet = self.workbook.active_sheet_mut();
        if let Some(ref cell) = new_cell {
            sheet.set_cell(addr.row, addr.col, cell.clone());
        } else {
            sheet.remove_cell(addr.row, addr.col);
        }

        self.undo_stack.push(UndoEntry::CellChange {
            addr,
            old,
            new: new_cell,
        });
        self.modified = true;
        self.edit_buffer.clear();
        self.edit_cursor = 0;
        self.mode = Mode::Normal;
        self.status_message = None;

        self.recalc_dependents(addr);
    }

    pub fn cancel_edit(&mut self) {
        if let Some(origin) = self.formula_origin_sheet.take() {
            self.save_sheet_state();
            self.workbook.active_sheet = origin as usize;
            self.restore_sheet_state(origin as usize);
        }
        self.edit_buffer.clear();
        self.edit_cursor = 0;
        self.mode = Mode::Normal;
        self.status_message = None;
    }

    // --- Formula evaluation ---

    fn eval_formula(&self, formula_text: &str, cell_addr: CellAddr) -> CellValue {
        match xlcli_formulas::parse(formula_text) {
            Ok(expr) => {
                let ctx = WorkbookEvalContext {
                    workbook: &self.workbook,
                    cell_addr,
                };
                xlcli_formulas::evaluate(&expr, &ctx, &self.formula_registry)
            }
            Err(_) => CellValue::Error(xlcli_core::types::CellError::Name),
        }
    }

    fn build_dep_graph(&mut self) {
        for sheet_idx in 0..self.workbook.sheets.len() {
            let cells_with_formulas: Vec<(u32, u16, String)> = self.workbook.sheets[sheet_idx]
                .cells_iter()
                .filter_map(|(&(row, col), cell)| {
                    cell.formula.as_ref().map(|f| (row, col, f.clone()))
                })
                .collect();

            for (row, col, formula) in cells_with_formulas {
                let addr = CellAddr::new(sheet_idx as u16, row, col);
                if let Ok(expr) = xlcli_formulas::parse(&formula) {
                    let sheets = &self.workbook.sheets;
                    let deps = extract_refs_with_resolver(&expr, addr.sheet, |name: &str| {
                        sheets.iter().enumerate()
                            .find(|(_, s)| s.name.eq_ignore_ascii_case(name))
                            .map(|(i, _)| i as u16)
                    });
                    self.dep_graph.set_dependencies(addr, deps);
                }
            }
        }
    }

    pub fn recalc_all(&mut self) {
        let sheet_count = self.workbook.sheets.len();
        for sheet_idx in 0..sheet_count {
            let cells_with_formulas: Vec<(u32, u16, String)> = self.workbook.sheets[sheet_idx]
                .cells_iter()
                .filter_map(|(&(row, col), cell)| {
                    cell.formula.as_ref().map(|f| (row, col, f.clone()))
                })
                .collect();

            for (row, col, formula) in cells_with_formulas {
                let addr = CellAddr::new(sheet_idx as u16, row, col);
                let value = self.eval_formula(&formula, addr);
                self.workbook.sheets[sheet_idx].set_cell(
                    row,
                    col,
                    Cell::with_formula(value, formula),
                );
            }
        }
    }

    fn recalc_dependents(&mut self, changed: CellAddr) {
        let to_recalc = self.dep_graph.dependents_toposorted(changed);
        for addr in to_recalc {
            let formula = self.workbook.sheets.get(addr.sheet as usize)
                .and_then(|s| s.get_cell(addr.row, addr.col))
                .and_then(|c| c.formula.clone());

            if let Some(formula) = formula {
                let value = self.eval_formula(&formula, addr);
                if let Some(sheet) = self.workbook.sheets.get_mut(addr.sheet as usize) {
                    sheet.set_cell(addr.row, addr.col, Cell::with_formula(value, formula));
                }
            }
        }
    }

    // --- Undo/Redo ---

    pub fn undo(&mut self) {
        if let Some(entry) = self.undo_stack.undo() {
            self.apply_undo_entry(&entry);
            self.modified = true;
            self.status_message = Some("Undo".to_string());
            self.recalc_all();
        } else {
            self.status_message = Some("Nothing to undo".to_string());
        }
    }

    pub fn redo(&mut self) {
        if let Some(entry) = self.undo_stack.redo() {
            self.apply_redo_entry(&entry);
            self.modified = true;
            self.status_message = Some("Redo".to_string());
            self.recalc_all();
        } else {
            self.status_message = Some("Nothing to redo".to_string());
        }
    }

    fn apply_undo_entry(&mut self, entry: &UndoEntry) {
        match entry {
            UndoEntry::CellChange { addr, old, .. } => {
                let sheet = &mut self.workbook.sheets[addr.sheet as usize];
                if let Some(cell) = old {
                    sheet.set_cell(addr.row, addr.col, cell.clone());
                } else {
                    sheet.remove_cell(addr.row, addr.col);
                }
            }
            UndoEntry::Batch(entries) => {
                for e in entries.iter().rev() {
                    self.apply_undo_entry(e);
                }
            }
        }
    }

    fn apply_redo_entry(&mut self, entry: &UndoEntry) {
        match entry {
            UndoEntry::CellChange { addr, new, .. } => {
                let sheet = &mut self.workbook.sheets[addr.sheet as usize];
                if let Some(cell) = new {
                    sheet.set_cell(addr.row, addr.col, cell.clone());
                } else {
                    sheet.remove_cell(addr.row, addr.col);
                }
            }
            UndoEntry::Batch(entries) => {
                for e in entries {
                    self.apply_redo_entry(e);
                }
            }
        }
    }

    // --- Clipboard ---

    pub fn yank(&mut self) {
        let cell = self
            .workbook
            .active_sheet()
            .get_cell(self.cursor.row, self.cursor.col);
        self.clipboard.yank_single(cell, self.cursor.row, self.cursor.col);
        self.status_message = Some("Yanked 1 cell".to_string());
    }

    pub fn yank_visual(&mut self) {
        if let Some(anchor) = self.visual_anchor {
            let min_row = anchor.row.min(self.cursor.row);
            let max_row = anchor.row.max(self.cursor.row);
            let min_col = anchor.col.min(self.cursor.col);
            let max_col = anchor.col.max(self.cursor.col);

            let sheet = self.workbook.active_sheet();
            let mut cells = Vec::new();
            for r in min_row..=max_row {
                for c in min_col..=max_col {
                    if let Some(cell) = sheet.get_cell(r, c) {
                        cells.push((r - min_row, (c - min_col), cell.clone()));
                    }
                }
            }
            let rows = max_row - min_row + 1;
            let cols = max_col - min_col + 1;
            self.clipboard.yank_range(cells, rows, cols, min_row, min_col);
            let count = rows as usize * cols as usize;
            self.status_message = Some(format!("Yanked {} cells", count));
        }
        self.visual_anchor = None;
        self.mode = Mode::Normal;
    }

    pub fn paste(&mut self) {
        let relative = self.clipboard.relative;
        let origin_row = self.clipboard.origin_row;
        let origin_col = self.clipboard.origin_col;

        if let Some(ref clip) = self.clipboard.content.clone() {
            let mut undo_entries = Vec::new();
            let base_row = self.cursor.row;
            let base_col = self.cursor.col;
            let sheet_idx = self.workbook.active_sheet as u16;

            for &(rel_row, rel_col, ref cell) in &clip.cells {
                let row = base_row + rel_row;
                let col = base_col + rel_col;
                let addr = CellAddr::new(sheet_idx, row, col);

                let paste_cell = if relative {
                    if let Some(ref formula) = cell.formula {
                        let drow = row as i32 - (origin_row + rel_row) as i32;
                        let dcol = col as i32 - (origin_col + rel_col) as i32;
                        if let Some(adjusted) = xlcli_formulas::adjust_formula(formula, drow, dcol) {
                            let value = self.eval_formula(&adjusted, addr);
                            Cell::with_formula(value, adjusted)
                        } else {
                            cell.clone()
                        }
                    } else {
                        cell.clone()
                    }
                } else {
                    cell.clone()
                };

                let sheet = self.workbook.active_sheet_mut();
                let old = sheet.get_cell(row, col).cloned();
                sheet.set_cell(row, col, paste_cell.clone());

                undo_entries.push(UndoEntry::CellChange {
                    addr,
                    old,
                    new: Some(paste_cell),
                });
            }

            if undo_entries.len() == 1 {
                self.undo_stack.push(undo_entries.pop().unwrap());
            } else {
                self.undo_stack.push(UndoEntry::Batch(undo_entries));
            }
            self.modified = true;
            let mode_str = if relative { " (relative)" } else { "" };
            self.status_message = Some(format!("Pasted {}x{}{}", clip.rows + 1, clip.cols + 1, mode_str));
            self.recalc_all();
        } else {
            self.status_message = Some("Nothing to paste".to_string());
        }
    }

    pub fn delete_cell(&mut self) {
        let addr = self.cursor;
        let sheet = self.workbook.active_sheet_mut();
        let old = sheet.get_cell(addr.row, addr.col).cloned();

        if old.is_some() {
            sheet.remove_cell(addr.row, addr.col);
            self.undo_stack.push(UndoEntry::CellChange {
                addr,
                old,
                new: None,
            });
            self.modified = true;
            self.status_message = Some("Deleted".to_string());
            self.recalc_all();
        }
    }

    pub fn delete_visual(&mut self) {
        if let Some(anchor) = self.visual_anchor {
            let min_row = anchor.row.min(self.cursor.row);
            let max_row = anchor.row.max(self.cursor.row);
            let min_col = anchor.col.min(self.cursor.col);
            let max_col = anchor.col.max(self.cursor.col);
            let sheet_idx = self.workbook.active_sheet as u16;

            let mut undo_entries = Vec::new();
            let sheet = self.workbook.active_sheet_mut();
            for r in min_row..=max_row {
                for c in min_col..=max_col {
                    let old = sheet.get_cell(r, c).cloned();
                    if old.is_some() {
                        sheet.remove_cell(r, c);
                        undo_entries.push(UndoEntry::CellChange {
                            addr: CellAddr::new(sheet_idx, r, c),
                            old,
                            new: None,
                        });
                    }
                }
            }

            if !undo_entries.is_empty() {
                self.undo_stack.push(UndoEntry::Batch(undo_entries));
                self.modified = true;
            }
            self.status_message = Some("Deleted range".to_string());
            self.recalc_all();
        }
        self.visual_anchor = None;
        self.mode = Mode::Normal;
    }

    // --- Row/Col Insert/Delete ---

    pub fn insert_row_above(&mut self) {
        let row = self.cursor.row;
        self.workbook.active_sheet_mut().insert_row(row);
        self.modified = true;
        self.status_message = Some(format!("Inserted row above {}", row + 1));
    }

    pub fn insert_row_below(&mut self) {
        let row = self.cursor.row + 1;
        self.workbook.active_sheet_mut().insert_row(row);
        self.modified = true;
        self.cursor.row = row;
        self.status_message = Some(format!("Inserted row below"));
    }

    pub fn delete_row(&mut self) {
        let row = self.cursor.row;
        self.workbook.active_sheet_mut().delete_row(row);
        self.modified = true;
        let max = self.workbook.active_sheet().row_count().saturating_sub(1);
        if self.cursor.row > max {
            self.cursor.row = max;
        }
        self.status_message = Some(format!("Deleted row {}", row + 1));
    }

    pub fn insert_col_left(&mut self) {
        let col = self.cursor.col;
        self.workbook.active_sheet_mut().insert_col(col);
        self.modified = true;
        self.status_message = Some(format!("Inserted column at {}", CellAddr::col_name(col)));
    }

    pub fn insert_col_right(&mut self) {
        let col = self.cursor.col + 1;
        self.workbook.active_sheet_mut().insert_col(col);
        self.modified = true;
        self.cursor.col = col;
        self.status_message = Some(format!("Inserted column at {}", CellAddr::col_name(col)));
    }

    pub fn delete_col(&mut self) {
        let col = self.cursor.col;
        self.workbook.active_sheet_mut().delete_col(col);
        self.modified = true;
        let max = self.workbook.active_sheet().col_count().saturating_sub(1);
        if self.cursor.col > max {
            self.cursor.col = max;
        }
        self.status_message = Some(format!("Deleted column {}", CellAddr::col_name(col)));
    }

    // --- Freeze ---

    fn parse_freeze(&mut self, args: &str) {
        let parts: Vec<&str> = args.split_whitespace().collect();
        if parts.is_empty() || parts.len() > 2 {
            self.status_message = Some("Usage: :freeze <rows> [cols] (e.g. :freeze 1, :freeze 1 A, :freeze 1-3 B-C)".to_string());
            return;
        }

        // Parse row spec: "1" or "1-3"
        let freeze_rows = if let Some((start, end)) = parts[0].split_once('-') {
            let s: u32 = match start.parse() {
                Ok(v) if v >= 1 => v,
                _ => { self.status_message = Some("Invalid row number".to_string()); return; }
            };
            let e: u32 = match end.parse() {
                Ok(v) if v >= s => v,
                _ => { self.status_message = Some("Invalid row range".to_string()); return; }
            };
            e
        } else {
            match parts[0].parse::<u32>() {
                Ok(v) if v >= 1 => v,
                _ => { self.status_message = Some("Invalid row number".to_string()); return; }
            }
        };

        // Parse col spec (optional): "A" or "B-C"
        let freeze_cols = if parts.len() == 2 {
            if let Some((start, end)) = parts[1].split_once('-') {
                let _s = match CellAddr::parse_col(&start.to_uppercase()) {
                    Some(v) => v,
                    None => { self.status_message = Some("Invalid column".to_string()); return; }
                };
                let e = match CellAddr::parse_col(&end.to_uppercase()) {
                    Some(v) => v,
                    None => { self.status_message = Some("Invalid column range".to_string()); return; }
                };
                e + 1
            } else {
                match CellAddr::parse_col(&parts[1].to_uppercase()) {
                    Some(v) => v + 1,
                    None => { self.status_message = Some("Invalid column".to_string()); return; }
                }
            }
        } else {
            0
        };

        self.workbook.active_sheet_mut().freeze = Some((freeze_rows, freeze_cols));
        if freeze_cols > 0 {
            self.status_message = Some(format!("Frozen {} rows, {} cols", freeze_rows, freeze_cols));
        } else {
            self.status_message = Some(format!("Frozen {} rows", freeze_rows));
        }
    }

    // --- Fill ---

    pub fn fill_down(&mut self) {
        if self.cursor.row == 0 {
            self.status_message = Some("No cell above".to_string());
            return;
        }
        let src_row = self.cursor.row - 1;
        let col = self.cursor.col;
        self.fill_from(src_row, col, self.cursor.row, col);
    }

    pub fn fill_right(&mut self) {
        if self.cursor.col == 0 {
            self.status_message = Some("No cell to the left".to_string());
            return;
        }
        let src_col = self.cursor.col - 1;
        let row = self.cursor.row;
        self.fill_from(row, src_col, row, self.cursor.col);
    }

    fn fill_from(&mut self, src_row: u32, src_col: u16, dst_row: u32, dst_col: u16) {
        let sheet = self.workbook.active_sheet();
        let src_cell = sheet.get_cell(src_row, src_col).cloned();

        let new_cell = match src_cell {
            Some(ref cell) => {
                if let Some(ref formula) = cell.formula {
                    let drow = dst_row as i32 - src_row as i32;
                    let dcol = dst_col as i32 - src_col as i32;
                    if let Some(adjusted) = xlcli_formulas::adjust_formula(formula, drow, dcol) {
                        let sheet_idx = self.workbook.active_sheet as u16;
                        let addr = CellAddr::new(sheet_idx, dst_row, dst_col);
                        let value = self.eval_formula(&adjusted, addr);
                        Some(Cell::with_formula(value, adjusted))
                    } else {
                        Some(cell.clone())
                    }
                } else {
                    Some(cell.clone())
                }
            }
            None => None,
        };

        let sheet_idx = self.workbook.active_sheet as u16;
        let addr = CellAddr::new(sheet_idx, dst_row, dst_col);
        let sheet = self.workbook.active_sheet_mut();
        let old = sheet.get_cell(dst_row, dst_col).cloned();

        if let Some(ref cell) = new_cell {
            sheet.set_cell(dst_row, dst_col, cell.clone());
        } else {
            sheet.remove_cell(dst_row, dst_col);
        }

        self.undo_stack.push(UndoEntry::CellChange { addr, old, new: new_cell });
        self.modified = true;
        self.status_message = Some("Filled".to_string());
        self.recalc_dependents(addr);
    }

    pub fn visual_fill(&mut self) {
        if let Some(anchor) = self.visual_anchor {
            let min_row = anchor.row.min(self.cursor.row);
            let max_row = anchor.row.max(self.cursor.row);
            let min_col = anchor.col.min(self.cursor.col);
            let max_col = anchor.col.max(self.cursor.col);
            let sheet_idx = self.workbook.active_sheet as u16;

            // Source is top-left cell of selection
            let src_row = min_row;
            let src_col = min_col;
            let sheet = self.workbook.active_sheet();
            let src_cell = sheet.get_cell(src_row, src_col).cloned();

            let mut undo_entries = Vec::new();

            for r in min_row..=max_row {
                for c in min_col..=max_col {
                    if r == src_row && c == src_col {
                        continue;
                    }
                    let new_cell = match src_cell {
                        Some(ref cell) => {
                            if let Some(ref formula) = cell.formula {
                                let drow = r as i32 - src_row as i32;
                                let dcol = c as i32 - src_col as i32;
                                if let Some(adjusted) = xlcli_formulas::adjust_formula(formula, drow, dcol) {
                                    let addr = CellAddr::new(sheet_idx, r, c);
                                    let value = self.eval_formula(&adjusted, addr);
                                    Some(Cell::with_formula(value, adjusted))
                                } else {
                                    Some(cell.clone())
                                }
                            } else {
                                Some(cell.clone())
                            }
                        }
                        None => None,
                    };

                    let addr = CellAddr::new(sheet_idx, r, c);
                    let sheet = self.workbook.active_sheet_mut();
                    let old = sheet.get_cell(r, c).cloned();

                    if let Some(ref cell) = new_cell {
                        sheet.set_cell(r, c, cell.clone());
                    } else {
                        sheet.remove_cell(r, c);
                    }

                    undo_entries.push(UndoEntry::CellChange { addr, old, new: new_cell });
                }
            }

            if !undo_entries.is_empty() {
                self.undo_stack.push(UndoEntry::Batch(undo_entries));
                self.modified = true;
                self.recalc_all();
            }

            let count = (max_row - min_row + 1) * (max_col - min_col + 1) as u32;
            self.status_message = Some(format!("Filled {} cells", count));
        }
        self.visual_anchor = None;
        self.mode = Mode::Normal;
    }

    // --- Sort ---

    pub fn execute_sort(&mut self, keys: &[SortKey], min_row: u32, max_row: u32, min_col: u16, max_col: u16, has_headers: bool) {
        let data_start = if has_headers { min_row + 1 } else { min_row };
        if data_start > max_row {
            self.status_message = Some("No data to sort".to_string());
            return;
        }

        let sheet = self.workbook.active_sheet();
        let mut row_indices: Vec<u32> = (data_start..=max_row).collect();
        let sheet_idx = self.workbook.active_sheet as u16;

        // Collect sort values for each key column
        let key_values: Vec<Vec<CellValue>> = keys.iter().map(|k| {
            (data_start..=max_row)
                .map(|r| sheet.get_cell_value(r, k.col).clone())
                .collect()
        }).collect();

        // Sort row indices
        row_indices.sort_by(|&a, &b| {
            let ai = (a - data_start) as usize;
            let bi = (b - data_start) as usize;
            for (ki, key) in keys.iter().enumerate() {
                let va = &key_values[ki][ai];
                let vb = &key_values[ki][bi];
                let ord = compare_cells(va, vb, key.sort_type);
                if ord != std::cmp::Ordering::Equal {
                    return ord;
                }
            }
            std::cmp::Ordering::Equal
        });

        // Build old and new cell data for undo
        let mut undo_entries = Vec::new();
        let sheet = self.workbook.active_sheet();
        let mut all_rows: Vec<Vec<Option<Cell>>> = Vec::new();
        for &ri in &row_indices {
            let mut row_cells = Vec::new();
            for c in min_col..=max_col {
                row_cells.push(sheet.get_cell(ri, c).cloned());
            }
            all_rows.push(row_cells);
        }

        // Write sorted rows back
        let sheet = self.workbook.active_sheet_mut();
        for (dest_idx, row_cells) in all_rows.iter().enumerate() {
            let dest_row = data_start + dest_idx as u32;
            for (col_offset, cell_opt) in row_cells.iter().enumerate() {
                let c = min_col + col_offset as u16;
                let addr = CellAddr::new(sheet_idx, dest_row, c);
                let old = sheet.get_cell(dest_row, c).cloned();
                if let Some(cell) = cell_opt {
                    sheet.set_cell(dest_row, c, cell.clone());
                } else {
                    sheet.remove_cell(dest_row, c);
                }
                undo_entries.push(UndoEntry::CellChange { addr, old, new: cell_opt.clone() });
            }
        }

        if !undo_entries.is_empty() {
            self.undo_stack.push(UndoEntry::Batch(undo_entries));
            self.modified = true;
        }
        let row_count = max_row - data_start + 1;
        self.status_message = Some(format!("Sorted {} rows", row_count));
        self.recalc_all();
    }

    pub fn open_sort_dialog(&mut self) {
        let (min_row, max_row, min_col, max_col) = if let Some(anchor) = self.visual_anchor {
            (
                anchor.row.min(self.cursor.row),
                anchor.row.max(self.cursor.row),
                anchor.col.min(self.cursor.col),
                anchor.col.max(self.cursor.col),
            )
        } else {
            if !self.config.sort.allow_full_sheet {
                self.status_message = Some("Select a range first (full-sheet sort disabled)".to_string());
                return;
            }
            let sheet = self.workbook.active_sheet();
            let rc = sheet.row_count().saturating_sub(1);
            let cc = sheet.col_count().saturating_sub(1);
            (0, rc, 0, cc)
        };

        let start_name = CellAddr::new(0, min_row, min_col).display_name();
        let end_name = CellAddr::new(0, max_row, max_col).display_name();

        self.sort_dialog = SortDialog {
            visible: true,
            focus: SortDialogFocus::SortByCol,
            range_text: format!("{}:{}", start_name, end_name),
            min_row, max_row, min_col, max_col,
            sort_by_col: min_col,
            sort_by_type: SortType::AZ,
            then_by_col: None,
            then_by_type: SortType::AZ,
            has_headers: false,
            open_dropdown: SortDropdown::None,
            dropdown_scroll: 0,
            screen_x: 0, screen_y: 0, screen_w: 42, screen_h: 11,
            dd_item_x: 0, dd_item_y: 0, dd_item_w: 16, dd_item_count: 0,
        };
        self.visual_anchor = None;
        self.mode = Mode::Normal;
    }

    pub fn confirm_sort_dialog(&mut self) {
        let d = &self.sort_dialog;
        let mut keys = vec![SortKey { col: d.sort_by_col, sort_type: d.sort_by_type }];
        if let Some(col) = d.then_by_col {
            keys.push(SortKey { col, sort_type: d.then_by_type });
        }
        let (min_row, max_row, min_col, max_col, has_headers) =
            (d.min_row, d.max_row, d.min_col, d.max_col, d.has_headers);
        self.sort_dialog.visible = false;
        self.execute_sort(&keys, min_row, max_row, min_col, max_col, has_headers);
    }

    pub fn parse_and_sort(&mut self, args: &str) {
        let (min_row, max_row, min_col, max_col) = if let Some(anchor) = self.visual_anchor {
            self.visual_anchor = None;
            self.mode = Mode::Normal;
            (
                anchor.row.min(self.cursor.row),
                anchor.row.max(self.cursor.row),
                anchor.col.min(self.cursor.col),
                anchor.col.max(self.cursor.col),
            )
        } else {
            if !self.config.sort.allow_full_sheet {
                self.status_message = Some("Select a range first (full-sheet sort disabled)".to_string());
                return;
            }
            let sheet = self.workbook.active_sheet();
            let rc = sheet.row_count().saturating_sub(1);
            let cc = sheet.col_count().saturating_sub(1);
            (0, rc, 0, cc)
        };

        let mut keys = Vec::new();
        let parts: Vec<&str> = args.split_whitespace().collect();
        let mut i = 0;
        while i < parts.len() {
            let col_str = parts[i].to_uppercase();
            if let Some(col) = CellAddr::parse_col(&col_str) {
                let sort_type = if i + 1 < parts.len() {
                    match parts[i + 1].to_lowercase().as_str() {
                        "desc" => { i += 1; SortType::ZA }
                        "num" => { i += 1; SortType::Numeric }
                        "numdesc" => { i += 1; SortType::NumericDesc }
                        "case" => { i += 1; SortType::CaseSensitive }
                        "natural" => { i += 1; SortType::Natural }
                        _ => SortType::AZ,
                    }
                } else {
                    SortType::AZ
                };
                keys.push(SortKey { col, sort_type });
            } else {
                self.status_message = Some(format!("Invalid column: {}", parts[i]));
                return;
            }
            i += 1;
        }

        if keys.is_empty() {
            self.status_message = Some("Usage: :sort <col> [desc|num|case|natural] [col2 ...]".to_string());
            return;
        }

        self.execute_sort(&keys, min_row, max_row, min_col, max_col, false);
    }

    // --- Sheet Management ---

    pub fn switch_sheet(&mut self, idx: usize) {
        if idx >= self.workbook.sheets.len() {
            self.status_message = Some(format!("No sheet {}", idx + 1));
            return;
        }
        if idx == self.workbook.active_sheet {
            return;
        }
        self.save_sheet_state();
        self.workbook.active_sheet = idx;
        self.restore_sheet_state(idx);
        self.status_message = Some(format!("Sheet: {}", self.workbook.sheets[idx].name));
    }

    pub fn add_sheet(&mut self) {
        let mut n = self.workbook.sheets.len() + 1;
        loop {
            let name = format!("Sheet{}", n);
            if self.workbook.sheet_by_name(&name).is_none() {
                let idx = self.workbook.add_sheet(name.clone());
                self.save_sheet_state();
                self.workbook.active_sheet = idx;
                self.cursor = CellAddr::new(idx as u16, 0, 0);
                self.viewport = Viewport::new();
                self.status_message = Some(format!("Added {}", name));
                self.modified = true;
                return;
            }
            n += 1;
        }
    }

    pub fn add_sheet_with_name(&mut self, name: &str) {
        if name.is_empty() {
            self.add_sheet();
            return;
        }
        if self.workbook.sheet_by_name(name).is_some() {
            self.status_message = Some(format!("Sheet '{}' already exists", name));
            return;
        }
        let idx = self.workbook.add_sheet(name);
        self.save_sheet_state();
        self.workbook.active_sheet = idx;
        self.cursor = CellAddr::new(idx as u16, 0, 0);
        self.viewport = Viewport::new();
        self.status_message = Some(format!("Added {}", name));
        self.modified = true;
    }

    pub fn delete_sheet(&mut self, idx: usize) {
        if self.workbook.sheets.len() <= 1 {
            self.status_message = Some("Cannot delete last sheet".to_string());
            return;
        }
        if idx >= self.workbook.sheets.len() {
            self.status_message = Some(format!("No sheet {}", idx + 1));
            return;
        }
        let name = self.workbook.sheets[idx].name.clone();
        self.sheet_states.remove(&idx);
        self.workbook.sheets.remove(idx);

        // Rekey sheet_states for indices above removed
        let mut new_states = HashMap::new();
        for (k, v) in self.sheet_states.drain() {
            if k > idx {
                new_states.insert(k - 1, v);
            } else {
                new_states.insert(k, v);
            }
        }
        self.sheet_states = new_states;

        if self.workbook.active_sheet >= self.workbook.sheets.len() {
            self.workbook.active_sheet = self.workbook.sheets.len() - 1;
        }
        self.restore_sheet_state(self.workbook.active_sheet);
        self.modified = true;
        self.status_message = Some(format!("Deleted {}", name));
    }

    pub fn rename_sheet(&mut self, old_name: &str, new_name: &str) {
        if let Some((idx, _)) = self.workbook.sheet_by_name(old_name) {
            if self.workbook.sheet_by_name(new_name).is_some() {
                self.status_message = Some(format!("Sheet '{}' already exists", new_name));
                return;
            }
            self.workbook.sheets[idx].name = new_name.to_string();
            self.modified = true;
            self.status_message = Some(format!("Renamed '{}' → '{}'", old_name, new_name));
        } else {
            self.status_message = Some(format!("No sheet '{}'", old_name));
        }
    }

    pub fn move_sheet(&mut self, from: usize, to: usize) {
        let len = self.workbook.sheets.len();
        if from >= len || to >= len {
            self.status_message = Some("Invalid sheet position".to_string());
            return;
        }
        if from == to {
            return;
        }
        let sheet = self.workbook.sheets.remove(from);
        self.workbook.sheets.insert(to, sheet);

        // Fix active_sheet index
        if self.workbook.active_sheet == from {
            self.workbook.active_sheet = to;
        } else if from < self.workbook.active_sheet && to >= self.workbook.active_sheet {
            self.workbook.active_sheet -= 1;
        } else if from > self.workbook.active_sheet && to <= self.workbook.active_sheet {
            self.workbook.active_sheet += 1;
        }

        // Rebuild sheet_states keys
        self.sheet_states.clear();
        self.modified = true;
        self.status_message = Some(format!("Moved sheet to position {}", to + 1));
    }

    fn save_sheet_state(&mut self) {
        let idx = self.workbook.active_sheet;
        self.sheet_states.insert(idx, SheetState {
            cursor: self.cursor,
            viewport: Viewport {
                top_row: self.viewport.top_row,
                left_col: self.viewport.left_col,
                visible_rows: self.viewport.visible_rows,
                visible_cols: self.viewport.visible_cols,
                col_width: self.viewport.col_width,
            },
        });
    }

    fn restore_sheet_state(&mut self, idx: usize) {
        if let Some(state) = self.sheet_states.get(&idx) {
            self.cursor = state.cursor;
            self.viewport.top_row = state.viewport.top_row;
            self.viewport.left_col = state.viewport.left_col;
        } else {
            self.cursor = CellAddr::new(idx as u16, 0, 0);
            self.viewport.top_row = 0;
            self.viewport.left_col = 0;
        }
    }

    pub fn switch_sheet_for_ref(&mut self, idx: usize) {
        if idx >= self.workbook.sheets.len() {
            self.status_message = Some(format!("No sheet {}", idx + 1));
            return;
        }
        if self.formula_origin_sheet.is_none() {
            self.formula_origin_sheet = Some(self.workbook.active_sheet as u16);
        }
        self.save_sheet_state();
        self.workbook.active_sheet = idx;
        self.restore_sheet_state(idx);
        self.status_message = Some(format!("Ref: {}", self.workbook.sheets[idx].name));
    }

    // --- Commands ---

    pub fn execute_command(&mut self) {
        let cmd = self.command_buffer.trim().to_string();
        self.command_buffer.clear();
        self.mode = Mode::Normal;

        if cmd == "q" || cmd == "quit" {
            if self.modified {
                self.status_message =
                    Some("Unsaved changes! Use :q! to force quit or :wq to save and quit".to_string());
            } else {
                self.should_quit = true;
            }
        } else if cmd == "q!" {
            self.should_quit = true;
        } else if cmd == "w" {
            self.save_file(None);
        } else if cmd.starts_with("w ") {
            let path = cmd[2..].trim().to_string();
            self.save_file(Some(path));
        } else if cmd == "wq" {
            self.save_file(None);
            if self.status_message.as_deref() != Some("Save failed") {
                self.should_quit = true;
            }
        } else if cmd.starts_with("c ") {
            let addr = cmd[2..].trim();
            if let Some((row, col)) = parse_cell_address(addr) {
                self.jump_to_cell(row, col);
                self.status_message = Some(format!("Jumped to {}", addr.to_uppercase()));
            } else {
                self.status_message = Some(format!("Invalid cell address: {}", addr));
            }
        } else if cmd == "ir" || cmd == "insert row" {
            self.insert_row_above();
        } else if cmd == "ir!" || cmd == "insert row below" {
            self.insert_row_below();
        } else if cmd == "dr" || cmd == "delete row" {
            self.delete_row();
        } else if cmd == "ic" || cmd == "insert col" {
            self.insert_col_left();
        } else if cmd == "ic!" || cmd == "insert col right" {
            self.insert_col_right();
        } else if cmd == "dc" || cmd == "delete col" {
            self.delete_col();
        } else if cmd == "unfreeze" {
            self.workbook.active_sheet_mut().freeze = None;
            self.status_message = Some("Unfrozen".to_string());
        } else if let Some(rest) = cmd.strip_prefix("freeze ") {
            self.parse_freeze(rest.trim());
        } else if cmd == "sort" {
            self.open_sort_dialog();
        } else if let Some(rest) = cmd.strip_prefix("sort ") {
            self.parse_and_sort(rest.trim());
        } else if let Some(rest) = cmd.strip_prefix("sheet ") {
            self.handle_sheet_command(rest.trim());
        } else if cmd == "headers" {
            if self.visual_anchor.is_some() {
                self.set_headers_from_visual();
            } else {
                self.apply_header_row(0);
            }
        } else if let Some(rest) = cmd.strip_prefix("headers ") {
            self.set_headers(rest.trim());
        } else if cmd == "unheaders" {
            let sheet = self.workbook.active_sheet_mut();
            if let Some(h) = sheet.header_row {
                // Remove auto-freeze if it matches header row
                if sheet.freeze.map(|(r, _)| r) == Some(h + 1) {
                    let freeze_cols = sheet.freeze.map(|(_, c)| c).unwrap_or(0);
                    if freeze_cols == 0 {
                        sheet.freeze = None;
                    } else {
                        sheet.freeze = Some((0, freeze_cols));
                    }
                }
            }
            self.workbook.active_sheet_mut().header_row = None;
            self.status_message = Some("Headers removed".to_string());
        } else if cmd == "filter" {
            self.open_filter_dialog(None);
        } else if cmd == "filter all" {
            self.set_filter_range_all();
            self.open_filter_dialog(None);
        } else if let Some(rest) = cmd.strip_prefix("filter ") {
            self.parse_and_filter(rest.trim());
        } else if cmd == "unfilter" {
            self.unfilter_visual();
        } else if cmd == "unfilter all" {
            self.unfilter_all();
        } else if let Some(rest) = cmd.strip_prefix("unfilter ") {
            self.unfilter_column(rest.trim());
        } else if cmd == "names" {
            self.list_named_ranges();
        } else if let Some(rest) = cmd.strip_prefix("name ") {
            self.handle_name_command(rest.trim());
        } else if cmd == "cf" {
            self.open_cf_dialog();
        } else if cmd == "cf list" {
            self.cf_list();
        } else if let Some(rest) = cmd.strip_prefix("cf ") {
            self.handle_cf_command(rest.trim());
        } else if let Some(rest) = cmd.strip_prefix("case ") {
            self.handle_case_command(rest.trim());
        } else {
            self.status_message = Some(format!("Unknown command: {}", cmd));
        }
    }

    fn handle_name_command(&mut self, rest: &str) {
        if let Some(name_to_del) = rest.strip_prefix("delete ") {
            let name_to_del = name_to_del.trim();
            let before = self.workbook.named_ranges.len();
            self.workbook.named_ranges.retain(|nr| !nr.name.eq_ignore_ascii_case(name_to_del));
            if self.workbook.named_ranges.len() < before {
                self.status_message = Some(format!("Deleted name: {}", name_to_del));
            } else {
                self.status_message = Some(format!("Name not found: {}", name_to_del));
            }
            return;
        }
        let parts: Vec<&str> = rest.splitn(2, ' ').collect();
        let sheet = self.workbook.active_sheet;
        // Single-arg + visual: use selection
        if parts.len() == 1 {
            if let Some(anchor) = self.visual_anchor {
                let name = parts[0].trim().to_string();
                if name.is_empty() {
                    self.status_message = Some("Usage: :name <Name>".to_string());
                    return;
                }
                let r1 = anchor.row.min(self.cursor.row);
                let r2 = anchor.row.max(self.cursor.row);
                let c1 = anchor.col.min(self.cursor.col);
                let c2 = anchor.col.max(self.cursor.col);
                let start = CellAddr::new(sheet as u16, r1, c1);
                let end = CellAddr::new(sheet as u16, r2, c2);
                self.workbook.named_ranges.retain(|nr| !nr.name.eq_ignore_ascii_case(&name));
                self.workbook.named_ranges.push(xlcli_core::range::NamedRange {
                    name: name.clone(),
                    range: xlcli_core::range::CellRange::new(start, end),
                });
                self.visual_anchor = None;
                self.mode = Mode::Normal;
                self.status_message = Some(format!("Named: {}", name));
                return;
            }
            self.status_message = Some("Usage: :name <Name> <A1:B10> (or select range in visual mode)".to_string());
            return;
        }
        let name = parts[0].trim().to_string();
        let range_str = parts[1].trim();
        let range_part = if let Some(idx) = range_str.find('!') {
            let sheet_name = &range_str[..idx];
            let sheet_name = sheet_name.trim_matches('\'');
            match self.workbook.sheets.iter().position(|s| s.name.eq_ignore_ascii_case(sheet_name)) {
                Some(_i) => {
                    // sheet override currently same-sheet only; store resolved below
                }
                None => {
                    self.status_message = Some(format!("Unknown sheet: {}", sheet_name));
                    return;
                }
            }
            &range_str[idx+1..]
        } else {
            range_str
        };
        let (start, end) = match range_part.find(':') {
            Some(i) => {
                let a = parse_cell_address(&range_part[..i]);
                let b = parse_cell_address(&range_part[i+1..]);
                match (a, b) {
                    (Some((r1, c1)), Some((r2, c2))) => (
                        CellAddr::new(sheet as u16, r1, c1),
                        CellAddr::new(sheet as u16, r2, c2),
                    ),
                    _ => {
                        self.status_message = Some("Invalid range".to_string());
                        return;
                    }
                }
            }
            None => match parse_cell_address(range_part) {
                Some((r, c)) => {
                    let a = CellAddr::new(sheet as u16, r, c);
                    (a, a)
                }
                None => {
                    self.status_message = Some("Invalid cell".to_string());
                    return;
                }
            },
        };
        self.workbook.named_ranges.retain(|nr| !nr.name.eq_ignore_ascii_case(&name));
        self.workbook.named_ranges.push(xlcli_core::range::NamedRange {
            name: name.clone(),
            range: xlcli_core::range::CellRange::new(start, end),
        });
        self.status_message = Some(format!("Named: {}", name));
    }

    pub fn open_cf_dialog(&mut self) {
        let range_text = if let Some(anchor) = self.visual_anchor {
            let r1 = anchor.row.min(self.cursor.row);
            let r2 = anchor.row.max(self.cursor.row);
            let c1 = anchor.col.min(self.cursor.col);
            let c2 = anchor.col.max(self.cursor.col);
            format!("{}{}:{}{}",
                CellAddr::col_name(c1), r1 + 1,
                CellAddr::col_name(c2), r2 + 1)
        } else {
            format!("{}{}", CellAddr::col_name(self.cursor.col), self.cursor.row + 1)
        };
        self.cf_dialog = CfDialog {
            visible: true,
            focus: CfDialogFocus::Range,
            editing_text: false,
            conditional: false,
            range_text,
            cond: CfCond::Gt,
            val1: String::new(),
            val2: String::new(),
            bold: false, italic: false, under: false, dunder: false, strike: false,
            bg_idx: 0, fg_idx: 0,
            bg_hex: String::new(), fg_hex: String::new(),
            open_dropdown: CfDropdown::None,
            dropdown_scroll: 0,
            rules_selected: 0,
            rules_scroll: 0,
            screen_x: 0, screen_y: 0, screen_w: 58, screen_h: 24,
            dd_item_x: 0, dd_item_y: 0, dd_item_w: 0, dd_item_count: 0,
            rect_conditional: (0,0,0),
            rect_range: (0,0,0), rect_cond: (0,0,0), rect_val1: (0,0,0), rect_val2: (0,0,0),
            rect_bold: (0,0,0), rect_italic: (0,0,0), rect_under: (0,0,0),
            rect_dunder: (0,0,0), rect_strike: (0,0,0),
            rect_bg: (0,0,0), rect_bg_hex: (0,0,0), rect_fg: (0,0,0), rect_fg_hex: (0,0,0), rect_rules: (0,0,0,0),
            rect_apply: (0,0,0), rect_base: (0,0,0), rect_dismiss: (0,0,0), rect_delete: (0,0,0),
            rect_cleanall: (0,0,0), rect_close: (0,0,0),
        };
        self.visual_anchor = None;
        self.cf_dialog_prefill_from_existing();
    }

    fn cf_dialog_prefill_from_existing(&mut self) {
        let sheet_idx = self.workbook.active_sheet as u16;
        let sel = match parse_range_arg(&self.cf_dialog.range_text, sheet_idx) {
            Some(r) => r,
            None => return,
        };
        let sheet = self.workbook.active_sheet();
        let mut matched: Vec<&xlcli_core::condfmt::CondRule> = sheet.cond_rules.iter()
            .filter(|r| {
                r.range.start.row == sel.start.row && r.range.end.row == sel.end.row
                && r.range.start.col == sel.start.col && r.range.end.col == sel.end.col
            })
            .collect();
        if matched.is_empty() {
            matched = sheet.cond_rules.iter()
                .filter(|r| {
                    sel.start.row >= r.range.start.row && sel.end.row <= r.range.end.row
                    && sel.start.col >= r.range.start.col && sel.end.col <= r.range.end.col
                })
                .collect();
        }
        let rule = match matched.last() { Some(r) => *r, None => return };
        let d = &mut self.cf_dialog;
        let o = match &rule.style {
            xlcli_core::condfmt::StyleSpec::Overlay(o) => o,
            _ => return,
        };
        if let Some(b) = o.bold { d.bold = b; }
        if let Some(b) = o.italic { d.italic = b; }
        if let Some(b) = o.underline { d.under = b; }
        if let Some(b) = o.double_underline { d.dunder = b; }
        if let Some(b) = o.strikethrough { d.strike = b; }
        if let Some(c) = o.bg_color {
            let idx = color_to_palette_idx(c);
            d.bg_idx = idx;
            if idx == 0 {
                if let Some(col) = c { d.bg_hex = format!("#{:02X}{:02X}{:02X}", col.r, col.g, col.b); }
            } else { d.bg_hex.clear(); }
        }
        if let Some(c) = o.fg_color {
            let idx = color_to_palette_idx(c);
            d.fg_idx = idx;
            if idx == 0 {
                if let Some(col) = c { d.fg_hex = format!("#{:02X}{:02X}{:02X}", col.r, col.g, col.b); }
            } else { d.fg_hex.clear(); }
        }
        use xlcli_core::condfmt::Condition;
        match &rule.cond {
            Condition::Always => { d.conditional = false; }
            Condition::Blanks => { d.conditional = true; d.cond = CfCond::Blanks; }
            Condition::NonBlanks => { d.conditional = true; d.cond = CfCond::NonBlanks; }
            Condition::Contains(s) => { d.conditional = true; d.cond = CfCond::Contains; d.val1 = s.clone(); }
            Condition::Between(a, b) => { d.conditional = true; d.cond = CfCond::Between; d.val1 = a.to_string(); d.val2 = b.to_string(); }
            Condition::Gt(n) => { d.conditional = true; d.cond = CfCond::Gt; d.val1 = n.to_string(); }
            Condition::Lt(n) => { d.conditional = true; d.cond = CfCond::Lt; d.val1 = n.to_string(); }
            Condition::Gte(n) => { d.conditional = true; d.cond = CfCond::Gte; d.val1 = n.to_string(); }
            Condition::Lte(n) => { d.conditional = true; d.cond = CfCond::Lte; d.val1 = n.to_string(); }
            Condition::Eq(n) => { d.conditional = true; d.cond = CfCond::Eq; d.val1 = n.to_string(); }
            Condition::Neq(n) => { d.conditional = true; d.cond = CfCond::Neq; d.val1 = n.to_string(); }
            _ => { d.conditional = true; }
        }
    }

    fn cf_dialog_build_overlay(&self) -> xlcli_core::condfmt::StyleOverlay {
        let d = &self.cf_dialog;
        let mut o = xlcli_core::condfmt::StyleOverlay::default();
        if d.bold { o.bold = Some(true); }
        if d.italic { o.italic = Some(true); }
        if d.under { o.underline = Some(true); }
        if d.dunder { o.double_underline = Some(true); }
        if d.strike { o.strikethrough = Some(true); }
        if let Some(c) = parse_hex_color(&d.bg_hex) {
            o.bg_color = Some(Some(c));
        } else if d.bg_idx > 0 {
            if let Ok(c) = color_by_name(CF_COLORS[d.bg_idx]) { o.bg_color = Some(c); }
        }
        if let Some(c) = parse_hex_color(&d.fg_hex) {
            o.fg_color = Some(Some(c));
        } else if d.fg_idx > 0 {
            if let Ok(c) = color_by_name(CF_COLORS[d.fg_idx]) { o.fg_color = Some(c); }
        }
        o
    }

    fn cf_dialog_build_cond(&self) -> Result<xlcli_core::condfmt::Condition, String> {
        use xlcli_core::condfmt::Condition;
        let d = &self.cf_dialog;
        match d.cond {
            CfCond::Blanks => Ok(Condition::Blanks),
            CfCond::NonBlanks => Ok(Condition::NonBlanks),
            CfCond::Contains => {
                if d.val1.is_empty() { return Err("Value required".into()); }
                Ok(Condition::Contains(d.val1.clone()))
            }
            CfCond::Between => {
                let a: f64 = d.val1.parse().map_err(|_| "Val1 not a number".to_string())?;
                let b: f64 = d.val2.parse().map_err(|_| "Val2 not a number".to_string())?;
                Ok(Condition::Between(a, b))
            }
            other => {
                let n: f64 = d.val1.parse().map_err(|_| "Val1 not a number".to_string())?;
                Ok(match other {
                    CfCond::Gt => Condition::Gt(n),
                    CfCond::Lt => Condition::Lt(n),
                    CfCond::Gte => Condition::Gte(n),
                    CfCond::Lte => Condition::Lte(n),
                    CfCond::Eq => Condition::Eq(n),
                    CfCond::Neq => Condition::Neq(n),
                    _ => unreachable!(),
                })
            }
        }
    }

    pub fn cf_dialog_apply(&mut self) {
        let sheet_idx = self.workbook.active_sheet as u16;
        let range = match parse_range_arg(&self.cf_dialog.range_text, sheet_idx) {
            Some(r) => r,
            None => {
                self.status_message = Some("Invalid range".to_string());
                return;
            }
        };
        let cond = if self.cf_dialog.conditional {
            match self.cf_dialog_build_cond() {
                Ok(c) => c,
                Err(e) => { self.status_message = Some(e); return; }
            }
        } else {
            xlcli_core::condfmt::Condition::Always
        };
        let overlay = self.cf_dialog_build_overlay();
        if overlay.is_empty() {
            self.status_message = Some("No style selected".to_string());
            return;
        }
        self.workbook.active_sheet_mut().cond_rules.push(xlcli_core::condfmt::CondRule {
            range, cond, style: xlcli_core::condfmt::StyleSpec::Overlay(overlay),
        });
        self.status_message = Some(if self.cf_dialog.conditional { "Rule added" } else { "Format applied" }.to_string());
    }

    pub fn cf_dialog_set_base(&mut self) {
        let overlay = self.cf_dialog_build_overlay();
        if overlay.is_empty() {
            self.status_message = Some("No style selected".to_string());
            return;
        }
        self.workbook.active_sheet_mut().base_style = overlay;
        self.status_message = Some("Base style set".to_string());
    }

    pub fn cf_dialog_delete_selected(&mut self) {
        let idx = self.cf_dialog.rules_selected;
        let sheet = self.workbook.active_sheet_mut();
        if idx < sheet.cond_rules.len() {
            sheet.cond_rules.remove(idx);
            if self.cf_dialog.rules_selected >= sheet.cond_rules.len()
                && self.cf_dialog.rules_selected > 0 {
                self.cf_dialog.rules_selected -= 1;
            }
            self.status_message = Some("Rule deleted".to_string());
        }
    }

    pub fn cf_dialog_clean_all(&mut self) {
        let sheet = self.workbook.active_sheet_mut();
        sheet.cond_rules.clear();
        sheet.base_style = xlcli_core::condfmt::StyleOverlay::default();
        self.status_message = Some("All rules cleared".to_string());
    }

    pub fn cf_dialog_dismiss(&mut self) {
        self.cf_dialog.visible = false;
    }

    fn handle_cf_command(&mut self, rest: &str) {
        if rest == "list" { self.cf_list(); return; }
        if rest == "diag" {
            self.status_message = Some(self.workbook.load_diagnostic.clone()
                .unwrap_or_else(|| "No load diagnostic".into()));
            return;
        }
        // :cf clean ...
        if rest == "clean" {
            self.cf_clean_visual();
            return;
        }
        if let Some(arg) = rest.strip_prefix("clean ") {
            let arg = arg.trim();
            if arg == "all" {
                let sheet = self.workbook.active_sheet_mut();
                let n = sheet.cond_rules.len();
                sheet.cond_rules.clear();
                sheet.base_style = xlcli_core::condfmt::StyleOverlay::default();
                self.status_message = Some(format!("Cleared {} rules + base", n));
            } else {
                self.cf_clean_range(arg);
            }
            return;
        }
        // :cf base <style>
        if let Some(rest) = rest.strip_prefix("base ") {
            match parse_style_overlay(rest.trim()) {
                Ok(overlay) => {
                    self.workbook.active_sheet_mut().base_style = overlay;
                    self.status_message = Some("Base style set".to_string());
                }
                Err(e) => self.status_message = Some(e),
            }
            return;
        }
        // Determine range: either first token looks like range OR visual selection
        let (range, rest_args) = if let Some((maybe_range, rest_args)) = rest.split_once(' ') {
            if looks_like_range(maybe_range) {
                match parse_range_arg(maybe_range, self.workbook.active_sheet as u16) {
                    Some(r) => (Some(r), rest_args.trim().to_string()),
                    None => {
                        self.status_message = Some("Invalid range".to_string());
                        return;
                    }
                }
            } else {
                (None, rest.to_string())
            }
        } else {
            (None, rest.to_string())
        };
        let range = if let Some(r) = range {
            r
        } else if let Some(anchor) = self.visual_anchor {
            let sheet = self.workbook.active_sheet as u16;
            let r1 = anchor.row.min(self.cursor.row);
            let r2 = anchor.row.max(self.cursor.row);
            let c1 = anchor.col.min(self.cursor.col);
            let c2 = anchor.col.max(self.cursor.col);
            xlcli_core::range::CellRange::new(
                CellAddr::new(sheet, r1, c1),
                CellAddr::new(sheet, r2, c2),
            )
        } else {
            self.status_message = Some("Select range in visual or give :cf <range> <cond> <style>".to_string());
            return;
        };
        // Parse cond + style
        let (cond, style_str) = match parse_condition(&rest_args) {
            Ok(c) => c,
            Err(e) => { self.status_message = Some(e); return; }
        };
        let style = match parse_style_overlay(&style_str) {
            Ok(s) => s,
            Err(e) => { self.status_message = Some(e); return; }
        };
        if style.is_empty() {
            self.status_message = Some("Empty style. Use bold, italic, bg=red, fg=blue, etc".to_string());
            return;
        }
        self.workbook.active_sheet_mut().cond_rules.push(xlcli_core::condfmt::CondRule {
            range,
            cond,
            style: xlcli_core::condfmt::StyleSpec::Overlay(style),
        });
        self.visual_anchor = None;
        self.status_message = Some("Rule added".to_string());
    }

    fn cf_list(&mut self) {
        self.cf_list = CfListDialog {
            visible: true,
            selected: 0,
            scroll: 0,
            pending_d: false,
            rect: (0, 0, 0, 0),
        };
    }

    pub fn cf_list_delete(&mut self) {
        let idx = self.cf_list.selected;
        let sheet = self.workbook.active_sheet_mut();
        if idx < sheet.cond_rules.len() {
            sheet.cond_rules.remove(idx);
            let n = sheet.cond_rules.len();
            if self.cf_list.selected >= n && self.cf_list.selected > 0 {
                self.cf_list.selected -= 1;
            }
            self.status_message = Some("Rule deleted".into());
        }
    }

    pub fn cf_list_edit(&mut self) {
        let idx = self.cf_list.selected;
        let sheet = self.workbook.active_sheet();
        if let Some(rule) = sheet.cond_rules.get(idx).cloned() {
            self.cf_list.visible = false;
            self.cf_dialog_open_prefill(&rule);
        }
    }

    fn cf_dialog_open_prefill(&mut self, rule: &xlcli_core::condfmt::CondRule) {
        let range_text = format!("{}{}:{}{}",
            CellAddr::col_name(rule.range.start.col), rule.range.start.row + 1,
            CellAddr::col_name(rule.range.end.col), rule.range.end.row + 1);
        self.cf_dialog = CfDialog {
            visible: true,
            focus: CfDialogFocus::Range,
            editing_text: false,
            conditional: false,
            range_text,
            cond: CfCond::Gt,
            val1: String::new(),
            val2: String::new(),
            bold: false, italic: false, under: false, dunder: false, strike: false,
            bg_idx: 0, fg_idx: 0,
            bg_hex: String::new(), fg_hex: String::new(),
            open_dropdown: CfDropdown::None,
            dropdown_scroll: 0,
            rules_selected: 0,
            rules_scroll: 0,
            screen_x: 0, screen_y: 0, screen_w: 58, screen_h: 24,
            dd_item_x: 0, dd_item_y: 0, dd_item_w: 0, dd_item_count: 0,
            rect_conditional: (0,0,0),
            rect_range: (0,0,0), rect_cond: (0,0,0), rect_val1: (0,0,0), rect_val2: (0,0,0),
            rect_bold: (0,0,0), rect_italic: (0,0,0), rect_under: (0,0,0),
            rect_dunder: (0,0,0), rect_strike: (0,0,0),
            rect_bg: (0,0,0), rect_bg_hex: (0,0,0), rect_fg: (0,0,0), rect_fg_hex: (0,0,0), rect_rules: (0,0,0,0),
            rect_apply: (0,0,0), rect_base: (0,0,0), rect_dismiss: (0,0,0), rect_delete: (0,0,0),
            rect_cleanall: (0,0,0), rect_close: (0,0,0),
        };
        self.cf_dialog_prefill_from_existing();
        let _ = rule;
    }

    pub fn cf_list_new(&mut self) {
        self.cf_list.visible = false;
        self.open_cf_dialog();
    }

    fn cf_clean_visual(&mut self) {
        if let Some(anchor) = self.visual_anchor {
            let sheet_idx = self.workbook.active_sheet as u16;
            let r1 = anchor.row.min(self.cursor.row);
            let r2 = anchor.row.max(self.cursor.row);
            let c1 = anchor.col.min(self.cursor.col);
            let c2 = anchor.col.max(self.cursor.col);
            let target = xlcli_core::range::CellRange::new(
                CellAddr::new(sheet_idx, r1, c1),
                CellAddr::new(sheet_idx, r2, c2),
            );
            let sheet = self.workbook.active_sheet_mut();
            let before = sheet.cond_rules.len();
            sheet.cond_rules.retain(|r| !ranges_overlap(&r.range, &target));
            let removed = before - sheet.cond_rules.len();
            self.visual_anchor = None;
            self.status_message = Some(format!("Cleaned {} rules", removed));
        } else {
            self.status_message = Some("No visual selection. Use :cf clean <range> or :cf clean all".to_string());
        }
    }

    fn cf_clean_range(&mut self, arg: &str) {
        let sheet_idx = self.workbook.active_sheet as u16;
        let target = match parse_range_arg(arg, sheet_idx) {
            Some(r) => r,
            None => {
                self.status_message = Some("Invalid range".to_string());
                return;
            }
        };
        let sheet = self.workbook.active_sheet_mut();
        let before = sheet.cond_rules.len();
        sheet.cond_rules.retain(|r| !ranges_overlap(&r.range, &target));
        let removed = before - sheet.cond_rules.len();
        self.status_message = Some(format!("Cleaned {} rules", removed));
    }

    fn handle_case_command(&mut self, arg: &str) {
        let mode = match arg.trim() {
            "upper" => CaseMode::Upper,
            "lower" => CaseMode::Lower,
            "title" => CaseMode::Title,
            "sentence" => CaseMode::Sentence,
            "toggle" => CaseMode::Toggle,
            _ => {
                self.status_message = Some("Usage: :case upper|lower|title|sentence|toggle".to_string());
                return;
            }
        };
        let (r1, r2, c1, c2) = if let Some(anchor) = self.visual_anchor {
            (
                anchor.row.min(self.cursor.row),
                anchor.row.max(self.cursor.row),
                anchor.col.min(self.cursor.col),
                anchor.col.max(self.cursor.col),
            )
        } else {
            (self.cursor.row, self.cursor.row, self.cursor.col, self.cursor.col)
        };
        let sheet = self.workbook.active_sheet_mut();
        let mut count = 0;
        for row in r1..=r2 {
            for col in c1..=c2 {
                if let Some(cell) = sheet.get_cell(row, col) {
                    if let xlcli_core::cell::CellValue::String(s) = &cell.value {
                        let new_s = transform_case(s, mode);
                        let mut new_cell = cell.clone();
                        new_cell.value = xlcli_core::cell::CellValue::String(new_s.into());
                        sheet.set_cell(row, col, new_cell);
                        count += 1;
                    }
                }
            }
        }
        self.visual_anchor = None;
        self.modified = true;
        self.status_message = Some(format!("Case changed on {} cells", count));
    }

    fn list_named_ranges(&mut self) {
        if self.workbook.named_ranges.is_empty() {
            self.status_message = Some("No named ranges".to_string());
            return;
        }
        let list: Vec<String> = self.workbook.named_ranges.iter()
            .map(|nr| {
                let s = nr.range.start;
                let e = nr.range.end;
                format!("{}={}{}:{}{}", nr.name,
                    CellAddr::col_name(s.col), s.row + 1,
                    CellAddr::col_name(e.col), e.row + 1)
            })
            .collect();
        self.status_message = Some(list.join(", "));
    }

    fn handle_sheet_command(&mut self, rest: &str) {
        let parts: Vec<&str> = rest.splitn(3, ' ').collect();
        match parts[0] {
            "add" => {
                if parts.len() > 1 {
                    let name = parts[1..].join(" ");
                    self.add_sheet_with_name(&name);
                } else {
                    self.add_sheet();
                }
            }
            "delete" => {
                let idx = if parts.len() > 1 {
                    match parts[1].parse::<usize>() {
                        Ok(n) if n >= 1 => n - 1,
                        _ => {
                            self.status_message = Some("Invalid sheet number".to_string());
                            return;
                        }
                    }
                } else {
                    self.workbook.active_sheet
                };
                self.delete_sheet(idx);
            }
            "rename" => {
                if parts.len() < 3 {
                    self.status_message = Some("Usage: :sheet rename <old> <new>".to_string());
                } else {
                    let old = parts[1].to_string();
                    let new_name = parts[2].to_string();
                    self.rename_sheet(&old, &new_name);
                }
            }
            "move" => {
                if parts.len() == 2 {
                    match parts[1] {
                        "left" => {
                            let cur = self.workbook.active_sheet;
                            if cur > 0 {
                                self.move_sheet(cur, cur - 1);
                            } else {
                                self.status_message = Some("Already first sheet".to_string());
                            }
                        }
                        "right" => {
                            let cur = self.workbook.active_sheet;
                            if cur + 1 < self.workbook.sheets.len() {
                                self.move_sheet(cur, cur + 1);
                            } else {
                                self.status_message = Some("Already last sheet".to_string());
                            }
                        }
                        n => {
                            if let Ok(pos) = n.parse::<usize>() {
                                if pos >= 1 && pos <= self.workbook.sheets.len() {
                                    self.move_sheet(self.workbook.active_sheet, pos - 1);
                                } else {
                                    self.status_message = Some("Invalid position".to_string());
                                }
                            } else {
                                self.status_message = Some("Usage: :sheet move <left|right|N> [M]".to_string());
                            }
                        }
                    }
                } else if parts.len() == 3 {
                    let from_str = parts[1];
                    let to_str = parts[2];
                    match (from_str.parse::<usize>(), to_str.parse::<usize>()) {
                        (Ok(f), Ok(t)) if f >= 1 && t >= 1 => {
                            let len = self.workbook.sheets.len();
                            if f <= len && t <= len {
                                self.move_sheet(f - 1, t - 1);
                            } else {
                                self.status_message = Some("Invalid position".to_string());
                            }
                        }
                        _ => self.status_message = Some("Usage: :sheet move <from> <to>".to_string()),
                    }
                } else {
                    self.status_message = Some("Usage: :sheet move <left|right|N> [M]".to_string());
                }
            }
            _ => {
                // :sheet N — switch to sheet by 1-indexed number
                if let Ok(n) = parts[0].parse::<usize>() {
                    if n >= 1 {
                        self.switch_sheet(n - 1);
                    } else {
                        self.status_message = Some("Sheet numbers start at 1".to_string());
                    }
                } else {
                    // Try switching by name
                    if let Some((idx, _)) = self.workbook.sheet_by_name(parts[0]) {
                        self.switch_sheet(idx);
                    } else {
                        self.status_message = Some(format!("Unknown sheet command: {}", rest));
                    }
                }
            }
        }
    }

    fn save_file(&mut self, path: Option<String>) {
        let save_path = path
            .or_else(|| self.workbook.file_path.clone());

        if let Some(p) = save_path {
            match xlcli_io::writer::write_file(&self.workbook, std::path::Path::new(&p)) {
                Ok(()) => {
                    self.workbook.file_path = Some(p.clone());
                    self.modified = false;
                    self.status_message = Some(format!("Saved to {}", p));
                }
                Err(e) => {
                    self.status_message = Some(format!("Save failed: {}", e));
                }
            }
        } else {
            self.status_message = Some("No file path. Use :w <filename>".to_string());
        }
    }

    // --- Headers ---

    fn set_headers_from_visual(&mut self) {
        if let Some(anchor) = self.visual_anchor {
            let row = anchor.row.min(self.cursor.row);
            self.apply_header_row(row);
            self.visual_anchor = None;
            self.mode = Mode::Normal;
        } else {
            self.status_message = Some("Select a row first or use :headers <row>".to_string());
        }
    }

    fn set_headers(&mut self, args: &str) {
        match args.parse::<u32>() {
            Ok(n) if n >= 1 => {
                self.apply_header_row(n - 1);
            }
            _ => {
                self.status_message = Some("Usage: :headers <row-number>".to_string());
            }
        }
    }

    fn apply_header_row(&mut self, row: u32) {
        let sheet = self.workbook.active_sheet_mut();
        sheet.header_row = Some(row);
        // Auto-freeze: freeze rows up to and including header row
        let freeze_cols = sheet.freeze.map(|(_, c)| c).unwrap_or(0);
        sheet.freeze = Some((row + 1, freeze_cols));
        self.status_message = Some(format!("Header row {} (frozen)", row + 1));
    }

    // --- Filter ---

    fn ensure_filter_range(&mut self) {
        let sheet = self.workbook.active_sheet();
        if sheet.filter_range.is_some() {
            return;
        }
        // Default: use visual selection or whole sheet
        if let Some(anchor) = self.visual_anchor {
            let min_row = anchor.row.min(self.cursor.row);
            let max_row = anchor.row.max(self.cursor.row);
            let min_col = anchor.col.min(self.cursor.col);
            let max_col = anchor.col.max(self.cursor.col);
            self.workbook.active_sheet_mut().filter_range = Some((min_row, max_row, min_col, max_col));
            self.visual_anchor = None;
            self.mode = Mode::Normal;
        } else {
            let rc = sheet.row_count().saturating_sub(1);
            let cc = sheet.col_count().saturating_sub(1);
            self.workbook.active_sheet_mut().filter_range = Some((0, rc, 0, cc));
        }
    }

    fn set_filter_range_all(&mut self) {
        let sheet = self.workbook.active_sheet();
        let rc = sheet.row_count().saturating_sub(1);
        let cc = sheet.col_count().saturating_sub(1);
        self.workbook.active_sheet_mut().filter_range = Some((0, rc, 0, cc));
    }

    pub fn open_filter_dialog(&mut self, preset_col: Option<u16>) {
        self.ensure_filter_range();
        let sheet = self.workbook.active_sheet();
        let (_, _, min_col, max_col) = sheet.filter_range.unwrap_or((0, 0, 0, 0));

        let col = preset_col.unwrap_or(self.cursor.col).clamp(min_col, max_col);
        let raw_values = self.collect_column_values(col);
        let all_values: Vec<(String, bool)> = raw_values.into_iter().map(|v| (v, true)).collect();
        let filtered_values: Vec<usize> = (0..all_values.len()).collect();
        self.filter_dialog = FilterDialog {
            visible: true,
            focus: if preset_col.is_some() { FilterDialogFocus::FilterType } else { FilterDialogFocus::Column },
            col,
            filter_type: FilterType::Eq,
            value_buf: String::new(),
            min_col, max_col,
            open_dropdown: FilterDropdown::None,
            dropdown_scroll: 0,
            screen_x: 0, screen_y: 0, screen_w: 44, screen_h: 11,
            dd_item_x: 0, dd_item_y: 0, dd_item_w: 0, dd_item_count: 0,
            all_values,
            filtered_values,
            value_selected: 0,
            value_scroll: 0,
            value_search: String::new(),
            all_checked: true,
        };
        self.visual_anchor = None;
        if self.mode == Mode::Visual {
            self.mode = Mode::Normal;
        }
    }

    fn collect_column_values(&self, col: u16) -> Vec<String> {
        let sheet = self.workbook.active_sheet();
        let (min_row, max_row, _, _) = sheet.filter_range.unwrap_or((0, 0, 0, 0));
        let header = sheet.header_row;
        let start = if let Some(h) = header {
            if h >= min_row && h <= max_row { h + 1 } else { min_row }
        } else {
            min_row
        };

        let mut seen = std::collections::HashSet::new();
        let mut vals = Vec::new();
        for row in start..=max_row {
            let v = sheet.get_cell_value(row, col);
            if !v.is_empty() {
                let s = v.display_value();
                if seen.insert(s.clone()) {
                    vals.push(s);
                }
            }
        }
        vals.sort_by(|a, b| a.to_lowercase().cmp(&b.to_lowercase()));
        vals
    }

    pub fn update_filter_value_list(&mut self) {
        let query = self.filter_dialog.value_search.to_lowercase();
        if query.is_empty() {
            self.filter_dialog.filtered_values = (0..self.filter_dialog.all_values.len()).collect();
        } else {
            self.filter_dialog.filtered_values = self.filter_dialog.all_values.iter()
                .enumerate()
                .filter(|(_, (v, _))| v.to_lowercase().contains(&query))
                .map(|(i, _)| i)
                .collect();
        }
        self.filter_dialog.value_selected = 0;
        self.filter_dialog.value_scroll = 0;
    }

    pub fn refresh_filter_column_values(&mut self) {
        let col = self.filter_dialog.col;
        let raw = self.collect_column_values(col);
        self.filter_dialog.all_values = raw.into_iter().map(|v| (v, true)).collect();
        self.filter_dialog.filtered_values = (0..self.filter_dialog.all_values.len()).collect();
        self.filter_dialog.value_search.clear();
        self.filter_dialog.value_selected = 0;
        self.filter_dialog.value_scroll = 0;
        self.filter_dialog.all_checked = true;
    }

    pub fn toggle_filter_value(&mut self, filtered_idx: usize) {
        if let Some(&real_idx) = self.filter_dialog.filtered_values.get(filtered_idx) {
            self.filter_dialog.all_values[real_idx].1 = !self.filter_dialog.all_values[real_idx].1;
            self.filter_dialog.all_checked = self.filter_dialog.all_values.iter().all(|(_, c)| *c);
        }
    }

    pub fn toggle_filter_all(&mut self) {
        let new_state = !self.filter_dialog.all_checked;
        self.filter_dialog.all_checked = new_state;
        for (_, checked) in &mut self.filter_dialog.all_values {
            *checked = new_state;
        }
    }

    pub fn confirm_filter_dialog(&mut self) {
        let col = self.filter_dialog.col;
        let ft = self.filter_dialog.filter_type;
        let val = self.filter_dialog.value_buf.clone();

        let cond = match ft {
            FilterType::Eq => {
                if self.filter_dialog.all_checked {
                    self.workbook.active_sheet_mut().filters.remove(&col);
                    self.workbook.active_sheet_mut().apply_filters();
                    self.filter_dialog.visible = false;
                    self.status_message = Some(format!("Filter on {} removed (all selected)", CellAddr::col_name(col)));
                    return;
                }
                let set: std::collections::HashSet<String> = self.filter_dialog.all_values.iter()
                    .filter(|(_, checked)| *checked)
                    .map(|(v, _)| v.clone())
                    .collect();
                if set.is_empty() {
                    self.status_message = Some("No values selected".to_string());
                    return;
                }
                FilterCondition::ValueSet(set)
            }
            FilterType::NotEq => FilterCondition::NotEq(val),
            FilterType::Gt => match val.parse::<f64>() {
                Ok(n) => FilterCondition::Gt(n),
                Err(_) => { self.status_message = Some("Value must be a number".to_string()); return; }
            },
            FilterType::Lt => match val.parse::<f64>() {
                Ok(n) => FilterCondition::Lt(n),
                Err(_) => { self.status_message = Some("Value must be a number".to_string()); return; }
            },
            FilterType::Gte => match val.parse::<f64>() {
                Ok(n) => FilterCondition::Gte(n),
                Err(_) => { self.status_message = Some("Value must be a number".to_string()); return; }
            },
            FilterType::Lte => match val.parse::<f64>() {
                Ok(n) => FilterCondition::Lte(n),
                Err(_) => { self.status_message = Some("Value must be a number".to_string()); return; }
            },
            FilterType::Contains => FilterCondition::Contains(val),
            FilterType::Blanks => FilterCondition::Blanks,
            FilterType::NonBlanks => FilterCondition::NonBlanks,
            FilterType::TopN => match val.parse::<usize>() {
                Ok(n) if n > 0 => FilterCondition::TopN(n),
                _ => { self.status_message = Some("Value must be a positive integer".to_string()); return; }
            },
            FilterType::BottomN => match val.parse::<usize>() {
                Ok(n) if n > 0 => FilterCondition::BottomN(n),
                _ => { self.status_message = Some("Value must be a positive integer".to_string()); return; }
            },
        };

        let col_name = CellAddr::col_name(col);
        self.workbook.active_sheet_mut().filters.insert(col, cond);
        self.workbook.active_sheet_mut().apply_filters();
        self.filter_dialog.visible = false;
        self.status_message = Some(format!("Filter on {} applied", col_name));
    }

    fn parse_and_filter(&mut self, args: &str) {
        // :filter B = hello, :filter B > 100, :filter B contains test, :filter B blanks
        self.ensure_filter_range();

        let parts: Vec<&str> = args.splitn(3, ' ').collect();
        if parts.is_empty() {
            self.status_message = Some("Usage: :filter <col> <op> [value]".to_string());
            return;
        }

        // Check for "all" — already handled in execute_command, but handle :filter all B = x
        if parts[0].eq_ignore_ascii_case("all") {
            self.set_filter_range_all();
            if parts.len() > 1 {
                let rest = args.strip_prefix(parts[0]).unwrap().trim();
                self.parse_and_filter(rest);
            } else {
                self.open_filter_dialog(None);
            }
            return;
        }

        let col_str = parts[0].to_uppercase();
        let col = match CellAddr::parse_col(&col_str) {
            Some(c) => c,
            None => { self.status_message = Some(format!("Invalid column: {}", parts[0])); return; }
        };

        if parts.len() < 2 {
            self.open_filter_dialog(Some(col));
            return;
        }

        let op = parts[1].to_lowercase();
        let val = if parts.len() > 2 { parts[2].to_string() } else { String::new() };

        let cond = match op.as_str() {
            "=" | "eq" => FilterCondition::Eq(val),
            "!=" | "neq" => FilterCondition::NotEq(val),
            ">" | "gt" => match val.parse::<f64>() {
                Ok(n) => FilterCondition::Gt(n),
                Err(_) => { self.status_message = Some("Value must be a number".to_string()); return; }
            },
            "<" | "lt" => match val.parse::<f64>() {
                Ok(n) => FilterCondition::Lt(n),
                Err(_) => { self.status_message = Some("Value must be a number".to_string()); return; }
            },
            ">=" | "gte" => match val.parse::<f64>() {
                Ok(n) => FilterCondition::Gte(n),
                Err(_) => { self.status_message = Some("Value must be a number".to_string()); return; }
            },
            "<=" | "lte" => match val.parse::<f64>() {
                Ok(n) => FilterCondition::Lte(n),
                Err(_) => { self.status_message = Some("Value must be a number".to_string()); return; }
            },
            "contains" => FilterCondition::Contains(val),
            "blanks" => FilterCondition::Blanks,
            "nonblanks" => FilterCondition::NonBlanks,
            "top" => match val.parse::<usize>() {
                Ok(n) if n > 0 => FilterCondition::TopN(n),
                _ => { self.status_message = Some("Usage: :filter B top 10".to_string()); return; }
            },
            "bottom" => match val.parse::<usize>() {
                Ok(n) if n > 0 => FilterCondition::BottomN(n),
                _ => { self.status_message = Some("Usage: :filter B bottom 10".to_string()); return; }
            },
            _ => { self.status_message = Some(format!("Unknown filter op: {}", op)); return; }
        };

        let col_name = CellAddr::col_name(col);
        self.workbook.active_sheet_mut().filters.insert(col, cond);
        self.workbook.active_sheet_mut().apply_filters();
        self.status_message = Some(format!("Filter on {} applied", col_name));
    }

    fn unfilter_visual(&mut self) {
        if let Some(anchor) = self.visual_anchor {
            let min_col = anchor.col.min(self.cursor.col);
            let max_col = anchor.col.max(self.cursor.col);
            let sheet = self.workbook.active_sheet_mut();
            for c in min_col..=max_col {
                sheet.filters.remove(&c);
            }
            sheet.apply_filters();
            self.visual_anchor = None;
            self.mode = Mode::Normal;
            self.status_message = Some("Filters removed for selection".to_string());
        } else {
            self.unfilter_all();
        }
    }

    fn unfilter_all(&mut self) {
        let sheet = self.workbook.active_sheet_mut();
        sheet.filters.clear();
        sheet.hidden_rows.clear();
        sheet.filter_range = None;
        self.status_message = Some("All filters removed".to_string());
    }

    fn unfilter_column(&mut self, args: &str) {
        let col_str = args.to_uppercase();
        if let Some(col) = CellAddr::parse_col(&col_str) {
            let sheet = self.workbook.active_sheet_mut();
            if sheet.filters.remove(&col).is_some() {
                sheet.apply_filters();
                self.status_message = Some(format!("Filter on {} removed", col_str));
            } else {
                self.status_message = Some(format!("No filter on {}", col_str));
            }
        } else {
            self.status_message = Some(format!("Invalid column: {}", args));
        }
    }

    pub fn is_header_row(&self, row: u32) -> bool {
        self.workbook.active_sheet().header_row == Some(row)
    }
}

struct WorkbookEvalContext<'a> {
    workbook: &'a Workbook,
    cell_addr: CellAddr,
}

impl<'a> EvalContext for WorkbookEvalContext<'a> {
    fn get_cell_value(&self, addr: CellAddr) -> CellValue {
        if let Some(sheet) = self.workbook.sheets.get(addr.sheet as usize) {
            sheet.get_cell_value(addr.row, addr.col).clone()
        } else {
            CellValue::Error(xlcli_core::types::CellError::Ref)
        }
    }

    fn current_cell(&self) -> CellAddr {
        self.cell_addr
    }

    fn current_sheet(&self) -> u16 {
        self.cell_addr.sheet
    }

    fn resolve_sheet(&self, name: &str) -> Option<u16> {
        self.workbook.sheets.iter().enumerate()
            .find(|(_, s)| s.name.eq_ignore_ascii_case(name))
            .map(|(i, _)| i as u16)
    }

    fn resolve_named_range(&self, name: &str) -> Option<(CellAddr, CellAddr)> {
        self.workbook.named_ranges.iter()
            .find(|nr| nr.name.eq_ignore_ascii_case(name))
            .map(|nr| (nr.range.start, nr.range.end))
    }
}

#[derive(Clone, Copy)]
enum CaseMode { Upper, Lower, Title, Sentence, Toggle }

fn transform_case(s: &str, mode: CaseMode) -> String {
    match mode {
        CaseMode::Upper => s.to_uppercase(),
        CaseMode::Lower => s.to_lowercase(),
        CaseMode::Toggle => s.chars().map(|c| {
            if c.is_uppercase() { c.to_lowercase().next().unwrap_or(c) }
            else if c.is_lowercase() { c.to_uppercase().next().unwrap_or(c) }
            else { c }
        }).collect(),
        CaseMode::Title => {
            let mut out = String::with_capacity(s.len());
            let mut at_word_start = true;
            for ch in s.chars() {
                if ch.is_whitespace() || !ch.is_alphanumeric() {
                    out.push(ch);
                    at_word_start = true;
                } else if at_word_start {
                    out.extend(ch.to_uppercase());
                    at_word_start = false;
                } else {
                    out.extend(ch.to_lowercase());
                }
            }
            out
        }
        CaseMode::Sentence => {
            let lower = s.to_lowercase();
            let mut out = String::with_capacity(lower.len());
            let mut capitalize = true;
            for ch in lower.chars() {
                if capitalize && ch.is_alphabetic() {
                    out.extend(ch.to_uppercase());
                    capitalize = false;
                } else {
                    out.push(ch);
                    if matches!(ch, '.' | '!' | '?') { capitalize = true; }
                }
            }
            out
        }
    }
}

fn looks_like_range(s: &str) -> bool {
    // Range: contains ':' or is pure A1 cell
    if s.contains(':') { return true; }
    let s_upper = s.to_uppercase();
    let mut bytes = s_upper.bytes();
    let mut saw_alpha = false;
    let mut saw_digit = false;
    while let Some(b) = bytes.next() {
        if b.is_ascii_alphabetic() {
            if saw_digit { return false; }
            saw_alpha = true;
        } else if b.is_ascii_digit() {
            if !saw_alpha { return false; }
            saw_digit = true;
        } else {
            return false;
        }
    }
    saw_alpha && saw_digit
}

fn parse_range_arg(s: &str, sheet: u16) -> Option<xlcli_core::range::CellRange> {
    if let Some((a, b)) = s.split_once(':') {
        let (r1, c1) = parse_cell_address(a)?;
        let (r2, c2) = parse_cell_address(b)?;
        Some(xlcli_core::range::CellRange::new(
            CellAddr::new(sheet, r1.min(r2), c1.min(c2)),
            CellAddr::new(sheet, r1.max(r2), c1.max(c2)),
        ))
    } else {
        let (r, c) = parse_cell_address(s)?;
        let a = CellAddr::new(sheet, r, c);
        Some(xlcli_core::range::CellRange::new(a, a))
    }
}

fn ranges_overlap(a: &xlcli_core::range::CellRange, b: &xlcli_core::range::CellRange) -> bool {
    if a.start.sheet != b.start.sheet { return false; }
    a.start.row <= b.end.row && a.end.row >= b.start.row
        && a.start.col <= b.end.col && a.end.col >= b.start.col
}

fn parse_condition(args: &str) -> Result<(xlcli_core::condfmt::Condition, String), String> {
    use xlcli_core::condfmt::Condition;
    let toks: Vec<&str> = args.split_whitespace().collect();
    if toks.is_empty() { return Err("Missing condition".into()); }
    let op = toks[0].to_lowercase();
    match op.as_str() {
        "blanks" => Ok((Condition::Blanks, toks[1..].join(" "))),
        "nonblanks" => Ok((Condition::NonBlanks, toks[1..].join(" "))),
        "between" => {
            if toks.len() < 3 { return Err("Usage: between <a> <b>".into()); }
            let a: f64 = toks[1].parse().map_err(|_| "Bad number".to_string())?;
            let b: f64 = toks[2].parse().map_err(|_| "Bad number".to_string())?;
            Ok((Condition::Between(a, b), toks[3..].join(" ")))
        }
        "contains" => {
            if toks.len() < 2 { return Err("Usage: contains <text>".into()); }
            Ok((Condition::Contains(toks[1].to_string()), toks[2..].join(" ")))
        }
        "gt" | "lt" | "gte" | "lte" | "eq" | "neq" => {
            if toks.len() < 2 { return Err(format!("Usage: {} <num>", op)); }
            let n: f64 = toks[1].parse().map_err(|_| "Bad number".to_string())?;
            let c = match op.as_str() {
                "gt" => Condition::Gt(n),
                "lt" => Condition::Lt(n),
                "gte" => Condition::Gte(n),
                "lte" => Condition::Lte(n),
                "eq" => Condition::Eq(n),
                "neq" => Condition::Neq(n),
                _ => unreachable!(),
            };
            Ok((c, toks[2..].join(" ")))
        }
        _ => Err(format!("Unknown condition: {}", op)),
    }
}

fn parse_style_overlay(s: &str) -> Result<xlcli_core::condfmt::StyleOverlay, String> {
    let mut overlay = xlcli_core::condfmt::StyleOverlay::default();
    for tok in s.split_whitespace() {
        let (negate, key) = if let Some(rest) = tok.strip_prefix('!') {
            (true, rest)
        } else {
            (false, tok)
        };
        let val_bool = !negate;
        if let Some((k, v)) = key.split_once('=') {
            match k {
                "bg" => overlay.bg_color = Some(color_by_name(v)?),
                "fg" => overlay.fg_color = Some(color_by_name(v)?),
                _ => return Err(format!("Unknown style key: {}", k)),
            }
        } else {
            match key.to_lowercase().as_str() {
                "bold" => overlay.bold = Some(val_bool),
                "italic" => overlay.italic = Some(val_bool),
                "under" | "underline" => overlay.underline = Some(val_bool),
                "dunder" | "doubleunder" => overlay.double_underline = Some(val_bool),
                "strike" | "strikethrough" => overlay.strikethrough = Some(val_bool),
                _ => return Err(format!("Unknown style: {}", key)),
            }
        }
    }
    Ok(overlay)
}

fn color_by_name(name: &str) -> Result<Option<xlcli_core::style::Color>, String> {
    use xlcli_core::style::Color;
    let c = match name.to_lowercase().as_str() {
        "none" | "default" => return Ok(None),
        "red" => Color::new(220, 50, 50),
        "green" => Color::new(50, 200, 80),
        "blue" => Color::new(70, 130, 230),
        "yellow" => Color::new(230, 210, 50),
        "cyan" => Color::new(70, 200, 210),
        "magenta" => Color::new(210, 80, 200),
        "orange" => Color::new(230, 140, 50),
        "white" => Color::new(240, 240, 240),
        "black" => Color::new(20, 20, 20),
        "gray" | "grey" => Color::new(130, 130, 130),
        _ => return Err(format!("Unknown color: {}", name)),
    };
    Ok(Some(c))
}

pub fn color_by_name_pub(name: &str) -> Result<Option<xlcli_core::style::Color>, String> {
    color_by_name(name)
}

pub fn parse_hex_color(s: &str) -> Option<xlcli_core::style::Color> {
    let s = s.trim().trim_start_matches('#');
    if s.len() != 6 { return None; }
    let r = u8::from_str_radix(&s[0..2], 16).ok()?;
    let g = u8::from_str_radix(&s[2..4], 16).ok()?;
    let b = u8::from_str_radix(&s[4..6], 16).ok()?;
    Some(xlcli_core::style::Color::new(r, g, b))
}

fn color_to_palette_idx(c: Option<xlcli_core::style::Color>) -> usize {
    let c = match c { Some(c) => c, None => return 0 };
    for (i, name) in CF_COLORS.iter().enumerate() {
        if i == 0 { continue; }
        if let Ok(Some(pc)) = color_by_name(name) {
            if pc.r == c.r && pc.g == c.g && pc.b == c.b { return i; }
        }
    }
    0
}

fn parse_cell_address(addr: &str) -> Option<(u32, u16)> {
    let addr = addr.trim().to_uppercase();
    let col_end = addr.bytes().position(|b| b.is_ascii_digit())?;
    if col_end == 0 {
        return None;
    }
    let col_str = &addr[..col_end];
    let row_str = &addr[col_end..];

    let col = CellAddr::parse_col(col_str)?;
    let row: u32 = row_str.parse().ok()?;
    if row == 0 {
        return None;
    }
    Some((row - 1, col))
}

fn compare_cells(a: &CellValue, b: &CellValue, sort_type: SortType) -> std::cmp::Ordering {
    use std::cmp::Ordering;

    match sort_type {
        SortType::AZ => {
            let sa = a.display_value();
            let sb = b.display_value();
            sa.to_lowercase().cmp(&sb.to_lowercase())
        }
        SortType::ZA => {
            let sa = a.display_value();
            let sb = b.display_value();
            sb.to_lowercase().cmp(&sa.to_lowercase())
        }
        SortType::Numeric => {
            let na = a.as_f64().unwrap_or(0.0);
            let nb = b.as_f64().unwrap_or(0.0);
            na.partial_cmp(&nb).unwrap_or(Ordering::Equal)
        }
        SortType::NumericDesc => {
            let na = a.as_f64().unwrap_or(0.0);
            let nb = b.as_f64().unwrap_or(0.0);
            nb.partial_cmp(&na).unwrap_or(Ordering::Equal)
        }
        SortType::CaseSensitive => {
            let sa = a.display_value();
            let sb = b.display_value();
            sa.cmp(&sb)
        }
        SortType::Natural => {
            let sa = a.display_value();
            let sb = b.display_value();
            natural_cmp(&sa, &sb)
        }
    }
}

fn natural_cmp(a: &str, b: &str) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    let mut ai = a.chars().peekable();
    let mut bi = b.chars().peekable();

    loop {
        match (ai.peek(), bi.peek()) {
            (None, None) => return Ordering::Equal,
            (None, Some(_)) => return Ordering::Less,
            (Some(_), None) => return Ordering::Greater,
            (Some(&ac), Some(&bc)) => {
                if ac.is_ascii_digit() && bc.is_ascii_digit() {
                    let mut an = String::new();
                    while ai.peek().map_or(false, |c| c.is_ascii_digit()) {
                        an.push(ai.next().unwrap());
                    }
                    let mut bn = String::new();
                    while bi.peek().map_or(false, |c| c.is_ascii_digit()) {
                        bn.push(bi.next().unwrap());
                    }
                    let na: u64 = an.parse().unwrap_or(0);
                    let nb: u64 = bn.parse().unwrap_or(0);
                    let ord = na.cmp(&nb);
                    if ord != Ordering::Equal {
                        return ord;
                    }
                } else {
                    let ord = ac.to_lowercase().cmp(bc.to_lowercase());
                    if ord != Ordering::Equal {
                        return ord;
                    }
                    ai.next();
                    bi.next();
                }
            }
        }
    }
}

fn parse_input_value(s: &str) -> CellValue {
    let s = s.trim();
    if s.is_empty() {
        return CellValue::Empty;
    }
    if let Ok(n) = s.parse::<f64>() {
        return CellValue::Number(n);
    }
    match s.to_uppercase().as_str() {
        "TRUE" => CellValue::Boolean(true),
        "FALSE" => CellValue::Boolean(false),
        _ => CellValue::String(s.into()),
    }
}
