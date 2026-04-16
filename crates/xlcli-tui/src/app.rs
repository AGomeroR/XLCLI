use xlcli_core::cell::{Cell, CellValue};
use xlcli_core::types::CellAddr;
use xlcli_core::workbook::Workbook;
use xlcli_formulas::{FunctionRegistry, EvalContext};

use crate::clipboard::Clipboard;
use crate::config::Config;
use crate::mode::Mode;
use crate::undo::{UndoEntry, UndoStack};
use crate::viewport::Viewport;

pub const MAX_ROW: u32 = 1_048_575;
pub const MAX_COL: u16 = 16_383;

pub struct App {
    pub workbook: Workbook,
    pub cursor: CellAddr,
    pub viewport: Viewport,
    pub mode: Mode,
    pub command_buffer: String,
    pub should_quit: bool,
    pub status_message: Option<String>,
    pub pending_key: Option<char>,
    pub edit_buffer: String,
    pub undo_stack: UndoStack,
    pub clipboard: Clipboard,
    pub visual_anchor: Option<CellAddr>,
    pub modified: bool,
    pub config: Config,
    pub formula_registry: FunctionRegistry,
    pub autocomplete: AutocompleteState,
}

pub struct AutocompleteState {
    pub visible: bool,
    pub matches: Vec<&'static str>,
    pub selected: usize,
    pub active_function: Option<String>,
}

impl App {
    pub fn new(workbook: Workbook, config: Config) -> Self {
        let mut app = Self {
            workbook,
            cursor: CellAddr::new(0, 0, 0),
            viewport: Viewport::new(),
            mode: Mode::Normal,
            command_buffer: String::new(),
            should_quit: false,
            status_message: None,
            pending_key: None,
            edit_buffer: String::new(),
            undo_stack: UndoStack::new(),
            clipboard: Clipboard::new(),
            visual_anchor: None,
            modified: false,
            config,
            formula_registry: FunctionRegistry::default(),
            autocomplete: AutocompleteState { visible: false, matches: Vec::new(), selected: 0, active_function: None },
        };
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
        self.mode = Mode::Insert;
    }

    pub fn enter_insert_clear(&mut self) {
        self.edit_buffer.clear();
        self.mode = Mode::Insert;
    }

    pub fn enter_insert_with(&mut self, prefix: &str) {
        self.edit_buffer = prefix.to_string();
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
        self.autocomplete.visible = false;
        self.autocomplete.matches.clear();
        self.autocomplete.selected = 0;
        self.update_autocomplete();
    }

    pub fn confirm_edit(&mut self) {
        let addr = self.cursor;
        let sheet = self.workbook.active_sheet_mut();
        let old = sheet.get_cell(addr.row, addr.col).cloned();

        let input = self.edit_buffer.trim().to_string();
        let new_cell = if input.is_empty() {
            None
        } else if let Some(formula_text) = input.strip_prefix('=') {
            let formula_text = formula_text.to_string();
            let value = self.eval_formula(&formula_text, addr);
            Some(Cell::with_formula(value, formula_text))
        } else {
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
        self.mode = Mode::Normal;
        self.status_message = None;

        self.recalc_all();
    }

    pub fn cancel_edit(&mut self) {
        self.edit_buffer.clear();
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
        self.clipboard.yank_single(cell);
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
            self.clipboard.yank_range(cells, rows, cols);
            let count = rows as usize * cols as usize;
            self.status_message = Some(format!("Yanked {} cells", count));
        }
        self.visual_anchor = None;
        self.mode = Mode::Normal;
    }

    pub fn paste(&mut self) {
        if let Some(ref clip) = self.clipboard.content.clone() {
            let mut undo_entries = Vec::new();
            let base_row = self.cursor.row;
            let base_col = self.cursor.col;
            let sheet_idx = self.workbook.active_sheet as u16;

            for &(rel_row, rel_col, ref cell) in &clip.cells {
                let row = base_row + rel_row;
                let col = base_col + rel_col;
                let addr = CellAddr::new(sheet_idx, row, col);

                let sheet = self.workbook.active_sheet_mut();
                let old = sheet.get_cell(row, col).cloned();
                sheet.set_cell(row, col, cell.clone());

                undo_entries.push(UndoEntry::CellChange {
                    addr,
                    old,
                    new: Some(cell.clone()),
                });
            }

            if undo_entries.len() == 1 {
                self.undo_stack.push(undo_entries.pop().unwrap());
            } else {
                self.undo_stack.push(UndoEntry::Batch(undo_entries));
            }
            self.modified = true;
            self.status_message = Some("Pasted".to_string());
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
        } else {
            self.status_message = Some(format!("Unknown command: {}", cmd));
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
