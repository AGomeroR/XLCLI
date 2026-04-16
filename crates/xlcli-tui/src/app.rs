use std::collections::HashMap;

use xlcli_core::cell::{Cell, CellValue};
use xlcli_core::dep_graph::DepGraph;
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

pub struct SheetState {
    pub cursor: CellAddr,
    pub viewport: Viewport,
}

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
    pub dep_graph: DepGraph,
    pub sheet_states: HashMap<usize, SheetState>,
    pub formula_origin_sheet: Option<u16>,
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
            dep_graph: DepGraph::new(),
            sheet_states: HashMap::new(),
            formula_origin_sheet: None,
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

    pub fn resolve_sheet_name(&self, name: &str) -> Option<u16> {
        self.workbook.sheets.iter().enumerate()
            .find(|(_, s)| s.name.eq_ignore_ascii_case(name))
            .map(|(i, _)| i as u16)
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
        } else if let Some(rest) = cmd.strip_prefix("sheet ") {
            self.handle_sheet_command(rest.trim());
        } else {
            self.status_message = Some(format!("Unknown command: {}", cmd));
        }
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
