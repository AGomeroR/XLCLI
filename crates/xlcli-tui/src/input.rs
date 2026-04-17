use crossterm::event::{Event, KeyCode, KeyEvent, KeyModifiers, MouseButton, MouseEvent, MouseEventKind};

use xlcli_core::types::CellAddr;

use crate::app::{App, CF_COLORS, CfCond, CfDialogFocus, CfDropdown, FilterDialogFocus, FilterDropdown, FilterType, SortDialogFocus, SortDropdown, SortType};
use crate::mode::Mode;

pub fn handle_event(app: &mut App, event: Event) {
    if app.filter_dialog.visible {
        match event {
            Event::Key(key) => handle_filter_dialog(app, key),
            Event::Mouse(mouse) => {
                if let MouseEventKind::Down(MouseButton::Left) = mouse.kind {
                    handle_filter_dialog_mouse(app, mouse.column, mouse.row);
                }
            }
            _ => {}
        }
        return;
    }

    if app.sort_dialog.visible {
        match event {
            Event::Key(key) => handle_sort_dialog(app, key),
            Event::Mouse(mouse) => {
                if let MouseEventKind::Down(MouseButton::Left) = mouse.kind {
                    handle_sort_dialog_mouse(app, mouse.column, mouse.row);
                }
            }
            _ => {}
        }
        return;
    }

    if app.cf_dialog.visible {
        match event {
            Event::Key(key) => handle_cf_dialog(app, key),
            Event::Mouse(mouse) => {
                if let MouseEventKind::Down(MouseButton::Left) = mouse.kind {
                    handle_cf_dialog_mouse(app, mouse.column, mouse.row);
                }
            }
            _ => {}
        }
        return;
    }

    match event {
        Event::Key(key) => match &app.mode {
            Mode::Normal => handle_normal(app, key),
            Mode::Command => handle_command(app, key),
            Mode::Insert => handle_insert(app, key),
            Mode::Visual => handle_visual(app, key),
        },
        Event::Mouse(mouse) => handle_mouse(app, mouse),
        _ => {}
    }
}

fn handle_normal(app: &mut App, key: KeyEvent) {
    match (key.modifiers, key.code) {
        // Navigation
        (KeyModifiers::NONE, KeyCode::Char('h')) | (KeyModifiers::NONE, KeyCode::Left) => {
            app.move_cursor(0, -1);
        }
        (KeyModifiers::NONE, KeyCode::Char('j')) | (KeyModifiers::NONE, KeyCode::Down) => {
            app.move_cursor(1, 0);
        }
        (KeyModifiers::NONE, KeyCode::Char('k')) | (KeyModifiers::NONE, KeyCode::Up) => {
            app.move_cursor(-1, 0);
        }
        (KeyModifiers::NONE, KeyCode::Char('l')) | (KeyModifiers::NONE, KeyCode::Right) => {
            app.move_cursor(0, 1);
        }
        (KeyModifiers::CONTROL, KeyCode::Char('d')) => {
            let max_row = app.active_sheet().row_count().saturating_sub(1);
            let half = app.viewport.visible_rows / 2;
            app.cursor.row = (app.cursor.row + half).min(max_row);
            app.viewport.scroll_half_page_down(max_row);
        }
        (KeyModifiers::CONTROL, KeyCode::Char('u')) => {
            let half = app.viewport.visible_rows / 2;
            app.cursor.row = app.cursor.row.saturating_sub(half);
            app.viewport.scroll_half_page_up();
        }
        (KeyModifiers::NONE, KeyCode::Char('g')) => {
            if app.pending_key == Some('g') {
                app.jump_to_start();
                app.pending_key = None;
            } else {
                app.pending_key = Some('g');
            }
            return;
        }
        (KeyModifiers::CONTROL, KeyCode::Char(c @ '1'..='9')) => {
            let idx = (c as usize) - ('1' as usize);
            app.switch_sheet(idx);
        }
        (KeyModifiers::SHIFT, KeyCode::Char('G')) => {
            app.jump_to_last_row();
        }
        (KeyModifiers::NONE, KeyCode::Char('0')) => {
            app.jump_to_col_start();
        }
        (KeyModifiers::SHIFT, KeyCode::Char('$')) => {
            app.jump_to_col_end();
        }

        // Enter Insert mode
        (KeyModifiers::NONE, KeyCode::Enter) => {
            app.enter_insert();
        }
        (KeyModifiers::NONE, KeyCode::Char('i')) => {
            app.enter_insert_clear();
        }
        (KeyModifiers::NONE, KeyCode::Char('a')) => {
            app.enter_insert();
        }
        (KeyModifiers::NONE, KeyCode::Char('s')) => {
            app.enter_insert_clear();
        }
        (KeyModifiers::NONE, KeyCode::Char('=')) => {
            app.enter_insert_with("=");
        }

        // Visual mode
        (KeyModifiers::NONE, KeyCode::Char('v')) => {
            app.visual_anchor = Some(app.cursor);
            app.mode = Mode::Visual;
            app.status_message = Some("-- VISUAL --".to_string());
        }

        // Command mode
        (KeyModifiers::NONE, KeyCode::Char(':')) => {
            app.mode = Mode::Command;
            app.command_buffer.clear();
        }

        // Undo/Redo
        (KeyModifiers::NONE, KeyCode::Char('u')) => {
            app.undo();
        }
        (KeyModifiers::CONTROL, KeyCode::Char('r')) => {
            app.redo();
        }

        // Clipboard
        (KeyModifiers::NONE, KeyCode::Char('y')) => {
            if app.pending_key == Some('y') {
                app.yank();
                app.pending_key = None;
            } else {
                app.pending_key = Some('y');
            }
            return;
        }
        (KeyModifiers::NONE, KeyCode::Char('p')) => {
            app.paste();
        }
        (KeyModifiers::NONE, KeyCode::Char('d')) => {
            if app.pending_key == Some('d') {
                app.delete_cell();
                app.pending_key = None;
            } else {
                app.pending_key = Some('d');
            }
            return;
        }
        (KeyModifiers::NONE, KeyCode::Char('x')) => {
            app.delete_cell();
        }

        // Tab to move right, Shift+Tab to move left
        (KeyModifiers::NONE, KeyCode::Tab) => {
            app.move_cursor(0, 1);
        }
        (KeyModifiers::SHIFT, KeyCode::BackTab) => {
            app.move_cursor(0, -1);
        }

        // Fill down / Fill right
        (KeyModifiers::ALT, KeyCode::Char('d')) => {
            app.fill_down();
        }
        (KeyModifiers::ALT, KeyCode::Char('r')) => {
            app.fill_right();
        }

        // Relative yank (Y) — yank with relative-paste flag
        (KeyModifiers::SHIFT, KeyCode::Char('Y')) => {
            if app.pending_key == Some('Y') {
                app.yank();
                app.clipboard.set_relative(true);
                app.status_message = Some("Yanked 1 cell (relative)".to_string());
                app.pending_key = None;
            } else {
                app.pending_key = Some('Y');
            }
            return;
        }

        // New sheet — opens command box with prefilled name prompt
        (KeyModifiers::NONE, KeyCode::Char('t')) => {
            app.mode = Mode::Command;
            app.command_buffer = "sheet add ".to_string();
        }

        // Shift+Enter on header row — open filter dialog for cursor column
        (KeyModifiers::SHIFT, KeyCode::Enter) => {
            if app.is_header_row(app.cursor.row) {
                app.open_filter_dialog(Some(app.cursor.col));
            }
        }

        // Quit shortcut
        (KeyModifiers::NONE, KeyCode::Char('q')) => {
            if !app.modified {
                app.should_quit = true;
            } else {
                app.status_message =
                    Some("Unsaved changes! Use :q! or :wq".to_string());
            }
        }

        _ => {}
    }
    app.pending_key = None;
}

fn handle_command(app: &mut App, key: KeyEvent) {
    match key.code {
        KeyCode::Enter => {
            app.execute_command();
        }
        KeyCode::Esc => {
            app.command_buffer.clear();
            app.mode = Mode::Normal;
            app.visual_anchor = None;
        }
        KeyCode::Backspace => {
            app.command_buffer.pop();
            if app.command_buffer.is_empty() {
                app.mode = Mode::Normal;
                app.visual_anchor = None;
            }
        }
        KeyCode::Char(c) => {
            app.command_buffer.push(c);
        }
        _ => {}
    }
}

fn handle_insert(app: &mut App, key: KeyEvent) {
    // Ctrl+1-9: cross-sheet ref browsing (only during formula edit)
    if key.modifiers.contains(KeyModifiers::CONTROL) {
        if let KeyCode::Char(c @ '1'..='9') = key.code {
            if app.edit_buffer.starts_with('=') {
                let idx = (c as usize) - ('1' as usize);
                app.switch_sheet_for_ref(idx);
            }
            return;
        }
    }

    if app.autocomplete.visible {
        match key.code {
            KeyCode::Tab | KeyCode::Enter => {
                app.accept_autocomplete();
                return;
            }
            KeyCode::Down => {
                let len = app.autocomplete.matches.len();
                if len > 0 {
                    app.autocomplete.selected = (app.autocomplete.selected + 1) % len;
                }
                return;
            }
            KeyCode::Up => {
                let len = app.autocomplete.matches.len();
                if len > 0 {
                    app.autocomplete.selected = (app.autocomplete.selected + len - 1) % len;
                }
                return;
            }
            KeyCode::Esc => {
                app.autocomplete.visible = false;
                app.autocomplete.matches.clear();
                return;
            }
            _ => {}
        }
    }

    match key.code {
        KeyCode::Esc => {
            app.autocomplete.visible = false;
            app.cancel_edit();
        }
        KeyCode::Enter => {
            app.confirm_edit();
            app.move_cursor(1, 0);
        }
        KeyCode::Tab => {
            app.confirm_edit();
            app.move_cursor(0, 1);
            app.enter_insert_clear();
        }
        KeyCode::Backspace => {
            app.edit_buffer.pop();
            app.update_autocomplete();
        }
        KeyCode::Char(c) => {
            app.edit_buffer.push(c);
            app.update_autocomplete();
        }
        _ => {}
    }
}

fn handle_visual(app: &mut App, key: KeyEvent) {
    if let Some(pending) = app.pending_key {
        match (pending, key.code) {
            ('g', KeyCode::Char('g')) => {
                app.cursor.row = 0;
                app.pending_key = None;
                return;
            }
            _ => { app.pending_key = None; }
        }
    }

    match (key.modifiers, key.code) {
        (_, KeyCode::Esc) => {
            app.visual_anchor = None;
            app.mode = Mode::Normal;
            app.status_message = None;
        }
        (KeyModifiers::CONTROL, KeyCode::Char(c @ '1'..='9')) => {
            let idx = (c as usize) - ('1' as usize);
            app.switch_sheet(idx);
        }
        (KeyModifiers::NONE, KeyCode::Char('h')) | (_, KeyCode::Left) => {
            app.move_cursor(0, -1);
        }
        (KeyModifiers::NONE, KeyCode::Char('j')) | (_, KeyCode::Down) => {
            app.move_cursor(1, 0);
        }
        (KeyModifiers::NONE, KeyCode::Char('k')) | (_, KeyCode::Up) => {
            app.move_cursor(-1, 0);
        }
        (KeyModifiers::NONE, KeyCode::Char('l')) | (_, KeyCode::Right) => {
            app.move_cursor(0, 1);
        }
        (KeyModifiers::SHIFT, KeyCode::Char('G')) => {
            let last_row = app.workbook.active_sheet().row_count().saturating_sub(1);
            app.cursor.row = last_row;
        }
        (KeyModifiers::NONE, KeyCode::Char('g')) => {
            app.pending_key = Some('g');
        }
        (KeyModifiers::NONE, KeyCode::Char('0')) => {
            app.jump_to_col_start();
        }
        (KeyModifiers::SHIFT, KeyCode::Char('$')) => {
            app.jump_to_col_end();
        }
        (KeyModifiers::CONTROL, KeyCode::Char('d')) => {
            let half = (app.viewport.visible_rows / 2) as i32;
            app.move_cursor(half, 0);
        }
        (KeyModifiers::CONTROL, KeyCode::Char('u')) => {
            let half = (app.viewport.visible_rows / 2) as i32;
            app.move_cursor(-half, 0);
        }
        (KeyModifiers::NONE, KeyCode::Char('y')) => {
            app.yank_visual();
        }
        (KeyModifiers::NONE, KeyCode::Char('d')) | (KeyModifiers::NONE, KeyCode::Char('x')) => {
            app.delete_visual();
        }
        // Fill visual selection from top-left cell
        (KeyModifiers::NONE, KeyCode::Char('f')) => {
            app.visual_fill();
        }
        (KeyModifiers::NONE, KeyCode::Char(':')) => {
            app.mode = Mode::Command;
            app.command_buffer.clear();
        }
        _ => {}
    }
}

fn handle_mouse(app: &mut App, mouse: MouseEvent) {
    if app.sort_dialog.visible {
        if let MouseEventKind::Down(MouseButton::Left) = mouse.kind {
            handle_sort_dialog_mouse(app, mouse.column, mouse.row);
        }
        return;
    }

    match mouse.kind {
        MouseEventKind::Down(MouseButton::Left) => {
            let x = mouse.column;
            let y = mouse.row;

            // Check sheet tabs (last row - 1 before status bar)
            let term_height = app.viewport.visible_rows as u16 + 4; // approx
            if y == term_height.saturating_sub(2) {
                if let Some(idx) = sheet_tab_at(app, x) {
                    if app.mode == Mode::Insert && app.edit_buffer.starts_with('=') {
                        app.switch_sheet_for_ref(idx);
                    } else {
                        app.switch_sheet(idx);
                    }
                    return;
                }
            }

            // Click on grid cell
            if let Some((data_row, data_col)) = screen_to_cell(app, x, y) {
                if app.mode == Mode::Insert && app.edit_buffer.starts_with('=') {
                    let viewing_sheet = app.workbook.active_sheet as u16;
                    let formula_sheet = app.formula_origin_sheet.unwrap_or(viewing_sheet);
                    let cell_ref = if viewing_sheet != formula_sheet {
                        let sheet_name = &app.workbook.sheets[viewing_sheet as usize].name;
                        if sheet_name.contains(' ') {
                            format!("'{}'!{}", sheet_name, CellAddr::new(viewing_sheet, data_row, data_col).display_name())
                        } else {
                            format!("{}!{}", sheet_name, CellAddr::new(viewing_sheet, data_row, data_col).display_name())
                        }
                    } else {
                        CellAddr::new(viewing_sheet, data_row, data_col).display_name()
                    };
                    app.edit_buffer.push_str(&cell_ref);
                    app.update_autocomplete();
                    return;
                }
                if app.mode == Mode::Insert {
                    app.confirm_edit();
                }
                // Click on header row opens filter dialog
                if app.is_header_row(data_row) && app.mode != Mode::Visual {
                    app.cursor.row = data_row;
                    app.cursor.col = data_col;
                    app.open_filter_dialog(Some(data_col));
                    return;
                }
                app.cursor.row = data_row;
                app.cursor.col = data_col;
                if app.mode == Mode::Visual {
                    app.visual_anchor = None;
                    app.mode = Mode::Normal;
                }
            }
        }
        MouseEventKind::ScrollUp => {
            app.move_cursor(-3, 0);
        }
        MouseEventKind::ScrollDown => {
            app.move_cursor(3, 0);
        }
        _ => {}
    }
}

fn screen_to_cell(app: &App, x: u16, y: u16) -> Option<(u32, u16)> {
    let row_num_width: u16 = 6;
    let grid_start_row: u16 = 2; // formula bar + col headers

    if y < grid_start_row || x < row_num_width {
        return None;
    }

    let (freeze_rows, freeze_cols) = app.workbook.active_sheet().freeze.unwrap_or((0, 0));
    let has_freeze_row = freeze_rows > 0;
    let has_freeze_col = freeze_cols > 0;

    // Row mapping
    let screen_row = (y - grid_start_row) as u32;
    let freeze_border_row = if has_freeze_row { 1u32 } else { 0 };

    let data_row = if screen_row < freeze_rows {
        screen_row
    } else if has_freeze_row && screen_row == freeze_rows {
        return None; // clicked on freeze border line
    } else {
        let scroll_screen_row = screen_row - freeze_rows - freeze_border_row;
        app.viewport.top_row + scroll_screen_row
    };

    // Column mapping
    let col_px = x - row_num_width;
    let frozen_px = freeze_cols as u16 * app.viewport.col_width;
    let freeze_border_px = if has_freeze_col { 1u16 } else { 0 };

    let data_col = if col_px < frozen_px {
        col_px / app.viewport.col_width
    } else if has_freeze_col && col_px < frozen_px + freeze_border_px {
        return None; // clicked on freeze border
    } else {
        let scroll_px = col_px - frozen_px - freeze_border_px;
        app.viewport.left_col + scroll_px / app.viewport.col_width
    };

    Some((data_row, data_col))
}

fn sheet_tab_at(app: &App, x: u16) -> Option<usize> {
    let mut offset: u16 = 0;
    for (i, sheet) in app.workbook.sheets.iter().enumerate() {
        let tab_width = sheet.name.len() as u16 + 2 + 1; // " name " + separator
        if x >= offset && x < offset + tab_width {
            return Some(i);
        }
        offset += tab_width;
    }
    None
}

fn handle_sort_dialog(app: &mut App, key: KeyEvent) {
    // If a dropdown is open, handle it first
    if app.sort_dialog.open_dropdown != SortDropdown::None {
        handle_sort_dropdown_key(app, key);
        return;
    }

    match key.code {
        KeyCode::Esc => {
            app.sort_dialog.visible = false;
        }
        KeyCode::Tab | KeyCode::Down | KeyCode::Char('j') => {
            app.sort_dialog.focus = app.sort_dialog.focus.next();
        }
        KeyCode::BackTab | KeyCode::Up | KeyCode::Char('k') => {
            app.sort_dialog.focus = app.sort_dialog.focus.prev();
        }
        KeyCode::Enter | KeyCode::Char(' ') => {
            match app.sort_dialog.focus {
                SortDialogFocus::BtnSort => app.confirm_sort_dialog(),
                SortDialogFocus::BtnCancel => app.sort_dialog.visible = false,
                SortDialogFocus::HasHeaders => {
                    app.sort_dialog.has_headers = !app.sort_dialog.has_headers;
                }
                SortDialogFocus::SortByCol => {
                    app.sort_dialog.open_dropdown = SortDropdown::SortByCol;
                    app.sort_dialog.dropdown_scroll = 0;
                }
                SortDialogFocus::SortByType => {
                    app.sort_dialog.open_dropdown = SortDropdown::SortByType;
                    app.sort_dialog.dropdown_scroll = 0;
                }
                SortDialogFocus::ThenByCol => {
                    app.sort_dialog.open_dropdown = SortDropdown::ThenByCol;
                    app.sort_dialog.dropdown_scroll = 0;
                }
                SortDialogFocus::ThenByType => {
                    app.sort_dialog.open_dropdown = SortDropdown::ThenByType;
                    app.sort_dialog.dropdown_scroll = 0;
                }
                SortDialogFocus::Range => {}
            }
        }
        KeyCode::Left | KeyCode::Char('h') => {
            match app.sort_dialog.focus {
                SortDialogFocus::SortByCol => {
                    if app.sort_dialog.sort_by_col > app.sort_dialog.min_col {
                        app.sort_dialog.sort_by_col -= 1;
                    }
                }
                SortDialogFocus::SortByType => {
                    app.sort_dialog.sort_by_type = app.sort_dialog.sort_by_type.prev();
                }
                SortDialogFocus::ThenByCol => {
                    let min = app.sort_dialog.min_col;
                    app.sort_dialog.then_by_col = Some(match app.sort_dialog.then_by_col {
                        Some(c) if c > min => c - 1,
                        Some(c) => c,
                        None => min,
                    });
                }
                SortDialogFocus::ThenByType => {
                    app.sort_dialog.then_by_type = app.sort_dialog.then_by_type.prev();
                }
                _ => {}
            }
        }
        KeyCode::Right | KeyCode::Char('l') => {
            match app.sort_dialog.focus {
                SortDialogFocus::SortByCol => {
                    if app.sort_dialog.sort_by_col < app.sort_dialog.max_col {
                        app.sort_dialog.sort_by_col += 1;
                    }
                }
                SortDialogFocus::SortByType => {
                    app.sort_dialog.sort_by_type = app.sort_dialog.sort_by_type.next();
                }
                SortDialogFocus::ThenByCol => {
                    let max = app.sort_dialog.max_col;
                    app.sort_dialog.then_by_col = Some(match app.sort_dialog.then_by_col {
                        Some(c) if c < max => c + 1,
                        Some(c) => c,
                        None => app.sort_dialog.min_col,
                    });
                }
                SortDialogFocus::ThenByType => {
                    app.sort_dialog.then_by_type = app.sort_dialog.then_by_type.next();
                }
                _ => {}
            }
        }
        KeyCode::Backspace => {
            if app.sort_dialog.focus == SortDialogFocus::ThenByCol {
                app.sort_dialog.then_by_col = None;
            }
        }
        _ => {}
    }
}

fn handle_sort_dropdown_key(app: &mut App, key: KeyEvent) {
    match key.code {
        KeyCode::Esc => {
            app.sort_dialog.open_dropdown = SortDropdown::None;
        }
        KeyCode::Up | KeyCode::Char('k') => {
            if app.sort_dialog.dropdown_scroll > 0 {
                app.sort_dialog.dropdown_scroll -= 1;
            }
            select_sort_dropdown_item(app, app.sort_dialog.dropdown_scroll);
        }
        KeyCode::Down | KeyCode::Char('j') => {
            app.sort_dialog.dropdown_scroll += 1;
            let max = sort_dropdown_item_count(app);
            if app.sort_dialog.dropdown_scroll >= max {
                app.sort_dialog.dropdown_scroll = max.saturating_sub(1);
            }
            select_sort_dropdown_item(app, app.sort_dialog.dropdown_scroll);
        }
        KeyCode::Enter | KeyCode::Char(' ') => {
            select_sort_dropdown_item(app, app.sort_dialog.dropdown_scroll);
            app.sort_dialog.open_dropdown = SortDropdown::None;
        }
        _ => {}
    }
}

fn sort_dropdown_item_count(app: &App) -> usize {
    match app.sort_dialog.open_dropdown {
        SortDropdown::SortByCol | SortDropdown::ThenByCol => {
            let extra = if app.sort_dialog.open_dropdown == SortDropdown::ThenByCol { 1 } else { 0 };
            (app.sort_dialog.max_col - app.sort_dialog.min_col + 1) as usize + extra
        }
        SortDropdown::SortByType | SortDropdown::ThenByType => SortType::all().len(),
        SortDropdown::None => 0,
    }
}

fn select_sort_dropdown_item(app: &mut App, idx: usize) {
    match app.sort_dialog.open_dropdown {
        SortDropdown::SortByCol => {
            let col = app.sort_dialog.min_col + idx as u16;
            if col <= app.sort_dialog.max_col {
                app.sort_dialog.sort_by_col = col;
            }
        }
        SortDropdown::SortByType => {
            if let Some(t) = SortType::all().get(idx) {
                app.sort_dialog.sort_by_type = *t;
            }
        }
        SortDropdown::ThenByCol => {
            if idx == 0 {
                app.sort_dialog.then_by_col = None;
            } else {
                let col = app.sort_dialog.min_col + (idx as u16 - 1);
                if col <= app.sort_dialog.max_col {
                    app.sort_dialog.then_by_col = Some(col);
                }
            }
        }
        SortDropdown::ThenByType => {
            if let Some(t) = SortType::all().get(idx) {
                app.sort_dialog.then_by_type = *t;
            }
        }
        SortDropdown::None => {}
    }
}

pub fn handle_sort_dialog_mouse(app: &mut App, x: u16, y: u16) -> bool {
    if !app.sort_dialog.visible {
        return false;
    }

    let dx = app.sort_dialog.screen_x;
    let dy = app.sort_dialog.screen_y;
    let area_w = app.sort_dialog.screen_w;
    let area_h = app.sort_dialog.screen_h;

    let inner_x = dx + 1;
    let inner_y = dy + 1;

    // If a dropdown is open, check clicks on dropdown items first
    if app.sort_dialog.open_dropdown != SortDropdown::None {
        let item_x = app.sort_dialog.dd_item_x;
        let item_y = app.sort_dialog.dd_item_y;
        let item_w = app.sort_dialog.dd_item_w;
        let item_count = app.sort_dialog.dd_item_count;

        if x >= item_x && x < item_x + item_w && y >= item_y && y < item_y + item_count {
            let idx = (y - item_y) as usize;
            select_sort_dropdown_item(app, idx);
            app.sort_dialog.open_dropdown = SortDropdown::None;
            return true;
        }
        // Click outside dropdown closes it
        app.sort_dialog.open_dropdown = SortDropdown::None;
        return true;
    }

    // Check if click is inside dialog
    if x < dx || x >= dx + area_w || y < dy || y >= dy + area_h {
        app.sort_dialog.visible = false;
        return true;
    }

    // Relative to inner area
    if x < inner_x || y < inner_y {
        return true;
    }
    let rel_x = x - inner_x;
    let rel_y = y - inner_y;

    // Row 0: Range
    // Row 2: Sort by col | Sort by type
    // Row 3: Then by col | Then by type
    // Row 5: Headers
    // Row 7: Sort | Cancel buttons

    match rel_y {
        0 => {
            app.sort_dialog.focus = SortDialogFocus::Range;
        }
        2 => {
            if rel_x >= 11 && rel_x < 20 {
                app.sort_dialog.focus = SortDialogFocus::SortByCol;
                app.sort_dialog.open_dropdown = SortDropdown::SortByCol;
                app.sort_dialog.dropdown_scroll = (app.sort_dialog.sort_by_col - app.sort_dialog.min_col) as usize;
            } else if rel_x >= 20 {
                app.sort_dialog.focus = SortDialogFocus::SortByType;
                app.sort_dialog.open_dropdown = SortDropdown::SortByType;
                app.sort_dialog.dropdown_scroll = SortType::all().iter().position(|t| *t == app.sort_dialog.sort_by_type).unwrap_or(0);
            }
        }
        3 => {
            if rel_x >= 11 && rel_x < 20 {
                app.sort_dialog.focus = SortDialogFocus::ThenByCol;
                app.sort_dialog.open_dropdown = SortDropdown::ThenByCol;
                app.sort_dialog.dropdown_scroll = match app.sort_dialog.then_by_col {
                    Some(c) => (c - app.sort_dialog.min_col + 1) as usize,
                    None => 0,
                };
            } else if rel_x >= 20 {
                app.sort_dialog.focus = SortDialogFocus::ThenByType;
                app.sort_dialog.open_dropdown = SortDropdown::ThenByType;
                app.sort_dialog.dropdown_scroll = SortType::all().iter().position(|t| *t == app.sort_dialog.then_by_type).unwrap_or(0);
            }
        }
        5 => {
            if rel_x >= 11 {
                app.sort_dialog.focus = SortDialogFocus::HasHeaders;
                app.sort_dialog.has_headers = !app.sort_dialog.has_headers;
            }
        }
        7 => {
            if rel_x >= 5 && rel_x < 13 {
                app.confirm_sort_dialog();
            } else if rel_x >= 17 && rel_x < 25 {
                app.sort_dialog.visible = false;
            }
        }
        _ => {}
    }

    true
}

// --- Filter Dialog ---

fn handle_filter_dialog(app: &mut App, key: KeyEvent) {
    if app.filter_dialog.open_dropdown != FilterDropdown::None {
        handle_filter_dropdown_key(app, key);
        return;
    }

    // Eq filter: multi-select value list with search
    if app.filter_dialog.focus == FilterDialogFocus::Value && app.filter_dialog.filter_type == FilterType::Eq {
        // value_selected 0 = "All" row, 1..N = actual values
        let list_len = app.filter_dialog.filtered_values.len() + 1; // +1 for "All"
        match key.code {
            KeyCode::Enter | KeyCode::Char(' ') => {
                if app.filter_dialog.value_selected == 0 {
                    app.toggle_filter_all();
                } else {
                    app.toggle_filter_value(app.filter_dialog.value_selected - 1);
                }
                return;
            }
            KeyCode::Char(c) => {
                app.filter_dialog.value_search.push(c);
                app.update_filter_value_list();
                return;
            }
            KeyCode::Backspace => {
                app.filter_dialog.value_search.pop();
                app.update_filter_value_list();
                return;
            }
            KeyCode::Down => {
                if app.filter_dialog.value_selected + 1 < list_len {
                    app.filter_dialog.value_selected += 1;
                }
                let max_visible: usize = 8;
                if app.filter_dialog.value_selected >= app.filter_dialog.value_scroll + max_visible {
                    app.filter_dialog.value_scroll = app.filter_dialog.value_selected - max_visible + 1;
                }
                return;
            }
            KeyCode::Up => {
                if app.filter_dialog.value_selected > 0 {
                    app.filter_dialog.value_selected -= 1;
                }
                if app.filter_dialog.value_selected < app.filter_dialog.value_scroll {
                    app.filter_dialog.value_scroll = app.filter_dialog.value_selected;
                }
                return;
            }
            KeyCode::Esc => {
                app.filter_dialog.visible = false;
                return;
            }
            KeyCode::Tab => {}
            KeyCode::BackTab => {}
            _ => { return; }
        }
    }

    // Non-Eq filter types: text input for value
    if app.filter_dialog.focus == FilterDialogFocus::Value && app.filter_dialog.filter_type.needs_value() && app.filter_dialog.filter_type != FilterType::Eq {
        match key.code {
            KeyCode::Char(c) => {
                app.filter_dialog.value_buf.push(c);
                return;
            }
            KeyCode::Backspace => {
                app.filter_dialog.value_buf.pop();
                return;
            }
            KeyCode::Esc => {
                app.filter_dialog.visible = false;
                return;
            }
            KeyCode::Tab => {}
            KeyCode::BackTab => {}
            _ => { return; }
        }
    }

    match key.code {
        KeyCode::Esc => {
            app.filter_dialog.visible = false;
        }
        KeyCode::Tab | KeyCode::Down | KeyCode::Char('j') => {
            app.filter_dialog.focus = app.filter_dialog.focus.next();
            app.filter_dialog.value_selected = 0;
            app.filter_dialog.value_scroll = 0;
        }
        KeyCode::BackTab | KeyCode::Up | KeyCode::Char('k') => {
            app.filter_dialog.focus = app.filter_dialog.focus.prev();
            app.filter_dialog.value_selected = 0;
            app.filter_dialog.value_scroll = 0;
        }
        KeyCode::Enter | KeyCode::Char(' ') => {
            match app.filter_dialog.focus {
                FilterDialogFocus::BtnApply => app.confirm_filter_dialog(),
                FilterDialogFocus::BtnCancel => app.filter_dialog.visible = false,
                FilterDialogFocus::Column => {
                    app.filter_dialog.open_dropdown = FilterDropdown::Column;
                    app.filter_dialog.dropdown_scroll = (app.filter_dialog.col - app.filter_dialog.min_col) as usize;
                }
                FilterDialogFocus::FilterType => {
                    app.filter_dialog.open_dropdown = FilterDropdown::FilterType;
                    app.filter_dialog.dropdown_scroll = FilterType::all().iter().position(|t| *t == app.filter_dialog.filter_type).unwrap_or(0);
                }
                FilterDialogFocus::Value => {}
            }
        }
        KeyCode::Left | KeyCode::Char('h') => {
            match app.filter_dialog.focus {
                FilterDialogFocus::Column => {
                    if app.filter_dialog.col > app.filter_dialog.min_col {
                        app.filter_dialog.col -= 1;
                        app.refresh_filter_column_values();
                    }
                }
                FilterDialogFocus::FilterType => {
                    let all = FilterType::all();
                    let idx = all.iter().position(|t| *t == app.filter_dialog.filter_type).unwrap_or(0);
                    app.filter_dialog.filter_type = all[(idx + all.len() - 1) % all.len()];
                }
                _ => {}
            }
        }
        KeyCode::Right | KeyCode::Char('l') => {
            match app.filter_dialog.focus {
                FilterDialogFocus::Column => {
                    if app.filter_dialog.col < app.filter_dialog.max_col {
                        app.filter_dialog.col += 1;
                        app.refresh_filter_column_values();
                    }
                }
                FilterDialogFocus::FilterType => {
                    let all = FilterType::all();
                    let idx = all.iter().position(|t| *t == app.filter_dialog.filter_type).unwrap_or(0);
                    app.filter_dialog.filter_type = all[(idx + 1) % all.len()];
                }
                _ => {}
            }
        }
        _ => {}
    }
}

fn handle_filter_dropdown_key(app: &mut App, key: KeyEvent) {
    match key.code {
        KeyCode::Esc => {
            app.filter_dialog.open_dropdown = FilterDropdown::None;
        }
        KeyCode::Up | KeyCode::Char('k') => {
            if app.filter_dialog.dropdown_scroll > 0 {
                app.filter_dialog.dropdown_scroll -= 1;
            }
        }
        KeyCode::Down | KeyCode::Char('j') => {
            let max = filter_dropdown_item_count(app);
            if app.filter_dialog.dropdown_scroll + 1 < max {
                app.filter_dialog.dropdown_scroll += 1;
            }
        }
        KeyCode::Enter | KeyCode::Char(' ') => {
            select_filter_dropdown_item(app, app.filter_dialog.dropdown_scroll);
            app.filter_dialog.open_dropdown = FilterDropdown::None;
        }
        _ => {}
    }
}

fn filter_dropdown_item_count(app: &App) -> usize {
    match app.filter_dialog.open_dropdown {
        FilterDropdown::Column => (app.filter_dialog.max_col - app.filter_dialog.min_col + 1) as usize,
        FilterDropdown::FilterType => FilterType::all().len(),
        FilterDropdown::None => 0,
    }
}

fn select_filter_dropdown_item(app: &mut App, idx: usize) {
    match app.filter_dialog.open_dropdown {
        FilterDropdown::Column => {
            let col = app.filter_dialog.min_col + idx as u16;
            if col <= app.filter_dialog.max_col {
                app.filter_dialog.col = col;
                app.refresh_filter_column_values();
            }
        }
        FilterDropdown::FilterType => {
            if let Some(t) = FilterType::all().get(idx) {
                app.filter_dialog.filter_type = *t;
            }
        }
        FilterDropdown::None => {}
    }
}

fn handle_filter_dialog_mouse(app: &mut App, x: u16, y: u16) {
    if !app.filter_dialog.visible {
        return;
    }

    let dx = app.filter_dialog.screen_x;
    let dy = app.filter_dialog.screen_y;
    let area_w = app.filter_dialog.screen_w;
    let area_h = app.filter_dialog.screen_h;
    let inner_x = dx + 1;
    let inner_y = dy + 1;

    // Dropdown open — check dropdown clicks first
    if app.filter_dialog.open_dropdown != FilterDropdown::None {
        let item_x = app.filter_dialog.dd_item_x;
        let item_y = app.filter_dialog.dd_item_y;
        let item_w = app.filter_dialog.dd_item_w;
        let item_count = app.filter_dialog.dd_item_count;

        if x >= item_x && x < item_x + item_w && y >= item_y && y < item_y + item_count {
            let idx = (y - item_y) as usize;
            select_filter_dropdown_item(app, idx);
            app.filter_dialog.open_dropdown = FilterDropdown::None;
            return;
        }
        app.filter_dialog.open_dropdown = FilterDropdown::None;
        return;
    }

    // Outside dialog
    if x < dx || x >= dx + area_w || y < dy || y >= dy + area_h {
        app.filter_dialog.visible = false;
        return;
    }

    if x < inner_x || y < inner_y {
        return;
    }
    let rel_x = x - inner_x;
    let rel_y = y - inner_y;

    // Row 1: Column
    // Row 2: Type
    // Row 3: Value
    // Row 6: Apply | Cancel
    match rel_y {
        1 => {
            if rel_x >= 10 {
                app.filter_dialog.focus = FilterDialogFocus::Column;
                app.filter_dialog.open_dropdown = FilterDropdown::Column;
                app.filter_dialog.dropdown_scroll = (app.filter_dialog.col - app.filter_dialog.min_col) as usize;
            }
        }
        2 => {
            if rel_x >= 10 {
                app.filter_dialog.focus = FilterDialogFocus::FilterType;
                app.filter_dialog.open_dropdown = FilterDropdown::FilterType;
                app.filter_dialog.dropdown_scroll = FilterType::all().iter().position(|t| *t == app.filter_dialog.filter_type).unwrap_or(0);
            }
        }
        3 => {
            if rel_x >= 10 {
                app.filter_dialog.focus = FilterDialogFocus::Value;
            }
        }
        6 => {
            if rel_x >= 5 && rel_x < 14 {
                app.confirm_filter_dialog();
            } else if rel_x >= 18 && rel_x < 26 {
                app.filter_dialog.visible = false;
            }
        }
        _ => {}
    }
}

fn handle_cf_dialog(app: &mut App, key: KeyEvent) {
    // Dropdown open takes precedence
    if app.cf_dialog.open_dropdown != CfDropdown::None {
        handle_cf_dropdown_key(app, key);
        return;
    }

    // Text input only when explicitly editing a text field
    let is_text_field = matches!(
        app.cf_dialog.focus,
        CfDialogFocus::Range | CfDialogFocus::Val1 | CfDialogFocus::Val2
    );
    if app.cf_dialog.editing_text && is_text_field {
        match key.code {
            KeyCode::Char(c) => {
                let buf = match app.cf_dialog.focus {
                    CfDialogFocus::Range => &mut app.cf_dialog.range_text,
                    CfDialogFocus::Val1 => &mut app.cf_dialog.val1,
                    CfDialogFocus::Val2 => &mut app.cf_dialog.val2,
                    _ => unreachable!(),
                };
                buf.push(c);
                return;
            }
            KeyCode::Backspace => {
                let buf = match app.cf_dialog.focus {
                    CfDialogFocus::Range => &mut app.cf_dialog.range_text,
                    CfDialogFocus::Val1 => &mut app.cf_dialog.val1,
                    CfDialogFocus::Val2 => &mut app.cf_dialog.val2,
                    _ => unreachable!(),
                };
                buf.pop();
                return;
            }
            KeyCode::Esc | KeyCode::Enter => {
                app.cf_dialog.editing_text = false;
                return;
            }
            KeyCode::Tab => {
                app.cf_dialog.editing_text = false;
                app.cf_dialog.focus = app.cf_dialog.focus.next(app.cf_dialog.conditional);
                return;
            }
            KeyCode::BackTab => {
                app.cf_dialog.editing_text = false;
                app.cf_dialog.focus = app.cf_dialog.focus.prev(app.cf_dialog.conditional);
                return;
            }
            _ => return,
        }
    }

    match key.code {
        KeyCode::Esc => { app.cf_dialog.visible = false; }
        KeyCode::Tab | KeyCode::Down | KeyCode::Char('j') => {
            if app.cf_dialog.focus == CfDialogFocus::RulesList {
                let n = app.workbook.active_sheet().cond_rules.len();
                if n > 0 && app.cf_dialog.rules_selected + 1 < n {
                    app.cf_dialog.rules_selected += 1;
                } else {
                    app.cf_dialog.focus = app.cf_dialog.focus.next(app.cf_dialog.conditional);
                }
            } else {
                app.cf_dialog.focus = app.cf_dialog.focus.next(app.cf_dialog.conditional);
            }
        }
        KeyCode::BackTab | KeyCode::Up | KeyCode::Char('k') => {
            if app.cf_dialog.focus == CfDialogFocus::RulesList {
                if app.cf_dialog.rules_selected > 0 {
                    app.cf_dialog.rules_selected -= 1;
                } else {
                    app.cf_dialog.focus = app.cf_dialog.focus.prev(app.cf_dialog.conditional);
                }
            } else {
                app.cf_dialog.focus = app.cf_dialog.focus.prev(app.cf_dialog.conditional);
            }
        }
        KeyCode::Left | KeyCode::Char('h') => {
            match app.cf_dialog.focus {
                CfDialogFocus::Cond => {
                    let all = CfCond::all();
                    let idx = all.iter().position(|t| *t == app.cf_dialog.cond).unwrap_or(0);
                    app.cf_dialog.cond = all[(idx + all.len() - 1) % all.len()];
                }
                CfDialogFocus::Bg => {
                    if app.cf_dialog.bg_idx == 0 {
                        app.cf_dialog.bg_idx = CF_COLORS.len() - 1;
                    } else {
                        app.cf_dialog.bg_idx -= 1;
                    }
                }
                CfDialogFocus::Fg => {
                    if app.cf_dialog.fg_idx == 0 {
                        app.cf_dialog.fg_idx = CF_COLORS.len() - 1;
                    } else {
                        app.cf_dialog.fg_idx -= 1;
                    }
                }
                _ => {}
            }
        }
        KeyCode::Right | KeyCode::Char('l') => {
            match app.cf_dialog.focus {
                CfDialogFocus::Cond => {
                    let all = CfCond::all();
                    let idx = all.iter().position(|t| *t == app.cf_dialog.cond).unwrap_or(0);
                    app.cf_dialog.cond = all[(idx + 1) % all.len()];
                }
                CfDialogFocus::Bg => {
                    app.cf_dialog.bg_idx = (app.cf_dialog.bg_idx + 1) % CF_COLORS.len();
                }
                CfDialogFocus::Fg => {
                    app.cf_dialog.fg_idx = (app.cf_dialog.fg_idx + 1) % CF_COLORS.len();
                }
                _ => {}
            }
        }
        KeyCode::Enter | KeyCode::Char(' ') => {
            match app.cf_dialog.focus {
                CfDialogFocus::Range | CfDialogFocus::Val1 | CfDialogFocus::Val2 => {
                    app.cf_dialog.editing_text = true;
                }
                CfDialogFocus::Conditional => {
                    app.cf_dialog.conditional = !app.cf_dialog.conditional;
                }
                CfDialogFocus::Bold => app.cf_dialog.bold = !app.cf_dialog.bold,
                CfDialogFocus::Italic => app.cf_dialog.italic = !app.cf_dialog.italic,
                CfDialogFocus::Under => app.cf_dialog.under = !app.cf_dialog.under,
                CfDialogFocus::DUnder => app.cf_dialog.dunder = !app.cf_dialog.dunder,
                CfDialogFocus::Strike => app.cf_dialog.strike = !app.cf_dialog.strike,
                CfDialogFocus::Over => app.cf_dialog.over = !app.cf_dialog.over,
                CfDialogFocus::Cond => {
                    app.cf_dialog.open_dropdown = CfDropdown::Cond;
                    let idx = CfCond::all().iter().position(|t| *t == app.cf_dialog.cond).unwrap_or(0);
                    app.cf_dialog.dropdown_scroll = idx;
                }
                CfDialogFocus::Bg => {
                    app.cf_dialog.open_dropdown = CfDropdown::Bg;
                    app.cf_dialog.dropdown_scroll = app.cf_dialog.bg_idx;
                }
                CfDialogFocus::Fg => {
                    app.cf_dialog.open_dropdown = CfDropdown::Fg;
                    app.cf_dialog.dropdown_scroll = app.cf_dialog.fg_idx;
                }
                CfDialogFocus::BtnApply => app.cf_dialog_apply(),
                CfDialogFocus::BtnBase => app.cf_dialog_set_base(),
                CfDialogFocus::BtnDelete => app.cf_dialog_delete_selected(),
                CfDialogFocus::BtnCleanAll => app.cf_dialog_clean_all(),
                CfDialogFocus::BtnClose => app.cf_dialog.visible = false,
                _ => {}
            }
        }
        _ => {}
    }
}

fn handle_cf_dropdown_key(app: &mut App, key: KeyEvent) {
    match key.code {
        KeyCode::Esc => {
            app.cf_dialog.open_dropdown = CfDropdown::None;
        }
        KeyCode::Up | KeyCode::Char('k') => {
            if app.cf_dialog.dropdown_scroll > 0 {
                app.cf_dialog.dropdown_scroll -= 1;
            }
        }
        KeyCode::Down | KeyCode::Char('j') => {
            let max = cf_dropdown_item_count(app);
            if app.cf_dialog.dropdown_scroll + 1 < max {
                app.cf_dialog.dropdown_scroll += 1;
            }
        }
        KeyCode::Enter | KeyCode::Char(' ') => {
            select_cf_dropdown_item(app, app.cf_dialog.dropdown_scroll);
            app.cf_dialog.open_dropdown = CfDropdown::None;
        }
        _ => {}
    }
}

fn cf_dropdown_item_count(app: &App) -> usize {
    match app.cf_dialog.open_dropdown {
        CfDropdown::Cond => CfCond::all().len(),
        CfDropdown::Bg | CfDropdown::Fg => CF_COLORS.len(),
        CfDropdown::None => 0,
    }
}

fn select_cf_dropdown_item(app: &mut App, idx: usize) {
    match app.cf_dialog.open_dropdown {
        CfDropdown::Cond => {
            if let Some(c) = CfCond::all().get(idx) {
                app.cf_dialog.cond = *c;
            }
        }
        CfDropdown::Bg => {
            if idx < CF_COLORS.len() { app.cf_dialog.bg_idx = idx; }
        }
        CfDropdown::Fg => {
            if idx < CF_COLORS.len() { app.cf_dialog.fg_idx = idx; }
        }
        CfDropdown::None => {}
    }
}

fn handle_cf_dialog_mouse(app: &mut App, x: u16, y: u16) {
    if !app.cf_dialog.visible { return; }

    // Dropdown open: check dropdown clicks first
    if app.cf_dialog.open_dropdown != CfDropdown::None {
        let ix = app.cf_dialog.dd_item_x;
        let iy = app.cf_dialog.dd_item_y;
        let iw = app.cf_dialog.dd_item_w;
        let count = app.cf_dialog.dd_item_count;
        if x >= ix && x < ix + iw && y >= iy && y < iy + count {
            let idx = (y - iy) as usize;
            select_cf_dropdown_item(app, idx);
            app.cf_dialog.open_dropdown = CfDropdown::None;
            return;
        }
        app.cf_dialog.open_dropdown = CfDropdown::None;
        return;
    }

    let hit = |rect: (u16,u16,u16), x: u16, y: u16| -> bool {
        x >= rect.0 && x < rect.0 + rect.2 && y == rect.1
    };
    let d = &mut app.cf_dialog;
    if hit(d.rect_conditional, x, y) {
        d.focus = CfDialogFocus::Conditional;
        d.conditional = !d.conditional;
        return;
    }
    if hit(d.rect_range, x, y) { d.focus = CfDialogFocus::Range; d.editing_text = true; return; }
    if hit(d.rect_cond, x, y) {
        d.focus = CfDialogFocus::Cond;
        d.open_dropdown = CfDropdown::Cond;
        d.dropdown_scroll = CfCond::all().iter().position(|c| *c == d.cond).unwrap_or(0);
        return;
    }
    if hit(d.rect_val1, x, y) { d.focus = CfDialogFocus::Val1; d.editing_text = true; return; }
    if hit(d.rect_val2, x, y) { d.focus = CfDialogFocus::Val2; d.editing_text = true; return; }
    if hit(d.rect_bold, x, y) { d.focus = CfDialogFocus::Bold; d.bold = !d.bold; return; }
    if hit(d.rect_italic, x, y) { d.focus = CfDialogFocus::Italic; d.italic = !d.italic; return; }
    if hit(d.rect_under, x, y) { d.focus = CfDialogFocus::Under; d.under = !d.under; return; }
    if hit(d.rect_dunder, x, y) { d.focus = CfDialogFocus::DUnder; d.dunder = !d.dunder; return; }
    if hit(d.rect_strike, x, y) { d.focus = CfDialogFocus::Strike; d.strike = !d.strike; return; }
    if hit(d.rect_over, x, y) { d.focus = CfDialogFocus::Over; d.over = !d.over; return; }
    if hit(d.rect_bg, x, y) {
        d.focus = CfDialogFocus::Bg;
        d.open_dropdown = CfDropdown::Bg;
        d.dropdown_scroll = d.bg_idx;
        return;
    }
    if hit(d.rect_fg, x, y) {
        d.focus = CfDialogFocus::Fg;
        d.open_dropdown = CfDropdown::Fg;
        d.dropdown_scroll = d.fg_idx;
        return;
    }
    let (rx, ry, rw, rh) = d.rect_rules;
    if x >= rx && x < rx + rw && y >= ry && y < ry + rh {
        d.focus = CfDialogFocus::RulesList;
        let idx = (y - ry) as usize + d.rules_scroll;
        let n = app.workbook.active_sheet().cond_rules.len();
        if idx < n {
            app.cf_dialog.rules_selected = idx;
        }
        return;
    }
    if hit(d.rect_apply, x, y) { d.focus = CfDialogFocus::BtnApply; app.cf_dialog_apply(); return; }
    if hit(d.rect_base, x, y) { d.focus = CfDialogFocus::BtnBase; app.cf_dialog_set_base(); return; }
    if hit(d.rect_delete, x, y) { d.focus = CfDialogFocus::BtnDelete; app.cf_dialog_delete_selected(); return; }
    if hit(d.rect_cleanall, x, y) { d.focus = CfDialogFocus::BtnCleanAll; app.cf_dialog_clean_all(); return; }
    if hit(d.rect_close, x, y) { app.cf_dialog.visible = false; return; }
}
