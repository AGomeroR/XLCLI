use crossterm::event::{Event, KeyCode, KeyEvent, KeyModifiers, MouseButton, MouseEvent, MouseEventKind};

use crate::app::App;
use crate::mode::Mode;

pub fn handle_event(app: &mut App, event: Event) {
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
        }
        KeyCode::Backspace => {
            app.command_buffer.pop();
            if app.command_buffer.is_empty() {
                app.mode = Mode::Normal;
            }
        }
        KeyCode::Char(c) => {
            app.command_buffer.push(c);
        }
        _ => {}
    }
}

fn handle_insert(app: &mut App, key: KeyEvent) {
    match key.code {
        KeyCode::Esc => {
            app.confirm_edit();
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
        }
        KeyCode::Char(c) => {
            app.edit_buffer.push(c);
        }
        KeyCode::Left => {
            // cursor within edit buffer — not implemented yet, just ignore
        }
        KeyCode::Right => {}
        _ => {}
    }
}

fn handle_visual(app: &mut App, key: KeyEvent) {
    match (key.modifiers, key.code) {
        (_, KeyCode::Esc) => {
            app.visual_anchor = None;
            app.mode = Mode::Normal;
            app.status_message = None;
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
        (KeyModifiers::NONE, KeyCode::Char('y')) => {
            app.yank_visual();
        }
        (KeyModifiers::NONE, KeyCode::Char('d')) | (KeyModifiers::NONE, KeyCode::Char('x')) => {
            app.delete_visual();
        }
        _ => {}
    }
}

fn handle_mouse(app: &mut App, mouse: MouseEvent) {
    match mouse.kind {
        MouseEventKind::Down(MouseButton::Left) => {
            let x = mouse.column;
            let y = mouse.row;

            // Check sheet tabs (last row - 1 before status bar)
            let term_height = app.viewport.visible_rows as u16 + 4; // approx
            if y == term_height.saturating_sub(2) {
                if let Some(idx) = sheet_tab_at(app, x) {
                    app.workbook.active_sheet = idx;
                    app.cursor.sheet = idx as u16;
                    app.status_message = Some(format!("Sheet: {}", app.workbook.sheets[idx].name));
                    return;
                }
            }

            // Click on grid cell
            if let Some((data_row, data_col)) = screen_to_cell(app, x, y) {
                if app.mode == Mode::Insert {
                    app.confirm_edit();
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

    let screen_row = (y - grid_start_row) as u32;
    let data_row = app.viewport.top_row + screen_row;

    let screen_col_px = x - row_num_width;
    let data_col = app.viewport.left_col + screen_col_px / app.viewport.col_width;

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
