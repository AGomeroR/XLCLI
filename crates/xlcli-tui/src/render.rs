use ratatui::layout::{Constraint, Direction, Layout, Rect};
use ratatui::style::{Color, Modifier, Style};
use ratatui::text::{Line, Span};
use ratatui::widgets::{Block, Borders, Clear, Paragraph};
use ratatui::Frame;

use xlcli_core::types::CellAddr;

use crate::app::{App, CF_COLORS, CfCond, CfDialogFocus, CfDropdown, FilterDialogFocus, FilterDropdown, FilterType, SortDialogFocus, SortDropdown, SortType};
use crate::mode::Mode;

pub fn render(frame: &mut Frame, app: &mut App) {
    app.recompute_formula_error();
    let size = frame.area();
    app.viewport.update_dimensions(size.width, size.height);
    let (freeze_rows, freeze_cols) = app.workbook.active_sheet().freeze.unwrap_or((0, 0));
    app.viewport.update_for_cursor_with_freeze(&app.cursor, freeze_rows, freeze_cols);

    let chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Length(1), // formula bar
            Constraint::Length(1), // column headers
            Constraint::Min(1),   // grid
            Constraint::Length(1), // sheet tabs
            Constraint::Length(1), // status bar
        ])
        .split(size);

    render_formula_bar(frame, app, chunks[0]);
    render_column_headers(frame, app, chunks[1]);
    render_grid(frame, app, chunks[2]);
    render_sheet_tabs(frame, app, chunks[3]);
    render_status_bar(frame, app, chunks[4]);

    if app.mode == Mode::Command && app.config.command_palette.enabled {
        render_command_palette(frame, app, size);
    }

    if app.mode == Mode::Insert && app.autocomplete.visible && app.config.formula_autocomplete.enabled {
        render_autocomplete(frame, app, size);
    }

    if app.mode == Mode::Insert && app.config.formula_autocomplete.show_signature {
        if let Some(ref func_name) = app.autocomplete.active_function {
            render_signature_tooltip(frame, app, func_name, size);
        }
    }

    if app.sort_dialog.visible {
        render_sort_dialog(frame, app, size);
    }

    if app.filter_dialog.visible {
        render_filter_dialog(frame, app, size);
    }

    if app.cf_dialog.visible {
        render_cf_dialog(frame, app, size);
    }

    if app.cf_list.visible {
        render_cf_list(frame, app, size);
    }

    if app.mode == Mode::Insert {
        if let Some(err) = app.formula_error.clone() {
            render_formula_error(frame, app, size, &err);
        }
    }
}

fn render_formula_error(frame: &mut Frame, app: &App, size: Rect, err: &xlcli_formulas::ParseError) {
    if size.height < 3 || size.width < 10 {
        return;
    }
    // Formula bar layout: " {:>6} " (8) + " " (1) + edit_buffer. Formula body starts at col 9 + 1 ('=').
    let body = app.edit_buffer.strip_prefix('=').unwrap_or("");
    let off = err.offset.min(body.len());
    let char_col = body[..off].chars().count() as u16;
    let caret_col = 9u16.saturating_add(1).saturating_add(char_col);
    let caret_col = caret_col.min(size.width.saturating_sub(1));

    let area = Rect { x: 0, y: 1, width: size.width, height: 2 };
    frame.render_widget(Clear, area);

    let mut caret_line = String::new();
    for _ in 0..caret_col { caret_line.push(' '); }
    caret_line.push('^');
    let caret_para = Paragraph::new(Line::from(vec![
        Span::styled(caret_line, Style::default().fg(Color::Red).add_modifier(Modifier::BOLD)),
    ]));
    let msg_para = Paragraph::new(Line::from(vec![
        Span::styled(format!(" {}", err.message), Style::default().fg(Color::Red)),
    ]));
    let caret_area = Rect { x: 0, y: 1, width: size.width, height: 1 };
    let msg_area = Rect { x: 0, y: 2, width: size.width, height: 1 };
    frame.render_widget(caret_para, caret_area);
    frame.render_widget(msg_para, msg_area);
}

fn render_formula_bar(frame: &mut Frame, app: &App, area: Rect) {
    let cell_name = app.cursor.display_name();

    let cell_content = if app.mode == Mode::Insert {
        app.edit_buffer.clone()
    } else {
        app.workbook
            .active_sheet()
            .get_cell(app.cursor.row, app.cursor.col)
            .map(|c| {
                c.formula
                    .as_deref()
                    .map(|f| format!("={}", f))
                    .unwrap_or_else(|| c.value.display_value())
            })
            .unwrap_or_default()
    };

    let name_style = if app.mode == Mode::Insert {
        Style::default().fg(Color::Black).bg(Color::Blue)
    } else {
        Style::default().fg(Color::Black).bg(Color::Gray)
    };

    let line = if app.mode == Mode::Insert {
        let cur = app.edit_cursor.min(cell_content.len());
        let (before, after) = cell_content.split_at(cur);
        let (under, rest) = if after.is_empty() {
            (" ".to_string(), String::new())
        } else {
            let mut end = 1;
            while end < after.len() && !after.is_char_boundary(end) { end += 1; }
            (after[..end].to_string(), after[end..].to_string())
        };
        Line::from(vec![
            Span::styled(format!(" {:>6} ", cell_name), name_style),
            Span::raw(" "),
            Span::raw(before.to_string()),
            Span::styled(under, Style::default().fg(Color::Black).bg(Color::White)),
            Span::raw(rest),
        ])
    } else {
        Line::from(vec![
            Span::styled(format!(" {:>6} ", cell_name), name_style),
            Span::raw(" "),
            Span::raw(cell_content),
        ])
    };
    frame.render_widget(Paragraph::new(line), area);
}

fn render_column_headers(frame: &mut Frame, app: &App, area: Rect) {
    let vp = &app.viewport;
    let row_num_width = 6u16;
    let sheet = app.workbook.active_sheet();
    let (_freeze_rows, freeze_cols) = sheet.freeze.unwrap_or((0, 0));
    let has_filter = |col: u16| sheet.filters.contains_key(&col);

    let mut spans = vec![Span::styled(
        format!("{:>width$}", "", width = row_num_width as usize),
        Style::default().fg(Color::DarkGray),
    )];

    let col_header_span = |col: u16, col_width: u16| -> Span<'static> {
        let col_name = CellAddr::col_name(col);
        let filtered = has_filter(col);
        let label = if filtered {
            format!("{}\u{25BC}", col_name)
        } else {
            col_name
        };
        let in_visual = is_col_in_visual(app, col);
        let style = if filtered {
            Style::default().fg(Color::Cyan).add_modifier(Modifier::BOLD)
        } else if col == app.cursor.col || in_visual {
            Style::default().fg(Color::Yellow).add_modifier(Modifier::BOLD)
        } else {
            Style::default().fg(Color::DarkGray)
        };
        Span::styled(format!("{:^width$}", label, width = col_width as usize), style)
    };

    for col in 0..freeze_cols {
        spans.push(col_header_span(col, vp.col_width));
    }

    if freeze_cols > 0 {
        spans.push(Span::styled("\u{2502}", Style::default().fg(Color::Cyan)));
    }

    let scroll_cols = if freeze_cols > 0 {
        let used = freeze_cols as u16 * vp.col_width + 1;
        let remaining_px = area.width.saturating_sub(row_num_width + used);
        (remaining_px / vp.col_width).max(1)
    } else {
        vp.visible_cols
    };

    for i in 0..scroll_cols {
        let col = vp.left_col + i;
        spans.push(col_header_span(col, vp.col_width));
    }

    frame.render_widget(Paragraph::new(Line::from(spans)), area);
}

fn render_grid(frame: &mut Frame, app: &App, area: Rect) {
    let vp = &app.viewport;
    let sheet = app.workbook.active_sheet();
    let row_num_width = 6u16;
    let (freeze_rows, freeze_cols) = sheet.freeze.unwrap_or((0, 0));

    let has_freeze_col = freeze_cols > 0;
    let has_freeze_row = freeze_rows > 0;

    let scroll_cols = if has_freeze_col {
        let used = freeze_cols as u16 * vp.col_width + 1;
        let remaining_px = area.width.saturating_sub(row_num_width + used);
        (remaining_px / vp.col_width).max(1)
    } else {
        vp.visible_cols
    };

    let col_width = vp.col_width;
    let left_col = vp.left_col;
    let max_screen_rows = area.height as u32;
    let mut lines = Vec::new();
    let mut rows_rendered: u32 = 0;

    // Frozen rows
    for fr in 0..freeze_rows {
        if rows_rendered >= max_screen_rows { break; }
        let spans = build_grid_row(app, sheet, fr, freeze_cols, has_freeze_col, scroll_cols, col_width, left_col, row_num_width);
        lines.push(Line::from(spans));
        rows_rendered += 1;
    }

    // Freeze row border
    if has_freeze_row && rows_rendered < max_screen_rows {
        let spans = build_freeze_border_row(freeze_cols, has_freeze_col, scroll_cols, col_width, row_num_width);
        lines.push(Line::from(spans));
        rows_rendered += 1;
    }

    // Scrollable rows — skip hidden rows
    let mut data_row = vp.top_row;
    while rows_rendered < max_screen_rows {
        if sheet.hidden_rows.contains(&data_row) {
            data_row += 1;
            continue;
        }
        let spans = build_grid_row(app, sheet, data_row, freeze_cols, has_freeze_col, scroll_cols, col_width, left_col, row_num_width);
        lines.push(Line::from(spans));
        rows_rendered += 1;
        data_row += 1;
    }

    frame.render_widget(Paragraph::new(lines), area);
}

fn build_freeze_border_row(
    freeze_cols: u16,
    has_freeze_col: bool,
    scroll_cols: u16,
    col_width: u16,
    row_num_width: u16,
) -> Vec<Span<'static>> {
    let border_style = Style::default().fg(Color::Cyan);
    let h_line = "\u{2500}".repeat(col_width as usize);
    let mut spans = Vec::new();
    spans.push(Span::styled(
        "\u{2500}".repeat(row_num_width as usize),
        border_style,
    ));
    for _ in 0..freeze_cols {
        spans.push(Span::styled(h_line.clone(), border_style));
    }
    if has_freeze_col {
        spans.push(Span::styled("\u{253C}", border_style));
    }
    for _ in 0..scroll_cols {
        spans.push(Span::styled(h_line.clone(), border_style));
    }
    spans
}

fn build_grid_row<'a>(
    app: &'a App,
    sheet: &xlcli_core::sheet::Sheet,
    data_row: u32,
    freeze_cols: u16,
    has_freeze_col: bool,
    scroll_cols: u16,
    col_width: u16,
    left_col: u16,
    row_num_width: u16,
) -> Vec<Span<'a>> {
    let mut spans = Vec::new();

    let in_visual_row = is_row_in_visual(app, data_row);
    let row_style = if data_row == app.cursor.row || in_visual_row {
        Style::default().fg(Color::Yellow).add_modifier(Modifier::BOLD)
    } else {
        Style::default().fg(Color::DarkGray)
    };
    spans.push(Span::styled(
        format!("{:>width$}", data_row + 1, width = row_num_width as usize),
        row_style,
    ));

    for col in 0..freeze_cols {
        render_cell_span(app, sheet, data_row, col, col_width, &mut spans);
    }

    if has_freeze_col {
        spans.push(Span::styled("\u{2502}", Style::default().fg(Color::Cyan)));
    }

    for i in 0..scroll_cols {
        let data_col = left_col + i;
        render_cell_span(app, sheet, data_row, data_col, col_width, &mut spans);
    }

    spans
}

fn render_cell_span<'a>(
    app: &'a App,
    sheet: &xlcli_core::sheet::Sheet,
    data_row: u32,
    data_col: u16,
    col_width: u16,
    spans: &mut Vec<Span<'a>>,
) {
    let is_cursor = data_row == app.cursor.row && data_col == app.cursor.col;
    let in_visual = is_in_visual_selection(app, data_row, data_col);

    let cell_text = if is_cursor && app.mode == Mode::Insert {
        app.edit_buffer.clone()
    } else {
        sheet
            .get_cell(data_row, data_col)
            .map(|c| c.value.display_value())
            .unwrap_or_default()
    };

    let max_chars = col_width as usize;
    let char_count = cell_text.chars().count();
    let display = if char_count > max_chars - 1 {
        let truncated: String = cell_text.chars().take(max_chars - 2).collect();
        format!("{:<width$}", truncated, width = max_chars - 1)
    } else {
        format!("{:<width$}", cell_text, width = max_chars)
    };

    let is_header = app.is_header_row(data_row);

    let base_style = sheet.get_cell(data_row, data_col)
        .map(|c| app.workbook.style_pool.get(c.style_id).clone())
        .unwrap_or_default();
    let effective = sheet.effective_style(data_row, data_col, &base_style);

    let mut style = if is_cursor {
        if app.mode == Mode::Insert {
            Style::default().fg(Color::White).bg(Color::Blue)
        } else {
            Style::default().fg(Color::Black).bg(Color::White)
        }
    } else if in_visual {
        Style::default().fg(Color::Black).bg(Color::Magenta)
    } else if is_header {
        Style::default().fg(Color::White).bg(Color::DarkGray)
    } else {
        let mut s = Style::default().fg(Color::White);
        if let Some(c) = effective.fg_color {
            s = s.fg(Color::Rgb(c.r, c.g, c.b));
        }
        if let Some(c) = effective.bg_color {
            s = s.bg(Color::Rgb(c.r, c.g, c.b));
        }
        s
    };
    if is_header { style = style.add_modifier(Modifier::BOLD); }
    if effective.bold { style = style.add_modifier(Modifier::BOLD); }
    if effective.italic { style = style.add_modifier(Modifier::ITALIC); }
    if effective.underline || effective.double_underline { style = style.add_modifier(Modifier::UNDERLINED); }
    if effective.strikethrough { style = style.add_modifier(Modifier::CROSSED_OUT); }

    spans.push(Span::styled(display, style));
}

fn render_sheet_tabs(frame: &mut Frame, app: &App, area: Rect) {
    let mut spans = Vec::new();
    for (i, sheet) in app.workbook.sheets.iter().enumerate() {
        let style = if i == app.workbook.active_sheet {
            Style::default()
                .fg(Color::Black)
                .bg(Color::White)
                .add_modifier(Modifier::BOLD)
        } else {
            Style::default().fg(Color::DarkGray)
        };
        spans.push(Span::styled(format!(" {} ", sheet.name), style));
        spans.push(Span::raw(" "));
    }
    frame.render_widget(Paragraph::new(Line::from(spans)), area);
}

fn render_status_bar(frame: &mut Frame, app: &App, area: Rect) {
    let mode_style = match app.mode {
        Mode::Normal => Style::default().fg(Color::Black).bg(Color::Green),
        Mode::Insert => Style::default().fg(Color::Black).bg(Color::Blue),
        Mode::Visual => Style::default().fg(Color::Black).bg(Color::Magenta),
        Mode::Command => Style::default().fg(Color::Black).bg(Color::Yellow),
    };

    let cell_name = app.cursor.display_name();
    let sheet_name = &app.workbook.active_sheet().name;
    let file_name = app.workbook.file_path.as_deref().unwrap_or("[new]");
    let modified_marker = if app.modified { " [+]" } else { "" };

    let status = if app.mode == Mode::Command && !app.config.command_palette.enabled {
        format!(":{}", app.command_buffer)
    } else if let Some(msg) = &app.status_message {
        msg.clone()
    } else {
        String::new()
    };

    let undo_indicator = if app.undo_stack.can_undo() { "u" } else { "-" };
    let redo_indicator = if app.undo_stack.can_redo() { "r" } else { "-" };

    let line = Line::from(vec![
        Span::styled(format!(" {} ", app.mode.label()), mode_style),
        Span::raw(" "),
        Span::styled(cell_name, Style::default().fg(Color::Cyan)),
        Span::raw(" | "),
        Span::styled(sheet_name.as_str(), Style::default().fg(Color::DarkGray)),
        Span::raw(" | "),
        Span::styled(file_name, Style::default().fg(Color::DarkGray)),
        Span::styled(modified_marker, Style::default().fg(Color::Red)),
        Span::raw(" "),
        Span::styled(
            format!("[{}{}]", undo_indicator, redo_indicator),
            Style::default().fg(Color::DarkGray),
        ),
        Span::raw("  "),
        Span::raw(status),
    ]);

    frame.render_widget(Paragraph::new(line), area);
}

fn is_in_visual_selection(app: &App, row: u32, col: u16) -> bool {
    if app.mode != Mode::Visual {
        return false;
    }
    if let Some(anchor) = app.visual_anchor {
        let min_row = anchor.row.min(app.cursor.row);
        let max_row = anchor.row.max(app.cursor.row);
        let min_col = anchor.col.min(app.cursor.col);
        let max_col = anchor.col.max(app.cursor.col);
        row >= min_row && row <= max_row && col >= min_col && col <= max_col
    } else {
        false
    }
}

fn is_row_in_visual(app: &App, row: u32) -> bool {
    if app.mode != Mode::Visual {
        return false;
    }
    if let Some(anchor) = app.visual_anchor {
        let min_row = anchor.row.min(app.cursor.row);
        let max_row = anchor.row.max(app.cursor.row);
        row >= min_row && row <= max_row
    } else {
        false
    }
}

fn is_col_in_visual(app: &App, col: u16) -> bool {
    if app.mode != Mode::Visual {
        return false;
    }
    if let Some(anchor) = app.visual_anchor {
        let min_col = anchor.col.min(app.cursor.col);
        let max_col = anchor.col.max(app.cursor.col);
        col >= min_col && col <= max_col
    } else {
        false
    }
}

pub fn fuzzy_score(query: &str, target: &str) -> Option<i32> {
    if query.is_empty() { return Some(0); }
    let t = target.to_lowercase();
    if t.starts_with(query) { return Some(1000 - target.len() as i32); }
    let mut score = 0i32;
    let mut last = 0usize;
    let mut prev_match = false;
    for qc in query.chars() {
        match t[last..].find(qc) {
            Some(pos) => {
                let abs = last + pos;
                score += if prev_match { 20 } else { 5 };
                if abs == 0 { score += 10; }
                last = abs + qc.len_utf8();
                prev_match = true;
            }
            None => return None,
        }
    }
    Some(score - target.len() as i32)
}

pub const COMMAND_LIST: &[(&str, &str)] = &[
    ("w", "save file"),
    ("w ", "save as <path>"),
    ("wq", "save and quit"),
    ("q", "quit"),
    ("q!", "force quit"),
    ("c ", "jump to cell <addr>"),
    ("ir", "insert row above"),
    ("ir!", "insert row below"),
    ("dr", "delete row"),
    ("ic", "insert col left"),
    ("ic!", "insert col right"),
    ("dc", "delete col"),
    ("freeze ", "freeze rows/cols"),
    ("unfreeze", "unfreeze panes"),
    ("sort", "open sort dialog"),
    ("sort ", "sort <args>"),
    ("sheet ", "sheet add/del/rename"),
    ("headers", "set header row"),
    ("headers ", "set headers <range>"),
    ("unheaders", "remove header row"),
    ("filter", "open filter dialog"),
    ("filter all", "filter whole sheet"),
    ("filter ", "filter <args>"),
    ("unfilter", "clear filter"),
    ("unfilter all", "clear all filters"),
    ("unfilter ", "clear filter <col>"),
    ("names", "list named ranges"),
    ("name ", "name <Name> [range]"),
    ("cf", "open format dialog"),
    ("cf list", "list rules"),
    ("cf clean", "clean rules in selection"),
    ("cf clean all", "clean all rules"),
    ("cf base ", "set base style"),
    ("cf ", "cf <range> <cond> <style>"),
    ("case ", "case upper/lower/title"),
];

fn render_command_palette(frame: &mut Frame, app: &App, area: Rect) {
    use crate::config::PalettePosition;

    let cfg = &app.config.command_palette;
    let width_pct = cfg.width_percent.clamp(20, 90);
    let palette_width = ((area.width as u32 * width_pct as u32) / 100)
        .max(20)
        .min(area.width.saturating_sub(2) as u32) as u16;
    let palette_height: u16 = 3;

    let x = (area.width.saturating_sub(palette_width)) / 2;
    let y = match cfg.position {
        PalettePosition::Top => 1,
        PalettePosition::Center => (area.height.saturating_sub(palette_height)) / 2,
        PalettePosition::Bottom => area.height.saturating_sub(palette_height + 2),
    };

    let palette_area = Rect::new(x, y, palette_width, palette_height);

    frame.render_widget(Clear, palette_area);

    let input_text = format!(":{}", app.command_buffer);
    let block = Block::default()
        .borders(Borders::ALL)
        .border_style(Style::default().fg(Color::Yellow))
        .title(Span::styled(
            " Command ",
            Style::default()
                .fg(Color::Yellow)
                .add_modifier(Modifier::BOLD),
        ));

    let inner = block.inner(palette_area);
    frame.render_widget(block, palette_area);

    let line = Line::from(vec![
        Span::styled(&input_text, Style::default().fg(Color::White)),
        Span::styled("█", Style::default().fg(Color::Yellow)),
    ]);
    frame.render_widget(Paragraph::new(line), inner);

    if app.command_buffer.is_empty() {
        return;
    }

    let q = app.command_buffer.to_lowercase();
    let mut scored: Vec<(i32, (&str, &str))> = COMMAND_LIST
        .iter()
        .filter_map(|e| fuzzy_score(&q, e.0).map(|s| (s, *e)))
        .collect();
    scored.sort_by(|a, b| b.0.cmp(&a.0).then_with(|| a.1.0.len().cmp(&b.1.0.len())));
    let matches: Vec<(&str, &str)> = scored.into_iter().take(10).map(|(_, e)| e).collect();
    if matches.is_empty() {
        return;
    }

    let sug_h = matches.len() as u16 + 2;
    let sug_y = palette_area.y + palette_area.height;
    if sug_y + sug_h > area.height {
        return;
    }
    let sug_area = Rect::new(palette_area.x, sug_y, palette_width, sug_h);
    frame.render_widget(Clear, sug_area);
    let sug_block = Block::default()
        .borders(Borders::ALL)
        .border_style(Style::default().fg(Color::Cyan));
    let sug_inner = sug_block.inner(sug_area);
    frame.render_widget(sug_block, sug_area);

    let sel = app.cmd_palette_selected.min(matches.len().saturating_sub(1));
    let lines: Vec<Line> = matches
        .iter()
        .enumerate()
        .map(|(i, (c, d))| {
            let (cmd_st, desc_st) = if i == sel {
                (Style::default().fg(Color::Black).bg(Color::Yellow).add_modifier(Modifier::BOLD),
                 Style::default().fg(Color::Black).bg(Color::Yellow))
            } else {
                (Style::default().fg(Color::Yellow), Style::default().fg(Color::Gray))
            };
            Line::from(vec![
                Span::styled(format!(" :{:<14}", c), cmd_st),
                Span::styled(format!(" {}", d), desc_st),
            ])
        })
        .collect();
    frame.render_widget(Paragraph::new(lines), sug_inner);

    let (sel_cmd, sel_desc) = matches[sel];
    let desc_y = sug_area.y + sug_area.height;
    let desc_h: u16 = 4;
    if desc_y + desc_h > area.height {
        return;
    }
    let desc_area = Rect::new(palette_area.x, desc_y, palette_width, desc_h);
    frame.render_widget(Clear, desc_area);
    let desc_block = Block::default()
        .borders(Borders::ALL)
        .border_style(Style::default().fg(Color::Cyan))
        .title(Span::styled(
            format!(" :{} ", sel_cmd),
            Style::default().fg(Color::Yellow).add_modifier(Modifier::BOLD),
        ));
    let desc_inner = desc_block.inner(desc_area);
    frame.render_widget(desc_block, desc_area);
    frame.render_widget(
        Paragraph::new(Line::from(Span::styled(
            sel_desc.to_string(),
            Style::default().fg(Color::White).add_modifier(Modifier::ITALIC),
        )))
        .wrap(ratatui::widgets::Wrap { trim: true }),
        desc_inner,
    );
}

fn render_autocomplete(frame: &mut Frame, app: &App, area: Rect) {
    let ac = &app.autocomplete;
    if ac.matches.is_empty() {
        return;
    }

    let show_desc = app.config.formula_autocomplete.show_description;
    let max_show = 8.min(ac.matches.len());
    let name_width = ac.matches.iter().map(|s| s.len()).max().unwrap_or(10) + 4;

    let popup_width = (name_width as u16).clamp(12, 40);
    let popup_height = max_show as u16 + 2;

    let row_num_width: u16 = 6;
    let grid_start_row: u16 = 2;
    let (freeze_rows, freeze_cols) = app.workbook.active_sheet().freeze.unwrap_or((0, 0));
    let freeze_border_row = if freeze_rows > 0 { 1u16 } else { 0 };
    let freeze_border_col = if freeze_cols > 0 { 1u16 } else { 0 };

    let cursor_screen_row = if app.cursor.row < freeze_rows {
        grid_start_row + app.cursor.row as u16
    } else {
        grid_start_row + freeze_rows as u16 + freeze_border_row + (app.cursor.row - app.viewport.top_row) as u16
    };
    let cursor_screen_col = if app.cursor.col < freeze_cols {
        row_num_width + app.cursor.col * app.viewport.col_width
    } else {
        row_num_width + freeze_cols as u16 * app.viewport.col_width + freeze_border_col + (app.cursor.col - app.viewport.left_col) * app.viewport.col_width
    };

    let x = cursor_screen_col.min(area.width.saturating_sub(popup_width));
    let y = (cursor_screen_row + 1).min(area.height.saturating_sub(popup_height));

    let popup_area = Rect::new(x, y, popup_width, popup_height);
    frame.render_widget(Clear, popup_area);

    let block = Block::default()
        .borders(Borders::ALL)
        .border_style(Style::default().fg(Color::Cyan));

    let inner = block.inner(popup_area);
    frame.render_widget(block, popup_area);

    let scroll_offset = if ac.selected >= max_show {
        ac.selected - max_show + 1
    } else {
        0
    };

    let mut lines = Vec::new();
    for (i, name) in ac.matches.iter().enumerate().skip(scroll_offset).take(max_show) {
        let style = if i == ac.selected {
            Style::default().fg(Color::Black).bg(Color::Cyan)
        } else {
            Style::default().fg(Color::White)
        };
        lines.push(Line::from(Span::styled(
            format!(" {:<width$}", name, width = popup_width as usize - 3),
            style,
        )));
    }

    frame.render_widget(Paragraph::new(lines), inner);

    if show_desc && ac.selected < ac.matches.len() {
        let name = ac.matches[ac.selected];
        let desc = app.formula_registry.description(name);
        if !desc.is_empty() {
            let desc_right_x = popup_area.x + popup_area.width;
            let available = area.width.saturating_sub(desc_right_x);
            if available >= 10 {
                let desc_width = (desc.len() as u16 + 4).min(available);
                let desc_height = 3u16;
                let desc_y = popup_area.y;

                let desc_area = Rect::new(desc_right_x, desc_y, desc_width, desc_height);
                frame.render_widget(Clear, desc_area);

                let desc_block = Block::default()
                    .borders(Borders::ALL)
                    .border_style(Style::default().fg(Color::DarkGray));

                let desc_inner = desc_block.inner(desc_area);
                frame.render_widget(desc_block, desc_area);

                frame.render_widget(
                    Paragraph::new(Line::from(Span::styled(
                        format!(" {}", desc),
                        Style::default().fg(Color::DarkGray).add_modifier(Modifier::ITALIC),
                    ))),
                    desc_inner,
                );
            }
        }
    }
}

fn render_signature_tooltip(frame: &mut Frame, app: &App, func_name: &str, area: Rect) {
    let syntax = app.formula_registry.syntax(func_name);
    if syntax.is_empty() {
        return;
    }

    let text = format!(" {} ", syntax);
    let width = (text.len() as u16 + 2).min(area.width);
    if width < 6 {
        return;
    }

    let x = area.width.saturating_sub(width);
    let y = 1;

    let tooltip_area = Rect::new(x, y, width, 3);
    frame.render_widget(Clear, tooltip_area);

    let block = Block::default()
        .borders(Borders::ALL)
        .border_style(Style::default().fg(Color::Yellow));

    let inner = block.inner(tooltip_area);
    frame.render_widget(block, tooltip_area);

    frame.render_widget(
        Paragraph::new(Line::from(Span::styled(
            text,
            Style::default().fg(Color::White).bg(Color::Black).add_modifier(Modifier::BOLD),
        ))),
        inner,
    );
}

fn render_sort_dialog(frame: &mut Frame, app: &mut App, area: Rect) {
    use xlcli_core::types::CellAddr;

    let width: u16 = 42;
    let height: u16 = 11;

    let x = (area.width.saturating_sub(width)) / 2;
    let y = 1u16; // Same position as command palette (top)
    let dialog_area = Rect::new(x, y, width, height);

    app.sort_dialog.screen_x = x;
    app.sort_dialog.screen_y = y;
    app.sort_dialog.screen_w = width;
    app.sort_dialog.screen_h = height;

    let d = &app.sort_dialog;

    frame.render_widget(Clear, dialog_area);

    let block = Block::default()
        .borders(Borders::ALL)
        .border_style(Style::default().fg(Color::Yellow))
        .title(Span::styled(
            " Sort ",
            Style::default().fg(Color::Yellow).add_modifier(Modifier::BOLD),
        ));

    let inner = block.inner(dialog_area);
    frame.render_widget(block, dialog_area);

    let focus_style = |f: SortDialogFocus| -> Style {
        if d.focus == f {
            Style::default().fg(Color::Black).bg(Color::Cyan)
        } else {
            Style::default().fg(Color::White)
        }
    };

    let dropdown_indicator = |f: SortDialogFocus, dd: SortDropdown| -> &'static str {
        let matches = match (f, dd) {
            (SortDialogFocus::SortByCol, SortDropdown::SortByCol) => true,
            (SortDialogFocus::SortByType, SortDropdown::SortByType) => true,
            (SortDialogFocus::ThenByCol, SortDropdown::ThenByCol) => true,
            (SortDialogFocus::ThenByType, SortDropdown::ThenByType) => true,
            _ => false,
        };
        if matches && d.open_dropdown == dd { " \u{25b2}" } else { " \u{25bc}" }
    };

    let label = Style::default().fg(Color::Yellow);

    let sort_by_col_name = CellAddr::col_name(d.sort_by_col);
    let sort_by_type_label = d.sort_by_type.label();

    let then_by_col_name = d.then_by_col
        .map(|c| CellAddr::col_name(c))
        .unwrap_or_else(|| "(none)".to_string());
    let then_by_type_label = d.then_by_type.label();

    let headers_mark = if d.has_headers { "[x]" } else { "[ ]" };

    let lines = vec![
        Line::from(vec![
            Span::styled(" Range:    ", label),
            Span::styled(format!(" {} ", d.range_text), focus_style(SortDialogFocus::Range)),
        ]),
        Line::from(Span::raw("")),
        Line::from(vec![
            Span::styled(" Sort by:  ", label),
            Span::styled(
                format!(" {:<6}{}", sort_by_col_name, dropdown_indicator(SortDialogFocus::SortByCol, SortDropdown::SortByCol)),
                focus_style(SortDialogFocus::SortByCol),
            ),
            Span::raw(" "),
            Span::styled(
                format!(" {:<12}{}", sort_by_type_label, dropdown_indicator(SortDialogFocus::SortByType, SortDropdown::SortByType)),
                focus_style(SortDialogFocus::SortByType),
            ),
        ]),
        Line::from(vec![
            Span::styled(" Then by:  ", label),
            Span::styled(
                format!(" {:<6}{}", then_by_col_name, dropdown_indicator(SortDialogFocus::ThenByCol, SortDropdown::ThenByCol)),
                focus_style(SortDialogFocus::ThenByCol),
            ),
            Span::raw(" "),
            Span::styled(
                format!(" {:<12}{}", then_by_type_label, dropdown_indicator(SortDialogFocus::ThenByType, SortDropdown::ThenByType)),
                focus_style(SortDialogFocus::ThenByType),
            ),
        ]),
        Line::from(Span::raw("")),
        Line::from(vec![
            Span::styled(" Headers:  ", label),
            Span::styled(format!(" {} ", headers_mark), focus_style(SortDialogFocus::HasHeaders)),
        ]),
        Line::from(Span::raw("")),
        Line::from(vec![
            Span::raw("     "),
            Span::styled("  Sort  ", focus_style(SortDialogFocus::BtnSort)),
            Span::raw("    "),
            Span::styled(" Cancel ", focus_style(SortDialogFocus::BtnCancel)),
        ]),
    ];

    frame.render_widget(Paragraph::new(lines), inner);

    // Render open dropdown overlay
    let open_dd = d.open_dropdown;
    let dd_scroll = d.dropdown_scroll;
    let dd_min_col = d.min_col;
    let dd_max_col = d.max_col;
    let _ = d;

    if open_dd != SortDropdown::None {
        let (dd_x, dd_y, items) = match open_dd {
            SortDropdown::SortByCol => {
                let items: Vec<String> = (dd_min_col..=dd_max_col)
                    .map(|c| CellAddr::col_name(c))
                    .collect();
                (inner.x + 11, inner.y + 3, items)
            }
            SortDropdown::SortByType => {
                let items: Vec<String> = SortType::all().iter()
                    .map(|t| t.label().to_string())
                    .collect();
                (inner.x + 20, inner.y + 3, items)
            }
            SortDropdown::ThenByCol => {
                let mut items = vec!["(none)".to_string()];
                items.extend((dd_min_col..=dd_max_col).map(|c| CellAddr::col_name(c)));
                (inner.x + 11, inner.y + 4, items)
            }
            SortDropdown::ThenByType => {
                let items: Vec<String> = SortType::all().iter()
                    .map(|t| t.label().to_string())
                    .collect();
                (inner.x + 20, inner.y + 4, items)
            }
            SortDropdown::None => unreachable!(),
        };

        let max_show = items.len().min(8) as u16;
        let dd_width: u16 = 16;
        let dd_height = max_show + 2;

        let dd_area = Rect::new(
            dd_x.min(area.width.saturating_sub(dd_width)),
            dd_y.min(area.height.saturating_sub(dd_height)),
            dd_width,
            dd_height,
        );

        frame.render_widget(Clear, dd_area);

        let dd_block = Block::default()
            .borders(Borders::ALL)
            .border_style(Style::default().fg(Color::Cyan));

        let dd_inner = dd_block.inner(dd_area);
        frame.render_widget(dd_block, dd_area);

        // Store exact item positions for mouse hit-testing
        app.sort_dialog.dd_item_x = dd_inner.x;
        app.sort_dialog.dd_item_y = dd_inner.y;
        app.sort_dialog.dd_item_w = dd_inner.width;
        app.sort_dialog.dd_item_count = max_show;

        let mut dd_lines = Vec::new();
        for (i, item) in items.iter().enumerate().take(max_show as usize) {
            let style = if i == dd_scroll {
                Style::default().fg(Color::Black).bg(Color::Cyan)
            } else {
                Style::default().fg(Color::White)
            };
            dd_lines.push(Line::from(Span::styled(
                format!(" {:<width$}", item, width = dd_width as usize - 3),
                style,
            )));
        }

        frame.render_widget(Paragraph::new(dd_lines), dd_inner);
    }
}

fn render_filter_dialog(frame: &mut Frame, app: &mut App, area: Rect) {
    let is_eq = app.filter_dialog.filter_type == FilterType::Eq;
    let width: u16 = 44;
    let height: u16 = if is_eq { 20 } else { 11 };

    let x = (area.width.saturating_sub(width)) / 2;
    let y = 1u16;
    let dialog_area = Rect::new(x, y, width, height.min(area.height.saturating_sub(1)));

    app.filter_dialog.screen_x = x;
    app.filter_dialog.screen_y = y;
    app.filter_dialog.screen_w = width;
    app.filter_dialog.screen_h = height;

    let d = &app.filter_dialog;

    frame.render_widget(Clear, dialog_area);

    let block = Block::default()
        .borders(Borders::ALL)
        .border_style(Style::default().fg(Color::Green))
        .title(Span::styled(
            " Filter ",
            Style::default().fg(Color::Green).add_modifier(Modifier::BOLD),
        ));

    let inner = block.inner(dialog_area);
    frame.render_widget(block, dialog_area);

    let focus_style = |f: FilterDialogFocus| -> Style {
        if d.focus == f {
            Style::default().fg(Color::Black).bg(Color::Cyan)
        } else {
            Style::default().fg(Color::White)
        }
    };

    let dd_indicator = |f: FilterDialogFocus, dd: FilterDropdown| -> &'static str {
        let matches = match (f, dd) {
            (FilterDialogFocus::Column, FilterDropdown::Column) => true,
            (FilterDialogFocus::FilterType, FilterDropdown::FilterType) => true,
            _ => false,
        };
        if matches && d.open_dropdown == dd { " \u{25b2}" } else { " \u{25bc}" }
    };

    let label = Style::default().fg(Color::Green);
    let col_name = CellAddr::col_name(d.col);
    let type_label = d.filter_type.label();

    let mut lines = vec![
        Line::from(Span::raw("")),
        Line::from(vec![
            Span::styled(" Column:  ", label),
            Span::styled(
                format!(" {:<6}{}", col_name, dd_indicator(FilterDialogFocus::Column, FilterDropdown::Column)),
                focus_style(FilterDialogFocus::Column),
            ),
        ]),
        Line::from(vec![
            Span::styled(" Type:    ", label),
            Span::styled(
                format!(" {:<14}{}", type_label, dd_indicator(FilterDialogFocus::FilterType, FilterDropdown::FilterType)),
                focus_style(FilterDialogFocus::FilterType),
            ),
        ]),
    ];

    if is_eq {
        let search_display = format!(" \u{1f50d} {} ", d.value_search);
        let value_focused = d.focus == FilterDialogFocus::Value;
        let search_style = if value_focused {
            Style::default().fg(Color::Black).bg(Color::DarkGray)
        } else {
            Style::default().fg(Color::DarkGray)
        };
        lines.push(Line::from(vec![
            Span::styled(" Search:  ", label),
            Span::styled(format!("{:<28}", search_display), search_style),
        ]));
        lines.push(Line::from(Span::styled(
            " \u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}",
            Style::default().fg(Color::DarkGray),
        )));

        let max_visible: usize = (inner.height as usize).saturating_sub(8);
        let scroll = d.value_scroll;
        let selected = d.value_selected;

        let all_check = if d.all_checked { "\u{2611}" } else { "\u{2610}" };
        let all_style = if value_focused && selected == 0 {
            Style::default().fg(Color::Black).bg(Color::Green).add_modifier(Modifier::BOLD)
        } else {
            Style::default().fg(Color::Yellow).add_modifier(Modifier::BOLD)
        };
        lines.push(Line::from(Span::styled(
            format!(" {} (All)", all_check),
            all_style,
        )));

        let filtered = &d.filtered_values;
        for (vi, &real_idx) in filtered.iter().enumerate().skip(scroll).take(max_visible.saturating_sub(1)) {
            let display_idx = vi + 1;
            let (val, checked) = &d.all_values[real_idx];
            let check = if *checked { "\u{2611}" } else { "\u{2610}" };
            let truncated: String = val.chars().take(35).collect();
            let style = if value_focused && display_idx == selected {
                Style::default().fg(Color::Black).bg(Color::Green)
            } else if *checked {
                Style::default().fg(Color::White)
            } else {
                Style::default().fg(Color::DarkGray)
            };
            lines.push(Line::from(Span::styled(
                format!(" {} {}", check, truncated),
                style,
            )));
        }

        lines.push(Line::from(Span::raw("")));
    } else {
        let value_display = if d.filter_type.needs_value() {
            format!(" {} ", d.value_buf)
        } else {
            " (n/a) ".to_string()
        };
        lines.push(Line::from(vec![
            Span::styled(" Value:   ", label),
            Span::styled(
                format!("{:<20}", value_display),
                focus_style(FilterDialogFocus::Value),
            ),
            if d.focus == FilterDialogFocus::Value && d.filter_type.needs_value() {
                Span::styled("\u{2588}", Style::default().fg(Color::Green))
            } else {
                Span::raw("")
            },
        ]));
        lines.push(Line::from(Span::raw("")));
        lines.push(Line::from(Span::raw("")));
    }

    lines.push(Line::from(vec![
        Span::raw("     "),
        Span::styled("  Apply  ", focus_style(FilterDialogFocus::BtnApply)),
        Span::raw("    "),
        Span::styled(" Cancel ", focus_style(FilterDialogFocus::BtnCancel)),
    ]));

    frame.render_widget(Paragraph::new(lines), inner);

    // Render dropdown overlay
    let open_dd = d.open_dropdown;
    let dd_scroll = d.dropdown_scroll;
    let dd_min_col = d.min_col;
    let dd_max_col = d.max_col;
    let _ = d;

    if open_dd != FilterDropdown::None {
        let (dd_x, dd_y, items) = match open_dd {
            FilterDropdown::Column => {
                let items: Vec<String> = (dd_min_col..=dd_max_col)
                    .map(|c| CellAddr::col_name(c))
                    .collect();
                (inner.x + 11, inner.y + 2, items)
            }
            FilterDropdown::FilterType => {
                let items: Vec<String> = FilterType::all().iter()
                    .map(|t| t.label().to_string())
                    .collect();
                (inner.x + 11, inner.y + 3, items)
            }
            FilterDropdown::None => unreachable!(),
        };

        let max_show = items.len().min(8) as u16;
        let dd_width: u16 = 18;
        let dd_height = max_show + 2;

        let dd_area = Rect::new(
            dd_x.min(area.width.saturating_sub(dd_width)),
            dd_y.min(area.height.saturating_sub(dd_height)),
            dd_width,
            dd_height,
        );

        frame.render_widget(Clear, dd_area);

        let dd_block = Block::default()
            .borders(Borders::ALL)
            .border_style(Style::default().fg(Color::Cyan));

        let dd_inner = dd_block.inner(dd_area);
        frame.render_widget(dd_block, dd_area);

        app.filter_dialog.dd_item_x = dd_inner.x;
        app.filter_dialog.dd_item_y = dd_inner.y;
        app.filter_dialog.dd_item_w = dd_inner.width;
        app.filter_dialog.dd_item_count = max_show;

        let mut dd_lines = Vec::new();
        for (i, item) in items.iter().enumerate().take(max_show as usize) {
            let style = if i == dd_scroll {
                Style::default().fg(Color::Black).bg(Color::Cyan)
            } else {
                Style::default().fg(Color::White)
            };
            dd_lines.push(Line::from(Span::styled(
                format!(" {:<width$}", item, width = dd_width as usize - 3),
                style,
            )));
        }

        frame.render_widget(Paragraph::new(dd_lines), dd_inner);
    }
}

fn render_cf_dialog(frame: &mut Frame, app: &mut App, area: Rect) {
    let width: u16 = 44;
    // Content rows: 1 blank + 1 Range + 1 blank + 1 Styles label + 6 styles + 1 blank
    // + 1 BG + 1 FG + 1 blank + 1 Conditional + (2 if conditional) + 1 blank + 4 buttons
    let base_rows: u16 = 1 + 1 + 1 + 1 + 5 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 1 + 4;
    let extra = if app.cf_dialog.conditional { 2 } else { 0 };
    let height: u16 = base_rows + extra + 2; // +2 borders

    let x = (area.width.saturating_sub(width)) / 2;
    let y = 1u16;
    let dialog_area = Rect::new(
        x,
        y,
        width.min(area.width),
        height.min(area.height.saturating_sub(1)),
    );

    app.cf_dialog.screen_x = x;
    app.cf_dialog.screen_y = y;
    app.cf_dialog.screen_w = dialog_area.width;
    app.cf_dialog.screen_h = dialog_area.height;

    frame.render_widget(Clear, dialog_area);

    let block = Block::default()
        .borders(Borders::ALL)
        .border_style(Style::default().fg(Color::Reset))
        .title(Span::styled(
            " Conditional Formatting ",
            Style::default().fg(Color::Reset).add_modifier(Modifier::BOLD),
        ));

    let inner = block.inner(dialog_area);
    frame.render_widget(block, dialog_area);

    let d_focus = app.cf_dialog.focus.clone();
    let focus_style = |f: CfDialogFocus| -> Style {
        if d_focus == f {
            Style::default().fg(Color::Black).bg(Color::Cyan)
        } else {
            Style::default().fg(Color::White)
        }
    };
    let label = Style::default().fg(Color::Cyan);

    let d = &app.cf_dialog;
    let conditional = d.conditional;

    let checkbox = |b: bool| -> &'static str { if b { "[x]" } else { "[ ]" } };

    let cond_str = d.cond.label();
    let bg_name = CF_COLORS[d.bg_idx];
    let fg_name = CF_COLORS[d.fg_idx];

    let cond_toggle = format!("{} Conditional", checkbox(conditional));
    let range_field = format!(" {:<30}", format!("[{}]", d.range_text));
    let cond_field = format!(" [{:<12}\u{25bc}]", cond_str);
    let val1_field = format!(" [{:<8}]", d.val1);
    let val2_field = format!(" [{:<8}]", d.val2);
    let bg_field = format!(" [{:<8}\u{25bc}]", bg_name);
    let fg_field = format!(" [{:<8}\u{25bc}]", fg_name);
    let bg_hex_field = format!(" [{:<9}]", if d.bg_hex.is_empty() { "#" } else { d.bg_hex.as_str() });
    let fg_hex_field = format!(" [{:<9}]", if d.fg_hex.is_empty() { "#" } else { d.fg_hex.as_str() });

    let mut lines: Vec<Line> = Vec::with_capacity(inner.height as usize);
    let mut yrow: u16 = 0;

    lines.push(Line::from(Span::raw("")));
    yrow += 1;

    // Range row (top)
    let range_label = " Range:  ";
    lines.push(Line::from(vec![
        Span::styled(range_label, label),
        Span::styled(range_field.clone(), focus_style(CfDialogFocus::Range)),
    ]));
    let range_y = yrow;
    yrow += 1;

    lines.push(Line::from(Span::raw("")));
    yrow += 1;

    // Styles label
    lines.push(Line::from(Span::styled(" Styles:", label)));
    yrow += 1;

    let bold_s = format!("{} Bold", checkbox(d.bold));
    let italic_s = format!("{} Italic", checkbox(d.italic));
    let under_s = format!("{} Underline", checkbox(d.under));
    let dunder_s = format!("{} Double Underline", checkbox(d.dunder));
    let strike_s = format!("{} Strikethrough", checkbox(d.strike));

    let push_style_row = |lines: &mut Vec<Line>, text: String, f: CfDialogFocus| {
        lines.push(Line::from(vec![
            Span::raw("  "),
            Span::styled(text, focus_style(f)),
        ]));
    };

    push_style_row(&mut lines, bold_s.clone(), CfDialogFocus::Bold);
    let bold_y = yrow; yrow += 1;
    push_style_row(&mut lines, italic_s.clone(), CfDialogFocus::Italic);
    let italic_y = yrow; yrow += 1;
    push_style_row(&mut lines, under_s.clone(), CfDialogFocus::Under);
    let under_y = yrow; yrow += 1;
    push_style_row(&mut lines, dunder_s.clone(), CfDialogFocus::DUnder);
    let dunder_y = yrow; yrow += 1;
    push_style_row(&mut lines, strike_s.clone(), CfDialogFocus::Strike);
    let strike_y = yrow; yrow += 1;

    lines.push(Line::from(Span::raw("")));
    yrow += 1;

    lines.push(Line::from(vec![
        Span::styled(" BG:     ", label),
        Span::styled(bg_field.clone(), focus_style(CfDialogFocus::Bg)),
    ]));
    let bg_y = yrow;
    yrow += 1;

    lines.push(Line::from(vec![
        Span::styled(" Hex:    ", label),
        Span::styled(bg_hex_field.clone(), focus_style(CfDialogFocus::BgHex)),
    ]));
    let bg_hex_y = yrow;
    yrow += 1;

    lines.push(Line::from(vec![
        Span::styled(" FG:     ", label),
        Span::styled(fg_field.clone(), focus_style(CfDialogFocus::Fg)),
    ]));
    let fg_y = yrow;
    yrow += 1;

    lines.push(Line::from(vec![
        Span::styled(" Hex:    ", label),
        Span::styled(fg_hex_field.clone(), focus_style(CfDialogFocus::FgHex)),
    ]));
    let fg_hex_y = yrow;
    yrow += 1;

    // Preview swatch row
    {
        let resolve = |hex: &str, idx: usize| -> Option<xlcli_core::style::Color> {
            if let Some(c) = crate::app::parse_hex_color(hex) { return Some(c); }
            if idx > 0 {
                if let Ok(Some(c)) = crate::app::color_by_name_pub(CF_COLORS[idx]) { return Some(c); }
            }
            None
        };
        let bg = resolve(&d.bg_hex, d.bg_idx);
        let fg = resolve(&d.fg_hex, d.fg_idx);
        let mut sw = Style::default();
        if let Some(c) = fg { sw = sw.fg(Color::Rgb(c.r, c.g, c.b)); }
        if let Some(c) = bg { sw = sw.bg(Color::Rgb(c.r, c.g, c.b)); }
        if d.bold { sw = sw.add_modifier(Modifier::BOLD); }
        if d.italic { sw = sw.add_modifier(Modifier::ITALIC); }
        if d.under { sw = sw.add_modifier(Modifier::UNDERLINED); }
        if d.strike { sw = sw.add_modifier(Modifier::CROSSED_OUT); }
        lines.push(Line::from(vec![
            Span::styled(" Preview:", label),
            Span::raw(" "),
            Span::styled("\u{25CF} Sample", sw),
        ]));
        yrow += 1;
    }

    lines.push(Line::from(Span::raw("")));
    yrow += 1;

    // Conditional toggle (after styles)
    lines.push(Line::from(vec![
        Span::raw("  "),
        Span::styled(cond_toggle.clone(), focus_style(CfDialogFocus::Conditional)),
    ]));
    let conditional_y = yrow;
    yrow += 1;

    let mut cond_y: u16 = 0;
    let mut val_y: u16 = 0;
    let val1_label = " Val1:  ";
    let val2_sep = "   Val2:";
    let cond_label = " Cond:   ";

    if conditional {
        lines.push(Line::from(vec![
            Span::styled(cond_label, label),
            Span::styled(cond_field.clone(), focus_style(CfDialogFocus::Cond)),
        ]));
        cond_y = yrow;
        yrow += 1;
        lines.push(Line::from(vec![
            Span::styled(val1_label, label),
            Span::styled(val1_field.clone(), focus_style(CfDialogFocus::Val1)),
            Span::styled(val2_sep, label),
            Span::styled(val2_field.clone(), focus_style(CfDialogFocus::Val2)),
        ]));
        val_y = yrow;
        yrow += 1;
    }

    lines.push(Line::from(Span::raw("")));
    yrow += 1;

    lines.push(Line::from(vec![
        Span::raw("  "),
        Span::styled(" SetBase ", focus_style(CfDialogFocus::BtnBase)),
    ]));
    let base_btn_y = yrow; yrow += 1;
    lines.push(Line::from(vec![
        Span::raw("  "),
        Span::styled(" Dismiss ", focus_style(CfDialogFocus::BtnDismiss)),
    ]));
    let dismiss_btn_y = yrow; yrow += 1;
    lines.push(Line::from(vec![
        Span::raw("  "),
        Span::styled(" Apply   ", focus_style(CfDialogFocus::BtnApply)),
    ]));
    let apply_btn_y = yrow; yrow += 1;
    lines.push(Line::from(vec![
        Span::raw("  "),
        Span::styled(" Apply & Close ", focus_style(CfDialogFocus::BtnClose)),
    ]));
    let close_btn_y = yrow;

    frame.render_widget(Paragraph::new(lines), inner);

    let ix = inner.x;
    let iy = inner.y;
    let d = &mut app.cf_dialog;
    d.rect_conditional = (ix + 2, iy + conditional_y, cond_toggle.chars().count() as u16);
    d.rect_range = (ix + range_label.chars().count() as u16, iy + range_y, range_field.chars().count() as u16);
    if conditional {
        d.rect_cond = (ix + cond_label.chars().count() as u16, iy + cond_y, cond_field.chars().count() as u16);
        d.rect_val1 = (ix + val1_label.chars().count() as u16, iy + val_y, val1_field.chars().count() as u16);
        let v2_start = val1_label.chars().count() as u16 + val1_field.chars().count() as u16 + val2_sep.chars().count() as u16;
        d.rect_val2 = (ix + v2_start, iy + val_y, val2_field.chars().count() as u16);
    } else {
        d.rect_cond = (0, 0, 0);
        d.rect_val1 = (0, 0, 0);
        d.rect_val2 = (0, 0, 0);
    }
    let col_x = ix + 2;
    d.rect_bold   = (col_x, iy + bold_y, bold_s.chars().count() as u16);
    d.rect_italic = (col_x, iy + italic_y, italic_s.chars().count() as u16);
    d.rect_under  = (col_x, iy + under_y, under_s.chars().count() as u16);
    d.rect_dunder = (col_x, iy + dunder_y, dunder_s.chars().count() as u16);
    d.rect_strike = (col_x, iy + strike_y, strike_s.chars().count() as u16);
    d.rect_bg = (ix + 9, iy + bg_y, bg_field.chars().count() as u16);
    d.rect_bg_hex = (ix + 9, iy + bg_hex_y, bg_hex_field.chars().count() as u16);
    d.rect_fg = (ix + 9, iy + fg_y, fg_field.chars().count() as u16);
    d.rect_fg_hex = (ix + 9, iy + fg_hex_y, fg_hex_field.chars().count() as u16);
    d.rect_rules = (0, 0, 0, 0);
    d.rect_base     = (ix + 2, iy + base_btn_y, 9);
    d.rect_dismiss  = (ix + 2, iy + dismiss_btn_y, 9);
    d.rect_apply    = (ix + 2, iy + apply_btn_y, 9);
    d.rect_close    = (ix + 2, iy + close_btn_y, 9);
    d.rect_delete   = (0, 0, 0);
    d.rect_cleanall = (0, 0, 0);

    let open_dd = d.open_dropdown;
    let dd_scroll = d.dropdown_scroll;
    let _ = d;

    if open_dd != CfDropdown::None {
        let (dd_x, dd_y, items) = match open_dd {
            CfDropdown::Cond => {
                let items: Vec<String> = CfCond::all().iter().map(|c| c.label().to_string()).collect();
                (inner.x + 10, inner.y + cond_y + 1, items)
            }
            CfDropdown::Bg => {
                let items: Vec<String> = CF_COLORS.iter().map(|s| s.to_string()).collect();
                (inner.x + 10, inner.y + bg_y + 1, items)
            }
            CfDropdown::Fg => {
                let items: Vec<String> = CF_COLORS.iter().map(|s| s.to_string()).collect();
                (inner.x + 10, inner.y + fg_y + 1, items)
            }
            CfDropdown::None => unreachable!(),
        };

        let max_show = items.len().min(10) as u16;
        let dd_width: u16 = 18;
        let dd_height = max_show + 2;

        let dd_area = Rect::new(
            dd_x.min(area.width.saturating_sub(dd_width)),
            dd_y.min(area.height.saturating_sub(dd_height)),
            dd_width,
            dd_height,
        );

        frame.render_widget(Clear, dd_area);

        let dd_block = Block::default()
            .borders(Borders::ALL)
            .border_style(Style::default().fg(Color::Cyan));

        let dd_inner = dd_block.inner(dd_area);
        frame.render_widget(dd_block, dd_area);

        app.cf_dialog.dd_item_x = dd_inner.x;
        app.cf_dialog.dd_item_y = dd_inner.y;
        app.cf_dialog.dd_item_w = dd_inner.width;
        app.cf_dialog.dd_item_count = max_show;

        let mut dd_lines = Vec::new();
        for (i, item) in items.iter().enumerate().take(max_show as usize) {
            let style = if i == dd_scroll {
                Style::default().fg(Color::Black).bg(Color::Cyan)
            } else {
                Style::default().fg(Color::White)
            };
            dd_lines.push(Line::from(Span::styled(
                format!(" {:<width$}", item, width = dd_width as usize - 3),
                style,
            )));
        }
        frame.render_widget(Paragraph::new(dd_lines), dd_inner);
    }
}

fn cf_range_str(r: &xlcli_core::range::CellRange) -> String {
    format!("{}{}:{}{}",
        CellAddr::col_name(r.start.col), r.start.row + 1,
        CellAddr::col_name(r.end.col), r.end.row + 1)
}

fn cf_cond_str(c: &xlcli_core::condfmt::Condition) -> String {
    use xlcli_core::condfmt::Condition::*;
    match c {
        Always => "always".into(),
        Gt(n) => format!("gt {}", n),
        Lt(n) => format!("lt {}", n),
        Gte(n) => format!("gte {}", n),
        Lte(n) => format!("lte {}", n),
        Eq(n) => format!("eq {}", n),
        Neq(n) => format!("neq {}", n),
        Between(a, b) => format!("between {} {}", a, b),
        NotBetween(a, b) => format!("notbetween {} {}", a, b),
        Contains(s) => format!("contains \"{}\"", s),
        NotContains(s) => format!("notcontains \"{}\"", s),
        BeginsWith(s) => format!("begins \"{}\"", s),
        EndsWith(s) => format!("ends \"{}\"", s),
        Blanks => "blanks".into(),
        NonBlanks => "nonblanks".into(),
        ContainsErrors => "errors".into(),
        NotContainsErrors => "no-errors".into(),
        DuplicateValues => "duplicates".into(),
        UniqueValues => "unique".into(),
        Top { count, percent, bottom } => format!("{}{} {}", if *bottom {"bottom"} else {"top"}, if *percent {"%"} else {""}, count),
        Average { above, .. } => format!("avg {}", if *above {"above"} else {"below"}),
        TimePeriod(p) => format!("time {:?}", p),
        Expression(s) => format!("expr {}", s),
    }
}

fn cf_style_str(s: &xlcli_core::condfmt::StyleSpec) -> String {
    use xlcli_core::condfmt::StyleSpec::*;
    match s {
        Overlay(o) => cf_overlay_str(o),
        ColorScale(stops) => format!("color-scale ({} stops)", stops.len()),
        DataBar { .. } => "data-bar".into(),
        IconSet { kind, .. } => format!("icons {:?}", kind),
    }
}

fn cf_overlay_str(o: &xlcli_core::condfmt::StyleOverlay) -> String {
    let mut p = Vec::new();
    if o.bold == Some(true) { p.push("bold"); }
    if o.italic == Some(true) { p.push("italic"); }
    if o.underline == Some(true) { p.push("under"); }
    if o.double_underline == Some(true) { p.push("dunder"); }
    if o.strikethrough == Some(true) { p.push("strike"); }
    let mut s = p.join(" ");
    if let Some(c) = o.bg_color {
        if !s.is_empty() { s.push(' '); }
        s.push_str(&format!("bg={}", color_name_or_hex(c)));
    }
    if let Some(c) = o.fg_color {
        if !s.is_empty() { s.push(' '); }
        s.push_str(&format!("fg={}", color_name_or_hex(c)));
    }
    s
}

fn color_name_or_hex(c: Option<xlcli_core::style::Color>) -> String {
    match c {
        None => "none".into(),
        Some(c) => format!("#{:02X}{:02X}{:02X}", c.r, c.g, c.b),
    }
}

fn render_cf_list(frame: &mut Frame, app: &mut App, area: Rect) {
    let rules = &app.workbook.active_sheet().cond_rules;
    let width: u16 = 70.min(area.width.saturating_sub(4));
    let inner_h = (rules.len().max(1) as u16).min(12);
    let height: u16 = inner_h + 4; // borders + header + footer hint
    let x = (area.width.saturating_sub(width)) / 2;
    let y = (area.height.saturating_sub(height)) / 3;
    let dialog_area = Rect::new(x, y, width.min(area.width), height.min(area.height.saturating_sub(1)));

    frame.render_widget(Clear, dialog_area);
    let block = Block::default()
        .borders(Borders::ALL)
        .border_style(Style::default().fg(Color::Reset))
        .title(Span::styled(" Conditional Rules ",
            Style::default().fg(Color::Reset).add_modifier(Modifier::BOLD)));
    let inner = block.inner(dialog_area);
    frame.render_widget(block, dialog_area);

    app.cf_list.rect = (inner.x, inner.y, inner.width, inner.height);

    let sel = app.cf_list.selected;
    let scroll = app.cf_list.scroll;
    let mut lines: Vec<Line> = Vec::new();

    if rules.is_empty() {
        lines.push(Line::from(Span::styled("  (no rules)", Style::default().fg(Color::DarkGray))));
    } else {
        for (i, r) in rules.iter().enumerate().skip(scroll).take(inner_h as usize) {
            let s = format!(" #{:<2} {:<14} {:<16} {}",
                i, cf_range_str(&r.range), cf_cond_str(&r.cond), cf_style_str(&r.style));
            let style = if i == sel {
                Style::default().fg(Color::Black).bg(Color::Cyan)
            } else {
                Style::default().fg(Color::White)
            };
            lines.push(Line::from(Span::styled(s, style)));
        }
    }
    lines.push(Line::from(Span::raw("")));
    lines.push(Line::from(Span::styled(
        " j/k nav  Enter edit  dd delete  a new  Esc close",
        Style::default().fg(Color::DarkGray))));

    frame.render_widget(Paragraph::new(lines), inner);
}
