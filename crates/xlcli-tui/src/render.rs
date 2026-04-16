use ratatui::layout::{Constraint, Direction, Layout, Rect};
use ratatui::style::{Color, Modifier, Style};
use ratatui::text::{Line, Span};
use ratatui::widgets::{Block, Borders, Clear, Paragraph};
use ratatui::Frame;

use xlcli_core::types::CellAddr;

use crate::app::App;
use crate::mode::Mode;

pub fn render(frame: &mut Frame, app: &mut App) {
    let size = frame.area();
    app.viewport.update_dimensions(size.width, size.height);
    app.viewport.update_for_cursor(&app.cursor);

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

    let line = Line::from(vec![
        Span::styled(format!(" {:>6} ", cell_name), name_style),
        Span::raw(" "),
        Span::raw(cell_content),
        if app.mode == Mode::Insert {
            Span::styled("█", Style::default().fg(Color::White))
        } else {
            Span::raw("")
        },
    ]);
    frame.render_widget(Paragraph::new(line), area);
}

fn render_column_headers(frame: &mut Frame, app: &App, area: Rect) {
    let vp = &app.viewport;
    let row_num_width = 6u16;

    let mut spans = vec![Span::styled(
        format!("{:>width$}", "", width = row_num_width as usize),
        Style::default().fg(Color::DarkGray),
    )];

    for i in 0..vp.visible_cols {
        let col = vp.left_col + i;
        let col_name = CellAddr::col_name(col);
        let in_visual = is_col_in_visual(app, col);
        let style = if col == app.cursor.col || in_visual {
            Style::default()
                .fg(Color::Yellow)
                .add_modifier(Modifier::BOLD)
        } else {
            Style::default().fg(Color::DarkGray)
        };
        spans.push(Span::styled(
            format!("{:^width$}", col_name, width = vp.col_width as usize),
            style,
        ));
    }

    frame.render_widget(Paragraph::new(Line::from(spans)), area);
}

fn render_grid(frame: &mut Frame, app: &App, area: Rect) {
    let vp = &app.viewport;
    let sheet = app.workbook.active_sheet();
    let row_num_width = 6u16;

    let mut lines = Vec::new();

    for screen_row in 0..vp.visible_rows.min(area.height as u32) {
        let data_row = vp.top_row + screen_row;
        let mut spans = Vec::new();

        let in_visual_row = is_row_in_visual(app, data_row);
        let row_style = if data_row == app.cursor.row || in_visual_row {
            Style::default()
                .fg(Color::Yellow)
                .add_modifier(Modifier::BOLD)
        } else {
            Style::default().fg(Color::DarkGray)
        };
        spans.push(Span::styled(
            format!("{:>width$}", data_row + 1, width = row_num_width as usize),
            row_style,
        ));

        for screen_col in 0..vp.visible_cols {
            let data_col = vp.left_col + screen_col;
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

            let display = if cell_text.len() > vp.col_width as usize - 1 {
                format!(
                    "{:<width$}",
                    &cell_text[..vp.col_width as usize - 2],
                    width = vp.col_width as usize - 1
                )
            } else {
                format!("{:<width$}", cell_text, width = vp.col_width as usize)
            };

            let style = if is_cursor {
                if app.mode == Mode::Insert {
                    Style::default().fg(Color::White).bg(Color::Blue)
                } else {
                    Style::default().fg(Color::Black).bg(Color::White)
                }
            } else if in_visual {
                Style::default().fg(Color::Black).bg(Color::Magenta)
            } else {
                Style::default().fg(Color::White)
            };

            spans.push(Span::styled(display, style));
        }

        lines.push(Line::from(spans));
    }

    frame.render_widget(Paragraph::new(lines), area);
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

    let line = Line::from(vec![
        Span::styled(format!(" {} ", app.mode.label()), mode_style),
        Span::raw(" "),
        Span::styled(cell_name, Style::default().fg(Color::Cyan)),
        Span::raw(" | "),
        Span::styled(sheet_name.as_str(), Style::default().fg(Color::DarkGray)),
        Span::raw(" | "),
        Span::styled(file_name, Style::default().fg(Color::DarkGray)),
        Span::styled(modified_marker, Style::default().fg(Color::Red)),
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
}
