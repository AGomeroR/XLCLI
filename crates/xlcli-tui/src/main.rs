mod app;
mod clipboard;
mod config;
mod input;
mod mode;
mod render;
mod undo;
mod viewport;

use std::io;
use std::time::Duration;

use anyhow::Result;
use clap::Parser;
use crossterm::event::{
    self, DisableMouseCapture, EnableMouseCapture,
    KeyboardEnhancementFlags, PopKeyboardEnhancementFlags, PushKeyboardEnhancementFlags,
};
use crossterm::execute;
use crossterm::terminal::{
    disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen,
    supports_keyboard_enhancement,
};
use ratatui::backend::CrosstermBackend;
use ratatui::Terminal;

use xlcli_core::workbook::Workbook;

use crate::app::App;
use crate::config::Config;

#[derive(Parser)]
#[command(name = "xlcli", version, about = "Terminal spreadsheet editor")]
struct Cli {
    /// File to open (xlsx, csv, tsv, ods)
    file: Option<String>,
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    Config::ensure_default_exists();
    let config = Config::load();

    let workbook = if let Some(path) = &cli.file {
        xlcli_io::reader::read_file(std::path::Path::new(path))?
    } else {
        Workbook::new()
    };

    let mut app = App::new(workbook, config);

    enable_raw_mode()?;
    let mut stdout = io::stdout();
    execute!(stdout, EnterAlternateScreen, EnableMouseCapture)?;

    let keyboard_enhanced = matches!(supports_keyboard_enhancement(), Ok(true));
    if keyboard_enhanced {
        execute!(
            stdout,
            PushKeyboardEnhancementFlags(
                KeyboardEnhancementFlags::DISAMBIGUATE_ESCAPE_CODES
            )
        )?;
    }

    let backend = CrosstermBackend::new(stdout);
    let mut terminal = Terminal::new(backend)?;

    let result = run_loop(&mut terminal, &mut app);

    if keyboard_enhanced {
        execute!(terminal.backend_mut(), PopKeyboardEnhancementFlags)?;
    }
    disable_raw_mode()?;
    execute!(terminal.backend_mut(), LeaveAlternateScreen, DisableMouseCapture)?;

    result
}

fn run_loop(terminal: &mut Terminal<CrosstermBackend<io::Stdout>>, app: &mut App) -> Result<()> {
    loop {
        terminal.draw(|f| render::render(f, app))?;

        if event::poll(Duration::from_millis(16))? {
            let evt = event::read()?;
            input::handle_event(app, evt);
        }

        if app.should_quit {
            return Ok(());
        }
    }
}
