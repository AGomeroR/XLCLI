use xlcli_core::cell::Cell;

#[derive(Debug, Clone)]
pub struct ClipboardEntry {
    pub cells: Vec<(u32, u16, Cell)>,
    pub rows: u32,
    pub cols: u16,
}

pub struct Clipboard {
    pub content: Option<ClipboardEntry>,
}

impl Clipboard {
    pub fn new() -> Self {
        Self { content: None }
    }

    pub fn yank_single(&mut self, cell: Option<&Cell>) {
        let cell = cell.cloned().unwrap_or_default();
        self.content = Some(ClipboardEntry {
            cells: vec![(0, 0, cell)],
            rows: 1,
            cols: 1,
        });
    }

    pub fn yank_range(
        &mut self,
        cells: Vec<(u32, u16, Cell)>,
        rows: u32,
        cols: u16,
    ) {
        self.content = Some(ClipboardEntry { cells, rows, cols });
    }
}
