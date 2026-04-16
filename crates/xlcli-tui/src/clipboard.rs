use xlcli_core::cell::Cell;

#[derive(Debug, Clone)]
pub struct ClipboardEntry {
    pub cells: Vec<(u32, u16, Cell)>,
    pub rows: u32,
    pub cols: u16,
}

pub struct Clipboard {
    pub content: Option<ClipboardEntry>,
    pub relative: bool,
    pub origin_row: u32,
    pub origin_col: u16,
}

impl Clipboard {
    pub fn new() -> Self {
        Self { content: None, relative: false, origin_row: 0, origin_col: 0 }
    }

    pub fn yank_single(&mut self, cell: Option<&Cell>, origin_row: u32, origin_col: u16) {
        let cell = cell.cloned().unwrap_or_default();
        self.content = Some(ClipboardEntry {
            cells: vec![(0, 0, cell)],
            rows: 1,
            cols: 1,
        });
        self.relative = false;
        self.origin_row = origin_row;
        self.origin_col = origin_col;
    }

    pub fn yank_range(
        &mut self,
        cells: Vec<(u32, u16, Cell)>,
        rows: u32,
        cols: u16,
        origin_row: u32,
        origin_col: u16,
    ) {
        self.content = Some(ClipboardEntry { cells, rows, cols });
        self.relative = false;
        self.origin_row = origin_row;
        self.origin_col = origin_col;
    }

    pub fn set_relative(&mut self, relative: bool) {
        self.relative = relative;
    }
}
