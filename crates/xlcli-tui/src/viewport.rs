use xlcli_core::types::CellAddr;

pub struct Viewport {
    pub top_row: u32,
    pub left_col: u16,
    pub visible_rows: u32,
    pub visible_cols: u16,
    pub col_width: u16,
}

impl Viewport {
    pub fn new() -> Self {
        Self {
            top_row: 0,
            left_col: 0,
            visible_rows: 20,
            visible_cols: 10,
            col_width: 12,
        }
    }

    pub fn update_for_cursor(&mut self, cursor: &CellAddr) {
        if cursor.row < self.top_row {
            self.top_row = cursor.row;
        }
        if cursor.row >= self.top_row + self.visible_rows {
            self.top_row = cursor.row - self.visible_rows + 1;
        }
        if cursor.col < self.left_col {
            self.left_col = cursor.col;
        }
        if cursor.col >= self.left_col + self.visible_cols {
            self.left_col = cursor.col - self.visible_cols + 1;
        }
    }

    pub fn update_dimensions(&mut self, width: u16, height: u16) {
        let row_num_width: u16 = 6;
        let header_height: u16 = 3;
        let footer_height: u16 = 2;

        let grid_width = width.saturating_sub(row_num_width);
        let grid_height = height.saturating_sub(header_height + footer_height);

        self.visible_cols = (grid_width / self.col_width).max(1);
        self.visible_rows = grid_height as u32;
    }

    pub fn scroll_half_page_down(&mut self, max_row: u32) {
        let half = self.visible_rows / 2;
        self.top_row = (self.top_row + half).min(max_row.saturating_sub(self.visible_rows));
    }

    pub fn scroll_half_page_up(&mut self) {
        let half = self.visible_rows / 2;
        self.top_row = self.top_row.saturating_sub(half);
    }
}
