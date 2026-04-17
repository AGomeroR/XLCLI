use std::collections::HashMap;

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct CellStyle {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub double_underline: bool,
    pub strikethrough: bool,
    pub overline: bool,
    pub fg_color: Option<Color>,
    pub bg_color: Option<Color>,
    pub number_format: Option<String>,
    pub h_align: HAlign,
    pub v_align: VAlign,
}

impl Default for CellStyle {
    fn default() -> Self {
        Self {
            bold: false,
            italic: false,
            underline: false,
            double_underline: false,
            strikethrough: false,
            overline: false,
            fg_color: None,
            bg_color: None,
            number_format: None,
            h_align: HAlign::General,
            v_align: VAlign::Bottom,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Color {
    pub r: u8,
    pub g: u8,
    pub b: u8,
}

impl Color {
    pub fn new(r: u8, g: u8, b: u8) -> Self {
        Self { r, g, b }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum HAlign {
    #[default]
    General,
    Left,
    Center,
    Right,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum VAlign {
    Top,
    Center,
    #[default]
    Bottom,
}

#[derive(Debug)]
pub struct StylePool {
    styles: Vec<CellStyle>,
    dedup: HashMap<CellStyle, u32>,
}

impl StylePool {
    pub fn new() -> Self {
        let default_style = CellStyle::default();
        let mut dedup = HashMap::new();
        dedup.insert(default_style.clone(), 0);
        Self {
            styles: vec![default_style],
            dedup,
        }
    }

    pub fn get_or_insert(&mut self, style: CellStyle) -> u32 {
        if let Some(&id) = self.dedup.get(&style) {
            return id;
        }
        let id = self.styles.len() as u32;
        self.dedup.insert(style.clone(), id);
        self.styles.push(style);
        id
    }

    pub fn get(&self, id: u32) -> &CellStyle {
        &self.styles[id as usize]
    }
}

impl Default for StylePool {
    fn default() -> Self {
        Self::new()
    }
}
