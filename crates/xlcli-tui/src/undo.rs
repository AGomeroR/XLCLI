use xlcli_core::cell::Cell;
use xlcli_core::types::CellAddr;

#[derive(Debug, Clone)]
pub enum UndoEntry {
    CellChange {
        addr: CellAddr,
        old: Option<Cell>,
        new: Option<Cell>,
    },
    Batch(Vec<UndoEntry>),
}

pub struct UndoStack {
    undo: Vec<UndoEntry>,
    redo: Vec<UndoEntry>,
}

impl UndoStack {
    pub fn new() -> Self {
        Self {
            undo: Vec::new(),
            redo: Vec::new(),
        }
    }

    pub fn push(&mut self, entry: UndoEntry) {
        self.undo.push(entry);
        self.redo.clear();
    }

    pub fn undo(&mut self) -> Option<UndoEntry> {
        let entry = self.undo.pop()?;
        let reversed = reverse_entry(&entry);
        self.redo.push(reversed);
        Some(entry)
    }

    pub fn redo(&mut self) -> Option<UndoEntry> {
        let entry = self.redo.pop()?;
        let reversed = reverse_entry(&entry);
        self.undo.push(reversed);
        Some(entry)
    }

    pub fn can_undo(&self) -> bool {
        !self.undo.is_empty()
    }

    pub fn can_redo(&self) -> bool {
        !self.redo.is_empty()
    }
}

fn reverse_entry(entry: &UndoEntry) -> UndoEntry {
    match entry {
        UndoEntry::CellChange { addr, old, new } => UndoEntry::CellChange {
            addr: *addr,
            old: new.clone(),
            new: old.clone(),
        },
        UndoEntry::Batch(entries) => {
            UndoEntry::Batch(entries.iter().rev().map(reverse_entry).collect())
        }
    }
}
