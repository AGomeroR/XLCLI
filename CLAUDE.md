# xlcli

Terminal spreadsheet editor in Rust. Vim-compatible, fast, full Excel formula support.

## Build

```bash
cargo build          # debug
cargo build --release # release
cargo run -- file.xlsx  # open file
cargo run               # new blank workbook
```

## Workspace Structure

- `crates/xlcli-core/` — cell storage, types, dependency graph, styles
- `crates/xlcli-formulas/` — formula lexer, parser, evaluator (Phase 3)
- `crates/xlcli-io/` — file readers/writers (xlsx, csv, tsv, ods)
- `crates/xlcli-python/` — Python scripting via PyO3 (Phase 5)
- `crates/xlcli-db/` — database connectivity via sqlx (Phase 6)
- `crates/xlcli-tui/` — terminal UI, binary crate

## Conventions

- Sparse HashMap for cell storage: `HashMap<(u32, u16), Cell>`
- CellAddr is 8 bytes: (sheet: u16, row: u32, col: u16)
- Virtual scrolling: only render visible viewport
- No tokio in TUI — synchronous event loop, 60fps cap
- `thiserror` for library errors, `anyhow` for application errors
- `compact_str` for cell string values
