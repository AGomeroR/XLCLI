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

## TODO (iterate across sessions)

Status: [x]=done [ ]=pending [~]=in progress

### Format dialog (`:cf`)
- [x] Toggle "Conditional" — when off, hide cond/val1/val2 (apply as plain format via `Condition::Always`)
- [x] Popup recolored Cyan to match other popups (was Magenta)
- [x] Title renamed " Format " (was " Conditional Formatting ")
- [x] Layout: Range top → Styles (Bold/Italic/etc) → BG/FG → Conditional toggle → cond fields if on
- [x] Range/Val1/Val2 no longer auto-capture keyboard. Vim-style nav (Tab/hjkl) by default; Enter on field activates edit, Esc/Enter exits edit. Mouse click activates edit.
- [ ] Persist last-used style across dialog opens
- [ ] Custom colors (hex) input besides preset list
- [ ] Edit-existing-rule (load rule into form on Enter from RulesList)
- [ ] Inline preview swatch of style choice

### Command palette (`:`)
- [x] Live suggestions box appears on first character (filtered by prefix)
- [ ] Tab to autocomplete top match
- [ ] Up/Down nav of suggestions
- [ ] Show description of currently-highlighted suggestion bigger
- [ ] Fuzzy match (not just prefix)

### Pending broader items
- [ ] Phase 4: charts, pivots
- [ ] Phase 5: Python scripting (PyO3 + bundled CPython)
- [ ] Phase 6: DB connectivity (sqlx)
- [ ] Save/load conditional formatting rules to xlsx
