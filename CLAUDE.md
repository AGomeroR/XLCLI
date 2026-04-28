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

## Popup conventions (apply to ALL popups in this project)

1. **Border color**: prefer terminal-default foreground; fall back to Cyan if not feasible. Never a unique per-popup color.
2. **Conditional / optional sections**: gate behind a checkbox. Hide fields when off; show when on. Popup auto-resizes.
3. **Vim nav default**: text fields do NOT auto-capture keyboard. Tab/BackTab/hjkl move focus. Enter activates edit mode on a field; Esc/Enter exits edit mode (does NOT close dialog). Mouse click on a text field also activates edit.
4. **Style checkboxes stacked vertically** — one per row.
5. **Buttons stacked vertically** in column on the right/bottom.
6. **Last button = "Apply & Close"** — applies form then dismisses.
7. **Dismiss button** — closes popup, does nothing else (no apply, no clear).
8. **`Shift+G`** jumps focus to last button. **`gg`** jumps focus to first field.
9. **Auto-sized**: width fixed; height = base, grows only when optional sections are toggled on. Never larger than needed.
10. **Range default**: visual selection if any, else current cursor cell. Never empty.
11. **Prefill from existing**: on open, scan existing rules/state covering the range and populate fields (last match wins).
12. **Cursor block at insertion point** in any text field being edited. Left/Right/Home/End/Delete supported, mid-string insert.

## Rules

1. Always ask questions, asume nothing.
2. Procede until you end the current task and it compiles perfectly with no errors, even if they are no impact ones.
3. Run always with /caveman skill.
4. Always update CLAUDE.md with what you have done.
5. If you introduce new commands, update CHEATSHEET.md with them

## TODO (iterate across sessions)

Status: [x]=done [ ]=pending [~]=in progress

### Format dialog (`:cf`)
- [x] Toggle "Conditional" — when off, hide cond/val1/val2 (apply as plain format via `Condition::Always`)
- [x] Popup recolored Cyan to match other popups (was Magenta)
- [x] Layout: Range top → Styles (Bold/Italic/etc) → BG/FG → Conditional toggle → cond fields if on
- [x] Range/Val1/Val2 no longer auto-capture keyboard. Vim nav (Tab/hjkl) default; Enter activates edit, Esc/Enter exits edit. Mouse click activates edit.
- [x] Style checkboxes stacked vertically (Bold / Italic / Underline / Double Underline / Strikethrough / Overline) — one per row
- [x] Rules list removed from Format popup (data still in `:cf list`)
- [x] Bold/Italic/Underline/Strike modifiers applied even on cursor/visual/header cells (was being masked)
- [x] `Shift+G` jumps focus to Close button (vim bottom). `gg` jumps focus to Range (vim top)
- [x] Buttons stacked vertically in column: SetBase → Dismiss → Apply → Apply & Close (Delete/CleanAll dropped from popup, still callable via `:cf` commands)
- [x] Last button "Apply & Close" — applies form then dismisses dialog
- [x] Popup auto-sized: width 44, height = base content rows (no Conditional) and grows by 2 rows when Conditional toggled on
- [x] Prefill dialog from existing rule on open: when selection matches/is contained by a rule's range, populate styles + conditional fields from that rule (last match wins)
- [x] No visual selection on `:cf` open → range defaults to current cursor cell (was empty)
- [~] Persist last-used style across dialog opens — superseded by full xlsx round-trip (see below)
- [x] Custom colors (hex) input besides preset list — BG/FG `[#RRGGBB]` on dedicated row below preset. Selecting preset (dropdown or h/l cycle) writes the preset's hex into the field so user can edit from there. Hex overrides preset when both present.
- [x] Drop `Overline` — removed from `CellStyle`, `StyleOverlay`, `:cf` dialog, parser, serializer, xlsx_read (xlsx has no overline primitive)
- [x] Inline preview swatch of style choice — "Preview: ● Sample" row reflects BG/FG + bold/italic/under/strike
- [x] Reachable rules editor (`:cf list` opens dedicated popup) — j/k nav, Enter edit, dd delete, a new, gg/G first/last
- [x] Italic renders correctly in Ghostty (confirmed)

### Conditional formatting xlsx round-trip
- [x] Writer: `xlcli-io/src/xlsx_write.rs` emits `CondRule`s via `rust_xlsxwriter::ConditionalFormat{Cell,Text,Blank,Formula}` + `Format` (bold/italic/underline/strike/fg/bg). `Always` → `expression` w/ formula `TRUE`. double_underline uses `FormatUnderline::Double`. Overline dropped (xlsx has no overline).
- [x] Reader: `xlcli-io/src/xlsx_read.rs` parses `xl/styles.xml` `<dxfs>`, `xl/workbook.xml` sheet order + `xl/_rels/workbook.xml.rels`, then `xl/worksheets/sheetN.xml` `<conditionalFormatting>` blocks via `quick-xml`. Supports cellIs operators, containsText, containsBlanks/notContainsBlanks, expression(TRUE)=Always. Unsupported types (color scales, data bars, icon sets, formula-based other than TRUE) silently skipped.
- [x] `parse_hex_color` helper in `xlcli-tui/src/app.rs` (pub); public `color_by_name_pub` wrapper for render swatch resolution.

### Formula bar / Insert mode
- [x] Arrow keys move text cursor inside `edit_buffer` (Left/Right/Home/End/Delete + insertion at cursor)
- [x] Cursor rendered as inverted block at insertion point (was always at end)
- [x] Formula error popup: `ParseError { offset, message }` in `xlcli-formulas::parser`; parser tracks token byte spans. Live recompute in `App::recompute_formula_error()` called per render frame while in Insert mode w/ `=`-prefixed buffer. `render_formula_error` draws 2-row overlay below formula bar: red `^` caret at formula-relative column + red error message on next row. Offsets clamp to body length for EOF errors.

### Command palette (`:`)
- [x] Live suggestions box appears on first character
- [x] Tab autocompletes selected (or top) match into buffer
- [x] Up/Down nav cycles through suggestions; selected row highlighted
- [x] Fuzzy match (subsequence + prefix bonus)
- [x] Description of highlighted suggestion shown in dedicated 4-row Cyan-bordered panel below suggestions; title displays `:cmd`, body wraps full description italic

### Repo / packaging
- [x] Pushed to GitHub (`git@github.com:AGomeroR/XLCLI.git`)
- [x] Un-gitignore `Cargo.lock` (committed for reproducible package builds)
- [x] Bump workspace version to 1.0.0 (first functional release)
- [x] `xlcli completions <shell>` subcommand via `clap_complete::generate` (Bash/Zsh/Fish/PowerShell/Elvish); hidden from `--help`
- [x] `assets/xlcli.desktop` — Type=Application, Terminal=true, MimeType for xlsx/xls/ods/csv/tsv
- [x] `assets/xlcli.svg` — placeholder grid+caret icon (Catppuccin palette; user can replace)
- [x] `Makefile` — DESTDIR/PREFIX-aware `build`/`completions`/`install`/`uninstall`/`deb`/`clean`. Auto-regenerates completions before install.
- [x] `packaging/arch/PKGBUILD` + `xlcli.install` (post hooks for `update-desktop-database` + `gtk-update-icon-cache`); installs via `make install DESTDIR="$pkgdir" PREFIX=/usr`
- [x] `packaging/build-deb.sh` — auto-installs `cargo-deb`, runs `make build && make completions && cargo deb --no-build`. `[package.metadata.deb]` in `crates/xlcli-tui/Cargo.toml` ships binary, .desktop, icon, completions, README, LICENSE
- [x] `README.md` with install instructions for Arch, Ubuntu, from-source
- [ ] Drop `*.xlsx` from `.gitignore` if sample fixtures should ship; otherwise add a `tests/fixtures/` exception
- [ ] Add `rust-toolchain.toml` pinning compiler (currently relies on user-installed Rust ≥ 1.75)
- [ ] CI workflow (cargo build + clippy + test + package builds on push)

### Search (vim-style)
- [x] `Mode::Search` added; `/` enters, `Enter` executes, `Esc` cancels
- [x] Case-insensitive substring match against `display_value()`, active sheet only
- [x] `n` next / `N` prev with wrap-around
- [x] All matches highlighted yellow bg; current match bold light-yellow at cursor
- [x] `Esc` in Normal clears highlights when search active
- [x] Status line shows `/query [idx/total]`
- [x] `App` state: `search_buffer`, `search_results: Vec<(u32,u16)>`, `search_idx`, `search_active`
- [x] `render_search_bar` renders bottom-center 60-wide Cyan box during Search mode

### Pending broader items
- [ ] Phase 4: charts, pivots
- [ ] Phase 5: Python scripting (PyO3 + bundled CPython)
- [ ] Phase 6: DB connectivity (sqlx)
- [x] Save/load conditional formatting rules to xlsx (round-trip with `rust_xlsxwriter` + quick-xml reader)
- [x] xlsx writer handles full `StyleSpec` enum: `Overlay` (dxf), `ColorScale` (2/3-stop via `ConditionalFormat{2,3}ColorScale`), `DataBar` (`ConditionalFormatDataBar`), `IconSet` (`ConditionalFormatIconSet` + `ConditionalFormatCustomIcon` thresholds). `IconSetKind` mapped to `ConditionalFormatIconType` (e.g. `ThreeSymbols2 → ThreeSymbolsCircled`, `FourRating → FourHistograms`, `FiveQuarters → FiveQuadrants`). Extended `Condition` coverage: `NotBetween`, `NotContains`, `BeginsWith`, `EndsWith`, `Expression`. Unsupported conditions in overlay path silently skip (Top/Average/TimePeriod/Duplicate/Unique/Errors).
- [x] Cell-level styles (bold/italic set directly on a cell, not via cond rule) — reader parses `xl/styles.xml` fonts/fills/cellXfs into `StylePool`, tags `Cell.style_id` from sheet.xml `<c s="N">` attr; `render.rs` uses cell's pool style as base for `effective_style`; writer emits via `write_*_with_format` / `write_blank` preserving per-cell format
