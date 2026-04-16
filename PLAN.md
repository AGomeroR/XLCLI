# xlcli — CLI Spreadsheet Editor Implementation Plan

## Context

Building **xlcli**, a full-featured TUI spreadsheet editor in Rust. Goal: fast, snappy, vim-compatible Excel replacement in terminal. Must support all 500+ Excel formulas, multiple file formats (xlsx, csv, ods, tsv), advanced features (charts, pivots, macros), Python scripting with pandas (bundled runtime), and database connectivity (PostgreSQL, MySQL, SQLite). License: GPL-3.0.

---

## Cargo Workspace Layout

```
/home/agomeror/Work/CLI_excel/
├── Cargo.toml                  # workspace root
├── CLAUDE.md
├── crates/
│   ├── xlcli-core/             # engine: cell storage, dep graph, recalc
│   ├── xlcli-formulas/         # formula parser + evaluator
│   ├── xlcli-io/               # file format read/write (xlsx, csv, tsv, ods)
│   ├── xlcli-python/           # Python scripting bridge (PyO3 + bundled CPython+pandas)
│   ├── xlcli-db/               # database connectivity (PostgreSQL, MySQL, SQLite)
│   └── xlcli-tui/              # terminal UI (binary crate)
```

**Dependency graph**: `core` ← `formulas`, `core` ← `io`, `core` ← `python`, `core` ← `db`, all ← `tui`

### Key External Dependencies

| Crate | Purpose |
|---|---|
| `ratatui` 0.30 | TUI rendering |
| `crossterm` 0.29 | Terminal backend |
| `calamine` 0.34 | Read xlsx/xls/ods |
| `rust_xlsxwriter` 0.94 | Write xlsx |
| `csv` 1.3 | CSV/TSV I/O |
| `quick-xml` 0.39 + `zip` 2.6 | ODS writer |
| `logos` 0.16 | Formula lexer (compile-time DFA) |
| `petgraph` 0.8 | Dependency graph + toposort |
| `chrono` 0.4 | Date/time for formulas |
| `compact_str` 0.8 | Inline small strings (no heap for short cell values) |
| `rayon` 1.12 | Parallel formula recalc |
| `clap` 4.x | CLI arg parsing |
| `anyhow`/`thiserror` | Error handling |
| `pyo3` 0.23 | Python ↔ Rust bridge (embed CPython) |
| `sqlx` 0.8 | Async DB driver (Postgres, MySQL, SQLite) |
| `tokio` 1.x | Async runtime for sqlx DB queries only |

---

## Core Engine Architecture (`xlcli-core`)

### Cell Storage: Sparse HashMap

```rust
// sheet.rs
cells: HashMap<(u32, u16), Cell>  // (row, col) -> Cell
```

- O(1) random access. Memory proportional to filled cells only.
- `u16` for cols (max 65535, Excel caps at 16384).
- `CellValue` enum: Number(f64), String(CompactString), Boolean, DateTime, Error, Empty, Array.
- Shared style pool: cells store `u32` style_id, not full style data.

### Dependency Graph + Recalculation

- `petgraph::DiGraphMap<CellAddr, ()>` — edges from dependency → dependent.
- **Lazy-incremental**: on cell edit, BFS forward to mark dirty dependents, toposort dirty subgraph only, eval in order.
- Cycle detection via Tarjan's SCC → set cycle cells to `#REF!`.
- Batch mode on file load: toposort all formulas, eval with `rayon` for independent subgraphs.

### CellAddr: 8 bytes (sheet: u16, row: u32, col: u16)

---

## Formula Engine (`xlcli-formulas`)

### Pipeline: `logos` lexer → Pratt parser → AST → tree-walk evaluator

- **Lexer**: logos compile-time DFA. Zero-copy tokens.
- **Parser**: Pratt (precedence climbing). Handles prefix, infix, postfix operators. No backtracking.
- **AST**: `Expr` enum (Number, String, Bool, CellRef, Range, BinOp, UnaryOp, FnCall, Array).
- **Evaluator**: Function registry with `fn` pointers (no Box<dyn>). `EvalContext` trait for workbook access.

### 500+ Formulas — organized by category files:

| Category | ~Count | ~LOC |
|---|---|---|
| Math | 70 | 800 |
| Text | 50 | 600 |
| Logical | 15 | 150 |
| Lookup/Reference | 30 | 800 |
| Date/Time | 40 | 500 |
| Statistical | 60 | 1000 |
| Financial | 55 | 700 |
| Information | 25 | 200 |
| Engineering | 50 | 600 |
| Database | 12 | 300 |
| Dynamic Array | 10 | 400 |
| Other | ~30 | 300 |
| **Total** | **~500** | **~6400** |

---

## TUI Architecture (`xlcli-tui`)

### Layout (PLACEHOLDER — will be redesigned later)
```
┌─────────────────────────────────────┐
│ FormulaBar  [A1] [=SUM(A1:A10)   ] │
├────┬────────────────────────────────┤
│    │  A   │  B   │  C   │  D   │...│
├────┼──────┼──────┼──────┼──────┤   │
│  1 │ Hello│ 42   │      │=A1+B1│   │
│  2 │      │      │      │      │   │
│ ...│      │      │      │      │   │
├────┴──────┴──────┴──────┴──────┘   │
│ Sheet1 | Sheet2 | Sheet3 |    [+]  │
├─────────────────────────────────────┤
│ NORMAL | A1 | Sheet1 | file.xlsx   │
└─────────────────────────────────────┘
```
> **Note**: This layout is a temporary placeholder. Final TUI design will be decided later.

### Virtual Scrolling — only render visible viewport (~1000 cells). O(viewport), not O(sheet_size).

### Event Loop — synchronous, 60fps cap (16ms poll). No tokio needed.

### Vim Mode State Machine
- **Normal**: hjkl, gg, G, 0, $, Ctrl-d/u, y/p/d, u/Ctrl-r
- **Insert**: edit cell content, Enter confirm, Escape cancel
- **Visual**: range selection with movement
- **Command**: `:w`, `:q`, `:wq`, `:c AB17`, `:sort`, `:filter`, `:freeze`, etc.
- Multi-key sequences via pending-key buffer (gg, dd, yy)

### Undo/Redo — command pattern, `Vec<UndoEntry>` stack. BatchChange for multi-cell ops.

---

## File I/O (`xlcli-io`)

- **Read**: `FileReader` trait. calamine for xlsx/ods/xls, csv crate for csv/tsv.
- **Write**: `FileWriter` trait. rust_xlsxwriter for xlsx, csv crate for csv/tsv, quick-xml+zip for ods.
- **Convert pipeline**: read → Workbook (neutral) → write. Enables `xlcli convert in.xlsx out.csv`.

---

## Python Scripting (`xlcli-python`)

### Architecture: Bundled CPython + pandas via PyO3

- Ship full CPython runtime + pandas/numpy inside xlcli package (~150MB total)
- PyO3 embeds interpreter in Rust process. No system Python needed.
- ~200ms cold start on first `:py` command per session. After that, fast.

### User-facing API

```python
# Available in xlcli Python environment:
import xlcli
import pandas as pd

# Read sheet data into pandas DataFrame
df = xlcli.sheet("Sheet1").to_dataframe("A1:D100")

# Write DataFrame back to sheet
xlcli.sheet("Sheet1").from_dataframe(df, "A1")

# Read/write single cells
val = xlcli.cell("A1")
xlcli.set_cell("A1", 42)

# Run SQL on current workbook (sheets as tables)
df = xlcli.sql("SELECT * FROM Sheet1 WHERE col_A > 100")
```

### TUI commands
- `:py <expression>` — run one-liner
- `:pyfile <path>` — run script file
- `:pyrepl` — enter interactive Python REPL mode (Escape to exit)

### Build/packaging
- Use `maturin` or custom build script to bundle CPython + wheels
- Platform-specific: build for linux-x86_64, linux-aarch64, macos-x86_64, macos-aarch64
- Pre-compiled pandas/numpy wheels included in package

---

## Database Connectivity (`xlcli-db`)

### Supported: PostgreSQL, MySQL, SQLite

- **sqlx** crate for async queries (compile-time checked SQL)
- **tokio** runtime spawned in background for DB operations only
- Connection strings stored in config or entered via command

### TUI commands
- `:db connect <connection_string>` — connect to database
- `:db query <SQL>` — run query, load results into new sheet
- `:db tables` — list tables
- `:db export <sheet> <table>` — write sheet data to DB table
- `:db disconnect` — close connection

### Python integration
```python
import xlcli
conn = xlcli.db.connect("postgresql://user:pass@host/db")
df = conn.query("SELECT * FROM users")
xlcli.sheet("DBResults").from_dataframe(df, "A1")
```

### Live refresh (Phase 5+)
- `:db watch <SQL> <interval>` — re-run query every N seconds, update sheet

---

## Implementation Phases

### Phase 1: Foundation (weeks 1-3) — Open, view, navigate
- Workspace setup, all 4 crate skeletons
- Core types: CellValue, Cell, Sheet, Workbook, CellAddr
- xlsx/csv readers (calamine + csv crate)
- TUI: grid with virtual scroll, headers, status bar, formula bar (display)
- Vim Normal mode navigation
- `:q` quit, CLI `xlcli <file>` / `xlcli new`
- **Result**: Open 100k-row xlsx, scroll at 60fps

### Phase 2: Editing (weeks 4-6) — Edit cells, save
- Insert mode, cell editing
- xlsx writer (rust_xlsxwriter), `:w`, `:wq`
- Undo/redo, Visual mode, copy/paste
- Insert/delete rows/cols
- `:c AB17` cell jump
- **Result**: Full edit workflow, open → edit → save xlsx

### Phase 3: Formula Engine (weeks 7-12) — All formulas
- Lexer, parser, evaluator, function registry
- Dependency graph + incremental recalc
- All 500+ formulas implemented by category (6 weeks, ~1000 LOC/week)
- Formula bar shows formula text, grid shows result
- **Result**: Complete formula engine

### Phase 4: Multi-Sheet + Advanced (weeks 13-16)
- Sheet tabs, add/delete/rename/reorder sheets
- Cross-sheet references, named ranges
- Sort, AutoFilter, freeze panes
- Conditional formatting
- TSV + ODS readers
- **Result**: Full workbook editing

### Phase 5: Python Scripting (weeks 17-20)
- PyO3 integration, bundled CPython runtime
- `xlcli` Python module: to_dataframe, from_dataframe, cell access
- `:py`, `:pyfile`, `:pyrepl` commands
- Bundle pandas + numpy in package
- Build pipeline for multi-platform bundled Python
- **Result**: Python+pandas scripting works out of box

### Phase 6: Database Connectivity (weeks 21-24)
- `xlcli-db` crate with sqlx (Postgres, MySQL, SQLite)
- `:db connect/query/tables/export/disconnect` commands
- Query results → new sheet
- Python DB integration (`xlcli.db.connect()`)
- `:db watch` live refresh
- **Result**: Full DB connectivity

### Phase 7: Power Features (weeks 25-30)
- TUI charts (bar, line, pie)
- Pivot tables (command-driven)
- Data validation
- ODS writer
- Rich text display, hyperlinks
- Macro recorder (`:macro record/stop/run`)
- Search/replace (`:s/old/new/g`)
- `xlcli.sql()` — query workbook sheets as SQL tables
- Performance optimization pass (1M row profiling)
- **Result**: Feature-complete

### Phase 8: Polish (weeks 31-34)
- Help system, config file (`~/.config/xlcli/config.toml`)
- Mouse support, system clipboard
- CI, tests, docs, packaging
- `xlcli convert` subcommand
- Multi-platform release builds with bundled Python
- AUR, brew, cargo install packaging

---

## Performance Strategy (5 pillars)

1. **Sparse HashMap** — O(1) cell access, memory ∝ filled cells
2. **Virtual scroll** — render only visible ~1000 cells
3. **Incremental recalc** — dirty-subgraph toposort, not full sheet
4. **logos DFA + Pratt parser** — microsecond formula parsing
5. **compact_str** — inline short strings, reduce allocator pressure

---

## Verification

- Phase 1: open large xlsx (100k rows), measure scroll FPS, ensure ≥60fps
- Phase 2: round-trip test: open xlsx → edit → save → reopen → verify edits
- Phase 3: test formulas against Excel output for known test spreadsheets
- Each phase: `cargo test`, `cargo clippy`, manual TUI testing
- Perf gate at Phase 5: flamegraph on 1M-row file, optimize hot paths
