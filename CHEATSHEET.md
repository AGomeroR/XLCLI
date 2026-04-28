# xlcli Cheatsheet

## Modes

xlcli has 4 modes, like Vim:

| Mode      | Indicator | How to enter              | How to exit       |
|-----------|-----------|---------------------------|-------------------|
| Normal    | `NORMAL`  | Default / `Esc` from any  | —                 |
| Insert    | `INSERT`  | `Enter`, `i`, `a`, `s`, `=` | `Esc`, `Enter`, `Tab` |
| Visual    | `VISUAL`  | `v`                       | `Esc`, `y`, `d`   |
| Command   | `COMMAND` | `:`                       | `Enter`, `Esc`    |
| Search    | `SEARCH`  | `/`                       | `Enter`, `Esc`    |

### Search

| Key       | Action                                                  |
|-----------|---------------------------------------------------------|
| `/text`   | Enter search mode, type query, `Enter` to execute       |
| `n`       | Jump to next match (wraps at end)                       |
| `N`       | Jump to previous match (wraps at start)                 |
| `Esc`     | (in Normal) clear search highlights                     |

Match is case-insensitive substring against rendered cell value, active sheet only. All matches highlighted yellow; current match is bold light-yellow.

---

## Normal Mode

### Navigation

| Key          | Action                     |
|--------------|----------------------------|
| `h` / `Left` | Move left one cell         |
| `j` / `Down` | Move down one cell         |
| `k` / `Up`   | Move up one cell           |
| `l` / `Right`| Move right one cell        |
| `Tab`        | Move right one cell         |
| `Shift+Tab`  | Move left one cell          |
| `gg`         | Jump to first cell (A1)     |
| `G`          | Jump to last row            |
| `0`          | Jump to first column        |
| `$`          | Jump to last column         |
| `Ctrl+d`     | Scroll half page down       |
| `Ctrl+u`     | Scroll half page up         |

### Entering Insert Mode

| Key     | Action                                          |
|---------|-------------------------------------------------|
| `Enter` | Edit cell — keeps current value in buffer       |
| `i`     | Insert — clears cell, start typing fresh        |
| `a`     | Append — keeps current value in buffer          |
| `s`     | Substitute — clears cell, start typing fresh    |
| `=`     | Formula — clears cell, pre-fills `=` in buffer  |

### Clipboard

| Key  | Action                                   |
|------|------------------------------------------|
| `yy` | Yank (copy) current cell                 |
| `YY` | Yank with relative mode (adjusts refs)   |
| `p`  | Paste (exact after `yy`, relative after `YY`) |
| `dd` | Delete current cell                      |
| `x`  | Delete current cell                      |

### Fill

| Key     | Action                                          |
|---------|-------------------------------------------------|
| `Alt+d` | Fill down — copy formula from cell above, adjust refs |
| `Alt+r` | Fill right — copy formula from cell left, adjust refs |

### Undo / Redo

| Key      | Action |
|----------|--------|
| `u`      | Undo   |
| `Ctrl+r` | Redo   |

### Sheet Management

| Key              | Action                                              |
|------------------|-----------------------------------------------------|
| `t`              | New sheet — opens command box to set name            |
| `Ctrl+1`–`Ctrl+9`| Switch to sheet 1–9 (Normal, Visual, Insert+formula)|

During formula editing (`=`), `Ctrl+1`–`Ctrl+9` switches the view for cross-sheet ref browsing.
Click a cell on the other sheet to insert `Sheet2!A1` style references.
On confirm/cancel, returns to the formula's origin sheet.

### Other

| Key | Action                               |
|-----|--------------------------------------|
| `v` | Enter Visual selection mode          |
| `:` | Enter Command mode                   |
| `Shift+Enter` | Open filter dialog (on header row) |
| `q` | Quit (if no unsaved changes)         |

---

## Insert Mode

| Key         | Action                                    |
|-------------|-------------------------------------------|
| Any char    | Type into cell                            |
| `Backspace` | Delete last character                     |
| `Esc`       | Confirm edit, return to Normal            |
| `Enter`     | Confirm edit, move down one row           |
| `Tab`       | Confirm edit, move right, start new insert|

### Formula Entry

Type `=` followed by the formula. Examples:
```
=SUM(A1:A10)
=IF(B1>100,"high","low")
=VLOOKUP(A1,B1:D20,3,FALSE)
=A1+B1*2
```

The formula bar shows `=FORMULA` text; the grid shows the computed result.
When you press `Enter` on a cell with a formula, the formula text is loaded for editing.

---

## Visual Mode

Select a range of cells for bulk operations.

| Key          | Action                      |
|--------------|-----------------------------|
| `h/j/k/l`   | Extend selection            |
| Arrow keys   | Extend selection            |
| `y`          | Yank selection, exit Visual |
| `d` / `x`   | Delete selection, exit Visual|
| `f`          | Fill selection from top-left cell (adjusts refs) |
| `Esc`        | Cancel selection            |

---

## Command Mode

Press `:` to open the command palette (or inline status bar, depending on config).

### File Commands

| Command       | Action                                 |
|---------------|----------------------------------------|
| `:w`          | Save file                              |
| `:w filename` | Save to specific file                  |
| `:q`          | Quit (fails if unsaved changes)        |
| `:q!`         | Force quit, discard changes            |
| `:wq`         | Save and quit                          |

### Navigation Commands

| Command       | Action                                 |
|---------------|----------------------------------------|
| `:c A1`       | Jump to cell A1                        |
| `:c AB17`     | Jump to cell AB17                      |

### Row Operations

| Command            | Action                          |
|--------------------|---------------------------------|
| `:ir`              | Insert row above cursor         |
| `:insert row`      | Insert row above cursor         |
| `:ir!`             | Insert row below cursor         |
| `:insert row below`| Insert row below cursor         |
| `:dr`              | Delete current row              |
| `:delete row`      | Delete current row              |

### Sheet Commands

| Command                  | Action                                      |
|--------------------------|---------------------------------------------|
| `:sheet add`             | Add new sheet (auto-name Sheet2, Sheet3...) |
| `:sheet add MySheet`     | Add new sheet with given name               |
| `:sheet 1`               | Switch to sheet 1 (1-indexed)               |
| `:sheet 3`               | Switch to sheet 3                           |
| `:sheet MySheet`         | Switch to sheet by name                     |
| `:sheet delete`          | Delete current sheet                        |
| `:sheet delete 3`        | Delete sheet 3                              |
| `:sheet rename old new`  | Rename sheet from "old" to "new"            |
| `:sheet move left`       | Move current sheet one position left        |
| `:sheet move right`      | Move current sheet one position right       |
| `:sheet move 2`          | Move current sheet to position 2            |
| `:sheet move 1 3`        | Move sheet 1 to position 3                  |

### Sort

| Command                   | Action                                      |
|---------------------------|---------------------------------------------|
| `:sort`                   | Open sort dialog (popup)                    |
| `:sort B`                 | Sort by column B, A-Z                       |
| `:sort B desc`            | Sort by column B, Z-A                       |
| `:sort B num`             | Sort by column B, numeric ascending         |
| `:sort B numdesc`         | Sort by column B, numeric descending        |
| `:sort B case`            | Sort by column B, case-sensitive            |
| `:sort B natural`         | Sort by column B, natural order             |
| `:sort B desc A`          | Sort by B descending, then A ascending      |
| `:sort B A num`           | Sort by B (A-Z), then A (numeric)           |

In Visual mode, sort applies to selected range. Without selection, sorts entire sheet (configurable via `sort.allow_full_sheet` in config).

Sort dialog (`:sort` with no args): navigate with `Tab`/`Shift+Tab`, change values with `Left`/`Right` (`h`/`l`), toggle headers with `Space`/`Enter`, confirm with `Sort` button.

### Freeze Panes

| Command              | Action                                      |
|----------------------|---------------------------------------------|
| `:freeze 1`           | Freeze first row only                      |
| `:freeze 1 A`         | Freeze first row and column A              |
| `:freeze 1-3`         | Freeze rows 1-3 only                       |
| `:freeze 1-3 B-C`     | Freeze rows 1-3 and columns B-C            |
| `:unfreeze`            | Remove all freeze panes                    |

A border line separates frozen and scrollable regions. Frozen rows/columns stay visible while scrolling.

### Table Headers

| Command              | Action                                      |
|----------------------|---------------------------------------------|
| `:headers`           | Set row 1 as header (or visual selection row if active) |
| `:headers 1`         | Set row 1 as header row                     |
| `:unheaders`         | Remove header row                           |

### Named Ranges

| Command                     | Action                                      |
|-----------------------------|---------------------------------------------|
| `:name MyRange A1:B10`      | Define named range `MyRange` = A1:B10       |
| `:name Total B5`            | Define single-cell name `Total` = B5        |
| `:name MyRange` (visual)    | Name the current visual selection           |
| `:name delete MyRange`      | Delete named range                          |
| `:names`                    | List all named ranges                       |

Use names directly in formulas: `=SUM(MyRange)`, `=Total*2`.

### Conditional Formatting

Rules match a **condition** against cell value, then apply a **style overlay**.

| Command                                      | Action                                      |
|----------------------------------------------|---------------------------------------------|
| `:cf gt 100 bg=red`                          | Visual selection: red bg when cell > 100    |
| `:cf A1:B10 lt 0 fg=red bold`                | Explicit range                              |
| `:cf contains foo italic`                    | Rule on visual selection, italic if contains "foo" |
| `:cf between 0 10 bg=green`                  |                                             |
| `:cf blanks bg=gray`                         |                                             |
| `:cf base bold fg=white`                     | Sheet-wide base style                       |
| `:cf clean` (visual)                         | Remove rules overlapping selection          |
| `:cf clean A1:B5`                            | Remove rules overlapping range              |
| `:cf clean all`                              | Remove all rules + base                     |
| `:cf list`                                   | Open rules editor popup                     |

**Conditions**: `gt N`, `lt N`, `gte N`, `lte N`, `eq N`, `neq N`, `between A B`, `contains TEXT`, `blanks`, `nonblanks`.

**Style tokens**: `bold`, `italic`, `under`, `dunder`, `strike`, `over`, `bg=<color>`, `fg=<color>`. Negate with `!` (e.g., `!bold`).

**Colors**: `red green blue yellow cyan magenta orange white black gray none`, or hex `#RRGGBB` (in `:cf` dialog hex field).

**Cond-format is round-tripped to xlsx** — rules saved as native `<conditionalFormatting>` + dxf entries. Excel/LibreOffice see same formatting.

#### Rules editor popup (`:cf list`)

| Key        | Action                                 |
|------------|----------------------------------------|
| `j` / `k`  | Next / prev rule                        |
| `gg` / `G` | First / last rule                       |
| `Enter`    | Edit selected rule (opens `:cf` dialog) |
| `dd`       | Delete selected rule                    |
| `a`        | New rule (opens `:cf` dialog)           |
| `Esc`/`q`  | Close                                   |

#### Format dialog hex color

In the `:cf` popup, each of BG / FG has an optional `[#RRGGBB]` text field next to the preset picker. Non-empty valid hex overrides the preset. Empty → preset is used.

The dialog also shows a live **Preview: ● Sample** row that reflects current BG/FG + bold/italic/underline/strikethrough choices.

### Text Case

| Command              | Action                                      |
|----------------------|---------------------------------------------|
| `:case upper`        | UPPERCASE visual selection (or cursor cell) |
| `:case lower`        | lowercase                                   |
| `:case title`        | Title Case                                  |
| `:case sentence`     | Sentence case                               |
| `:case toggle`       | tOGGLE cASE                                 |

Header row displays bold with background highlight. Sort automatically skips header row. Clicking or pressing `Shift+Enter` on a header cell opens the filter dialog.

### AutoFilter

| Command                  | Action                                          |
|--------------------------|------------------------------------------------|
| `:filter`                | Open filter dialog (visual selection or sheet)  |
| `:filter all`            | Open filter dialog for entire sheet             |
| `:filter B = hello`      | Filter column B: exact match "hello"           |
| `:filter B != test`      | Filter column B: not equal "test"              |
| `:filter C > 100`        | Filter column C: greater than 100              |
| `:filter C < 50`         | Filter column C: less than 50                  |
| `:filter C >= 10`        | Filter column C: greater or equal 10           |
| `:filter C <= 99`        | Filter column C: less or equal 99              |
| `:filter B contains abc` | Filter column B: contains substring "abc"      |
| `:filter B blanks`       | Filter column B: show only blank cells         |
| `:filter B nonblanks`    | Filter column B: show only non-blank cells     |
| `:filter B top 10`       | Filter column B: top 10 numeric values         |
| `:filter B bottom 5`     | Filter column B: bottom 5 numeric values       |
| `:unfilter`              | Remove filters for visual selection columns    |
| `:unfilter all`          | Remove all filters                             |
| `:unfilter B`            | Remove filter on column B only                 |

Filters stack across columns (AND logic). Row numbers are preserved (like Excel). Column headers show `▼` indicator when filtered.

**Opening filter dialog:**
- `:filter` with no column argument
- `Shift+Enter` on a header cell (Normal mode)
- Mouse click on a header cell

**Filter dialog (Equals type):** Shows multi-select checkbox list of all unique column values. `(All)` at top toggles all on/off. `Space`/`Enter` toggles individual items. Type to search/filter the list. `Up`/`Down` navigates. `Tab` moves between fields. `Apply` confirms.

**Filter dialog (other types):** Type value in text field, confirm with `Apply` button. Navigate with `Tab`/`Shift+Tab`, change dropdowns with `Left`/`Right` (`h`/`l`).

### Column Operations

| Command              | Action                        |
|----------------------|-------------------------------|
| `:ic`                | Insert column left of cursor  |
| `:insert col`        | Insert column left of cursor  |
| `:ic!`               | Insert column right of cursor |
| `:insert col right`  | Insert column right of cursor |
| `:dc`                | Delete current column         |
| `:delete col`        | Delete current column         |

---

## Supported Formulas (506 total)

All formulas use `=` prefix. Ranges use colon syntax: `A1:B10`.
Arguments separated by `,` or `;`.

### Math (40)

```
SUM        AVERAGE     MIN         MAX         ABS
ROUND      ROUNDUP     ROUNDDOWN   INT         MOD
POWER      SQRT        LOG         LOG10       LN
EXP        PI          RAND        RANDBETWEEN CEILING
FLOOR      SIGN        PRODUCT     SUMPRODUCT  SIN
COS        TAN         ASIN        ACOS        ATAN
ATAN2      DEGREES     RADIANS     EVEN        ODD
FACT       GCD         LCM         TRUNC       QUOTIENT
```

### Text (24)

```
LEN         LEFT        RIGHT       MID         UPPER
LOWER       PROPER      TRIM        CLEAN       CONCATENATE
CONCAT      TEXTJOIN    REPT        SUBSTITUTE  REPLACE
FIND        SEARCH      TEXT        VALUE       EXACT
T           CHAR        CODE        NUMBERVALUE
```

### Logical (11)

```
IF          AND         OR          NOT         XOR
IFERROR     IFNA        IFS         SWITCH      TRUE
FALSE
```

### Lookup & Reference (13)

```
VLOOKUP     HLOOKUP     INDEX       MATCH       XLOOKUP
CHOOSE      ROW         COLUMN      ROWS        COLUMNS
ADDRESS     INDIRECT*   OFFSET*
```

`*` = stub, returns `#REF!` (requires dynamic range resolution)

### Statistical (19)

```
COUNT       COUNTA      COUNTBLANK  COUNTIF     SUMIF
AVERAGEIF   MEDIAN      MODE        STDEV       STDEVP
VAR         VARP        LARGE       SMALL       RANK
PERCENTILE  CORREL      MINIFS      MAXIFS
```

### Information (15)

```
ISBLANK     ISERROR     ISERR       ISNA        ISNUMBER
ISTEXT      ISLOGICAL   ISNONTEXT   ISEVEN      ISODD
TYPE        N           NA          ERROR.TYPE  SHEET
```

### Date & Time (18)

```
NOW         TODAY       DATE        TIME        YEAR
MONTH       DAY         HOUR        MINUTE      SECOND
WEEKDAY     WEEKNUM     DATEVALUE   DAYS        EDATE
EOMONTH     DATEDIF     ISOWEEKNUM
```

### Financial (28)

```
PMT         FV          PV          NPV         IRR
RATE        NPER        IPMT        PPMT        CUMIPMT
CUMPRINC    SLN         SYD         DB          DDB
EFFECT      NOMINAL     DOLLARDE    DOLLARFR    PDURATION
RRI         DISC        TBILLEQ     TBILLPRICE  TBILLYIELD
XNPV        MIRR        ISPMT
```

### Engineering (27)

```
DEC2BIN     DEC2OCT     DEC2HEX     BIN2DEC     BIN2OCT
BIN2HEX     OCT2DEC     OCT2BIN     OCT2HEX     HEX2DEC
HEX2BIN     HEX2OCT     BITAND      BITOR       BITXOR
BITLSHIFT   BITRSHIFT   COMPLEX     IMAGINARY   IMREAL
IMABS       IMSUM       DELTA       GESTEP      ERF
ERFC        CONVERT
```

---

## Formula Autocomplete

When typing a formula (after `=`), an autocomplete popup appears:

| Key         | Action                           |
|-------------|----------------------------------|
| Type chars  | Filter matching formulas         |
| `Down`      | Select next match                |
| `Up`        | Select previous match            |
| `Tab`/`Enter`| Accept selected formula + `(`  |
| `Esc`       | Dismiss autocomplete             |

Example: typing `=SU` shows `SUM`, `SUBSTITUTE`, `SUMIF`, `SUMPRODUCT`...

---

## Operators

Usable in formulas:

| Operator | Meaning              | Example          |
|----------|----------------------|------------------|
| `+`      | Addition             | `=A1+B1`         |
| `-`      | Subtraction          | `=A1-B1`         |
| `*`      | Multiplication       | `=A1*B1`         |
| `/`      | Division             | `=A1/B1`         |
| `^`      | Exponentiation       | `=2^10`          |
| `&`      | Text concatenation   | `=A1&" "&B1`     |
| `=`      | Equal                | `=A1=B1`         |
| `<>`     | Not equal            | `=A1<>B1`        |
| `<`      | Less than            | `=A1<B1`         |
| `>`      | Greater than         | `=A1>B1`         |
| `<=`     | Less or equal        | `=A1<=B1`        |
| `>=`     | Greater or equal     | `=A1>=B1`        |
| `%`      | Percent (postfix)    | `=50%` → `0.5`  |
| `-`      | Negation (prefix)    | `=-A1`           |

### Precedence (high to low)

1. `%` (postfix percent)
2. `-` `+` (unary prefix)
3. `^` (exponent, right-associative)
4. `*` `/`
5. `+` `-`
6. `&` (concatenation)
7. `=` `<>` `<` `>` `<=` `>=` (comparison)

---

## Mouse

| Action              | Effect                                  |
|---------------------|-----------------------------------------|
| Left click on cell  | Move cursor to that cell                |
| Left click on tab   | Switch to that sheet                    |
| Scroll up           | Move cursor up 3 rows                  |
| Scroll down         | Move cursor down 3 rows                |

Mouse works in all modes. Clicking while in Insert mode confirms the current edit first.

---

## Cell References

| Syntax              | Type             | Example                    |
|---------------------|------------------|----------------------------|
| `A1`                | Relative         | `=A1`                      |
| `$A1`               | Absolute column  | `=$A1`                     |
| `A$1`               | Absolute row     | `=A$1`                     |
| `$A$1`              | Fully absolute   | `=$A$1`                    |
| `A1:B10`            | Range            | `=SUM(A1:B10)`             |
| `Sheet2!A1`         | Cross-sheet ref  | `=Sheet2!A1`               |
| `'My Sheet'!A1`     | Quoted sheet ref | `='My Sheet'!A1`           |
| `Sheet2!A1:B10`     | Cross-sheet range| `=SUM(Sheet2!A1:B10)`      |

---

## Configuration

Config file: `~/.config/xlcli/config.toml`

```toml
[command_palette]
enabled = true            # true/false — floating palette or inline
position = "top"          # "top", "center", "bottom"
width_percent = 50        # 20-90
```

---

## CLI Usage

```bash
xlcli                     # New blank workbook
xlcli file.xlsx           # Open xlsx file
xlcli data.csv            # Open CSV file
xlcli sheet.tsv           # Open TSV file
xlcli book.ods            # Open ODS file
xlcli completions <shell> # Print shell completion script (bash|zsh|fish|powershell|elvish)
```

Supported formats: `.xlsx`, `.xls`, `.csv`, `.tsv`, `.ods`

## Install

- **Arch:** `cd packaging/arch && makepkg -si`
- **Ubuntu/Debian:** `bash packaging/build-deb.sh && sudo apt install ./target/debian/xlcli_*.deb`
- **From source:** `make build && sudo make install` (uninstall: `sudo make uninstall`)
