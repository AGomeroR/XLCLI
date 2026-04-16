# xlcli Cheatsheet

## Modes

xlcli has 4 modes, like Vim:

| Mode      | Indicator | How to enter              | How to exit       |
|-----------|-----------|---------------------------|-------------------|
| Normal    | `NORMAL`  | Default / `Esc` from any  | —                 |
| Insert    | `INSERT`  | `Enter`, `i`, `a`, `s`, `=` | `Esc`, `Enter`, `Tab` |
| Visual    | `VISUAL`  | `v`                       | `Esc`, `y`, `d`   |
| Command   | `COMMAND` | `:`                       | `Enter`, `Esc`    |

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

| Key  | Action                    |
|------|---------------------------|
| `yy` | Yank (copy) current cell  |
| `p`  | Paste                     |
| `dd` | Delete current cell       |
| `x`  | Delete current cell       |

### Undo / Redo

| Key      | Action |
|----------|--------|
| `u`      | Undo   |
| `Ctrl+r` | Redo   |

### Other

| Key | Action                               |
|-----|--------------------------------------|
| `v` | Enter Visual selection mode          |
| `:` | Enter Command mode                   |
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

## Supported Formulas (195 total)

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

| Syntax   | Type             | Example |
|----------|------------------|---------|
| `A1`     | Relative         | `=A1`   |
| `$A1`    | Absolute column  | `=$A1`  |
| `A$1`    | Absolute row     | `=A$1`  |
| `$A$1`   | Fully absolute   | `=$A$1` |
| `A1:B10` | Range            | `=SUM(A1:B10)` |

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
```

Supported formats: `.xlsx`, `.xls`, `.csv`, `.tsv`, `.ods`
