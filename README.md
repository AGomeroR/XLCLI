# xlcli

Terminal spreadsheet editor in Rust. Vim-compatible, fast, full Excel formula support.

## Features

- xlsx / csv / tsv / ods read & write (round-trip with conditional formatting)
- Vim-style modal editing (Normal / Insert / Visual / Command)
- Full formula engine with dependency graph + parallel evaluation
- Conditional formatting: classic styles, color scales, data bars, icon sets
- Cell-level styles (bold/italic/underline/strikethrough, fg/bg colors)
- Command palette with live suggestions and description panel
- Shell completions for bash / zsh / fish

---

# Install

xlcli works the same on every Linux distro after install: a single `xlcli` command
on your `$PATH`, an entry in your application launcher, and right-click → Open With
support for `.xlsx` / `.csv` / `.ods` files.

Pick the section that matches your system.

## Arch Linux (CachyOS, Manjaro, EndeavourOS)

### Step 1 — install Rust if you don't have it

```bash
sudo pacman -S --needed rust make base-devel
```

(Rust is bundled with `base-devel` on most Arch derivatives. Skip the ones you
already have.)

### Step 2 — get the source

```bash
git clone https://github.com/AGomeroR/XLCLI.git
cd XLCLI
```

### Step 3 — install

You have two options. Pick one.

**Option A — Quick install with `make` (no pacman tracking):**

```bash
make build
sudo make install PREFIX=/usr
```

Done. Run `xlcli` from any terminal, or find it in your application launcher.

To uninstall later:

```bash
sudo make uninstall PREFIX=/usr
```

**Option B — Real pacman package with `makepkg` (recommended for clean uninstall):**

The bundled `PKGBUILD` downloads source from a tagged GitHub release. If you've
cloned the repo locally and want to build from your current checkout instead,
edit `packaging/arch/PKGBUILD` and replace the `source=(...)` line with:

```bash
source=()
```

Then add a `prepare()` step that copies the parent repo, or just use Option A.

For the standard tagged build:

```bash
cd packaging/arch
makepkg -si
```

This compiles xlcli, packages it, and installs it via pacman. To uninstall:

```bash
sudo pacman -R xlcli
```

### Step 4 — verify

```bash
xlcli --version           # should print: xlcli 1.0.0
xlcli                     # opens an empty spreadsheet
```

In your file manager, right-click any `.xlsx` file → Open With → xlcli.

---

## Ubuntu / Debian / Linux Mint / Pop!_OS

### Step 1 — install Rust and build tools

```bash
sudo apt update
sudo apt install -y build-essential curl git
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y
source "$HOME/.cargo/env"
```

The Rust version in Ubuntu's repos is sometimes too old. Installing through
[rustup](https://rustup.rs) (the command above) gives you a current toolchain
in `~/.cargo/bin`.

Verify:

```bash
rustc --version           # should print 1.75 or newer
```

### Step 2 — get the source

```bash
git clone https://github.com/AGomeroR/XLCLI.git
cd XLCLI
```

### Step 3 — install

You have two options. Pick one.

**Option A — Quick install with `make` (no apt tracking):**

```bash
make build
sudo make install PREFIX=/usr/local
```

Done. Run `xlcli` from any terminal, or find it in your application launcher.

To uninstall later:

```bash
sudo make uninstall PREFIX=/usr/local
```

**Option B — Real `.deb` package with `cargo-deb` (recommended for clean uninstall):**

```bash
bash packaging/build-deb.sh
```

That script will:
1. Install `cargo-deb` if missing
2. Build xlcli in release mode
3. Generate shell completions
4. Produce a `.deb` package in `target/debian/`

Then install it:

```bash
sudo apt install ./target/debian/xlcli_*.deb
```

To uninstall:

```bash
sudo apt remove xlcli
```

### Step 4 — verify

```bash
xlcli --version           # should print: xlcli 1.0.0
xlcli                     # opens an empty spreadsheet
```

In your file manager (Nautilus / Nemo / Files), right-click any `.xlsx` file →
Open With → xlcli.

---

## Shell completions

Completions are installed automatically with both Option A and Option B on both
distros. To activate them:

- **Bash** — restart your shell, or run `source /usr/share/bash-completion/completions/xlcli`
- **Zsh** — restart your shell. If completions don't appear, run `compinit`
- **Fish** — works immediately

Try it: type `xlcli ` and press `<Tab>`.

---

## Troubleshooting

**"command not found: xlcli" after install**
Your shell is caching `$PATH`. Open a new terminal, or run `hash -r`.

**"xlcli not in app launcher"**
Refresh the desktop database:
```bash
sudo update-desktop-database /usr/share/applications
sudo gtk-update-icon-cache -f /usr/share/icons/hicolor
```
Log out and back in if your launcher still doesn't show it.

**"Open With" doesn't list xlcli for `.xlsx`**
Set it as the default explicitly:
```bash
xdg-mime default xlcli.desktop application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
xdg-mime default xlcli.desktop text/csv
xdg-mime default xlcli.desktop application/vnd.oasis.opendocument.spreadsheet
```

**Build fails with "could not find Cargo.lock"**
Run from the repository root, not from inside `packaging/` or `crates/`.

**Build fails on Ubuntu with "rust 1.7x required"**
The Ubuntu-packaged Rust is too old. Install via rustup (see Step 1 above).

---

## Usage quick reference

```bash
xlcli                     # new blank workbook
xlcli sheet.xlsx          # open file
xlcli data.csv            # open csv
xlcli completions bash    # print bash completion script
```

See [CHEATSHEET.md](CHEATSHEET.md) for keybindings and commands.

## Development

```bash
cargo build               # debug build
cargo build --release     # release build
cargo run -- file.xlsx    # run debug build with a file
```

Workspace layout and conventions in [CLAUDE.md](CLAUDE.md).

## License

GPL-3.0-only. See [LICENSE](LICENSE).
