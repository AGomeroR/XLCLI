#!/usr/bin/env bash
# Build xlcli .deb package for Ubuntu/Debian.
# Output: target/debian/xlcli_<ver>_<arch>.deb
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

if ! command -v cargo-deb >/dev/null 2>&1; then
  echo "Installing cargo-deb..."
  cargo install cargo-deb --locked
fi

make build
make completions
cargo deb -p xlcli-tui --no-build

echo
echo "Built:"
ls -la target/debian/*.deb
echo
echo "Install with:  sudo apt install ./target/debian/xlcli_*.deb"
