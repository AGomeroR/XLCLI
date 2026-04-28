PREFIX ?= /usr/local
DESTDIR ?=

BIN_SRC := target/release/xlcli
BIN_DST := $(DESTDIR)$(PREFIX)/bin
APPS_DST := $(DESTDIR)$(PREFIX)/share/applications
ICON_DST := $(DESTDIR)$(PREFIX)/share/icons/hicolor/scalable/apps
BASH_DST := $(DESTDIR)$(PREFIX)/share/bash-completion/completions
ZSH_DST := $(DESTDIR)$(PREFIX)/share/zsh/site-functions
FISH_DST := $(DESTDIR)$(PREFIX)/share/fish/vendor_completions.d
DOC_DST := $(DESTDIR)$(PREFIX)/share/doc/xlcli

.PHONY: build completions install uninstall clean deb

build:
	cargo build --release --locked

completions: build
	mkdir -p assets/completions
	./$(BIN_SRC) completions bash > assets/completions/xlcli.bash
	./$(BIN_SRC) completions zsh  > assets/completions/_xlcli
	./$(BIN_SRC) completions fish > assets/completions/xlcli.fish

install: build completions
	install -Dm755 $(BIN_SRC) $(BIN_DST)/xlcli
	install -Dm644 assets/xlcli.desktop $(APPS_DST)/xlcli.desktop
	install -Dm644 assets/xlcli.svg $(ICON_DST)/xlcli.svg
	install -Dm644 assets/completions/xlcli.bash $(BASH_DST)/xlcli
	install -Dm644 assets/completions/_xlcli $(ZSH_DST)/_xlcli
	install -Dm644 assets/completions/xlcli.fish $(FISH_DST)/xlcli.fish
	install -Dm644 README.md $(DOC_DST)/README.md
	install -Dm644 LICENSE $(DOC_DST)/LICENSE

uninstall:
	rm -f $(BIN_DST)/xlcli
	rm -f $(APPS_DST)/xlcli.desktop
	rm -f $(ICON_DST)/xlcli.svg
	rm -f $(BASH_DST)/xlcli
	rm -f $(ZSH_DST)/_xlcli
	rm -f $(FISH_DST)/xlcli.fish
	rm -rf $(DOC_DST)

deb: completions
	cargo deb -p xlcli-tui --no-build

clean:
	cargo clean
	rm -rf assets/completions
