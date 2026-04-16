use serde::{Deserialize, Serialize};
use std::path::PathBuf;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct Config {
    pub command_palette: CommandPaletteConfig,
    pub formula_autocomplete: FormulaAutocompleteConfig,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct FormulaAutocompleteConfig {
    pub enabled: bool,
    pub show_description: bool,
    pub show_signature: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct CommandPaletteConfig {
    pub enabled: bool,
    pub position: PalettePosition,
    pub width_percent: u16,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum PalettePosition {
    Top,
    Center,
    Bottom,
}

impl Default for Config {
    fn default() -> Self {
        Self {
            command_palette: CommandPaletteConfig::default(),
            formula_autocomplete: FormulaAutocompleteConfig::default(),
        }
    }
}

impl Default for FormulaAutocompleteConfig {
    fn default() -> Self {
        Self {
            enabled: true,
            show_description: true,
            show_signature: true,
        }
    }
}

impl Default for CommandPaletteConfig {
    fn default() -> Self {
        Self {
            enabled: true,
            position: PalettePosition::Top,
            width_percent: 50,
        }
    }
}

impl Config {
    pub fn load() -> Self {
        let path = config_path();
        if path.exists() {
            match std::fs::read_to_string(&path) {
                Ok(content) => match toml::from_str(&content) {
                    Ok(config) => return config,
                    Err(e) => eprintln!("Config parse error: {}. Using defaults.", e),
                },
                Err(e) => eprintln!("Config read error: {}. Using defaults.", e),
            }
        }
        Config::default()
    }

    pub fn save(&self) -> anyhow::Result<()> {
        let path = config_path();
        if let Some(parent) = path.parent() {
            std::fs::create_dir_all(parent)?;
        }
        let content = toml::to_string_pretty(self)?;
        std::fs::write(&path, content)?;
        Ok(())
    }

    pub fn ensure_default_exists() {
        let path = config_path();
        if !path.exists() {
            let config = Config::default();
            if let Err(e) = config.save() {
                eprintln!("Could not write default config: {}", e);
            }
        }
    }
}

fn config_path() -> PathBuf {
    dirs::config_dir()
        .unwrap_or_else(|| PathBuf::from("."))
        .join("xlcli")
        .join("config.toml")
}
