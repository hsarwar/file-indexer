use std::{
    fs,
    path::{Path, PathBuf},
};

use anyhow::{Context, Result};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub enum SortField {
    Name,
    Modified,
    Size,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub enum SortDirection {
    Asc,
    Desc,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FavoriteSearch {
    pub name: String,
    pub query: String,
    pub sort_field: SortField,
    pub sort_direction: SortDirection,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct AppConfig {
    pub selected_roots: Vec<String>,
    pub favorites: Vec<FavoriteSearch>,
}

impl Default for AppConfig {
    fn default() -> Self {
        Self {
            selected_roots: available_roots(),
            favorites: Vec::new(),
        }
    }
}

impl AppConfig {
    pub fn load(path: &Path) -> Result<Self> {
        if !path.exists() {
            return Ok(Self::default());
        }

        let raw = fs::read_to_string(path)
            .with_context(|| format!("failed to read config at {}", path.display()))?;
        let config = serde_json::from_str(&raw)
            .with_context(|| format!("failed to parse config at {}", path.display()))?;
        Ok(config)
    }

    pub fn save(&self, path: &Path) -> Result<()> {
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent)
                .with_context(|| format!("failed to create {}", parent.display()))?;
        }

        let raw = serde_json::to_string_pretty(self)?;
        fs::write(path, raw).with_context(|| format!("failed to write {}", path.display()))
    }
}

pub fn app_data_dir() -> Result<PathBuf> {
    let base = dirs::data_local_dir().context("failed to resolve local app data directory")?;
    Ok(base.join("file-indexer"))
}

pub fn storage_dir() -> Result<PathBuf> {
    let exe_dir = std::env::current_exe()
        .context("failed to resolve executable path")?
        .parent()
        .context("failed to resolve executable directory")?
        .to_path_buf();

    if is_writable_directory(&exe_dir) {
        return Ok(exe_dir);
    }

    app_data_dir()
}

pub fn config_path() -> Result<PathBuf> {
    Ok(storage_dir()?.join("config.json"))
}

pub fn database_path() -> Result<PathBuf> {
    Ok(storage_dir()?.join("index.sqlite3"))
}

pub fn available_roots() -> Vec<String> {
    let mut roots = Vec::new();
    for letter in b'A'..=b'Z' {
        let candidate = format!("{}:\\", letter as char);
        if Path::new(&candidate).exists() {
            roots.push(candidate);
        }
    }
    roots
}

fn is_writable_directory(path: &Path) -> bool {
    if fs::create_dir_all(path).is_err() {
        return false;
    }

    let probe = path.join(".file-indexer-write-test");
    match fs::write(&probe, b"ok") {
        Ok(()) => {
            let _ = fs::remove_file(probe);
            true
        }
        Err(_) => false,
    }
}
