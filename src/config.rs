use std::{
    fs,
    path::{Path, PathBuf},
};

use anyhow::{Context, Result};
use serde::{Deserialize, Serialize};

pub const DEFAULT_MIN_INDEX_SIZE_BYTES: u64 = 8 * 1024;

const DEFAULT_MEDIA_EXTENSIONS: &[&str] = &[
    "3fr", "3g2", "3gp", "8svx", "aa", "aac", "aax", "ac3", "act", "adt", "adts", "aif", "aifc",
    "aiff", "alac", "amr", "amv", "ape", "apng", "ari", "arw", "asf", "au", "avif", "avi", "avs",
    "bay", "bik", "bik2", "bmp", "caf", "cap", "cda", "cr2", "cr3", "crw", "cue", "dcm", "dcr",
    "dcs", "dff", "divx", "dng", "drf", "dsf", "dts", "dv", "eip", "erf", "evo", "f4a", "f4b",
    "f4p", "f4v", "fff", "fit", "fits", "flac", "fli", "flif", "flv", "gif", "gsm", "h264",
    "h265", "hdr", "heic", "heif", "hevc", "ico", "iiq", "it", "j2c", "j2k", "jfif", "jp2", "jpc",
    "jpe", "jpeg", "jpg", "jpm", "jpx", "jxl", "k25", "kdc", "m1v", "m2p", "m2t", "m2ts", "m2v",
    "m4a", "m4b", "m4p", "m4v", "mef", "mid", "midi", "mjpeg", "mjpg", "mk3d", "mka", "mkv", "mod",
    "mov", "mos", "mp1", "mp2", "mp3", "mp4", "mpa", "mpc", "mpe", "mpeg", "mpg", "mpv", "mrw",
    "mts", "mxf", "nef", "nrw", "nsv", "nut", "oga", "ogg", "ogm", "ogv", "opus", "orf", "pam",
    "pbm", "pcx", "pef", "pfm", "pgm", "png", "pnm", "ppm", "psb", "psd", "ptx", "pxn", "qoi",
    "qt", "ra", "raf", "ram", "raw", "rle", "rm", "rmvb", "roq", "rw2", "rwl", "rwz", "s3m",
    "snd", "sr2", "srf", "srw", "svg", "swf", "tak", "tga", "tif", "tiff", "tod", "ts", "tta",
    "voc", "vob", "vox", "w64", "wav", "wbmp", "webm", "webp", "wma", "wmv", "wv", "x3f", "xm",
    "xpm", "y4m",
];

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
    pub index_all_extensions: bool,
    pub indexed_extensions: String,
    pub min_index_size_bytes: u64,
}

impl Default for AppConfig {
    fn default() -> Self {
        Self {
            selected_roots: available_roots(),
            favorites: Vec::new(),
            index_all_extensions: false,
            indexed_extensions: default_indexed_extensions(),
            min_index_size_bytes: DEFAULT_MIN_INDEX_SIZE_BYTES,
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

pub fn default_indexed_extensions() -> String {
    DEFAULT_MEDIA_EXTENSIONS.join(" ")
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
