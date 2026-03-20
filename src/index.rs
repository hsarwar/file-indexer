use std::{
    collections::HashSet,
    path::{Path, PathBuf},
    time::{SystemTime, UNIX_EPOCH},
};

use anyhow::{Context, Result};
use rusqlite::{Connection, params};

use crate::config::{SortDirection, SortField};

#[derive(Debug, Clone)]
pub struct FileRecord {
    pub root: String,
    pub full_path: String,
    pub filename: String,
    pub normalized_filename: String,
    pub extension: String,
    pub modified_unix_secs: i64,
    pub size_bytes: i64,
}

#[derive(Debug, Clone)]
pub struct SearchResult {
    pub full_path: String,
    pub filename: String,
    pub extension: String,
    pub root: String,
    pub size_bytes: i64,
    pub modified_unix_secs: i64,
    pub score: i64,
}

#[derive(Debug, Clone)]
pub struct SearchPage {
    pub results: Vec<SearchResult>,
    pub total_matches: i64,
}

#[derive(Debug, Clone)]
pub struct RootScanInfo {
    pub root_path: String,
    pub last_scan_unix_secs: i64,
    pub file_count: i64,
}

pub struct IndexStore {
    db_path: PathBuf,
}

impl IndexStore {
    pub fn new(db_path: PathBuf) -> Result<Self> {
        let store = Self { db_path };
        store.initialize()?;
        Ok(store)
    }

    pub fn replace_all(&self, roots: &[String], records: &[FileRecord]) -> Result<()> {
        let mut conn = self.open()?;
        let tx = conn.transaction()?;

        tx.execute("DELETE FROM file_trigrams", [])?;
        tx.execute("DELETE FROM files", [])?;
        tx.execute("DELETE FROM roots", [])?;

        {
            let mut root_stmt =
                tx.prepare("INSERT INTO roots (root_path, last_scan_unix_secs) VALUES (?1, ?2)")?;
            let now = now_unix_secs();
            for root in roots {
                root_stmt.execute(params![root, now])?;
            }
        }

        {
            let mut file_stmt = tx.prepare(
                "INSERT INTO files (
                    root_path,
                    full_path,
                    filename,
                    normalized_filename,
                    extension,
                    modified_unix_secs,
                    size_bytes
                ) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7)",
            )?;

            let mut trigram_stmt =
                tx.prepare("INSERT INTO file_trigrams (trigram, file_id) VALUES (?1, ?2)")?;

            for record in records {
                file_stmt.execute(params![
                    record.root,
                    record.full_path,
                    record.filename,
                    record.normalized_filename,
                    record.extension,
                    record.modified_unix_secs,
                    record.size_bytes
                ])?;
                let file_id = tx.last_insert_rowid();

                for trigram in unique_trigrams(&record.normalized_filename) {
                    trigram_stmt.execute(params![trigram, file_id])?;
                }
            }
        }

        tx.commit()?;
        Ok(())
    }

    pub fn search(
        &self,
        query: &str,
        limit: usize,
        offset: usize,
        sort_field: &SortField,
        sort_direction: &SortDirection,
    ) -> Result<SearchPage> {
        let expression = parse_search_expression(query);
        if expression.is_empty() {
            return Ok(SearchPage {
                results: Vec::new(),
                total_matches: 0,
            });
        }

        let conn = self.open()?;
        let order_clause = order_clause(sort_field, sort_direction);
        let where_clause = search_where_clause(&expression);
        let count_sql = format!(
            "SELECT COUNT(*) FROM files WHERE {where_clause}"
        );
        let sql = format!(
            "SELECT
                full_path,
                filename,
                extension,
                root_path,
                size_bytes,
                modified_unix_secs,
                100 AS score
            FROM files
            WHERE {where_clause}
            ORDER BY {order_clause}
            LIMIT ? OFFSET ?"
        );

        let where_values = search_params(&expression);
        let total_matches: i64 = conn.query_row(
            &count_sql,
            rusqlite::params_from_iter(where_values.clone()),
            |row| row.get(0),
        )?;

        let mut values = where_values;
        values.push((limit as i64).into());
        values.push((offset as i64).into());

        let mut stmt = conn.prepare(&sql)?;
        let rows = stmt.query_map(rusqlite::params_from_iter(values), map_search)?;
        let results = rows
            .collect::<rusqlite::Result<Vec<_>>>()
            .map_err(anyhow::Error::from)?;
        Ok(SearchPage {
            results,
            total_matches,
        })
    }

    pub fn export_playlist_paths(
        &self,
        query: &str,
        sort_field: &SortField,
        sort_direction: &SortDirection,
    ) -> Result<Vec<String>> {
        let expression = parse_search_expression(query);
        if expression.is_empty() {
            return Ok(Vec::new());
        }

        let conn = self.open()?;
        let order_clause = order_clause(sort_field, sort_direction);
        let where_clause = search_where_clause(&expression);
        let sql = format!(
            "SELECT full_path, 100 AS score FROM files WHERE {where_clause} ORDER BY {order_clause}"
        );
        let values = search_params(&expression);
        let mut stmt = conn.prepare(&sql)?;
        let rows = stmt.query_map(rusqlite::params_from_iter(values), |row| row.get(0))?;
        rows.collect::<rusqlite::Result<Vec<_>>>()
            .map_err(anyhow::Error::from)
    }

    pub fn total_files(&self) -> Result<i64> {
        let conn = self.open()?;
        let count: i64 = conn.query_row("SELECT COUNT(*) FROM files", [], |row| row.get(0))?;
        Ok(count)
    }

    pub fn last_scan_unix_secs(&self) -> Result<Option<i64>> {
        let conn = self.open()?;
        let value = conn.query_row("SELECT MAX(last_scan_unix_secs) FROM roots", [], |row| {
            row.get::<_, Option<i64>>(0)
        })?;
        Ok(value)
    }

    pub fn root_scan_info(&self) -> Result<Vec<RootScanInfo>> {
        let conn = self.open()?;
        let mut stmt = conn.prepare(
            "SELECT
                r.root_path,
                r.last_scan_unix_secs,
                COUNT(f.id) AS file_count
            FROM roots r
            LEFT JOIN files f ON f.root_path = r.root_path
            GROUP BY r.root_path, r.last_scan_unix_secs
            ORDER BY r.root_path ASC",
        )?;
        let rows = stmt.query_map([], |row| {
            Ok(RootScanInfo {
                root_path: row.get(0)?,
                last_scan_unix_secs: row.get(1)?,
                file_count: row.get(2)?,
            })
        })?;
        rows.collect::<rusqlite::Result<Vec<_>>>()
            .map_err(Into::into)
    }

    fn initialize(&self) -> Result<()> {
        if let Some(parent) = self.db_path.parent() {
            std::fs::create_dir_all(parent)
                .with_context(|| format!("failed to create {}", parent.display()))?;
        }

        let conn = self.open()?;
        conn.execute_batch(
            "PRAGMA journal_mode = WAL;
            PRAGMA synchronous = NORMAL;
            PRAGMA temp_store = MEMORY;
            CREATE TABLE IF NOT EXISTS files (
                id INTEGER PRIMARY KEY,
                root_path TEXT NOT NULL,
                full_path TEXT NOT NULL UNIQUE,
                filename TEXT NOT NULL,
                normalized_filename TEXT NOT NULL,
                extension TEXT NOT NULL,
                modified_unix_secs INTEGER NOT NULL,
                size_bytes INTEGER NOT NULL
            );
            CREATE TABLE IF NOT EXISTS file_trigrams (
                trigram TEXT NOT NULL,
                file_id INTEGER NOT NULL,
                FOREIGN KEY(file_id) REFERENCES files(id) ON DELETE CASCADE
            );
            CREATE INDEX IF NOT EXISTS idx_files_normalized_filename
                ON files(normalized_filename);
            CREATE INDEX IF NOT EXISTS idx_file_trigrams_trigram
                ON file_trigrams(trigram);
            CREATE INDEX IF NOT EXISTS idx_file_trigrams_file_id
                ON file_trigrams(file_id);
            CREATE TABLE IF NOT EXISTS roots (
                root_path TEXT PRIMARY KEY,
                last_scan_unix_secs INTEGER NOT NULL
            );",
        )?;
        Ok(())
    }

    fn open(&self) -> Result<Connection> {
        Connection::open(&self.db_path)
            .with_context(|| format!("failed to open database at {}", self.db_path.display()))
    }
}

fn parse_search_expression(query: &str) -> Vec<Vec<String>> {
    query
        .split("||")
        .map(|or_group| {
            or_group
                .split("&&")
                .map(str::trim)
                .map(normalize)
                .filter(|term| !term.is_empty())
                .collect::<Vec<_>>()
        })
        .filter(|group| !group.is_empty())
        .collect()
}

fn search_where_clause(expression: &[Vec<String>]) -> String {
    expression
        .iter()
        .map(|group| {
            let and_clause = (0..group.len())
                .map(|_| "normalized_filename LIKE ?")
                .collect::<Vec<_>>()
                .join(" AND ");
            format!("({and_clause})")
        })
        .collect::<Vec<_>>()
        .join(" OR ")
}

fn search_params(expression: &[Vec<String>]) -> Vec<rusqlite::types::Value> {
    let mut values = Vec::new();
    for group in expression {
        for term in group {
            values.push(format!("%{term}%").into());
        }
    }
    values
}

fn order_clause(sort_field: &SortField, sort_direction: &SortDirection) -> String {
    let direction = match sort_direction {
        SortDirection::Asc => "ASC",
        SortDirection::Desc => "DESC",
    };

    match sort_field {
        SortField::Name => format!("filename {direction}, score DESC, full_path ASC"),
        SortField::Modified => {
            format!("modified_unix_secs {direction}, filename ASC, full_path ASC")
        }
        SortField::Size => format!("size_bytes {direction}, filename ASC, full_path ASC"),
    }
}

fn map_search(row: &rusqlite::Row<'_>) -> rusqlite::Result<SearchResult> {
    Ok(SearchResult {
        full_path: row.get(0)?,
        filename: row.get(1)?,
        extension: row.get(2)?,
        root: row.get(3)?,
        size_bytes: row.get(4)?,
        modified_unix_secs: row.get(5)?,
        score: row.get(6)?,
    })
}

pub fn build_record(
    root: &str,
    full_path: &Path,
    metadata: &std::fs::Metadata,
) -> Option<FileRecord> {
    let filename = full_path.file_name()?.to_string_lossy().to_string();
    let normalized_filename = normalize(&filename);
    let extension = full_path
        .extension()
        .map(|value| value.to_string_lossy().to_string())
        .unwrap_or_default();
    let modified_unix_secs = metadata
        .modified()
        .ok()
        .and_then(system_time_to_unix_secs)
        .unwrap_or_default();

    Some(FileRecord {
        root: root.to_string(),
        full_path: full_path.to_string_lossy().to_string(),
        filename,
        normalized_filename,
        extension,
        modified_unix_secs,
        size_bytes: metadata.len().try_into().ok()?,
    })
}

pub fn normalize(value: &str) -> String {
    value.to_lowercase()
}

pub fn unique_trigrams(value: &str) -> Vec<String> {
    let chars: Vec<char> = value.chars().collect();
    if chars.len() < 3 {
        return vec![value.to_string()];
    }

    let mut set = HashSet::new();
    for window in chars.windows(3) {
        let trigram: String = window.iter().collect();
        set.insert(trigram);
    }

    let mut items: Vec<_> = set.into_iter().collect();
    items.sort();
    items
}

fn system_time_to_unix_secs(time: SystemTime) -> Option<i64> {
    let duration = time.duration_since(UNIX_EPOCH).ok()?;
    i64::try_from(duration.as_secs()).ok()
}

fn now_unix_secs() -> i64 {
    system_time_to_unix_secs(SystemTime::now()).unwrap_or_default()
}
