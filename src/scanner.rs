use std::{collections::HashSet, path::PathBuf};

use walkdir::WalkDir;

use crate::index::{FileRecord, build_record};

#[derive(Debug, Clone, Default)]
pub struct ScanStats {
    pub indexed_files: usize,
    pub skipped_entries: usize,
}

#[derive(Debug, Clone)]
pub struct ScanFilter {
    pub extensions: HashSet<String>,
    pub min_size_bytes: u64,
}

pub fn scan_roots<F>(
    roots: &[String],
    filter: &ScanFilter,
    mut progress: F,
) -> (Vec<FileRecord>, ScanStats)
where
    F: FnMut(&str, usize),
{
    let mut records = Vec::new();
    let mut stats = ScanStats::default();

    for root in roots {
        let root_path = PathBuf::from(root);
        for entry in WalkDir::new(&root_path).follow_links(false).into_iter() {
            match entry {
                Ok(entry) => {
                    if !entry.file_type().is_file() {
                        continue;
                    }

                    match entry.metadata() {
                        Ok(metadata) => {
                            if metadata.len() < filter.min_size_bytes {
                                stats.skipped_entries += 1;
                                continue;
                            }

                            let extension = entry
                                .path()
                                .extension()
                                .and_then(|value| value.to_str())
                                .map(|value| value.to_ascii_lowercase())
                                .unwrap_or_default();
                            if !filter.extensions.is_empty() && !filter.extensions.contains(&extension)
                            {
                                stats.skipped_entries += 1;
                                continue;
                            }

                            if let Some(record) = build_record(root, entry.path(), &metadata) {
                                records.push(record);
                                stats.indexed_files += 1;
                                if stats.indexed_files % 1_000 == 0 {
                                    progress(root, stats.indexed_files);
                                }
                            } else {
                                stats.skipped_entries += 1;
                            }
                        }
                        Err(_) => stats.skipped_entries += 1,
                    }
                }
                Err(_) => stats.skipped_entries += 1,
            }
        }
    }

    (records, stats)
}
