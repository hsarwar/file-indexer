#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod config;
mod index;
mod scanner;

use std::{
    collections::{HashMap, hash_map::DefaultHasher},
    env, ffi::c_void, fs,
    hash::{Hash, Hasher},
    io::Cursor,
    mem::size_of,
    os::windows::ffi::OsStrExt,
    process::Command,
    sync::mpsc::{self, Receiver},
    thread,
    time::Instant,
};

use anyhow::Result;
use chrono::{DateTime, Local, Utc};
use config::{
    AppConfig, FavoriteSearch, SortDirection, SortField, available_roots, config_path,
    database_path,
};
use eframe::egui::{self, Color32, CornerRadius, RichText, Stroke, Vec2};
use image::{DynamicImage, GenericImageView, ImageFormat, ImageReader, RgbaImage};
use index::{FfmpegPreviewSettings, IndexStore, RootScanInfo, SearchResult};
use windows::{
    Win32::{
        Foundation::SIZE,
        Graphics::Gdi::{
            BI_RGB, BITMAP, BITMAPINFO, BITMAPINFOHEADER, CreateCompatibleDC, DIB_RGB_COLORS,
            DeleteDC, DeleteObject, GetDIBits, GetObjectW, HBITMAP, ReleaseDC,
        },
        System::Com::{COINIT_APARTMENTTHREADED, CoInitializeEx, CoUninitialize},
        UI::{
            Shell::{
                IShellItemImageFactory, SHCreateItemFromParsingName, SIIGBF_BIGGERSIZEOK,
                SIIGBF_RESIZETOFIT, SIIGBF_THUMBNAILONLY,
            },
        },
    },
    core::PCWSTR,
};
use scanner::ScanStats;

const PAGE_SIZE: usize = 100;
const H1_SIZE: f32 = 28.0;
const H2_SIZE: f32 = 21.0;
const H3_SIZE: f32 = 18.0;
const BODY_SIZE: f32 = 16.0;
const SMALL_SIZE: f32 = 14.0;
const MIN_WINDOW_WIDTH: f32 = 1280.0;
const MIN_WINDOW_HEIGHT: f32 = 800.0;
const DEFAULT_WINDOW_WIDTH: f32 = 1440.0;
const DEFAULT_WINDOW_HEIGHT: f32 = 900.0;
const DEFAULT_FFMPEG_PREVIEW_FRAME_COUNT: usize = 15;
const DEFAULT_FFMPEG_PREVIEW_INTERVAL_SECONDS: u32 = 120;

fn set_icon_pixel(
    rgba: &mut [u8],
    width: usize,
    height: usize,
    x: usize,
    y: usize,
    color: [u8; 4],
) {
    if x >= width || y >= height {
        return;
    }
    let idx = (y * width + x) * 4;
    rgba[idx..idx + 4].copy_from_slice(&color);
}

fn fill_icon_rect(
    rgba: &mut [u8],
    width: usize,
    height: usize,
    x0: usize,
    y0: usize,
    x1: usize,
    y1: usize,
    color: [u8; 4],
) {
    for y in y0.min(height)..y1.min(height) {
        for x in x0.min(width)..x1.min(width) {
            set_icon_pixel(rgba, width, height, x, y, color);
        }
    }
}

fn fill_icon_circle(
    rgba: &mut [u8],
    width: usize,
    height: usize,
    cx: i32,
    cy: i32,
    radius: i32,
    color: [u8; 4],
) {
    let r2 = radius * radius;
    for y in (cy - radius).max(0)..=(cy + radius).min(height as i32 - 1) {
        for x in (cx - radius).max(0)..=(cx + radius).min(width as i32 - 1) {
            let dx = x - cx;
            let dy = y - cy;
            if dx * dx + dy * dy <= r2 {
                set_icon_pixel(rgba, width, height, x as usize, y as usize, color);
            }
        }
    }
}

fn fill_icon_ring(
    rgba: &mut [u8],
    width: usize,
    height: usize,
    cx: i32,
    cy: i32,
    outer_radius: i32,
    inner_radius: i32,
    color: [u8; 4],
) {
    let outer2 = outer_radius * outer_radius;
    let inner2 = inner_radius * inner_radius;
    for y in (cy - outer_radius).max(0)..=(cy + outer_radius).min(height as i32 - 1) {
        for x in (cx - outer_radius).max(0)..=(cx + outer_radius).min(width as i32 - 1) {
            let dx = x - cx;
            let dy = y - cy;
            let dist2 = dx * dx + dy * dy;
            if dist2 <= outer2 && dist2 >= inner2 {
                set_icon_pixel(rgba, width, height, x as usize, y as usize, color);
            }
        }
    }
}

fn draw_icon_handle(
    rgba: &mut [u8],
    width: usize,
    height: usize,
    x: usize,
    y: usize,
    length: usize,
    thickness: usize,
    color: [u8; 4],
) {
    for offset in 0..length {
        fill_icon_rect(
            rgba,
            width,
            height,
            x + offset,
            y + offset,
            x + offset + thickness,
            y + offset + thickness,
            color,
        );
    }
}

fn app_icon_rgba(width: usize, height: usize) -> Vec<u8> {
    let mut rgba = vec![0_u8; width * height * 4];

    let folder = [224, 174, 84, 255];
    let folder_top = [244, 200, 111, 255];
    let folder_shade = [196, 145, 59, 255];
    let glass = [111, 184, 236, 255];
    let glass_center = [226, 244, 255, 255];
    let outline = [86, 112, 133, 255];
    let handle = [111, 184, 236, 255];

    let px = |v: f32, axis: usize| ((v * axis as f32).round() as usize).min(axis);

    fill_icon_rect(
        &mut rgba,
        width,
        height,
        px(0.18, width),
        px(0.22, height),
        px(0.44, width),
        px(0.33, height),
        folder_top,
    );
    fill_icon_rect(
        &mut rgba,
        width,
        height,
        px(0.15, width),
        px(0.29, height),
        px(0.69, width),
        px(0.57, height),
        folder,
    );
    fill_icon_rect(
        &mut rgba,
        width,
        height,
        px(0.15, width),
        px(0.51, height),
        px(0.69, width),
        px(0.57, height),
        folder_shade,
    );
    fill_icon_rect(
        &mut rgba,
        width,
        height,
        px(0.15, width),
        px(0.28, height),
        px(0.69, width),
        px(0.38, height),
        outline,
    );

    let cx = px(0.60, width) as i32;
    let cy = px(0.52, height) as i32;
    let outer = px(0.14, width.min(height)) as i32;
    let inner = px(0.08, width.min(height)) as i32;
    fill_icon_ring(&mut rgba, width, height, cx, cy, outer, inner, glass);
    fill_icon_circle(&mut rgba, width, height, cx, cy, inner - 1, glass_center);
    draw_icon_handle(
        &mut rgba,
        width,
        height,
        px(0.66, width),
        px(0.58, height),
        px(0.13, width),
        px(0.05, width.max(height)).max(2),
        handle,
    );

    rgba
}

fn app_icon() -> egui::IconData {
    let width = 128;
    let height = 128;
    egui::IconData {
        rgba: app_icon_rgba(width, height),
        width: width as u32,
        height: height as u32,
    }
}

fn main() -> Result<()> {
    let viewport = egui::ViewportBuilder::default()
        .with_icon(app_icon())
        .with_min_inner_size([MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT])
        .with_inner_size([DEFAULT_WINDOW_WIDTH, DEFAULT_WINDOW_HEIGHT]);

    let native_options = eframe::NativeOptions {
        viewport,
        ..Default::default()
    };
    eframe::run_native(
        "File Indexer",
        native_options,
        Box::new(|cc| Ok(Box::new(FileIndexerApp::new(cc)?))),
    )
    .map_err(|err| anyhow::anyhow!(err.to_string()))
}

struct FileIndexerApp {
    config: AppConfig,
    available_roots: Vec<String>,
    tabs: Vec<SearchTab>,
    active_tab: usize,
    next_tab_id: usize,
    db: IndexStore,
    status: String,
    scan_state: Option<ScanState>,
    total_files: i64,
    last_scan_label: String,
    db_path_label: String,
    root_scan_info: HashMap<String, RootScanInfo>,
    favorites_popup_open: bool,
    favorites_filter: String,
    drives_popup_open: bool,
    options_popup_open: bool,
    video_preview_backend: VideoPreviewBackend,
    ffmpeg_preview_settings: FfmpegPreviewSettings,
    ffmpeg_thumbnail_count_input: String,
    ffmpeg_interval_seconds_input: String,
    startup_maximize_delay_frames: u8,
    preview: PreviewState,
}

struct SearchTab {
    id: usize,
    title: String,
    query: String,
    results: Vec<SearchResult>,
    page: usize,
    total_matches: i64,
    sort_field: SortField,
    sort_direction: SortDirection,
}

struct ScanState {
    receiver: Receiver<ScanMessage>,
    indexed_files: usize,
}

struct PreviewTexture {
    texture: egui::TextureHandle,
}

struct PreviewFrameBytes {
    bytes: Vec<u8>,
}

struct PreviewLoadResult {
    path: String,
    frames: Vec<PreviewFrameBytes>,
    error: Option<String>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum VideoPreviewBackend {
    Ffmpeg,
    WindowsShell,
}

#[derive(Default)]
struct PreviewState {
    selected_path: Option<String>,
    selected_extension: String,
    rendered_path: Option<String>,
    textures: Vec<PreviewTexture>,
    error: Option<String>,
    loading: bool,
    receiver: Option<Receiver<PreviewLoadResult>>,
}

enum ScanMessage {
    Progress { root: String, indexed_files: usize },
    Completed { stats: ScanStats },
    Failed { message: String },
}

impl FileIndexerApp {
    fn new(cc: &eframe::CreationContext<'_>) -> Result<Self> {
        configure_theme(&cc.egui_ctx);
        let config = AppConfig::load(&config_path()?)?;
        let db_path = database_path()?;
        let db_path_label = db_path.display().to_string();
        let db = IndexStore::new(db_path)?;
        let ffmpeg_preview_settings = db
            .load_ffmpeg_preview_settings()
            .unwrap_or(FfmpegPreviewSettings {
                thumbnail_count: DEFAULT_FFMPEG_PREVIEW_FRAME_COUNT,
                interval_seconds: DEFAULT_FFMPEG_PREVIEW_INTERVAL_SECONDS,
            });
        let total_files = db.total_files().unwrap_or_default();
        let last_scan_label = format_last_scan(db.last_scan_unix_secs().ok().flatten());
        let root_scan_info = map_root_scan_info(db.root_scan_info().unwrap_or_default());

        Ok(Self {
            config,
            available_roots: available_roots(),
            tabs: vec![SearchTab::new(1)],
            active_tab: 0,
            next_tab_id: 2,
            db,
            status: "Ready".to_string(),
            scan_state: None,
            total_files,
            last_scan_label,
            db_path_label,
            root_scan_info,
            favorites_popup_open: false,
            favorites_filter: String::new(),
            drives_popup_open: false,
            options_popup_open: false,
            video_preview_backend: VideoPreviewBackend::WindowsShell,
            ffmpeg_thumbnail_count_input: ffmpeg_preview_settings.thumbnail_count.to_string(),
            ffmpeg_interval_seconds_input: ffmpeg_preview_settings.interval_seconds.to_string(),
            ffmpeg_preview_settings,
            startup_maximize_delay_frames: 8,
            preview: PreviewState::default(),
        })
    }

    fn active_tab(&self) -> &SearchTab {
        &self.tabs[self.active_tab]
    }

    fn active_tab_mut(&mut self) -> &mut SearchTab {
        &mut self.tabs[self.active_tab]
    }

    fn move_result_selection(&mut self, step: isize) {
        let results = self.active_tab().results.clone();
        if results.is_empty() {
            return;
        }

        let selected_index = self
            .preview
            .selected_path
            .as_deref()
            .and_then(|selected_path| {
                results
                    .iter()
                    .position(|result| result.full_path.as_str() == selected_path)
            });

        let next_index = match selected_index {
            Some(index) => {
                let shifted = index as isize + step;
                shifted.clamp(0, results.len().saturating_sub(1) as isize) as usize
            }
            None if step >= 0 => 0,
            None => results.len().saturating_sub(1),
        };

        self.set_preview_target(&results[next_index]);
    }

    fn set_preview_target(&mut self, result: &SearchResult) {
        if self.preview.selected_path.as_deref() == Some(result.full_path.as_str()) {
            return;
        }

        self.preview.selected_path = Some(result.full_path.clone());
        self.preview.selected_extension = result.extension.to_lowercase();
        self.preview.rendered_path = None;
        self.preview.textures.clear();
        self.preview.error = None;
        self.preview.loading = true;

        let (sender, receiver) = mpsc::channel();
        self.preview.receiver = Some(receiver);
        let path = result.full_path.clone();
        let extension = result.extension.to_lowercase();
        let video_preview_backend = self.video_preview_backend;
        let ffmpeg_preview_settings = self.ffmpeg_preview_settings;
        thread::spawn(move || {
            let preview_started_at = Instant::now();
            log_preview_timing(&path, "preview worker start", preview_started_at.elapsed());
            let message = match load_preview_bytes(
                &path,
                &extension,
                video_preview_backend,
                ffmpeg_preview_settings,
            ) {
                Ok(frames) => PreviewLoadResult {
                    path: path.clone(),
                    frames,
                    error: None,
                },
                Err(err) => PreviewLoadResult {
                    path: path.clone(),
                    frames: Vec::new(),
                    error: Some(err.to_string()),
                },
            };
            log_preview_timing(&path, "preview worker complete", preview_started_at.elapsed());
            let _ = sender.send(message);
        });
    }

    fn rerender_selected_preview(&mut self) {
        let Some(path) = self.preview.selected_path.clone() else {
            return;
        };
        if self.preview.selected_extension.is_empty() {
            return;
        }

        let preview_result = SearchResult {
            full_path: path.clone(),
            filename: std::path::Path::new(&path)
                .file_name()
                .and_then(|name| name.to_str())
                .unwrap_or(&path)
                .to_string(),
            extension: self.preview.selected_extension.clone(),
            root: String::new(),
            size_bytes: 0,
            modified_unix_secs: 0,
            score: 0,
        };

        self.preview.selected_path = None;
        self.set_preview_target(&preview_result);
    }

    fn poll_preview_loaded(&mut self, ctx: &egui::Context) {
        let Some(receiver) = self.preview.receiver.as_ref() else {
            return;
        };
        let Ok(result) = receiver.try_recv() else {
            return;
        };

        self.preview.receiver = None;
        self.preview.loading = false;

        if self.preview.selected_path.as_deref() != Some(result.path.as_str()) {
            return;
        }

        self.preview.textures.clear();
        self.preview.error = result.error;
        self.preview.rendered_path = Some(result.path.clone());

        if self.preview.error.is_none() {
            let texture_started_at = Instant::now();
            log_preview_timing(
                &result.path,
                "texture upload start",
                texture_started_at.elapsed(),
            );
            for (index, frame) in result.frames.iter().enumerate() {
                match load_preview_texture(
                    ctx,
                    &format!("preview://{}/{}", result.path, index),
                    &frame.bytes,
                ) {
                    Ok(texture) => self.preview.textures.push(PreviewTexture { texture }),
                    Err(err) => {
                        self.preview.error = Some(err.to_string());
                        self.preview.textures.clear();
                        break;
                    }
                }
            }
            log_preview_timing(
                &result.path,
                "texture upload complete",
                texture_started_at.elapsed(),
            );
        }
    }

    fn refresh_active_search(&mut self) {
        let (query, page, sort_field, sort_direction) = {
            let tab = self.active_tab();
            (
                tab.query.trim().to_string(),
                tab.page,
                tab.sort_field.clone(),
                tab.sort_direction.clone(),
            )
        };

        if query.is_empty() {
            let tab = self.active_tab_mut();
            tab.results.clear();
            tab.total_matches = 0;
            tab.title = tab_title(tab.id, &tab.query);
            self.status = "Ready".to_string();
            return;
        }

        match self.db.search(
            &query,
            PAGE_SIZE,
            page * PAGE_SIZE,
            &sort_field,
            &sort_direction,
        ) {
            Ok(page_data) => {
                let tab = self.active_tab_mut();
                tab.results = page_data.results;
                tab.total_matches = page_data.total_matches;
                tab.title = tab_title(tab.id, &tab.query);

                if tab.total_matches == 0 {
                    self.status = "No matches found".to_string();
                } else {
                    self.status = "Ready".to_string();
                }
            }
            Err(err) => {
                let tab = self.active_tab_mut();
                tab.results.clear();
                tab.total_matches = 0;
                self.status = format!("Search failed: {err}");
            }
        }
    }

    fn export_active_search_m3u(&mut self) {
        let (query, sort_field, sort_direction) = {
            let tab = self.active_tab();
            (
                tab.query.trim().to_string(),
                tab.sort_field.clone(),
                tab.sort_direction.clone(),
            )
        };

        if query.is_empty() {
            self.status = "Enter a search query before exporting".to_string();
            return;
        }

        let paths = match self
            .db
            .export_playlist_paths(&query, &sort_field, &sort_direction)
        {
            Ok(paths) => paths,
            Err(err) => {
                self.status = format!("Export failed: {err}");
                return;
            }
        };

        if paths.is_empty() {
            self.status = "No search results to export".to_string();
            return;
        }

        let export_dir = match env::current_dir() {
            Ok(dir) => dir.join("exports"),
            Err(err) => {
                self.status = format!("Export failed: {err}");
                return;
            }
        };
        if let Err(err) = fs::create_dir_all(&export_dir) {
            self.status = format!("Export failed: {err}");
            return;
        }

        let timestamp = Local::now().format("%Y%m%d-%H%M%S");
        let file_name = format!("{}-{timestamp}.m3u", sanitize_export_name(&query));
        let export_path = export_dir.join(file_name);

        let mut playlist = String::from("#EXTM3U\n");
        for path in &paths {
            playlist.push_str(path);
            playlist.push('\n');
        }

        match fs::write(&export_path, playlist) {
            Ok(()) => {
                self.status = format!(
                    "Exported {} paths to {}",
                    paths.len(),
                    export_path.display()
                );
            }
            Err(err) => {
                self.status = format!("Export failed: {err}");
            }
        }
    }

    fn save_config(&mut self) {
        match config_path().and_then(|path| self.config.save(&path)) {
            Ok(()) => {}
            Err(err) => self.status = format!("Failed to save config: {err}"),
        }
    }

    fn commit_ffmpeg_preview_settings(&mut self) {
        let parsed_count = self.ffmpeg_thumbnail_count_input.trim().parse::<usize>();
        let parsed_interval = self.ffmpeg_interval_seconds_input.trim().parse::<u32>();
        let (Ok(thumbnail_count), Ok(interval_seconds)) = (parsed_count, parsed_interval) else {
            return;
        };

        let requested = FfmpegPreviewSettings {
            thumbnail_count,
            interval_seconds,
        };
        let Ok(normalized) = self.db.save_ffmpeg_preview_settings(requested) else {
            self.status = "Failed to save FFmpeg preview settings".to_string();
            return;
        };

        let changed = normalized.thumbnail_count != self.ffmpeg_preview_settings.thumbnail_count
            || normalized.interval_seconds != self.ffmpeg_preview_settings.interval_seconds;
        self.ffmpeg_preview_settings = normalized;
        self.ffmpeg_thumbnail_count_input = normalized.thumbnail_count.to_string();
        self.ffmpeg_interval_seconds_input = normalized.interval_seconds.to_string();

        if changed
            && is_video_extension(&self.preview.selected_extension)
            && self.video_preview_backend == VideoPreviewBackend::Ffmpeg
        {
            self.rerender_selected_preview();
        }
    }

    fn reload_index_stats(&mut self) {
        self.total_files = self.db.total_files().unwrap_or_default();
        self.last_scan_label = format_last_scan(self.db.last_scan_unix_secs().ok().flatten());
        self.root_scan_info = map_root_scan_info(self.db.root_scan_info().unwrap_or_default());
    }

    fn set_active_page(&mut self, page: usize) {
        self.active_tab_mut().page = page;
        self.refresh_active_search();
    }

    fn add_tab(&mut self) {
        let id = self.next_tab_id;
        self.next_tab_id += 1;
        self.tabs.push(SearchTab::new(id));
        self.active_tab = self.tabs.len() - 1;
        self.status = "New search tab created".to_string();
    }

    fn open_favorite(&mut self, favorite: &FavoriteSearch) {
        let id = self.next_tab_id;
        self.next_tab_id += 1;
        self.tabs.push(SearchTab::from_favorite(id, favorite));
        self.active_tab = self.tabs.len() - 1;
        self.refresh_active_search();
        self.status = format!("Opened favorite '{}'", favorite.name);
        self.favorites_popup_open = false;
    }

    fn toggle_star_active_tab(&mut self) {
        let favorite = {
            let tab = self.active_tab();
            if tab.query.trim().is_empty() {
                self.status = "Cannot favorite an empty search".to_string();
                return;
            }

            FavoriteSearch {
                name: tab.query.trim().to_string(),
                query: tab.query.clone(),
                sort_field: tab.sort_field.clone(),
                sort_direction: tab.sort_direction.clone(),
            }
        };

        if let Some(index) = self.config.favorites.iter().position(|item| {
            item.query == favorite.query
                && item.sort_field == favorite.sort_field
                && item.sort_direction == favorite.sort_direction
        }) {
            let removed = self.config.favorites.remove(index);
            self.save_config();
            self.status = format!("Removed favorite '{}'", removed.name);
            return;
        }

        self.config.favorites.push(favorite.clone());
        self.config
            .favorites
            .sort_by(|a, b| a.name.to_lowercase().cmp(&b.name.to_lowercase()));
        self.save_config();
        self.status = format!("Saved favorite '{}'", favorite.name);
    }

    fn active_tab_is_favorite(&self) -> bool {
        let tab = self.active_tab();
        let query = tab.query.trim();
        if query.is_empty() {
            return false;
        }

        self.config.favorites.iter().any(|favorite| {
            favorite.query == tab.query
                && favorite.sort_field == tab.sort_field
                && favorite.sort_direction == tab.sort_direction
        })
    }

    fn close_tab(&mut self, index: usize) {
        if self.tabs.len() == 1 {
            return;
        }

        self.tabs.remove(index);
        if self.active_tab >= self.tabs.len() {
            self.active_tab = self.tabs.len() - 1;
        } else if index < self.active_tab {
            self.active_tab -= 1;
        }
        self.status = "Search tab closed".to_string();
    }

    fn start_scan(&mut self) {
        if self.scan_state.is_some() {
            return;
        }

        let selected_roots = self.config.selected_roots.clone();
        if selected_roots.is_empty() {
            self.status = "Select at least one drive before scanning".to_string();
            return;
        }

        self.save_config();
        let db_path = match database_path() {
            Ok(path) => path,
            Err(err) => {
                self.status = format!("Failed to resolve database path: {err}");
                return;
            }
        };

        let (sender, receiver) = mpsc::channel();
        self.status = "Starting scan...".to_string();
        self.scan_state = Some(ScanState {
            receiver,
            indexed_files: 0,
        });

        thread::spawn(move || {
            let result = (|| -> Result<ScanStats> {
                let (records, stats) =
                    scanner::scan_roots(&selected_roots, |root, indexed_files| {
                        let _ = sender.send(ScanMessage::Progress {
                            root: root.to_string(),
                            indexed_files,
                        });
                    });

                let db = IndexStore::new(db_path)?;
                db.replace_all(&selected_roots, &records)?;
                Ok(stats)
            })();

            match result {
                Ok(stats) => {
                    let _ = sender.send(ScanMessage::Completed { stats });
                }
                Err(err) => {
                    let _ = sender.send(ScanMessage::Failed {
                        message: err.to_string(),
                    });
                }
            }
        });
    }

    fn poll_scan(&mut self) {
        let mut clear_scan = false;
        let mut refresh_search = false;
        let mut reload_index_stats = false;

        if let Some(scan_state) = &mut self.scan_state {
            while let Ok(message) = scan_state.receiver.try_recv() {
                match message {
                    ScanMessage::Progress {
                        root,
                        indexed_files,
                    } => {
                        scan_state.indexed_files = indexed_files;
                        self.status = format!("Scanning {root} ({indexed_files} files indexed)");
                    }
                    ScanMessage::Completed { stats } => {
                        reload_index_stats = true;
                        self.status = format!(
                            "Scan complete: {} files indexed, {} skipped",
                            stats.indexed_files, stats.skipped_entries
                        );
                        refresh_search = true;
                        clear_scan = true;
                    }
                    ScanMessage::Failed { message } => {
                        self.status = format!("Scan failed: {message}");
                        clear_scan = true;
                    }
                }
            }
        }

        if clear_scan {
            self.scan_state = None;
        }

        if reload_index_stats {
            self.reload_index_stats();
        }

        if refresh_search {
            self.refresh_active_search();
        }
    }
}

impl eframe::App for FileIndexerApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        if self.startup_maximize_delay_frames > 0 {
            self.startup_maximize_delay_frames -= 1;
            if self.startup_maximize_delay_frames == 0 {
                ctx.send_viewport_cmd(egui::ViewportCommand::Maximized(true));
            }
        }

        self.poll_scan();
        ctx.request_repaint_after(std::time::Duration::from_millis(250));

        egui::TopBottomPanel::top("top_panel")
            .frame(
                egui::Frame::new()
                    .fill(Color32::from_rgb(18, 24, 30))
                    .inner_margin(egui::Margin::same(0)),
            )
            .show(ctx, |ui| {
                egui::Frame::new()
                    .fill(Color32::from_rgb(18, 24, 30))
                    .inner_margin(egui::Margin::same(12))
                    .show(ui, |ui| {
                        ui.horizontal(|ui| {
                            ui.with_layout(egui::Layout::left_to_right(egui::Align::BOTTOM), |ui| {
                                ui.label(
                                    RichText::new("File Indexer")
                                        .size(H1_SIZE)
                                        .strong()
                                        .color(Color32::from_rgb(242, 245, 247)),
                                );
                                ui.add_space(14.0);
                                ui.label(
                                    RichText::new(format!(
                                        "Indexed: {}  |  Last scan: {}",
                                        format_number(self.total_files),
                                        self.last_scan_label
                                    ))
                                    .size(BODY_SIZE)
                                    .color(Color32::from_rgb(220, 229, 236)),
                                );
                            });

                            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                                egui::Frame::new()
                                    .fill(Color32::from_rgb(24, 32, 40))
                                    .corner_radius(CornerRadius::same(10))
                                    .stroke(Stroke::new(1.0, Color32::from_rgb(50, 64, 76)))
                                    .inner_margin(egui::Margin::same(10))
                                    .show(ui, |ui| {
                                        ui.set_max_width(420.0);
                                        ui.label(
                                            RichText::new("SQLite DB")
                                                .strong()
                                                .color(Color32::from_rgb(220, 232, 241)),
                                        );
                                        ui.small(
                                            RichText::new(self.db_path_label.as_str())
                                                .size(SMALL_SIZE)
                                                .color(Color32::from_rgb(176, 192, 203)),
                                        );
                                    });

                                ui.add_space(10.0);
                                if ui
                                    .add(
                                        egui::Button::new(
                                            RichText::new("Options")
                                                .size(BODY_SIZE)
                                                .strong()
                                                .color(Color32::from_rgb(242, 245, 247)),
                                        )
                                        .min_size(egui::vec2(104.0, 34.0))
                                        .fill(Color32::from_rgb(44, 98, 86))
                                        .stroke(Stroke::new(1.0, Color32::from_rgb(82, 144, 129))),
                                    )
                                    .clicked()
                                {
                                    self.options_popup_open = true;
                                }

                                ui.add_space(10.0);
                                if ui
                                    .add(
                                        egui::Button::new(
                                            RichText::new("Drives")
                                                .size(BODY_SIZE)
                                                .strong()
                                                .color(Color32::from_rgb(242, 245, 247)),
                                        )
                                        .min_size(egui::vec2(104.0, 34.0))
                                        .fill(Color32::from_rgb(57, 83, 106))
                                        .stroke(Stroke::new(1.0, Color32::from_rgb(92, 115, 136))),
                                    )
                                    .clicked()
                                {
                                    self.drives_popup_open = true;
                                }

                            });
                        });
                    });
            });

        self.poll_preview_loaded(ctx);

        egui::SidePanel::right("preview_panel")
            .resizable(true)
            .default_width(395.0)
            .min_width(295.0)
            .frame(
                egui::Frame::new()
                    .fill(Color32::from_rgb(31, 37, 44))
                    .inner_margin(egui::Margin::same(12)),
            )
            .show(ctx, |ui| {
                let previous_backend = self.video_preview_backend;
                ui.horizontal(|ui| {
                    let backend_button = |selected: bool, label: &str| {
                        egui::Button::new(
                            RichText::new(label)
                                .size(SMALL_SIZE - 3.0)
                                .strong()
                                .color(Color32::from_rgb(242, 245, 247)),
                        )
                        .min_size(egui::vec2(68.0, 20.0))
                        .fill(if selected {
                            Color32::from_rgb(57, 83, 106)
                        } else {
                            Color32::from_rgb(27, 39, 49)
                        })
                        .stroke(Stroke::new(
                            1.0,
                            if selected {
                                Color32::from_rgb(132, 168, 198)
                            } else {
                                Color32::from_rgb(56, 78, 92)
                            },
                        ))
                    };

                    ui.label(
                        RichText::new("Preview")
                            .size(H2_SIZE)
                            .strong()
                            .color(Color32::from_rgb(242, 245, 247)),
                    );

                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        if ui
                            .add(backend_button(
                                self.video_preview_backend == VideoPreviewBackend::Ffmpeg,
                                "FFmpeg",
                            ))
                            .clicked()
                        {
                            self.video_preview_backend = VideoPreviewBackend::Ffmpeg;
                        }
                        if ui
                            .add(backend_button(
                                self.video_preview_backend == VideoPreviewBackend::WindowsShell,
                                "Windows",
                            ))
                            .clicked()
                        {
                            self.video_preview_backend = VideoPreviewBackend::WindowsShell;
                        }
                    });
                });
                ui.add_space(6.0);
                let preview_is_video = is_video_extension(&self.preview.selected_extension);
                if preview_is_video && self.video_preview_backend != previous_backend {
                    self.rerender_selected_preview();
                }
                ui.add_space(6.0);

                if self.preview.selected_path.is_some() {
                    ui.add_space(2.0);

                    if self.preview.loading {
                        ui.horizontal(|ui| {
                            ui.spinner();
                            ui.label(
                                RichText::new("Generating preview...")
                                    .size(BODY_SIZE)
                                    .color(Color32::from_rgb(196, 207, 216)),
                            );
                        });
                    } else if let Some(error) = self.preview.error.as_deref() {
                        ui.label(
                            RichText::new(error)
                                .size(BODY_SIZE)
                                .color(Color32::from_rgb(214, 160, 160)),
                        );
                        ui.small(
                            RichText::new(
                                "Video previews use bundled tools from tools/ffmpeg next to the app, or fall back to PATH. Image previews use local Windows decoding.",
                            )
                            .size(SMALL_SIZE)
                            .color(Color32::from_rgb(176, 192, 203)),
                        );
                    } else if self.preview.textures.is_empty() {
                        ui.label(
                            RichText::new("Select an image or video result to render a preview")
                                .size(BODY_SIZE)
                                .color(Color32::from_rgb(196, 207, 216)),
                        );
                    } else {
                        egui::ScrollArea::vertical().show(ui, |ui| {
                            for item in &self.preview.textures {
                                ui.add(
                                    egui::Image::from_texture(&item.texture)
                                        .max_width(ui.available_width())
                                        .corner_radius(CornerRadius::same(8)),
                                );
                                ui.add_space(10.0);
                            }
                        });
                    }
                } else {
                    ui.label(
                        RichText::new(
                            "Click a result card to preview supported image and video files here.",
                        )
                        .size(BODY_SIZE)
                        .color(Color32::from_rgb(196, 207, 216)),
                    );
                }
            });

        egui::CentralPanel::default()
            .frame(
                egui::Frame::new()
                    .fill(Color32::from_rgb(37, 43, 50))
                    .inner_margin(egui::Margin::same(12)),
            )
            .show(ctx, |ui| {
                ui.spacing_mut().item_spacing = Vec2::new(10.0, 10.0);

                ui.horizontal(|ui| {
                    let action_button = |label: &str, fill: Color32| {
                        egui::Button::new(
                            RichText::new(label)
                                .size(BODY_SIZE)
                                .strong()
                                .color(Color32::from_rgb(242, 245, 247)),
                        )
                        .min_size(egui::vec2(112.0, 32.0))
                        .fill(fill)
                        .stroke(Stroke::new(1.0, Color32::from_rgb(92, 115, 136)))
                    };

                    if ui
                        .add(action_button("New Tab", Color32::from_rgb(57, 83, 106)))
                        .clicked()
                    {
                        self.add_tab();
                    }

                    let star_label = if self.active_tab_is_favorite() {
                        "Unstar"
                    } else {
                        "Star"
                    };
                    let star_fill = if self.active_tab_is_favorite() {
                        Color32::from_rgb(173, 109, 28)
                    } else {
                        Color32::from_rgb(38, 122, 184)
                    };
                    if ui.add(action_button(star_label, star_fill)).clicked() {
                        self.toggle_star_active_tab();
                    }

                    if ui
                        .add(action_button("Favorites", Color32::from_rgb(47, 138, 92)))
                        .clicked()
                    {
                        self.favorites_popup_open = true;
                    }

                    let export_enabled = !self.active_tab().query.trim().is_empty();
                    if ui
                        .add_enabled(
                            export_enabled,
                            action_button("Export M3U", Color32::from_rgb(26, 143, 128)),
                        )
                        .clicked()
                    {
                        self.export_active_search_m3u();
                    }
                });

                ui.add_space(2.0);

                let mut close_index = None;
                egui::ScrollArea::horizontal()
                    .id_salt("tabs_scroll")
                    .auto_shrink([false, false])
                    .max_height(58.0)
                    .show(ui, |ui| {
                        ui.spacing_mut().item_spacing = Vec2::new(10.0, 0.0);
                        ui.with_layout(egui::Layout::left_to_right(egui::Align::Center), |ui| {
                            for index in 0..self.tabs.len() {
                                let selected = self.active_tab == index;
                                let tab_title = self.tabs[index].title.clone();
                                let tab_fill = if selected {
                                    Color32::from_rgb(57, 83, 106)
                                } else {
                                    Color32::from_rgb(28, 38, 49)
                                };
                                let tab_stroke = if selected {
                                    Stroke::new(1.0, Color32::from_rgb(132, 168, 198))
                                } else {
                                    Stroke::new(1.0, Color32::from_rgb(50, 70, 84))
                                };

                                egui::Frame::new()
                                    .fill(tab_fill)
                                    .corner_radius(CornerRadius::same(10))
                                    .stroke(tab_stroke)
                                    .inner_margin(egui::Margin::symmetric(8, 5))
                                    .show(ui, |ui| {
                                        ui.horizontal(|ui| {
                                            let tab_button = egui::Button::new(
                                                RichText::new(tab_title)
                                                    .size(BODY_SIZE)
                                                    .strong()
                                                    .color(Color32::from_rgb(236, 241, 244)),
                                            )
                                            .frame(false)
                                            .min_size(egui::vec2(104.0, 24.0));
                                            if ui.add(tab_button).clicked() {
                                                self.active_tab = index;
                                            }
                                            if self.tabs.len() > 1
                                                && ui
                                                    .link(
                                                        RichText::new("x")
                                                            .size(BODY_SIZE - 2.0)
                                                            .strong()
                                                            .color(Color32::from_rgb(224, 96, 96)),
                                                    )
                                                    .clicked()
                                            {
                                                close_index = Some(index);
                                            }
                                        });
                                    });
                            }
                        });
                    });

                if let Some(index) = close_index {
                    self.close_tab(index);
                }

                ui.add_space(2.0);

                let active_index = self.active_tab;
                let mut search_changed = false;

                egui::Frame::new()
                    .fill(Color32::from_rgb(20, 28, 36))
                    .corner_radius(CornerRadius::same(12))
                    .stroke(Stroke::new(1.0, Color32::from_rgb(46, 62, 74)))
                    .inner_margin(egui::Margin::same(12))
                    .show(ui, |ui| {
                        ui.scope(|ui| {
                            let visuals = ui.visuals_mut();
                            visuals.override_text_color = Some(Color32::from_rgb(236, 241, 244));
                            visuals.widgets.inactive.weak_bg_fill =
                                Color32::from_rgb(27, 39, 49);
                            visuals.widgets.inactive.bg_fill = Color32::from_rgb(27, 39, 49);
                            visuals.widgets.inactive.fg_stroke =
                                Stroke::new(1.4, Color32::from_rgb(236, 241, 244));
                            visuals.widgets.hovered.weak_bg_fill =
                                Color32::from_rgb(40, 52, 64);
                            visuals.widgets.hovered.bg_fill = Color32::from_rgb(40, 52, 64);
                            visuals.widgets.hovered.fg_stroke =
                                Stroke::new(1.4, Color32::from_rgb(248, 250, 252));
                            visuals.widgets.active.fg_stroke =
                                Stroke::new(1.4, Color32::from_rgb(250, 251, 252));
                            visuals.widgets.open.fg_stroke =
                                Stroke::new(1.4, Color32::from_rgb(248, 250, 252));

                            let sort_block_width = 324.0;
                            let block_spacing = 16.0;
                            let item_spacing = ui.spacing().item_spacing.x;
                            let search_label_width = 56.0;
                            let search_block_width =
                                (ui.available_width() - sort_block_width - block_spacing).max(220.0);
                            let (control_row_rect, _) = ui.allocate_exact_size(
                                egui::vec2(ui.available_width(), 32.0),
                                egui::Sense::hover(),
                            );

                            let search_label_rect = egui::Rect::from_min_size(
                                control_row_rect.min,
                                egui::vec2(search_label_width, 32.0),
                            );
                            let search_input_rect = egui::Rect::from_min_size(
                                egui::pos2(search_label_rect.max.x + item_spacing, control_row_rect.min.y),
                                egui::vec2(
                                    (search_block_width - search_label_width - item_spacing).max(120.0),
                                    32.0,
                                ),
                            );
                            let sort_label_rect = egui::Rect::from_min_size(
                                egui::pos2(search_input_rect.max.x + block_spacing, control_row_rect.min.y),
                                egui::vec2(44.0, 32.0),
                            );
                            let sort_field_rect = egui::Rect::from_min_size(
                                egui::pos2(sort_label_rect.max.x + item_spacing, control_row_rect.min.y),
                                egui::vec2(132.0, 32.0),
                            );
                            let sort_direction_rect = egui::Rect::from_min_size(
                                egui::pos2(sort_field_rect.max.x + item_spacing, control_row_rect.min.y),
                                egui::vec2(124.0, 32.0),
                            );

                            ui.scope_builder(
                                egui::UiBuilder::new()
                                    .max_rect(search_label_rect)
                                    .layout(egui::Layout::centered_and_justified(
                                        egui::Direction::LeftToRight,
                                    )),
                                |ui| {
                                    ui.label(
                                        RichText::new("Search")
                                            .strong()
                                            .color(Color32::from_rgb(236, 241, 244)),
                                    );
                                },
                            );

                            let response = ui
                                .scope_builder(
                                    egui::UiBuilder::new().max_rect(search_input_rect),
                                    |ui| {
                                        ui.add_sized(
                                            [search_input_rect.width(), 32.0],
                                            egui::TextEdit::singleline(
                                                &mut self.tabs[active_index].query,
                                            )
                                            .vertical_align(egui::Align::Center)
                                            .hint_text("Name or folder search: mp4 && ytd_ || trailer"),
                                        )
                                    },
                                )
                                .inner;
                            if response.changed() {
                                self.tabs[active_index].page = 0;
                                self.tabs[active_index].title = tab_title(
                                    self.tabs[active_index].id,
                                    &self.tabs[active_index].query,
                                );
                                search_changed = true;
                            }

                            let mut sort_field =
                                self.tabs[active_index].sort_field.clone();
                            let mut sort_direction =
                                self.tabs[active_index].sort_direction.clone();

                            ui.scope_builder(
                                egui::UiBuilder::new()
                                    .max_rect(sort_label_rect)
                                    .layout(egui::Layout::centered_and_justified(
                                        egui::Direction::LeftToRight,
                                    )),
                                |ui| {
                                    ui.label(
                                        RichText::new("Sort")
                                            .strong()
                                            .color(Color32::from_rgb(236, 241, 244)),
                                    );
                                },
                            );

                            ui.scope_builder(
                                egui::UiBuilder::new().max_rect(sort_field_rect),
                                |ui| {
                                    egui::ComboBox::from_id_salt("sort_field")
                                        .width(132.0)
                                        .selected_text(
                                            RichText::new(sort_field_label(&sort_field))
                                                .color(Color32::from_rgb(236, 241, 244)),
                                        )
                                        .show_ui(ui, |ui| {
                                            ui.selectable_value(
                                                &mut sort_field,
                                                SortField::Name,
                                                "Name",
                                            );
                                            ui.selectable_value(
                                                &mut sort_field,
                                                SortField::Modified,
                                                "Date",
                                            );
                                            ui.selectable_value(
                                                &mut sort_field,
                                                SortField::Size,
                                                "Size",
                                            );
                                        });
                                },
                            );

                            ui.scope_builder(
                                egui::UiBuilder::new().max_rect(sort_direction_rect),
                                |ui| {
                                    egui::ComboBox::from_id_salt("sort_direction")
                                        .width(124.0)
                                        .selected_text(
                                            RichText::new(sort_direction_label(&sort_direction))
                                                .color(Color32::from_rgb(236, 241, 244)),
                                        )
                                        .show_ui(ui, |ui| {
                                            ui.selectable_value(
                                                &mut sort_direction,
                                                SortDirection::Asc,
                                                "Ascending",
                                            );
                                            ui.selectable_value(
                                                &mut sort_direction,
                                                SortDirection::Desc,
                                                "Descending",
                                            );
                                        });
                                },
                            );

                            if sort_field != self.tabs[active_index].sort_field
                                || sort_direction != self.tabs[active_index].sort_direction
                            {
                                self.tabs[active_index].sort_field = sort_field;
                                self.tabs[active_index].sort_direction = sort_direction;
                                self.tabs[active_index].page = 0;
                                search_changed = true;
                            }
                        });
                    });

                if search_changed {
                    self.refresh_active_search();
                }

                let mut keyboard_selection_changed = false;
                if !ctx.wants_keyboard_input() {
                    let move_down = ctx.input(|input| input.key_pressed(egui::Key::ArrowDown));
                    let move_up = ctx.input(|input| input.key_pressed(egui::Key::ArrowUp));
                    if move_down {
                        self.move_result_selection(1);
                        keyboard_selection_changed = true;
                    } else if move_up {
                        self.move_result_selection(-1);
                        keyboard_selection_changed = true;
                    }
                }

                ui.add_space(8.0);

                let total_matches = self.active_tab().total_matches;
                ui.horizontal(|ui| {
                    ui.label(
                        RichText::new("Results")
                            .strong()
                            .size(H2_SIZE)
                            .color(Color32::from_rgb(242, 245, 247)),
                    );

                    if !self.active_tab().query.trim().is_empty() {
                        let total_pages = page_count(total_matches, PAGE_SIZE);
                        ui.with_layout(
                            egui::Layout::right_to_left(egui::Align::Center),
                            |ui| {
                                let next_enabled = self.active_tab().page + 1 < total_pages;
                                let next_button = egui::Button::new(
                                    RichText::new("Next").color(Color32::from_rgb(220, 229, 236)),
                                )
                                .fill(if next_enabled {
                                    Color32::from_rgb(27, 39, 49)
                                } else {
                                    Color32::from_rgb(46, 54, 62)
                                })
                                .stroke(Stroke::new(1.0, Color32::from_rgb(70, 88, 102)));
                                if ui.add_enabled(next_enabled, next_button).clicked() {
                                    self.set_active_page(self.active_tab().page + 1);
                                }

                                ui.label(
                                    RichText::new(format!("{} total matches", total_matches))
                                        .color(Color32::from_rgb(220, 229, 236)),
                                );
                                ui.label(
                                    RichText::new(format!(
                                        "Page {} of {}",
                                        self.active_tab().page + 1,
                                        total_pages.max(1)
                                    ))
                                    .color(Color32::from_rgb(220, 229, 236)),
                                );

                                let previous_enabled = self.active_tab().page > 0;
                                let previous_button = egui::Button::new(
                                    RichText::new("Previous").color(Color32::from_rgb(220, 229, 236)),
                                )
                                .fill(if previous_enabled {
                                    Color32::from_rgb(27, 39, 49)
                                } else {
                                    Color32::from_rgb(46, 54, 62)
                                })
                                .stroke(Stroke::new(1.0, Color32::from_rgb(70, 88, 102)));
                                if ui.add_enabled(previous_enabled, previous_button).clicked() {
                                    self.set_active_page(self.active_tab().page.saturating_sub(1));
                                }
                            },
                        );
                    }
                });
                ui.separator();

                egui::ScrollArea::vertical().show(ui, |ui| {
                    let mut pending_preview: Option<SearchResult> = None;
                    for result in &self.active_tab().results {
                        let preview_selected = self.preview.selected_path.as_deref()
                            == Some(result.full_path.as_str());
                        let preview_supported = is_image_extension(&result.extension)
                            || is_video_extension(&result.extension);
                        let card_fill = if preview_selected {
                            Color32::from_rgb(29, 41, 52)
                        } else {
                            Color32::from_rgb(22, 31, 40)
                        };
                        let card_stroke = if preview_selected {
                            Stroke::new(1.0, Color32::from_rgb(120, 184, 228))
                        } else {
                            Stroke::new(1.0, Color32::from_rgb(46, 62, 74))
                        };
                        let card = ui.scope_builder(
                            egui::UiBuilder::new().sense(egui::Sense::click()),
                            |ui| {
                                egui::Frame::new()
                                    .fill(card_fill)
                                    .corner_radius(CornerRadius::same(12))
                                    .stroke(card_stroke)
                                    .inner_margin(egui::Margin::same(12))
                                    .show(ui, |ui| {
                                        ui.set_min_width(ui.available_width());
                                        ui.vertical(|ui| {
                                            ui.horizontal(|ui| {
                                                if ui
                                                    .link(
                                                        RichText::new(&result.filename)
                                                            .size(H3_SIZE)
                                                            .strong()
                                                            .color(ui.visuals().hyperlink_color),
                                                    )
                                                    .clicked()
                                                {
                                                    let _ = open_with_registered_app(&result.full_path);
                                                }
                                                if !result.extension.is_empty() {
                                                    ui.label(
                                                        RichText::new(format!(".{}", result.extension))
                                                            .color(Color32::from_rgb(140, 186, 222)),
                                                    );
                                                }
                                            });
                                            ui.small(
                                                RichText::new(result.full_path.as_str())
                                                    .size(SMALL_SIZE)
                                                    .color(Color32::from_rgb(198, 208, 216)),
                                            );
                                            ui.small(
                                                RichText::new(format!(
                                                    "{} | {} bytes | {}",
                                                    result.root,
                                                    format_number(result.size_bytes),
                                                    format_unix_secs(result.modified_unix_secs)
                                                ))
                                                .size(SMALL_SIZE)
                                                .color(Color32::from_rgb(190, 203, 213)),
                                            );
                                            ui.add_space(6.0);
                                            ui.horizontal(|ui| {
                                                if ui
                                                    .add_enabled(preview_supported, egui::Button::new("Preview"))
                                                    .clicked()
                                                {
                                                    pending_preview = Some(result.clone());
                                                }
                                                if ui.button("Open File").clicked() {
                                                    let _ = open_with_registered_app(&result.full_path);
                                                }
                                                if ui.button("Reveal in Explorer").clicked() {
                                                    let _ = open_in_explorer(&result.full_path);
                                                }
                                            });
                                        });
                                    });
                            },
                        );
                        let card_response = card.response;
                        if preview_selected && keyboard_selection_changed {
                            card_response.scroll_to_me(Some(egui::Align::Center));
                        }
                        if card_response.clicked() && preview_supported {
                            pending_preview = Some(result.clone());
                        }
                        ui.add_space(8.0);
                    }

                    if let Some(result) = pending_preview {
                        self.set_preview_target(&result);
                    }

                    if !self.active_tab().query.is_empty() && self.active_tab().results.is_empty() {
                        ui.label("No matches found");
                    }
                });
            });

        if self.drives_popup_open {
            let mut popup_open = self.drives_popup_open;
            egui::Window::new("Drives")
                .open(&mut popup_open)
                .collapsible(false)
                .resizable(true)
                .default_width(460.0)
                .frame(
                    egui::Frame::new()
                        .fill(Color32::from_rgb(24, 32, 40))
                        .corner_radius(CornerRadius::same(12))
                        .stroke(Stroke::new(1.0, Color32::from_rgb(50, 64, 76)))
                        .inner_margin(egui::Margin::same(12)),
                )
                .show(ctx, |ui| {
                    ui.scope(|ui| {
                        let visuals = ui.visuals_mut();
                        visuals.override_text_color = Some(Color32::from_rgb(236, 241, 244));
                        visuals.widgets.inactive.weak_bg_fill = Color32::from_rgb(31, 40, 48);
                        visuals.widgets.inactive.bg_fill = Color32::from_rgb(31, 40, 48);
                        visuals.widgets.inactive.bg_stroke =
                            Stroke::new(1.0, Color32::from_rgb(70, 88, 102));
                        visuals.widgets.hovered.weak_bg_fill = Color32::from_rgb(42, 54, 66);
                        visuals.widgets.hovered.bg_fill = Color32::from_rgb(42, 54, 66);
                        visuals.widgets.hovered.bg_stroke =
                            Stroke::new(1.0, Color32::from_rgb(106, 128, 148));
                        visuals.widgets.active.weak_bg_fill = Color32::from_rgb(57, 83, 106);
                        visuals.widgets.active.bg_fill = Color32::from_rgb(57, 83, 106);
                        visuals.widgets.active.bg_stroke =
                            Stroke::new(1.0, Color32::from_rgb(132, 168, 198));

                        ui.label(
                            RichText::new("Drives")
                                .strong()
                                .size(H2_SIZE)
                                .color(Color32::from_rgb(242, 245, 247)),
                        );
                        ui.add_space(4.0);
                        ui.label(
                            RichText::new("Choose which drives are indexed, then rebuild when needed.")
                                .size(BODY_SIZE)
                                .color(Color32::from_rgb(220, 229, 236)),
                        );
                        ui.add_space(8.0);

                        egui::ScrollArea::vertical()
                            .max_height(320.0)
                            .show(ui, |ui| {
                                let available_roots = self.available_roots.clone();
                                for root in &available_roots {
                                    let root_info = self.root_scan_info.get(root);
                                    let scan_label = scan_label_for_root(root_info);
                                    let file_count_label = file_count_label_for_root(root_info);

                                    egui::Frame::new()
                                        .fill(Color32::from_rgb(23, 31, 39))
                                        .corner_radius(CornerRadius::same(10))
                                        .stroke(Stroke::new(1.0, Color32::from_rgb(46, 62, 74)))
                                        .inner_margin(egui::Margin::same(10))
                                        .show(ui, |ui| {
                                            ui.set_width(ui.available_width());
                                            let mut selected = self.config.selected_roots.contains(root);
                                            if ui
                                                .checkbox(
                                                    &mut selected,
                                                    RichText::new(root.as_str())
                                                        .color(Color32::from_rgb(232, 239, 244)),
                                                )
                                                .changed()
                                            {
                                                if selected {
                                                    if !self.config.selected_roots.contains(root) {
                                                        self.config.selected_roots.push(root.clone());
                                                        self.config.selected_roots.sort();
                                                    }
                                                } else {
                                                    self.config
                                                        .selected_roots
                                                        .retain(|value| value != root);
                                                }
                                                self.save_config();
                                            }

                                            ui.small(
                                                RichText::new(scan_label.clone())
                                                    .size(SMALL_SIZE)
                                                    .color(Color32::from_rgb(190, 203, 213)),
                                            );
                                            ui.small(
                                                RichText::new(file_count_label.clone())
                                                    .size(SMALL_SIZE)
                                                    .color(Color32::from_rgb(190, 203, 213)),
                                            );
                                        });
                                    ui.add_space(6.0);
                                }
                            });

                        ui.add_space(8.0);
                        let button_label = if self.scan_state.is_some() {
                            "Scanning..."
                        } else if self.total_files > 0 {
                            "Rebuild Index"
                        } else {
                            "Build Index"
                        };
                        if ui
                            .add_enabled(
                                self.scan_state.is_none(),
                                egui::Button::new(
                                    RichText::new(button_label)
                                        .size(BODY_SIZE)
                                        .strong()
                                        .color(Color32::from_rgb(242, 245, 247)),
                                )
                                .min_size(egui::vec2(140.0, 34.0))
                                .fill(Color32::from_rgb(57, 83, 106))
                                .stroke(Stroke::new(1.0, Color32::from_rgb(92, 115, 136))),
                            )
                            .clicked()
                        {
                            self.start_scan();
                        }
                    });
                });
            self.drives_popup_open = popup_open;
        }

        if self.options_popup_open {
            let mut popup_open = self.options_popup_open;
            egui::Window::new("Options")
                .open(&mut popup_open)
                .collapsible(false)
                .resizable(false)
                .default_width(420.0)
                .frame(
                    egui::Frame::new()
                        .fill(Color32::from_rgb(24, 32, 40))
                        .corner_radius(CornerRadius::same(12))
                        .stroke(Stroke::new(1.0, Color32::from_rgb(50, 64, 76)))
                        .inner_margin(egui::Margin::same(12)),
                )
                .show(ctx, |ui| {
                    ui.scope(|ui| {
                        let visuals = ui.visuals_mut();
                        visuals.override_text_color = Some(Color32::from_rgb(236, 241, 244));
                        visuals.widgets.inactive.weak_bg_fill = Color32::from_rgb(31, 40, 48);
                        visuals.widgets.inactive.bg_fill = Color32::from_rgb(31, 40, 48);
                        visuals.widgets.inactive.bg_stroke =
                            Stroke::new(1.0, Color32::from_rgb(70, 88, 102));
                        visuals.widgets.hovered.weak_bg_fill = Color32::from_rgb(42, 54, 66);
                        visuals.widgets.hovered.bg_fill = Color32::from_rgb(42, 54, 66);
                        visuals.widgets.hovered.bg_stroke =
                            Stroke::new(1.0, Color32::from_rgb(106, 128, 148));
                        visuals.widgets.active.weak_bg_fill = Color32::from_rgb(57, 83, 106);
                        visuals.widgets.active.bg_fill = Color32::from_rgb(57, 83, 106);
                        visuals.widgets.active.bg_stroke =
                            Stroke::new(1.0, Color32::from_rgb(132, 168, 198));

                        ui.label(
                            RichText::new("Options")
                                .strong()
                                .size(H2_SIZE)
                                .color(Color32::from_rgb(242, 245, 247)),
                        );
                        ui.add_space(4.0);
                        ui.label(
                            RichText::new("FFmpeg preview extraction")
                                .size(BODY_SIZE)
                                .color(Color32::from_rgb(220, 229, 236)),
                        );
                        ui.add_space(10.0);

                        let mut thumbs_response = None;
                        let mut interval_response = None;
                        egui::Grid::new("options_ffmpeg_grid")
                            .num_columns(2)
                            .spacing(egui::vec2(12.0, 8.0))
                            .show(ui, |ui| {
                                ui.label(
                                    RichText::new("Max images")
                                        .size(SMALL_SIZE - 1.0)
                                        .color(Color32::from_rgb(176, 192, 203)),
                                );
                                thumbs_response = Some(ui.add_sized(
                                    [56.0, 20.0],
                                    egui::TextEdit::singleline(
                                        &mut self.ffmpeg_thumbnail_count_input,
                                    ),
                                ));
                                ui.end_row();

                                ui.label(
                                    RichText::new("Seconds between images")
                                        .size(SMALL_SIZE - 1.0)
                                        .color(Color32::from_rgb(176, 192, 203)),
                                );
                                interval_response = Some(ui.add_sized(
                                    [64.0, 20.0],
                                    egui::TextEdit::singleline(
                                        &mut self.ffmpeg_interval_seconds_input,
                                    ),
                                ));
                                ui.end_row();
                            });

                        if let (Some(thumbs_response), Some(interval_response)) =
                            (thumbs_response, interval_response)
                        {
                            let enter_pressed =
                                ui.input(|input| input.key_pressed(egui::Key::Enter));
                            let commit_requested = thumbs_response.lost_focus()
                                || interval_response.lost_focus()
                                || (enter_pressed
                                    && (thumbs_response.has_focus()
                                        || interval_response.has_focus()));
                            if commit_requested {
                                self.commit_ffmpeg_preview_settings();
                            }
                        }
                    });
                });
            self.options_popup_open = popup_open;
        }

        if self.favorites_popup_open {
            let mut popup_open = self.favorites_popup_open;
            let filter = self.favorites_filter.to_lowercase();
            let filtered_favorites: Vec<FavoriteSearch> = self
                .config
                .favorites
                .iter()
                .filter(|favorite| {
                    filter.is_empty() || favorite.name.to_lowercase().contains(&filter)
                })
                .cloned()
                .collect();
            let mut favorite_to_open: Option<FavoriteSearch> = None;

            egui::Window::new("Favorite Searches")
                .open(&mut popup_open)
                .collapsible(false)
                .resizable(true)
                .default_width(420.0)
                .frame(
                    egui::Frame::new()
                        .fill(Color32::from_rgb(24, 32, 40))
                        .corner_radius(CornerRadius::same(12))
                        .stroke(Stroke::new(1.0, Color32::from_rgb(50, 64, 76)))
                        .inner_margin(egui::Margin::same(12)),
                )
                .show(ctx, |ui| {
                    ui.scope(|ui| {
                        let visuals = ui.visuals_mut();
                        visuals.override_text_color = Some(Color32::from_rgb(236, 241, 244));
                        visuals.widgets.inactive.weak_bg_fill = Color32::from_rgb(31, 40, 48);
                        visuals.widgets.inactive.bg_fill = Color32::from_rgb(31, 40, 48);
                        visuals.widgets.inactive.bg_stroke =
                            Stroke::new(1.0, Color32::from_rgb(70, 88, 102));
                        visuals.widgets.inactive.fg_stroke =
                            Stroke::new(1.4, Color32::from_rgb(236, 241, 244));
                        visuals.widgets.hovered.weak_bg_fill = Color32::from_rgb(42, 54, 66);
                        visuals.widgets.hovered.bg_fill = Color32::from_rgb(42, 54, 66);
                        visuals.widgets.hovered.bg_stroke =
                            Stroke::new(1.0, Color32::from_rgb(106, 128, 148));
                        visuals.widgets.hovered.fg_stroke =
                            Stroke::new(1.4, Color32::from_rgb(248, 250, 252));
                        visuals.widgets.active.weak_bg_fill = Color32::from_rgb(57, 83, 106);
                        visuals.widgets.active.bg_fill = Color32::from_rgb(57, 83, 106);
                        visuals.widgets.active.bg_stroke =
                            Stroke::new(1.0, Color32::from_rgb(132, 168, 198));
                        visuals.widgets.active.fg_stroke =
                            Stroke::new(1.4, Color32::from_rgb(250, 251, 252));
                        visuals.widgets.open.weak_bg_fill = Color32::from_rgb(42, 54, 66);
                        visuals.widgets.open.bg_fill = Color32::from_rgb(42, 54, 66);
                        visuals.widgets.open.bg_stroke =
                            Stroke::new(1.0, Color32::from_rgb(106, 128, 148));
                        visuals.widgets.open.fg_stroke =
                            Stroke::new(1.4, Color32::from_rgb(248, 250, 252));
                        visuals.extreme_bg_color = Color32::from_rgb(20, 27, 34);
                        visuals.faint_bg_color = Color32::from_rgb(31, 40, 48);

                        ui.label(
                            RichText::new("Favorite Searches")
                                .strong()
                                .size(H2_SIZE)
                                .color(Color32::from_rgb(242, 245, 247)),
                        );
                        ui.add_space(4.0);
                        ui.label(
                            RichText::new("Filter favorites by name:")
                                .size(BODY_SIZE)
                                .color(Color32::from_rgb(220, 229, 236)),
                        );
                        ui.add(
                            egui::TextEdit::singleline(&mut self.favorites_filter)
                                .hint_text("Start typing a favorite name..."),
                        );
                        ui.add_space(8.0);

                        if filtered_favorites.is_empty() {
                            ui.label(
                                RichText::new("No favorites match the current filter")
                                    .size(BODY_SIZE)
                                    .color(Color32::from_rgb(196, 207, 216)),
                            );
                        } else {
                            egui::ScrollArea::vertical()
                                .max_height(320.0)
                                .show(ui, |ui| {
                                    for favorite in &filtered_favorites {
                                        let label = format!(
                                            "{} [{} {}]",
                                            favorite.name,
                                            sort_field_label(&favorite.sort_field),
                                            sort_direction_label(&favorite.sort_direction)
                                        );
                                        let (rect, response) = ui.allocate_exact_size(
                                            egui::vec2(ui.available_width(), 34.0),
                                            egui::Sense::click(),
                                        );
                                        let fill = if response.hovered() {
                                            Color32::from_rgb(42, 54, 66)
                                        } else {
                                            Color32::from_rgb(31, 40, 48)
                                        };
                                        ui.painter().rect(
                                            rect,
                                            CornerRadius::same(6),
                                            fill,
                                            Stroke::new(1.0, Color32::from_rgb(70, 88, 102)),
                                            egui::StrokeKind::Outside,
                                        );
                                        ui.painter().text(
                                            rect.center(),
                                            egui::Align2::CENTER_CENTER,
                                            label,
                                            egui::FontId::new(
                                                BODY_SIZE,
                                                egui::FontFamily::Proportional,
                                            ),
                                            Color32::from_rgb(236, 241, 244),
                                        );
                                        if response.clicked() {
                                            favorite_to_open = Some(favorite.clone());
                                        }
                                    }
                                });
                        }
                    });
                });

            self.favorites_popup_open = popup_open;
            if let Some(favorite) = favorite_to_open {
                self.open_favorite(&favorite);
            }
        }
    }
}

impl SearchTab {
    fn new(id: usize) -> Self {
        Self {
            id,
            title: tab_title(id, ""),
            query: String::new(),
            results: Vec::new(),
            page: 0,
            total_matches: 0,
            sort_field: SortField::Name,
            sort_direction: SortDirection::Asc,
        }
    }

    fn from_favorite(id: usize, favorite: &FavoriteSearch) -> Self {
        Self {
            id,
            title: tab_title(id, &favorite.query),
            query: favorite.query.clone(),
            results: Vec::new(),
            page: 0,
            total_matches: 0,
            sort_field: favorite.sort_field.clone(),
            sort_direction: favorite.sort_direction.clone(),
        }
    }
}

fn sanitize_export_name(value: &str) -> String {
    let sanitized: String = value
        .chars()
        .map(|ch| {
            if ch.is_ascii_alphanumeric() {
                ch
            } else {
                '-'
            }
        })
        .collect();
    sanitized
        .trim_matches('-')
        .chars()
        .take(48)
        .collect::<String>()
        .if_empty_then("search-results")
}

trait IfEmptyThen {
    fn if_empty_then(self, fallback: &str) -> String;
}

impl IfEmptyThen for String {
    fn if_empty_then(self, fallback: &str) -> String {
        if self.is_empty() {
            fallback.to_string()
        } else {
            self
        }
    }
}

fn tab_title(id: usize, query: &str) -> String {
    let trimmed = query.trim();
    if trimmed.is_empty() {
        format!("Search {id}")
    } else {
        trimmed.chars().take(18).collect()
    }
}

fn sort_field_label(value: &SortField) -> &'static str {
    match value {
        SortField::Name => "Name",
        SortField::Modified => "Date",
        SortField::Size => "Size",
    }
}

fn sort_direction_label(value: &SortDirection) -> &'static str {
    match value {
        SortDirection::Asc => "Asc",
        SortDirection::Desc => "Desc",
    }
}

fn open_in_explorer(path: &str) -> Result<()> {
    Command::new("explorer.exe")
        .arg("/select,")
        .arg(path)
        .spawn()?;
    Ok(())
}

fn open_with_registered_app(path: &str) -> Result<()> {
    Command::new("rundll32.exe")
        .arg("url.dll,FileProtocolHandler")
        .arg(path)
        .spawn()?;
    Ok(())
}

fn format_last_scan(value: Option<i64>) -> String {
    value
        .map(format_unix_secs)
        .unwrap_or_else(|| "Never".to_string())
}

fn format_unix_secs(value: i64) -> String {
    if value <= 0 {
        return "Unknown".to_string();
    }

    let Some(utc_time) = DateTime::<Utc>::from_timestamp(value, 0) else {
        return "Unknown".to_string();
    };
    let local_time: DateTime<Local> = utc_time.with_timezone(&Local);
    local_time.format("%Y-%m-%d %I:%M %p").to_string()
}

fn page_count(total_matches: i64, page_size: usize) -> usize {
    if total_matches <= 0 {
        return 0;
    }

    (total_matches as usize).div_ceil(page_size)
}

fn map_root_scan_info(items: Vec<RootScanInfo>) -> HashMap<String, RootScanInfo> {
    items
        .into_iter()
        .map(|item| (item.root_path.clone(), item))
        .collect()
}

fn scan_label_for_root(info: Option<&RootScanInfo>) -> String {
    info.map(|value| format!("Last scan: {}", format_unix_secs(value.last_scan_unix_secs)))
        .unwrap_or_else(|| "Last scan: Never".to_string())
}

fn file_count_label_for_root(info: Option<&RootScanInfo>) -> String {
    info.map(|value| format!("Indexed files: {}", format_number(value.file_count)))
        .unwrap_or_else(|| "Indexed files: 0".to_string())
}

fn is_image_extension(extension: &str) -> bool {
    matches!(
        extension,
        "png" | "jpg" | "jpeg" | "bmp" | "gif" | "webp" | "tif" | "tiff"
    )
}

fn is_video_extension(extension: &str) -> bool {
    matches!(
        extension,
        "mp4" | "mkv" | "avi" | "mov" | "webm" | "wmv" | "m4v" | "mpg" | "mpeg"
    )
}

fn load_preview_bytes(
    path: &str,
    extension: &str,
    video_preview_backend: VideoPreviewBackend,
    ffmpeg_preview_settings: FfmpegPreviewSettings,
) -> Result<Vec<PreviewFrameBytes>> {
    if is_image_extension(extension) {
        let image = render_image_preview(path, extension)?;
        return Ok(vec![PreviewFrameBytes { bytes: image }]);
    }

    if is_video_extension(extension) {
        let frames = render_video_previews(path, video_preview_backend, ffmpeg_preview_settings)?;
        return Ok(frames);
    }

    anyhow::bail!("Preview is only available for common image and video files")
}

fn render_image_preview(path: &str, extension: &str) -> Result<Vec<u8>> {
    if extension == "bmp" {
        return fs::read(path).map_err(Into::into);
    }

    let mut image = ImageReader::open(path)?.decode()?;
    let (width, height) = image.dimensions();
    if width > 1600 || height > 1600 {
        image = DynamicImage::ImageRgba8(image.thumbnail(1600, 1600).to_rgba8());
    }

    let mut bytes = Vec::new();
    image.write_to(&mut Cursor::new(&mut bytes), ImageFormat::Bmp)?;
    Ok(bytes)
}

fn render_video_previews(
    path: &str,
    backend: VideoPreviewBackend,
    ffmpeg_preview_settings: FfmpegPreviewSettings,
) -> Result<Vec<PreviewFrameBytes>> {
    match backend {
        VideoPreviewBackend::WindowsShell => {
            let bytes = render_windows_shell_video_preview(path)?;
            Ok(vec![PreviewFrameBytes { bytes }])
        }
        VideoPreviewBackend::Ffmpeg => {
            let ffmpeg_path = find_media_tool("ffmpeg")?;
            let output_dir = preview_temp_dir(path, "video");
            if output_dir.exists() {
                let _ = fs::remove_dir_all(&output_dir);
            }
            fs::create_dir_all(&output_dir)?;

            run_ffmpeg_preview_frames(&ffmpeg_path, path, &output_dir, ffmpeg_preview_settings)?;

            let read_started_at = Instant::now();
            log_preview_timing(path, "ffmpeg frame read start", read_started_at.elapsed());
            let mut frame_paths = fs::read_dir(&output_dir)?
                .filter_map(|entry| entry.ok().map(|value| value.path()))
                .filter(|frame_path| {
                    frame_path
                        .extension()
                        .and_then(|ext| ext.to_str())
                        .map(|ext| ext.eq_ignore_ascii_case("bmp"))
                        .unwrap_or(false)
                })
                .collect::<Vec<_>>();
            frame_paths.sort();

            let frames = frame_paths
                .into_iter()
                .map(|frame_path| fs::read(frame_path).map(|bytes| PreviewFrameBytes { bytes }))
                .collect::<std::io::Result<Vec<_>>>()?;
            log_preview_timing(path, "ffmpeg frame read complete", read_started_at.elapsed());

            if frames.is_empty() {
                anyhow::bail!("ffmpeg did not render any preview frames for {path}");
            }

            Ok(frames)
        }
    }
}

fn preview_temp_dir(path: &str, kind: &str) -> std::path::PathBuf {
    let mut hasher = DefaultHasher::new();
    path.hash(&mut hasher);
    let hash = hasher.finish();
    env::temp_dir().join("file_indexer_previews").join(format!("{kind}_{hash}"))
}

fn log_preview_timing(path: &str, stage: &str, elapsed: std::time::Duration) {
    let timestamp = Local::now().format("%Y-%m-%d %H:%M:%S%.3f");
    eprintln!(
        "[{timestamp}] [preview] {stage} | elapsed={}ms | path={}",
        elapsed.as_millis(),
        path
    );
}

fn media_tool_candidates(name: &str) -> Vec<std::path::PathBuf> {
    let exe_name = format!("{name}.exe");
    let mut candidates = Vec::new();

    if let Ok(current_exe) = env::current_exe() {
        if let Some(exe_dir) = current_exe.parent() {
            candidates.push(exe_dir.join(&exe_name));
            candidates.push(exe_dir.join("tools").join("ffmpeg").join(&exe_name));
            candidates.push(exe_dir.join("ffmpeg").join(&exe_name));
            if let Some(parent_dir) = exe_dir.parent() {
                candidates.push(parent_dir.join("tools").join("ffmpeg").join(&exe_name));
            }
        }
    }

    if let Ok(cwd) = env::current_dir() {
        candidates.push(cwd.join(&exe_name));
        candidates.push(cwd.join("tools").join("ffmpeg").join(&exe_name));
        candidates.push(cwd.join("ffmpeg").join(&exe_name));
    }

    candidates
}

fn find_media_tool(name: &str) -> Result<std::path::PathBuf> {
    for candidate in media_tool_candidates(name) {
        if candidate.is_file() {
            return Ok(candidate);
        }
    }

    Ok(std::path::PathBuf::from(name))
}

fn run_ffmpeg_preview_frames(
    ffmpeg_path: &std::path::Path,
    path: &str,
    output_dir: &std::path::Path,
    settings: FfmpegPreviewSettings,
) -> Result<()> {
    let ffmpeg_started_at = Instant::now();
    log_preview_timing(path, "ffmpeg extraction start", ffmpeg_started_at.elapsed());

    let mut command = Command::new(ffmpeg_path);
    command.arg("-y").arg("-loglevel").arg("error");

    for index in 0..settings.thumbnail_count {
        let seek_secs = index as u32 * settings.interval_seconds;
        let output_path = output_dir.join(format!("frame_{:02}.bmp", index + 1));
        command
            .arg("-ss")
            .arg(seek_secs.to_string())
            .arg("-i")
            .arg(path)
            .arg("-map")
            .arg(format!("{index}:v:0"))
            .arg("-an")
            .arg("-frames:v")
            .arg("1")
            .arg("-vf")
            .arg("scale=360:-1:force_original_aspect_ratio=decrease")
            .arg(output_path);
    }

    let status = command.status();
    log_preview_timing(path, "ffmpeg extraction complete", ffmpeg_started_at.elapsed());

    match status {
        Ok(status) if status.success() => Ok(()),
        Ok(_) => anyhow::bail!("ffmpeg failed to render preview for {path}"),
        Err(err) if err.kind() == std::io::ErrorKind::NotFound => {
            anyhow::bail!("ffmpeg was not found. Put ffmpeg.exe in tools/ffmpeg next to the app or on PATH")
        }
        Err(err) => Err(err.into()),
    }
}

fn load_preview_texture(ctx: &egui::Context, uri: &str, bytes: &[u8]) -> Result<egui::TextureHandle> {
    let image = image::load_from_memory(bytes)?.to_rgba8();
    let (width, height) = image.dimensions();
    let color_image = egui::ColorImage::from_rgba_unmultiplied(
        [width as usize, height as usize],
        image.as_raw(),
    );
    Ok(ctx.load_texture(uri.to_string(), color_image, Default::default()))
}

fn render_windows_shell_video_preview(path: &str) -> Result<Vec<u8>> {
    let mut wide: Vec<u16> = std::ffi::OsStr::new(path).encode_wide().collect();
    wide.push(0);

    unsafe {
        let com_init = CoInitializeEx(None, COINIT_APARTMENTTHREADED).is_ok();
        let shell_item: IShellItemImageFactory =
            SHCreateItemFromParsingName(PCWSTR(wide.as_ptr()), None)?;
        let bitmap: HBITMAP = shell_item.GetImage(
            SIZE { cx: 360, cy: 240 },
            SIIGBF_THUMBNAILONLY | SIIGBF_BIGGERSIZEOK | SIIGBF_RESIZETOFIT,
        )?;

        let result = hbitmap_to_bmp_bytes(bitmap);
        let _ = DeleteObject(bitmap.into());
        if com_init {
            CoUninitialize();
        }
        result
    }
}

fn hbitmap_to_bmp_bytes(bitmap: HBITMAP) -> Result<Vec<u8>> {
    unsafe {
        let mut info = BITMAP::default();
        let object_size = i32::try_from(size_of::<BITMAP>())?;
        if GetObjectW(bitmap.into(), object_size, Some((&mut info as *mut BITMAP).cast::<c_void>())) == 0 {
            anyhow::bail!("GetObjectW failed for Windows preview bitmap")
        }

        let width = info.bmWidth;
        let height = info.bmHeight;
        if width <= 0 || height <= 0 {
            anyhow::bail!("Windows preview returned invalid bitmap dimensions")
        }

        let mut bitmap_info = BITMAPINFO::default();
        bitmap_info.bmiHeader = BITMAPINFOHEADER {
            biSize: u32::try_from(size_of::<BITMAPINFOHEADER>())?,
            biWidth: width,
            biHeight: -height,
            biPlanes: 1,
            biBitCount: 32,
            biCompression: BI_RGB.0,
            ..Default::default()
        };

        let mut bgra = vec![0_u8; usize::try_from(width * height * 4)?];
        let screen_dc = windows::Win32::Graphics::Gdi::GetDC(None);
        let memory_dc = CreateCompatibleDC(Some(screen_dc));
        let scan_lines = GetDIBits(
            memory_dc,
            bitmap,
            0,
            height as u32,
            Some(bgra.as_mut_ptr().cast::<c_void>()),
            &mut bitmap_info,
            DIB_RGB_COLORS,
        );
        let _ = DeleteDC(memory_dc);
        let _ = ReleaseDC(None, screen_dc);
        if scan_lines == 0 {
            anyhow::bail!("GetDIBits failed for Windows preview bitmap")
        }

        let mut rgba = bgra;
        for pixel in rgba.chunks_exact_mut(4) {
            pixel.swap(0, 2);
        }

        let image = RgbaImage::from_raw(width as u32, height as u32, rgba)
            .ok_or_else(|| anyhow::anyhow!("Failed to convert Windows preview bitmap"))?;
        let mut bytes = Vec::new();
        DynamicImage::ImageRgba8(image).write_to(&mut Cursor::new(&mut bytes), ImageFormat::Bmp)?;
        Ok(bytes)
    }
}

fn configure_theme(ctx: &egui::Context) {
    let mut style = (*ctx.style()).clone();
    style.spacing.item_spacing = Vec2::new(10.0, 10.0);
    style.spacing.button_padding = Vec2::new(12.0, 8.0);
    style.spacing.menu_margin = egui::Margin::same(10);
    style.spacing.indent = 18.0;
    style.text_styles = [
        (
            egui::TextStyle::Heading,
            egui::FontId::new(H1_SIZE, egui::FontFamily::Proportional),
        ),
        (
            egui::TextStyle::Body,
            egui::FontId::new(BODY_SIZE, egui::FontFamily::Proportional),
        ),
        (
            egui::TextStyle::Button,
            egui::FontId::new(BODY_SIZE, egui::FontFamily::Proportional),
        ),
        (
            egui::TextStyle::Small,
            egui::FontId::new(SMALL_SIZE, egui::FontFamily::Proportional),
        ),
        (
            egui::TextStyle::Monospace,
            egui::FontId::new(BODY_SIZE, egui::FontFamily::Monospace),
        ),
    ]
    .into();
    style.visuals = egui::Visuals::dark();
    style.visuals.override_text_color = Some(Color32::from_rgb(230, 236, 240));
    style.visuals.hyperlink_color = Color32::from_rgb(130, 165, 197);
    style.visuals.faint_bg_color = Color32::from_rgb(43, 49, 58);
    style.visuals.extreme_bg_color = Color32::from_rgb(25, 31, 38);
    style.visuals.panel_fill = Color32::from_rgb(37, 43, 50);
    style.visuals.window_fill = Color32::from_rgb(34, 40, 48);
    style.visuals.code_bg_color = Color32::from_rgb(31, 37, 45);
    style.visuals.widgets.noninteractive.weak_bg_fill = Color32::from_rgb(20, 28, 36);
    style.visuals.widgets.noninteractive.bg_fill = Color32::from_rgb(20, 28, 36);
    style.visuals.widgets.noninteractive.bg_stroke =
        Stroke::new(1.0, Color32::from_rgb(43, 60, 72));
    style.visuals.widgets.noninteractive.fg_stroke =
        Stroke::new(1.2, Color32::from_rgb(190, 203, 213));
    style.visuals.widgets.inactive.weak_bg_fill = Color32::from_rgb(27, 39, 49);
    style.visuals.widgets.inactive.bg_fill = Color32::from_rgb(27, 39, 49);
    style.visuals.widgets.inactive.bg_stroke = Stroke::new(1.0, Color32::from_rgb(56, 78, 92));
    style.visuals.widgets.inactive.fg_stroke = Stroke::new(1.4, Color32::from_rgb(240, 243, 245));
    style.visuals.widgets.hovered.weak_bg_fill = Color32::from_rgb(40, 52, 64);
    style.visuals.widgets.hovered.bg_fill = Color32::from_rgb(40, 52, 64);
    style.visuals.widgets.hovered.bg_stroke = Stroke::new(1.0, Color32::from_rgb(92, 115, 136));
    style.visuals.widgets.hovered.fg_stroke = Stroke::new(1.4, Color32::from_rgb(247, 249, 250));
    style.visuals.widgets.active.weak_bg_fill = Color32::from_rgb(57, 83, 106);
    style.visuals.widgets.active.bg_fill = Color32::from_rgb(57, 83, 106);
    style.visuals.widgets.active.bg_stroke = Stroke::new(1.0, Color32::from_rgb(132, 168, 198));
    style.visuals.widgets.active.fg_stroke = Stroke::new(1.4, Color32::from_rgb(250, 251, 252));
    style.visuals.widgets.open.weak_bg_fill = Color32::from_rgb(40, 52, 64);
    style.visuals.widgets.open.bg_fill = Color32::from_rgb(40, 52, 64);
    style.visuals.widgets.open.bg_stroke = Stroke::new(1.0, Color32::from_rgb(92, 115, 136));
    style.visuals.widgets.open.fg_stroke = Stroke::new(1.4, Color32::from_rgb(247, 249, 250));
    style.visuals.selection.bg_fill = Color32::from_rgb(57, 83, 106);
    style.visuals.selection.stroke = Stroke::new(1.0, Color32::from_rgb(180, 205, 226));
    style.visuals.window_corner_radius = CornerRadius::same(14);
    style.visuals.menu_corner_radius = CornerRadius::same(12);
    style.visuals.widgets.inactive.corner_radius = CornerRadius::same(10);
    style.visuals.widgets.hovered.corner_radius = CornerRadius::same(10);
    style.visuals.widgets.active.corner_radius = CornerRadius::same(10);
    style.visuals.widgets.open.corner_radius = CornerRadius::same(10);
    ctx.set_style(style);
}

fn format_number(value: i64) -> String {
    let negative = value < 0;
    let digits = value.abs().to_string();
    let mut out = String::new();
    for (index, ch) in digits.chars().rev().enumerate() {
        if index > 0 && index % 3 == 0 {
            out.push(',');
        }
        out.push(ch);
    }
    let formatted: String = out.chars().rev().collect();
    if negative {
        format!("-{formatted}")
    } else {
        formatted
    }
}
