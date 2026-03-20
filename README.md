# File Indexer

File Indexer is a vibe coded Windows desktop app for indexing your drives and finding files quickly by name or folder path. It is built for large local file collections where manual folder browsing is slow.

## Why Use It

- Index local drives once, then search filenames much faster than browsing folders manually.
- Narrow results with boolean filename search using `&&` and `||`.
- Preview images and videos directly inside the app.
- Move through results with the keyboard and export matching files to `.m3u`.
- Uses a minimum supported window size of `1280x800` to keep the UI usable.

## Features

- Fast local filename indexing
- Boolean filename search
- Folder-path search
- Image preview
- Video preview with Windows shell or FFmpeg
- Result-card navigation with `Up` / `Down`
- M3U export for current search results
- Windows executable icon embedded into the app

## Search Syntax

Search matches against filename and folder path.

- `mp4 && unforgiven`
  Matches filenames containing both `mp4` and `unforgiven`
- `ytd_ || trailer`
  Matches filenames containing either `ytd_` or `trailer`

`&&` acts as AND. `||` acts as OR.

## Build

Debug build:

```powershell
cargo build
```

Release build:

```powershell
cargo build --release
```

Output binaries:

- `target\debug\file-indexer.exe`
- `target\release\file-indexer.exe`

## Run

```powershell
cargo run
```

Or launch the built executable directly.

## Display

The app expects a minimum window resolution of `1280x800`. It is designed to start large on launch so the result cards and controls do not overlap.

## FFmpeg Preview

For FFmpeg video previews, place `ffmpeg.exe` next to the app in one of these locations:

- same folder as the executable
- `tools\ffmpeg\ffmpeg.exe`
- `ffmpeg\ffmpeg.exe`

If FFmpeg is not found there, the app will also try `PATH`.

## Tech Stack

- Rust
- `eframe` / `egui`
- SQLite via `rusqlite`
- Windows shell APIs
