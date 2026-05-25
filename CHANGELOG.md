# Changelog

All notable changes to Audiobook Maker PRO are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [1.3.0] — 2026-05-25

### Added
- **Playlist URL detection in mode 2**: when a playlist URL is pasted instead of a single video URL (common when copying a link from within a playlist), the tool detects it and offers two paths: process the whole playlist as one audiobook, or browse and pick individual videos from a numbered list. Ranges are supported (`1,3,5-8`). If multiple videos are picked, the user can combine them into one audiobook (each video = chapter) or make one separate audiobook per video.
- **Per-video output in mode 3**: after selecting which videos to process, mode 3 now asks whether to produce one combined audiobook (previous behaviour) or one `.m4b` per video.
- **Thumbnail cover default for per-video output**: when producing one audiobook per video, each book automatically uses that video's YouTube thumbnail as cover art — no prompt needed.

## [1.2.0] — 2026-05-25

### Added
- **YouTube chapter detection (mode 2)**: when a single YouTube video has chapters defined (via timestamps in the description), the tool now detects them, shows the count in the info box, and asks whether to use them as audiobook chapters. If accepted, the exact YouTube chapter timestamps and titles are embedded directly into the `.m4b` — no need to split the video into separate tracks.

---

## [1.1.0] — 2025-04-15

### Changed
- **Spreadsheet mode — fixed column headers**: file must now have a `title` column and a `link` column in the header row (row 1). Headers are matched case-insensitively and leading/trailing whitespace is stripped, so `"Title "`, `"TITLE"`, `" title"` all work. Cell values in both columns are also stripped of whitespace. Previously the user was prompted to identify columns at runtime; this is no longer needed.
- **Spreadsheet download UI**: replaced the single overall progress bar with a per-track bracket display that shows the track name, a live download bar with speed, and a running `x/n audio downloaded` counter after each track completes. Failed downloads show `✗  Download failed  (x/n so far)` so the count is always visible.

---

## [1.0.0] — 2025-04-15

### Added
- **Four input modes** in a single script: folder, YouTube video, YouTube playlist, spreadsheet (Excel/CSV)
- **Interactive UI** — clean boxed panels, step headers, no flags required
- **Per-track download progress bars** via tqdm with live speed readout
- **Overall progress bar** for multi-track modes (playlist / spreadsheet)
- **Encoding progress bar** driven by FFmpeg `-progress pipe:1`
- **Smart encoder detection** — tests `aac_at` (Apple hardware) → `libfdk_aac` → native `aac` in order; verifies each actually runs before selecting
- **Hardware/software encoding split** — hardware uses CBR at source + 10% buffer; software uses VBR quality level matched to source + 10% buffer
- **Upsampling guard** — warns and prompts user if computed target exceeds source by >30%
- **Chapter naming options** — original title / number only / number + title
- **Cover art** — YouTube thumbnail (first or last video) or custom image; auto-converted to JPEG
- **Rich M4B metadata** — title, artist, album, album_artist, composer (narrator), date, genre, comment
- **Playlist cache** — playlist metadata saved to `cache/` to skip re-fetching on repeat runs
- **Audio completion notification** — spoken phrase on macOS (Samantha voice), system beep on Windows/Linux
- **Spreadsheet mode** — reads Excel (.xlsx) or CSV; user selects title and URL columns at runtime
- **Cross-platform** — macOS, Linux, Windows; FFmpeg auto-detected in common paths

### Technical
- Two-pass M4B pipeline: encode first, then mux (cleaner than single-pass)
- Soft dependency check on startup with actionable install instructions
- Temporary directory used for all intermediates; auto-cleaned on success or failure
- CBR ladder: 64 / 96 / 128 / 160 / 192 / 256 / 320 kbps
- VBR quality map: quality 0 (~130 kbps) through quality 9 (~28 kbps)
