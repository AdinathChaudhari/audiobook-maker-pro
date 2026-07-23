# Changelog

All notable changes to Audiobook Maker PRO are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [1.6.0] — 2026-07-23

### Added
- **Recursive (nested) batch mode.** Batch mode now walks the whole folder tree, not
  just immediate children: every folder at any depth that directly contains audio
  becomes one audiobook. Titles are built from the hierarchy relative to the parent,
  joined by ` - ` (e.g. `Basics - Level 1`, `Intelligence - Level 8`). A folder that
  has both audio and sub-folders becomes a book *and* is descended into. Flat parents
  behave exactly as before (single-level titles) — fully backward-compatible.
- **Cover inheritance.** If a leaf folder has no cover image, the tool walks up toward
  the parent and uses the nearest single image — so one cover in a theme folder applies
  to all of that theme's sub-books.

---

## [1.5.0] — 2026-06-15

### Added
- **Long-book splitting for older devices.** Books longer than 20 hours can stop
  resuming playback on older hardware (iPod classic/nano, some car head units) —
  audio halts when the screen sleeps or the app is closed. Before writing the
  final file, the tool now detects this and offers to split the book into
  roughly-equal parts named `(Part 1)`, `(Part 2)`, … that you add as separate
  books. Part count is `ceil(total / 20h)` divided evenly (e.g. a 50h book →
  three ~16/17/16h parts, not 20/20/10), and you can override the count.
  - **No re-encode.** The audio is encoded exactly once; each part is produced by
    a stream copy (`-c copy`) of the already-encoded audio, so quality is
    byte-identical to the unsplit book.
  - **Cuts land on chapter boundaries** nearest each division, so no chapter is
    sliced; each part keeps its own chapters (re-based to 0:00) and the cover.
  - **Batch mode (mode 4)** auto-splits long books without prompting, matching its
    existing auto-adjust behaviour.
  - Tunable via the `SPLIT_THRESHOLD_MS` / `MAX_PART_MS` constants near the top of
    the script (e.g. lower `MAX_PART_MS` to 15h for shorter parts).

## [1.4.0] — 2026-05-25

### Changed
- **Mode order rewritten** to reflect a more logical flow — YouTube-first, then local files, then spreadsheet:
  - 1 → YouTube video (was 2)
  - 2 → YouTube playlist (was 3)
  - 3 → Folder of audio files (was 1)
  - 4 → Parent folder / batch (was 5)
  - 5 → Spreadsheet (was 4)
- Cookie prompt now correctly triggers for modes 1, 2, and 5 (the YouTube-sourced modes).

## [1.3.0] — 2026-05-25

### Added
- **Unified playlist handler** shared by modes 2 and 3 — the same three-option menu appears whether a playlist URL was auto-detected (mode 2) or entered directly (mode 3):
  - **P** — one combined audiobook (each video = one chapter)
  - **V** — browse the full video list and pick one or more (ranges like `1,3,5-8` supported); a single pick goes straight to the single-video flow, multiple picks offer combined or per-video
  - **E** — one separate audiobook per video (each gets its own thumbnail cover automatically)
- **Wording adapts to context**: mode 2 says "Playlist URL detected — this is not a single video" while mode 3 says "How would you like to process this playlist?" — same logic, appropriate framing.

### Changed
- Module docstring updated to reflect that mode 2 handles both single videos and accidental playlist URLs via the shared handler.

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
