# Audiobook Maker PRO

Convert anything into a chapterized `.m4b` audiobook — from your local files, a YouTube video, an entire playlist, or a spreadsheet of links.

One script. Fully interactive. No config files needed.

---

## Features

- **Five input sources** — folder of audio files, single YT video, YT playlist, Excel/CSV spreadsheet, parent folder of subfolders (batch)
- **Smart encoding** — detects hardware acceleration (Apple AudioToolbox) and uses CBR; falls back to software VBR automatically
- **Smart bitrate** — measures source bitrate, adds a 10% buffer, and snaps up to the next CBR rung
- **Chapter naming** — keep original titles, number-only, or number + title
- **Cover art** — auto-detects a JPG/PNG in the source folder; falls back to YouTube thumbnail or manual path
- **Rich metadata** — author, narrator, year, genre, description embedded in the M4B
- **Clean progress UI** — per-track download bars, overall progress bar, encoding progress bar
- **Audio notification** — spoken "Your audiobook is ready" on completion (macOS); system sound on Windows/Linux
- **Redo loop** — after each audiobook finishes, offers to make another without restarting the script
- **Playlist cache** — playlist metadata is cached so re-runs skip the slow extraction step
- **Cross-platform** — macOS, Linux, Windows

---

## Modes at a glance

| # | Input | Output |
|---|-------|--------|
| 1 | YouTube video URL | Single `.m4b` — YouTube chapters auto-detected. Playlist URL auto-detected with P/V prompt |
| 2 | YouTube playlist URL | Three options: one combined `.m4b`, pick specific videos, or one `.m4b` per video |
| 3 | Folder of `.mp3` / `.m4a` / `.wav` etc. | Single `.m4b` with chapters |
| 4 | Parent folder containing subfolders | One `.m4b` per subfolder (batch mode) |
| 5 | Excel / CSV — title column + URL column | Single `.m4b` — each row is one chapter |

### Batch mode (4) details

- Each subfolder becomes one audiobook; subfolder name becomes the title
- Shared metadata (author, narrator, year, genre) is asked **once** and applied to all
- Cover art is auto-detected per subfolder (single JPG/PNG inside it)
- Already-converted books (`.m4b` exists in parent) are skipped automatically
- Bitrate is detected per subfolder independently; if >30% above source it auto-adjusts without prompting

---

## Requirements

### System
- **Python** 3.10+
- **FFmpeg** (with `ffprobe`)

```bash
# macOS
brew install ffmpeg

# Ubuntu / Debian
sudo apt install ffmpeg

# Windows — download from https://ffmpeg.org/download.html
```

### Python packages

```bash
pip install yt-dlp tqdm pandas openpyxl
```

> `openpyxl` is only needed for `.xlsx` files (mode 4). You can skip it if you use CSV instead.

---

## Installation

```bash
git clone https://github.com/YOUR_USERNAME/audiobook-maker-pro.git
cd audiobook-maker-pro
pip install -r requirements.txt
```

---

## Usage

```bash
python audiobook_maker.py
```

The tool guides you through every decision interactively — no flags, no config files.

### Example session (mode 4 — spreadsheet)

```
  ──────────────────── SPREADSHEET → ONE AUDIOBOOK ────────────────────
  ┌─ EXPECTED FORMAT ──────────────────────────────────────────────────┐
  │ Your file must have exactly these two column headers               │
  │ in the first row (case-insensitive, extra spaces OK):              │
  │                                                                    │
  │    title   |   link                                                │
  │   ─────────┼──────────────────────────────────                    │
  │   Intro     │  https://youtube.com/watch?v=...                    │
  │   Chapter 1 │  https://youtube.com/watch?v=...                    │
  │   Chapter 2 │  https://youtube.com/watch?v=...                    │
  │                                                                    │
  │ Each row = one chapter.  Blank rows are skipped.                   │
  └────────────────────────────────────────────────────────────────────┘

  Path to Excel (.xlsx) or CSV file: ~/Books/my_list.xlsx
  ✓  12 chapter(s) ready

  ...chapter naming, title, cover, metadata prompts...

  ──────────────────── DOWNLOADING ────────────────────────────────────

  ┌─ [1/12]  Introduction
  │  Downloading  ████████████████████| 100%  3.1MB/s
  └─ ✓  1/12 audio downloaded

  ┌─ [2/12]  Part One — Laying Plans
  │  Downloading  ████████████████████| 100%  2.8MB/s
  └─ ✓  2/12 audio downloaded

  ...

  ✓  12 of 12 audio tracks downloaded successfully

  ──────────────────── ENCODING ───────────────────────────────────────
  Encoder  :  Apple AudioToolbox  (hardware, macOS)
  Mode     :  Hardware · CBR
  Chapters :  12

  Step 1/4  Analysing source files
    ████████████████████| 12/12
  ✓  Total duration : 06:14:30

  Step 3/4  Encoding  [CBR 128k]
    ████████████████████| 22470/22470 s
  ✓  Encoded  (CBR 128k)

  ┌─ DONE ─────────────────────────────────────────────────────────────┐
  │ AUDIOBOOK CREATED                                                  │
  │                                                                    │
  │ File      :  my_list.m4b                                           │
  │ Location  :  /Users/you                                            │
  │ Size      :  287.3 MB                                              │
  │ Duration  :  06:14:30                                              │
  │ Chapters  :  12                                                    │
  └────────────────────────────────────────────────────────────────────┘
```

---

## Spreadsheet format (mode 4)

Your file must have **exactly these two column headers** in the first row:

| title | link |
|-------|------|
| Introduction | https://youtube.com/watch?v=... |
| Chapter One | https://youtube.com/watch?v=... |
| Chapter Two | https://youtube.com/watch?v=... |

**Rules:**
- Headers are case-insensitive — `Title`, `TITLE`, `title`, `  Title ` all work
- Leading/trailing spaces in headers and cell values are stripped automatically
- Blank rows are skipped silently
- Column order doesn't matter as long as both headers are present
- Both `.xlsx` and `.csv` are supported; for `.xlsx` install `openpyxl`

The tool will show a preview of the first 3 rows after loading so you can confirm it parsed correctly before anything is downloaded.

---

## Encoding logic

| Condition | Encoder | Mode |
|-----------|---------|------|
| Apple hardware available | `aac_at` | CBR — snapped **up** to next rung ≥ `source × 1.10` |
| Fraunhofer FDK installed | `libfdk_aac` | VBR quality level matching `source × 1.10` |
| Fallback | `aac` (FFmpeg native) | VBR quality level matching `source × 1.10` |

CBR always rounds **up** to the next ladder rung, so the output is never below the source quality. The tool warns you and asks for confirmation if the computed target is more than 30% above the detected source bitrate.

CBR ladder: 64 / 96 / 128 / 160 / 192 / 256 / 320 kbps

---

## Output format

All output files are standard `.m4b` (MPEG-4 Audiobook) compatible with:
- Apple Books / iTunes
- VLC
- Prologue, Bound, BookPlayer (iOS)
- Any AAC-capable player

---

## FAQ

**The download is slow — is that normal?**
Download speed depends on your internet connection and YouTube's throttling. The progress bar shows live speed.

**Can I re-run after a partial download?**
For playlists, failed tracks are skipped with a warning and the rest are stitched together. There is no per-track resume yet.

**What if my folder has mixed audio formats?**
Supported: `.mp3 .m4a .m4b .mp4 .wav .flac .aac .ogg .opus .wma`. All are concatenated via FFmpeg's concat demuxer.

**Does it work with private YouTube videos?**
Yes — when you select a YouTube mode (2, 3, or 4), the tool asks which browser to pull cookies from. As long as you're logged into YouTube in that browser, private and Premium videos will download. Supports Safari, Chrome, Firefox, Edge, Brave, and Opera. Skip this prompt to download public videos only.

See [Browser cookie permissions](#browser-cookie-permissions) if you get a permission error.

**Can I disable the voice notification?**
The `notify()` function at the bottom of the script can be commented out, or you can call `main()` with a quick edit.

---

## Browser cookie permissions

When you pick a browser in the YouTube authentication prompt, yt-dlp reads the cookie database directly from disk. On macOS, Chrome, Brave, and Edge store their cookie database in a location that requires **Full Disk Access** — without it you'll see this error:

```
ERROR: Could not copy Chrome cookie database.
```

### Fix (macOS)

1. Open **System Settings → Privacy & Security → Full Disk Access**
2. Click the **+** button
3. Add your terminal app (Terminal.app, iTerm2, Warp, etc.)
4. Re-run the script

Safari does **not** require Full Disk Access and works out of the box — use it if you want to avoid the permission step.

### Windows

Chrome/Edge on Windows may show:

```
PermissionError: [Errno 13] Permission denied: '...Cookies'
```

Close Chrome/Edge completely before running the script, as the browser locks the cookie file while open.

### Linux

No special permissions needed. Firefox cookies work without any extra steps.

---

## License

MIT License — see [LICENSE](LICENSE).

---

## Acknowledgements

Built on top of:
- [yt-dlp](https://github.com/yt-dlp/yt-dlp) — YouTube downloading
- [FFmpeg](https://ffmpeg.org/) — audio processing
- [tqdm](https://github.com/tqdm/tqdm) — progress bars
- [pandas](https://pandas.pydata.org/) — spreadsheet parsing
