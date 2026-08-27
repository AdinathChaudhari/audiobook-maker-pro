#!/usr/bin/env python3
"""
Audiobook Maker — Universal Converter
======================================
Five workflows, one interactive tool:

  1. YT Video       → M4B  (single video; playlist URL auto-detected)
  2. YT Playlist    → M4B  (combined book / pick videos / one per video)
  3. Folder         → M4B  (local audio files as chapters)
  4. Parent Folder  → M4B  (each subfolder becomes one audiobook — batch mode)
  5. Spreadsheet    → M4B  (Excel/CSV with title + URL columns)

Books over 20 hours (which can stall playback on older iPods/car systems) are
offered for splitting into roughly-equal "(Part N)" books at chapter boundaries
— a stream copy of the already-encoded audio, never a re-encode.

Run:
    python audiobook_maker.py
"""

import os
import sys
import re
import json
import time
import shutil
import hashlib
import tempfile
import subprocess
from pathlib import Path
from datetime import timedelta

# ═════════════════════════════════════════════════════════════════════════════
#  BOOTSTRAP — check & install all dependencies before anything else runs
# ═════════════════════════════════════════════════════════════════════════════

def _bootstrap():
    """
    Verify Python version, pip packages, and ffmpeg/ffprobe.
    Installs missing pip packages automatically.
    Offers to install ffmpeg via Homebrew if missing.
    Exits with a clear message if anything cannot be resolved.
    """
    # ── Python version ────────────────────────────────────────────────────────
    if sys.version_info < (3, 8):
        print(f"\n  Python 3.8+ required (you have {sys.version.split()[0]}).")
        print("  Download: https://www.python.org/downloads/\n")
        sys.exit(1)

    # ── pip packages ─────────────────────────────────────────────────────────
    REQUIRED = [
        ('tqdm',     'tqdm>=4.66.0'),
        ('yt_dlp',   'yt-dlp>=2026.3.17'),
        ('pandas',   'pandas>=2.0.0'),
        ('openpyxl', 'openpyxl>=3.1.0'),
    ]

    import importlib
    missing = []
    for module, pip_spec in REQUIRED:
        try:
            importlib.import_module(module)
        except ImportError:
            missing.append(pip_spec)

    if missing:
        print("\n  ┌─ Installing missing packages " + "─" * 29 + "┐")
        for pkg in missing:
            print(f"  │   • {pkg:<52} │")
        print("  └" + "─" * 60 + "┘\n")

        result = subprocess.run(
            [sys.executable, '-m', 'pip', 'install', '--quiet', '--upgrade'] + missing,
            text=True
        )
        if result.returncode != 0:
            print("\n  pip install failed. Try manually:\n")
            print(f"    pip install {' '.join(missing)}\n")
            sys.exit(1)
        print(f"\n  ✓  Installed: {', '.join(missing)}\n")

    # ── ffmpeg / ffprobe ──────────────────────────────────────────────────────
    BREW_PATHS = ['/opt/homebrew/bin/ffmpeg', '/usr/local/bin/ffmpeg']

    def _which_ffmpeg():
        for c in [shutil.which('ffmpeg')] + BREW_PATHS:
            if c and os.path.exists(c):
                return c
        return None

    if not _which_ffmpeg():
        print("\n  ┌─ FFmpeg not found " + "─" * 41 + "┐")
        print("  │  FFmpeg is required for audio encoding.                    │")
        print("  └" + "─" * 60 + "┘\n")

        # Try to install via Homebrew automatically
        brew = shutil.which('brew')
        if brew:
            print("  Homebrew detected — installing ffmpeg automatically…\n")
            r = subprocess.run([brew, 'install', 'ffmpeg'])
            if r.returncode != 0 or not _which_ffmpeg():
                print("\n  Homebrew install failed. Try manually:\n")
                print("    brew install ffmpeg\n")
                sys.exit(1)
            print("\n  ✓  ffmpeg installed via Homebrew\n")
        else:
            print("  Install options:\n")
            print("    • Homebrew (recommended):  brew install ffmpeg")
            print("      Get Homebrew: https://brew.sh\n")
            print("    • Direct download: https://ffmpeg.org/download.html\n")
            sys.exit(1)

_bootstrap()

# ── Imports (guaranteed present after bootstrap) ──────────────────────────────
from tqdm import tqdm
import yt_dlp
import pandas as pd

# Browser to pull YouTube cookies from (set in main() for YT modes)
_COOKIE_BROWSER = None

# ═════════════════════════════════════════════════════════════════════════════
#  CONSTANTS & DETECTION
# ═════════════════════════════════════════════════════════════════════════════

def _find_ffmpeg():
    for c in [shutil.which('ffmpeg'), '/opt/homebrew/bin/ffmpeg', '/usr/local/bin/ffmpeg']:
        if c and os.path.exists(c):
            return c
    sys.exit("\n  FFmpeg not found. Run the script again and it will offer to install it.\n")

FFMPEG     = _find_ffmpeg()
FFPROBE    = shutil.which('ffprobe') or FFMPEG.replace('ffmpeg', 'ffprobe')
AUDIO_EXTS = {'.mp3', '.m4a', '.m4b', '.mp4', '.wav', '.flac', '.aac', '.ogg', '.opus', '.wma'}
CBR_LADDER = [64, 96, 128, 160, 192, 256, 320]

# ── Long-book splitting ───────────────────────────────────────────────────────
# Older devices (iPod classic/nano, some car head units) can stop resuming
# playback on very long books — audio halts when the screen sleeps or the app
# is closed. Above the threshold we offer to split into roughly-equal parts.
SPLIT_THRESHOLD_MS = 20 * 60 * 60 * 1000   # offer to split books longer than this
MAX_PART_MS        = 20 * 60 * 60 * 1000   # aim for each part to stay at/under this

VBR_QUALITY_MAP = {                    # AAC VBR: 0 = highest quality
    0: 130, 1: 120, 2: 105, 3: 90,
    4: 75,  5: 60,  6: 52,  7: 45,  8: 35, 9: 28,
}

try:
    W = max(62, min(os.get_terminal_size(1).columns - 4, 100))
except (AttributeError, OSError):
    W = 62

# ═════════════════════════════════════════════════════════════════════════════
#  UI HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def _box(lines, title=''):
    """Print a clean bordered box. Long lines are word-wrapped to fit."""
    import textwrap
    inner = W - 2
    max_content = inner - 2  # 1 char padding each side
    if title:
        print(f"  ┌─ {title} {'─' * max(0, inner - len(title) - 3)}┐")
    else:
        print(f"  ┌{'─' * inner}┐")
    for line in lines:
        if not line:
            print(f"  │{' ' * inner}│")
            continue
        for wrapped in textwrap.wrap(line, width=max_content) or [line]:
            pad = max_content - len(wrapped)
            print(f"  │ {wrapped}{' ' * pad} │")
    print(f"  └{'─' * inner}┘")

def _sep(label=''):
    if label:
        side = (W - len(label) - 2) // 2
        print(f"\n  {'─' * side} {label} {'─' * (W - side - len(label) - 2)}")
    else:
        print(f"\n  {'─' * W}")

def _ok(msg):   print(f"  \033[32m✓\033[0m  {msg}")
def _warn(msg): print(f"  \033[33m⚠\033[0m  {msg}")
def _info(msg): print(f"     {msg}")

def _ask(prompt, default=None, valid=None):
    hint = f" [{default}]" if default else ""
    while True:
        try:
            sys.stdout.flush()
            ans = input(f"\n  {prompt}{hint}: ").strip()
        except (EOFError, KeyboardInterrupt):
            print(); sys.exit(0)
        if not ans and default is not None:
            ans = default
        if valid and ans not in valid:
            print(f"  Please enter one of: {', '.join(valid)}")
            continue
        if ans:
            return ans

def _ask_optional(prompt):
    try:
        sys.stdout.flush()
        return input(f"\n  {prompt} (Enter to skip): ").strip() or None
    except (EOFError, KeyboardInterrupt):
        print(); sys.exit(0)

def _ask_path(prompt, is_dir=False):
    while True:
        p = os.path.expanduser(_ask(prompt).strip('\'"'))
        if is_dir and not os.path.isdir(p):
            print("  Directory not found, try again.")
            continue
        if not is_dir and not os.path.isfile(p):
            print("  File not found, try again.")
            continue
        return p

def _step(n, total, label):
    """Print a clean step header with progress fraction."""
    print(f"\n  \033[1mStep {n}/{total}\033[0m  {label}")

# ═════════════════════════════════════════════════════════════════════════════
#  NOTIFICATION  (audio feedback on completion)
# ═════════════════════════════════════════════════════════════════════════════

def notify():
    """Play a spoken completion notice + system bell, cross-platform."""
    print('\a'); sys.stdout.flush()
    try:
        if sys.platform == 'darwin':
            subprocess.run(
                ['say', '-v', 'Samantha', 'Your audiobook is ready.'],
                check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
        elif sys.platform == 'win32':
            import winsound
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        elif sys.platform.startswith('linux'):
            for cmd in [
                ['paplay', '/usr/share/sounds/freedesktop/stereo/complete.oga'],
                ['aplay',  '/usr/share/sounds/sound-icons/finish.wav'],
            ]:
                try:
                    subprocess.run(cmd, check=False, timeout=3,
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    break
                except Exception:
                    pass
    except Exception:
        pass

# ═════════════════════════════════════════════════════════════════════════════
#  ENCODER DETECTION
# ═════════════════════════════════════════════════════════════════════════════

def detect_encoder():
    """Return (name, is_hw, description) for the best available AAC encoder."""
    candidates = [
        ('aac_at',     True,  'Apple AudioToolbox  (hardware, macOS)'),
        ('libfdk_aac', False, 'Fraunhofer FDK AAC  (best software)'),
        ('aac',        False, 'FFmpeg native AAC   (software fallback)'),
    ]
    try:
        avail = subprocess.run(
            [FFMPEG, '-hide_banner', '-encoders'],
            capture_output=True, text=True, timeout=5
        ).stdout
    except Exception:
        return ('aac', False, 'FFmpeg native AAC (fallback)')

    for enc, hw, desc in candidates:
        if enc not in avail:
            continue
        try:
            ok = subprocess.run(
                [FFMPEG, '-f', 'lavfi', '-i', 'anullsrc=r=44100:cl=mono',
                 '-t', '0.1', '-c:a', enc, '-b:a', '64k', '-f', 'null', '-'],
                capture_output=True, timeout=5
            )
            if ok.returncode == 0:
                return (enc, hw, desc)
        except Exception:
            continue
    return ('aac', False, 'FFmpeg native AAC (ultimate fallback)')

# ═════════════════════════════════════════════════════════════════════════════
#  MEDIA HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def get_duration_ms(path):
    out = subprocess.run(
        [FFPROBE, '-v', 'error', '-show_entries', 'format=duration',
         '-of', 'default=noprint_wrappers=1:nokey=1', str(path)],
        capture_output=True, text=True
    ).stdout.strip()
    try:
        return int(float(out) * 1000)
    except ValueError:
        return 0

def get_bitrate_kbps(path):
    out = subprocess.run(
        [FFPROBE, '-v', 'error', '-select_streams', 'a:0',
         '-show_entries', 'stream=bit_rate',
         '-of', 'default=noprint_wrappers=1:nokey=1', str(path)],
        capture_output=True, text=True
    ).stdout.strip()
    try:
        return int(out) // 1000
    except ValueError:
        return 0

def pick_cbr_bitrate(source_kbps):
    """Hardware path: CBR at source + 10% buffer, snapped to ladder."""
    if source_kbps <= 0:
        return '192k'
    ceil = source_kbps * 1.10
    for k in CBR_LADDER:
        if k >= ceil:
            return f'{k}k'
    return f'{CBR_LADDER[-1]}k'

def pick_vbr_quality(source_kbps):
    """Software path: VBR quality level matching source + 10% buffer. Returns (q, kbps)."""
    if source_kbps <= 0:
        return 2, VBR_QUALITY_MAP[2]
    target = source_kbps * 1.10
    for q in range(10):
        kbps = VBR_QUALITY_MAP[q]
        if kbps <= target:
            return q, kbps
    return 9, VBR_QUALITY_MAP[9]

def ms_to_hms(ms):
    s = ms // 1000
    return f"{s//3600:02d}:{(s%3600)//60:02d}:{s%60:02d}"

def sanitize(name):
    return re.sub(r'[\\/*?:"<>|]', '_', name).strip()

def _check_output_path(output_path):
    """Warn and confirm before overwriting an existing file."""
    if output_path.exists():
        size_mb = output_path.stat().st_size / (1024 * 1024)
        _warn(f'Output file already exists: {output_path.name}  ({size_mb:.1f} MB)')
        ans = _ask('Overwrite?  (y/n)', default='no', valid=['yes', 'no', 'y', 'n'])
        if ans in ('no', 'n'):
            print('\n  Aborted — choose a different output path and re-run.\n')
            sys.exit(0)

# ═════════════════════════════════════════════════════════════════════════════
#  LONG-BOOK SPLITTING  (no re-encode — splits the already-encoded audio)
# ═════════════════════════════════════════════════════════════════════════════

def _num_parts(total_ms):
    """Number of roughly-equal parts so each stays at/under MAX_PART_MS (min 2)."""
    return max(2, -(-total_ms // MAX_PART_MS))   # -(-a//b) == ceil(a/b)

def _plan_split_points(chapter_ends_ms, n_parts):
    """
    Pick the chapter boundaries that divide a book into n_parts of roughly equal
    length. Each cut snaps to the chapter end nearest that part's ideal boundary,
    so a chapter is never sliced in half. Requires len(chapter_ends_ms) >= n_parts.
    Returns the last-chapter index (0-based) of every part except the last.
    """
    total      = chapter_ends_ms[-1]
    C          = len(chapter_ends_ms)
    cuts, prev = [], -1
    for k in range(1, n_parts):
        target = k * total / n_parts
        lo, hi = prev + 1, C - (n_parts - k)   # leave >=1 chapter for each remaining part
        best_i = min(range(lo, hi), key=lambda i: abs(chapter_ends_ms[i] - target))
        cuts.append(best_i)
        prev = best_i
    return cuts

def _chapters_for_window(chapters, win_start, win_end):
    """Chapters overlapping [win_start, win_end), clamped to the window and rebased to 0."""
    out = []
    for ch in chapters:
        s = max(ch['start'], win_start)
        e = min(ch['end'],   win_end)
        if e - s <= 0:
            continue
        out.append({'start': s - win_start, 'end': e - win_start, 'title': ch['title']})
    return out

# ═════════════════════════════════════════════════════════════════════════════
#  CHAPTER NAMING
# ═════════════════════════════════════════════════════════════════════════════

def ask_chapter_naming(sample_titles):
    """Interactive chapter naming setup. Returns callable f(idx, title) → chapter_title."""
    n = len(sample_titles)
    sample = sample_titles[0] if sample_titles else 'Example Title'
    width  = len(str(max(1, n)))

    _sep('CHAPTER NAMING')
    _box([
        f'1.  Keep original title as-is',
        f'    e.g.  "{sample}"',
        f'',
        f'2.  Chapter number only',
        f'    e.g.  "Chapter {"01".zfill(width)}"',
        f'',
        f'3.  Chapter number + original title',
        f'    e.g.  "Chapter {"01".zfill(width)} - {sample}"',
    ])
    choice = _ask('Choice', default='1', valid=['1', '2', '3'])

    if choice == '1':
        _ok('Using original titles')
        return lambda idx, title: title
    elif choice == '2':
        _ok('Using chapter numbers only')
        return lambda idx, title: f"Chapter {idx:0{width}d}"
    else:
        _ok('Using chapter number + original title')
        return lambda idx, title: f"Chapter {idx:0{width}d} - {title}"

# ═════════════════════════════════════════════════════════════════════════════
#  METADATA PROMPTS
# ═════════════════════════════════════════════════════════════════════════════

def ask_metadata():
    _sep('OPTIONAL METADATA')
    _info('Press Enter to skip any field.')
    return {
        'author':      _ask_optional('Author name'),
        'narrator':    _ask_optional('Narrator name'),
        'year':        _ask_optional('Year  (e.g. 2024)'),
        'genre':       _ask_optional('Genre / Category'),
        'description': _ask_optional('Description / Comment'),
    }

def ask_title(suggested):
    choice = _ask(f'Use title "{suggested}"?  (y/n)', default='yes',
                  valid=['yes', 'no', 'y', 'n'])
    if choice in ('yes', 'y'):
        return suggested
    return _ask('Enter title')

# ═════════════════════════════════════════════════════════════════════════════
#  CORE: BUILD M4B
# ═════════════════════════════════════════════════════════════════════════════

def build_m4b(audio_paths, chapter_titles, output_path, cover_path, metadata, encoder_info, batch=False, yt_chapters=None):
    """
    Four-pass M4B builder:
      1. Encode each source file individually  →  per-chapter .m4a files
      2. Measure durations of encoded files    →  exact chapter timestamps
      3. Concatenate encoded files             →  combined .m4a  (copy, no re-encode)
      4. Mux                                   →  final .m4b  (chapters + cover + tags)

    Encoding before concatenating guarantees chapter boundaries match the
    actual encoded audio — re-encoding after concat shifts frame boundaries
    and causes chapter start/end drift.
    """
    audio_paths = [Path(p) for p in audio_paths]
    output_path = Path(output_path)
    encoder, is_hw, enc_desc = encoder_info
    total_chapters = len(yt_chapters) if yt_chapters else len(audio_paths)

    # ── Long-book split decision (asked once, BEFORE the expensive encode) ─────
    # The split itself happens after encoding as a stream copy — audio is never
    # re-encoded. Here we only decide how many parts to produce.
    total_src_ms     = sum(get_duration_ms(p) for p in audio_paths)
    do_split, n_parts = False, 1
    if total_src_ms > SPLIT_THRESHOLD_MS:
        n_parts = _num_parts(total_src_ms)
        if batch:
            do_split = True
            _warn(f'Long book (~{ms_to_hms(total_src_ms)}) — auto-splitting into '
                  f'{n_parts} parts of ~{ms_to_hms(total_src_ms // n_parts)} each.')
        else:
            _sep('LONG AUDIOBOOK')
            _box([
                f'This audiobook is ~{ms_to_hms(total_src_ms)} long.',
                '',
                'Books over 20h can fail to resume on older devices',
                '(iPod classic/nano, some car systems): playback stops',
                'when the screen sleeps or the app is closed.',
                '',
                f'Split into {n_parts} parts of ~{ms_to_hms(total_src_ms // n_parts)} each?',
                'Parts are named "(Part 1)", "(Part 2)", … and added as',
                'separate books.  The audio is NOT re-encoded.',
            ])
            ans = _ask('Split into parts?  (y/n)', default='yes',
                       valid=['yes', 'no', 'y', 'n'])
            do_split = ans in ('yes', 'y')
            if do_split:
                custom = _ask('Number of parts', default=str(n_parts))
                try:
                    n_parts = max(2, int(custom))
                except ValueError:
                    pass

    # ── Header ────────────────────────────────────────────────────────────────
    _sep('ENCODING')
    mode_tag = 'Hardware · CBR' if is_hw else 'Software · VBR'
    _box([
        f'Encoder  :  {enc_desc}',
        f'Mode     :  {mode_tag}',
        f'Chapters :  {total_chapters}',
    ])

    # ── Source quality detection & encoding target ────────────────────────────
    src_kbps = get_bitrate_kbps(audio_paths[0])

    if src_kbps > 0:
        _info(f'Source bitrate : ~{src_kbps} kbps')
    else:
        _info('Source bitrate : unknown — using safe default')

    if is_hw:
        cbr       = pick_cbr_bitrate(src_kbps)
        cbr_kbps  = int(cbr.rstrip('k'))
        _info(f'Target         : {cbr} CBR  (source + 10% buffer)')

        if src_kbps > 0 and cbr_kbps > src_kbps * 1.30:
            _warn(f'{cbr} is >30% above detected source (~{src_kbps} kbps) — file will be larger with no quality gain.')
            sys.stdout.flush()
            if batch:
                cbr = next((f'{k}k' for k in CBR_LADDER if k >= src_kbps), f'{CBR_LADDER[-1]}k')
                _ok(f'Auto-adjusted to {cbr} (batch mode)')
            else:
                ans = input('\n  Keep target bitrate anyway? (y/n) [yes]: ').strip().lower() or 'yes'
                if ans not in ('yes', 'y'):
                    # Skip the 10% buffer — snap up to just cover the raw source
                    cbr = next((f'{k}k' for k in CBR_LADDER if k >= src_kbps), f'{CBR_LADDER[-1]}k')
                    _ok(f'Adjusted to {cbr}')

        enc_flags  = ['-b:a', cbr]
        mode_label = f'CBR {cbr}'

    else:
        vbr_q, vbr_kbps = pick_vbr_quality(src_kbps)
        _info(f'Target         : VBR quality {vbr_q}  (~{vbr_kbps} kbps, source + 10% buffer)')

        if src_kbps > 0 and vbr_kbps > src_kbps * 1.30:
            _warn(f'Quality {vbr_q} (~{vbr_kbps} kbps) is >30% above source (~{src_kbps} kbps).')
            sys.stdout.flush()
            if batch:
                vbr_q, vbr_kbps = pick_vbr_quality(src_kbps)
                _ok(f'Auto-adjusted to quality {vbr_q} (~{vbr_kbps} kbps) (batch mode)')
            else:
                ans = input('\n  Keep quality level anyway? [yes]: ').strip().lower() or 'yes'
                if ans not in ('yes', 'y'):
                    vbr_q, vbr_kbps = pick_vbr_quality(src_kbps * 0.95)
                    _ok(f'Adjusted to quality {vbr_q} (~{vbr_kbps} kbps)')

        enc_flags  = ['-q:a', str(vbr_q)]
        mode_label = f'VBR q{vbr_q}  (~{vbr_kbps} kbps)'

    with tempfile.TemporaryDirectory() as _tmp:
        tmp = Path(_tmp)

        # ── Step 1: Encode each chapter individually ──────────────────────────
        # Encoding first (before concat) preserves exact per-chapter boundaries.
        # A post-concat re-encode shifts AAC frame boundaries and causes drift.
        _step(1, 4, f'Encoding chapters  [{mode_label}]')
        encoded_files = []
        # total_src_ms was computed above (for the split decision) — reuse it.

        with tqdm(total=int(total_src_ms / 1000), unit='s', ncols=W + 4,
                  bar_format='  {l_bar}{bar}| {n_fmt}/{total_fmt} s') as pbar:
            for idx, src in enumerate(audio_paths):
                enc_out = tmp / f'ch_{idx:04d}.m4a'
                src_sec = get_duration_ms(src) / 1000
                enc_cmd = [FFMPEG, '-y', '-i', str(src), '-vn',
                           '-c:a', encoder] + enc_flags + [str(enc_out)]
                proc = subprocess.Popen(
                    enc_cmd + ['-progress', 'pipe:1', '-nostats'],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True
                )
                last = 0
                for line in proc.stdout:
                    if 'out_time_ms=' in line:
                        try:
                            cur = min(int(line.split('=')[1]) // 1000, int(src_sec))
                            pbar.update(cur - last); last = cur
                        except ValueError:
                            pass
                proc.wait()

                if proc.returncode != 0:
                    _warn(f'Primary encoder failed on chapter {idx+1} — retrying with native aac…')
                    fallback = [FFMPEG, '-y', '-i', str(src), '-vn',
                                '-c:a', 'aac'] + enc_flags + [str(enc_out)]
                    subprocess.run(fallback, check=True, capture_output=True)

                # Advance progress bar to the full duration of this chapter
                pbar.update(int(src_sec) - last)
                encoded_files.append(enc_out)

        _ok(f'Encoded {total_chapters} chapter(s)  ({mode_label})')

        # ── Step 2: Measure encoded durations (used for chapter timestamps) ───
        # Durations are read from the encoded files, not the originals, so
        # chapter START/END values exactly match what is in the audio stream.
        _step(2, 4, 'Measuring encoded chapter durations')
        durations_ms = []
        with tqdm(total=total_chapters, unit='file', ncols=W + 4,
                  bar_format='  {l_bar}{bar}| {n_fmt}/{total_fmt}') as pbar:
            for enc in encoded_files:
                durations_ms.append(get_duration_ms(enc))
                pbar.update(1)

        total_ms = sum(durations_ms)
        _ok(f'Total duration : {ms_to_hms(total_ms)}')

        # ── Step 3: Concatenate encoded chapters (copy — no re-encode) ────────
        _step(3, 4, 'Concatenating encoded chapters')
        concat_txt = tmp / 'concat.txt'
        with open(concat_txt, 'w', encoding='utf-8') as f:
            for enc in encoded_files:
                safe = str(enc).replace("'", "\\'")
                f.write(f"file '{safe}'\n")

        encoded = tmp / 'encoded.m4a'
        r = subprocess.run(
            [FFMPEG, '-y', '-f', 'concat', '-safe', '0',
             '-i', str(concat_txt), '-c', 'copy', str(encoded)],
            capture_output=True, text=True
        )
        if r.returncode != 0:
            print(f'\n  Concat error:\n{r.stderr[-500:]}')
            sys.exit(1)
        _ok(f'Concatenated {total_chapters} chapter(s)')

        # ── Build the unified chapter list (start/end in ms) ──────────────────
        if yt_chapters:
            chapters = [{'start': int(c['start_time'] * 1000),
                         'end':   int(c['end_time'] * 1000),
                         'title': c['title']} for c in yt_chapters]
        else:
            chapters, _cursor = [], 0
            for title_ch, dur in zip(chapter_titles, durations_ms):
                chapters.append({'start': _cursor, 'end': _cursor + dur, 'title': title_ch})
                _cursor += dur

        # ── Cover (extracted once, reused for every part) ─────────────────────
        cover_arg = []
        if cover_path and os.path.isfile(str(cover_path)):
            cover_jpg = tmp / 'cover.jpg'
            subprocess.run([FFMPEG, '-y', '-i', str(cover_path), str(cover_jpg)],
                           capture_output=True)
            if cover_jpg.exists():
                cover_arg = ['-i', str(cover_jpg)]

        # ── Work out the part windows (one window = one output file) ──────────
        # When splitting, cuts snap to the CHAPTER boundary nearest each equal
        # division, so no chapter is sliced. The cut is a stream copy (-c copy)
        # below — the encoded audio is reused as-is, never re-encoded.
        if do_split:
            ends = [c['end'] for c in chapters]
            if len(chapters) >= n_parts:
                cuts   = _plan_split_points(ends, n_parts)
                bounds = [0] + [chapters[i]['end'] for i in cuts] + [total_ms]
            else:
                # Fewer chapters than parts: fall back to equal time cuts.
                bounds = [round(k * total_ms / n_parts) for k in range(n_parts)] + [total_ms]
            windows = [(bounds[k], bounds[k + 1], k + 1) for k in range(n_parts)]
        else:
            windows = [(0, total_ms, None)]

        # ── Step 4: Mux each window → M4B ─────────────────────────────────────
        tag_map = {
            'author': 'artist', 'narrator': 'composer',
            'year': 'date',     'genre': 'genre',  'description': 'comment',
        }
        base_title = (metadata or {}).get('title') or output_path.stem
        results    = []

        for win_s, win_e, part_no in windows:
            if part_no is None:
                out_path   = output_path
                this_title = base_title
            else:
                out_path   = output_path.parent / f'{sanitize(output_path.stem)} (Part {part_no}).m4b'
                this_title = f'{base_title} (Part {part_no})'

            win_chaps = _chapters_for_window(chapters, win_s, win_e)
            win_dur   = win_e - win_s

            # ffmetadata for this window — global tags + its chapters, rebased to 0.
            meta_file = tmp / f'meta_{part_no or 0}.txt'
            with open(meta_file, 'w', encoding='utf-8') as mf:
                mf.write(';FFMETADATA1\n')
                if metadata:
                    for k, tag in tag_map.items():
                        v = metadata.get(k)
                        if v:
                            mf.write(f'{tag}={v}\n')
                mf.write(f'title={this_title}\n')
                mf.write(f'album={this_title}\n')
                if metadata and metadata.get('author'):
                    mf.write(f"album_artist={metadata['author']}\n")
                mf.write('\n')
                for ch in win_chaps:
                    mf.write(f'[CHAPTER]\nTIMEBASE=1/1000\n'
                             f'START={ch["start"]}\nEND={ch["end"]}\ntitle={ch["title"]}\n\n')

            label = out_path.name if part_no is None else f'Part {part_no}/{n_parts} → {out_path.name}'
            _step(4, 4, f'Muxing → {label}')
            out_path.parent.mkdir(parents=True, exist_ok=True)
            if not batch:
                _check_output_path(out_path)

            # -ss/-t as INPUT options for the encoded stream = copy-only cut.
            mux_cmd = [FFMPEG, '-y']
            if part_no is not None:
                mux_cmd += ['-ss', f'{win_s / 1000:.3f}', '-t', f'{win_dur / 1000:.3f}']
            mux_cmd += ['-i', str(encoded), '-f', 'ffmetadata', '-i', str(meta_file)]
            if cover_arg:
                mux_cmd += cover_arg
            mux_cmd += ['-map_metadata', '1', '-map_chapters', '1', '-map', '0:a']
            if cover_arg:
                mux_cmd += ['-map', '2:v', '-c:v', 'copy', '-disposition:v:0', 'attached_pic']
            mux_cmd += ['-c:a', 'copy', '-movflags', '+faststart',
                        '-brand', 'M4B ', '-f', 'mp4', str(out_path)]

            r = subprocess.run(mux_cmd, capture_output=True, text=True)
            if r.returncode != 0:
                print(f'\n  Mux error:\n{r.stderr[-500:]}')
                sys.exit(1)
            results.append((out_path, win_dur, len(win_chaps)))

    # ── Summary ───────────────────────────────────────────────────────────────
    _sep()
    if len(results) == 1:
        out_path, dur, nch = results[0]
        size_mb = out_path.stat().st_size / (1024 * 1024)
        _box([
            'AUDIOBOOK CREATED',
            '',
            f'File      :  {out_path.name}',
            f'Location  :  {out_path.parent}',
            f'Size      :  {size_mb:.1f} MB',
            f'Duration  :  {ms_to_hms(dur)}',
            f'Chapters  :  {nch}',
        ], title='DONE')
    else:
        lines = [f'AUDIOBOOK SPLIT INTO {len(results)} PARTS', '']
        for out_path, dur, nch in results:
            size_mb = out_path.stat().st_size / (1024 * 1024)
            lines.append(out_path.name)
            lines.append(f'    {ms_to_hms(dur)}  ·  {nch} chapters  ·  {size_mb:.1f} MB')
            lines.append('')
        lines.append(f'Location  :  {results[0][0].parent}')
        _box(lines, title='DONE')
    return [out for out, _, _ in results]

# ═════════════════════════════════════════════════════════════════════════════
#  MODE 3: FOLDER → AUDIOBOOK
# ═════════════════════════════════════════════════════════════════════════════

def mode_folder(encoder_info):
    _sep('FOLDER → AUDIOBOOK')
    folder = Path(_ask_path('Path to folder containing audio files', is_dir=True))

    files = sorted(
        [f for f in folder.iterdir() if f.suffix.lower() in AUDIO_EXTS],
        key=lambda f: (int(m.group(1)) if (m := re.match(r'^(\d+)', f.stem)) else 0, f.name)
    )
    if not files:
        print('\n  No audio files found in that folder.\n'); sys.exit(1)

    _box(
        [f'{i}. {f.name}' for i, f in enumerate(files[:6], 1)]
        + ([f'… and {len(files) - 6} more'] if len(files) > 6 else []),
        title=f'Found {len(files)} audio file(s)'
    )

    raw_titles = [re.sub(r'^\d+[\s\-_\.]*', '', f.stem).strip() or f.stem for f in files]
    namer          = ask_chapter_naming(raw_titles)
    chapter_titles = [namer(i + 1, t) for i, t in enumerate(raw_titles)]

    _sep('AUDIOBOOK DETAILS')
    title = ask_title(folder.name)

    IMAGE_EXTS = {'.jpg', '.jpeg', '.png'}
    cover_images = [f for f in folder.iterdir() if f.suffix.lower() in IMAGE_EXTS]
    cover_path = None
    if len(cover_images) == 1:
        cover_path = cover_images[0]
        _ok(f'Cover image detected: {cover_path.name}')
    else:
        choice = _ask('Do you have a cover image?  (y/n)', default='no',
                      valid=['yes', 'no', 'y', 'n'])
        if choice in ('yes', 'y'):
            cover_path = _ask_path('Path to cover image (JPG or PNG)')

    meta = ask_metadata()
    meta['title'] = title

    suggested_out = str(folder.parent / f'{sanitize(title)}.m4b')
    out_in = _ask_optional(f'Output path  [{suggested_out}]')
    output_path = Path(out_in or suggested_out)

    build_m4b(files, chapter_titles, output_path, cover_path, meta, encoder_info)

# ═════════════════════════════════════════════════════════════════════════════
#  MODE 4: PARENT FOLDER → ONE AUDIOBOOK PER SUBFOLDER  (batch)
# ═════════════════════════════════════════════════════════════════════════════

def _find_audio_folders(parent):
    """Recursively collect every folder under `parent` that DIRECTLY contains audio
    files. Enables nested batch: parent/theme/level/*.mp3 → one book per leaf folder.
    A folder with both audio and sub-folders becomes a book AND is descended into."""
    found = []
    def walk(d):
        try:
            entries = sorted(d.iterdir())
        except OSError:
            return
        if any(f.suffix.lower() in AUDIO_EXTS for f in entries if f.is_file()):
            found.append(d)
        for sub in entries:
            if sub.is_dir():
                walk(sub)
    for d in sorted(parent.iterdir()):
        if d.is_dir():
            walk(d)
    return found


def _folder_title(parent, folder):
    """Hierarchical title: the folder's path relative to parent, joined by ' - '.
    Immediate children keep their plain name (backward-compatible)."""
    return ' - '.join(folder.relative_to(parent).parts)


def _find_cover_upwards(folder, parent, IMAGE_EXTS):
    """A single image in the folder; else the nearest one walking up to `parent`
    (so a cover in a theme folder applies to all its level sub-books)."""
    d = folder
    while True:
        imgs = [f for f in d.iterdir() if f.is_file() and f.suffix.lower() in IMAGE_EXTS]
        if len(imgs) == 1:
            return imgs[0]
        if d == parent:
            return None
        d = d.parent


def mode_batch_folder(encoder_info):
    _sep('BATCH FOLDERS → AUDIOBOOKS')
    parent = Path(_ask_path('Path to parent folder containing subfolders', is_dir=True))

    IMAGE_EXTS = {'.jpg', '.jpeg', '.png'}

    # Discover every folder (at any depth) that directly contains audio.
    subfolders = _find_audio_folders(parent)
    if not subfolders:
        print('\n  No folders with audio files found.\n'); sys.exit(1)

    # If any book comes from a nested folder, let the user choose whether the
    # parent folder(s) are part of the book title.
    nested = any(len(d.relative_to(parent).parts) > 1 for d in subfolders)
    include_parents = True
    if nested:
        ex = next(d for d in subfolders if len(d.relative_to(parent).parts) > 1)
        _sep('BOOK NAMING')
        _box([
            f'1.  Include parent folder(s):  "{_folder_title(parent, ex)}"',
            f'2.  Leaf folder name only:     "{ex.name}"',
        ])
        include_parents = _ask('Choice', default='1', valid=['1', '2']) == '1'

    title_of = {d: (_folder_title(parent, d) if include_parents else d.name) for d in subfolders}

    # Guard: leaf-only titles can collide (e.g. many "Level 1"), which would
    # overwrite outputs — fall back to full hierarchy if that happens.
    if len(set(title_of.values())) != len(title_of):
        _warn('Leaf-only names collide — keeping parent folders so titles stay unique.')
        title_of = {d: _folder_title(parent, d) for d in subfolders}

    _box([title_of[d] for d in subfolders[:8]] + ([f'… and {len(subfolders) - 8} more'] if len(subfolders) > 8 else []),
         title=f'Found {len(subfolders)} folder(s) to convert')

    # Check which already have a .m4b (named by hierarchical title) and skip
    already_done = [d for d in subfolders if (parent / f'{sanitize(title_of[d])}.m4b').exists()]
    if already_done:
        _box([f'✓  {title_of[d]}.m4b' for d in already_done], title=f'{len(already_done)} already converted — will skip')
        subfolders = [d for d in subfolders if d not in already_done]
        if not subfolders:
            print('\n  All audiobooks already exist. Nothing to do.\n'); return

    # Shared metadata — asked once, applied to all
    _sep('SHARED METADATA  (applied to all audiobooks)')
    _info('Press Enter to skip any field.')
    shared_meta = {
        'author':      _ask_optional('Author name'),
        'narrator':    _ask_optional('Narrator name'),
        'year':        _ask_optional('Year  (e.g. 2024)'),
        'genre':       _ask_optional('Genre / Category'),
        'description': _ask_optional('Description / Comment'),
    }

    # Chapter naming — asked once
    _sep('CHAPTER NAMING')
    def _sort_key(f):
        mo = re.match(r'^(\d+)', f.stem)
        return (int(mo.group(1)) if mo else 0, f.name)
    sample_titles = [re.sub(r'^\d+[\s\-_\.]*', '', f.stem).strip() or f.stem
                     for f in sorted(subfolders[0].iterdir(), key=_sort_key)
                     if f.suffix.lower() in AUDIO_EXTS][:3]
    namer = ask_chapter_naming(sample_titles)

    # Process each subfolder
    done, skipped, failed = [], [], []
    for i, folder in enumerate(subfolders, 1):
        title = title_of[folder]
        _sep(f'[{i}/{len(subfolders)}]  {title}')

        files = sorted(
            [f for f in folder.iterdir() if f.suffix.lower() in AUDIO_EXTS],
            key=lambda f: (int(m.group(1)) if (m := re.match(r'^(\d+)', f.stem)) else 0, f.name)
        )

        raw_titles     = [re.sub(r'^\d+[\s\-_\.]*', '', f.stem).strip() or f.stem for f in files]
        chapter_titles = [namer(j + 1, t) for j, t in enumerate(raw_titles)]

        cover_path = _find_cover_upwards(folder, parent, IMAGE_EXTS)
        if cover_path:
            _ok(f'Cover detected: {cover_path.name}')

        meta          = {**shared_meta, 'title': title}
        output_path   = parent / f'{sanitize(title)}.m4b'

        try:
            build_m4b(files, chapter_titles, output_path, cover_path, meta, encoder_info, batch=True)
            done.append(folder.name)
        except Exception as e:
            _warn(f'Failed: {e}')
            failed.append(folder.name)

    # Final summary
    _sep('BATCH COMPLETE')
    _box(
        [f'✓  {n}' for n in done] +
        ([f'✗  {n}  (failed)' for n in failed] if failed else []),
        title=f'{len(done)} succeeded  |  {len(failed)} failed  |  {len(already_done)} skipped'
    )

# ═════════════════════════════════════════════════════════════════════════════
#  YOUTUBE HELPERS
# ═════════════════════════════════════════════════════════════════════════════

class _Q:
    def debug(self, m): pass
    def warning(self, m): pass
    def error(self, m): pass

def _cookie_opts():
    """Return yt-dlp cookiesfrombrowser option if a browser is selected."""
    if _COOKIE_BROWSER:
        return {'cookiesfrombrowser': (_COOKIE_BROWSER, None, None, None)}
    return {}

def _ytdlp_info(url, flat=False):
    opts = {'quiet': True, 'no_warnings': True, 'logger': _Q(), 'extract_flat': flat,
            'remote_components': ['ejs:github'], **_cookie_opts()}
    with yt_dlp.YoutubeDL(opts) as ydl:
        return ydl.extract_info(url, download=False)

def _is_original_audio(f):
    """True if this audio track is the video's original language, not a YouTube auto-dub.

    YouTube encodes auto-dubbed tracks at a HIGHER bitrate than the original, so any
    bitrate-first sort silently picks a dub. The original is flagged two ways:
    language_preference == 10, and 'original' in format_note.
    """
    if (f.get('language_preference') or -1) >= 10:
        return True
    return 'original' in (f.get('format_note') or '').lower()

def _best_audio_format(url):
    try:
        info   = _ytdlp_info(url)
        audio  = [f for f in info.get('formats', [])
                  if f.get('acodec') not in (None, 'none')
                  and f.get('vcodec') in (None, 'none', 'video only')]
        # Keep only the original-language track when the video carries auto-dubs.
        originals = [f for f in audio if _is_original_audio(f)]
        if originals:
            audio = originals
        pref   = {'opus': 0, 'aac': 1, 'mp4a': 1, 'vorbis': 2, 'mp3': 3}
        audio.sort(key=lambda f: (
            -(f.get('abr') or f.get('tbr') or 0),
            next((v for k, v in pref.items() if k in (f.get('acodec') or '').lower()), 99)
        ))
        if audio:
            best = audio[0]
            if originals:
                lang = best.get('language') or 'original'
                _info(f'Audio track     : {lang} (original) — auto-dubs skipped')
            return (f"{best['format_id']}/bestaudio",
                    best.get('abr') or best.get('tbr') or 0)
    except Exception:
        pass
    # Fallback: yt-dlp's own default sort puts language preference first, so this
    # also prefers the original track. Never bare 'best' — that can be a dub.
    return 'bestaudio/best', 0

def _download_audio(url, dest_path, label='', cache_stem=None):
    """
    Download best audio from url → dest_path (yt-dlp adds extension). Returns Path.

    If cache_stem is provided, the file is saved to the persistent downloads/
    folder so it can be reused on resume without re-downloading.
    """
    # Resume: return cached file immediately if it already exists
    if cache_stem:
        cached = _find_cached_download(cache_stem)
        if cached:
            _ok(f'Resuming from cache: {cached.name}')
            return cached
        dest_path = _dl_cache_dir() / cache_stem

    fmt, _ = _best_audio_format(url)

    pbar = tqdm(total=100, unit='%', ncols=W + 4,
                bar_format=f'  {label or "Download"} ' + '{bar}| {n_fmt}%  {postfix}',
                postfix='')

    def hook(d):
        if d['status'] == 'downloading':
            total = d.get('total_bytes') or d.get('total_bytes_estimate', 0)
            done  = d.get('downloaded_bytes', 0)
            if total:
                pct = int(done / total * 100)
                pbar.n = pct
                pbar.set_postfix_str(d.get('_speed_str', ''))
                pbar.refresh()
        elif d['status'] == 'finished':
            pbar.n = 100; pbar.refresh()

    def _attempt(with_cookies):
        opts = {
            'format': fmt,
            'outtmpl': str(dest_path) + '.%(ext)s',
            'ffmpeg_location': FFMPEG,
            'progress_hooks': [hook],
            'quiet': True, 'no_warnings': True, 'logger': _Q(),
            'postprocessors': [],
            'remote_components': ['ejs:github'],
            **(_cookie_opts() if with_cookies else {}),
        }
        with yt_dlp.YoutubeDL(opts) as ydl:
            ydl.download([url])

    try:
        _attempt(with_cookies=True)
    except Exception:
        if _COOKIE_BROWSER:
            _warn('Cookie download failed — retrying without cookies…')
            _attempt(with_cookies=False)
        else:
            raise
    pbar.close()

    parent = Path(dest_path).parent
    stem   = Path(dest_path).name
    for f in parent.iterdir():
        if f.stem == stem:
            return f
    return None

def _download_thumbnail(url, dest_folder, stem):
    """Download & convert thumbnail to JPG. Returns path or None."""
    out = str(Path(dest_folder) / stem)
    opts = {
        'skip_download': True, 'writethumbnail': True,
        'outtmpl': out + '.%(ext)s',
        'ffmpeg_location': FFMPEG,
        'quiet': True, 'no_warnings': True, 'logger': _Q(),
        'remote_components': ['ejs:github'],
        **_cookie_opts(),
    }
    with yt_dlp.YoutubeDL(opts) as ydl:
        try: ydl.download([url])
        except Exception: pass
    for ext in ('jpg', 'png', 'webp', 'bmp'):
        src = Path(f'{out}.{ext}')
        if src.exists():
            jpg = Path(f'{out}.jpg')
            if src != jpg:
                subprocess.run([FFMPEG, '-y', '-i', str(src), str(jpg)], capture_output=True)
                src.unlink(missing_ok=True)
            return str(jpg)
    return None

# ── Playlist cache ─────────────────────────────────────────────────────────

def _cache_path(url):
    h = hashlib.md5(url.encode()).hexdigest()
    d = Path('cache'); d.mkdir(exist_ok=True)
    return d / f'pl_{h}.json'

CACHE_MAX_AGE_S = 24 * 60 * 60  # 24 hours

def _load_cache(url):
    p = _cache_path(url)
    if p.exists():
        try:
            data = json.loads(p.read_text(encoding='utf-8'))
            if data.get('url') != url:
                return None, None
            age = time.time() - data.get('saved_at', 0)
            if age > CACHE_MAX_AGE_S:
                _info(f'Playlist cache is {int(age / 3600)}h old — refreshing…')
                p.unlink(missing_ok=True)
                return None, None
            return data.get('entries'), data.get('title')
        except Exception:
            pass
    return None, None

def _save_cache(url, entries, title):
    try:
        _cache_path(url).write_text(
            json.dumps({'url': url, 'entries': entries, 'title': title,
                        'saved_at': time.time()}, ensure_ascii=False),
            encoding='utf-8'
        )
    except Exception:
        pass

# ── Download cache (persist audio between runs for resume support) ─────────

def _dl_cache_dir():
    d = Path('downloads'); d.mkdir(exist_ok=True)
    return d

def _dl_cache_stem(url, index):
    """Stable filename for a given URL + playlist position."""
    h = hashlib.md5(url.encode()).hexdigest()[:12]
    return f'{index:04d}_{h}'

def _find_cached_download(stem):
    """Return the cached audio file if it already exists, else None."""
    for f in _dl_cache_dir().iterdir():
        if f.stem == stem and f.suffix.lower() in AUDIO_EXTS:
            return f
    return None

def _playlist_url(url: str) -> str:
    """Return a clean playlist-only URL (strips v= and index= so yt-dlp fetches the full list)."""
    from urllib.parse import urlparse, parse_qs, urlencode, urlunparse
    p  = urlparse(url)
    qs = parse_qs(p.query, keep_blank_values=True)
    if 'list' not in qs:
        return url
    clean = urlencode({'list': qs['list'][0]})
    return urlunparse(('https', 'www.youtube.com', '/playlist', '', clean, ''))


def _fetch_playlist(url):
    pl_url = _playlist_url(url)
    entries, title = _load_cache(pl_url)
    if entries:
        _ok(f'Loaded {len(entries)} entries from cache')
        return entries, title

    _info('Fetching playlist info…')
    flat     = _ytdlp_info(pl_url, flat=True)
    title    = flat.get('title', 'Playlist')
    flat_ids = {e['id'] for e in flat.get('entries', [])
                if e and e.get('title') not in ('Private video', 'Deleted video')}
    _ok(f'{len(flat_ids)} visible videos found')

    _info('Fetching full metadata (may take a moment)…')
    full    = _ytdlp_info(url, flat=False)
    entries = [e for e in full.get('entries', []) if e and e.get('id') in flat_ids]
    entries.sort(key=lambda e: int(e.get('playlist_index') or 0))
    _save_cache(url, entries, title)
    return entries, title

# ═════════════════════════════════════════════════════════════════════════════
#  MODE 1 / 2 SHARED HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def _is_playlist_url(url: str) -> bool:
    """True when url carries a playlist ID (index= param doesn't change that)."""
    from urllib.parse import urlparse, parse_qs
    qs = parse_qs(urlparse(url).query)
    return 'list' in qs


def _parse_selection(s: str, max_n: int) -> list:
    """Parse a selection string like '1,3,5-8' into a sorted list of 1-based indices."""
    indices = set()
    for part in s.split(','):
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            a, b = part.split('-', 1)
            indices.update(range(int(a), int(b) + 1))
        else:
            indices.add(int(part))
    return sorted(i for i in indices if 1 <= i <= max_n)


def _vurl(entry) -> str:
    return entry.get('webpage_url') or f"https://www.youtube.com/watch?v={entry['id']}"


def _run_playlist_to_one(entries, pl_title, encoder_info):
    """Download all entries and stitch into a single M4B (shared by mode 2 & 3)."""
    _sep('AUDIOBOOK DETAILS')
    title = ask_title(pl_title)

    raw_titles     = [e.get('title', f'Video {i+1}') for i, e in enumerate(entries)]
    namer          = ask_chapter_naming(raw_titles)
    chapter_titles = [namer(i + 1, t) for i, t in enumerate(raw_titles)]

    cv = _ask('Cover — first thumbnail (F), last thumbnail (L), or your own (C)?',
              default='F', valid=['F', 'f', 'L', 'l', 'C', 'c'])

    meta = ask_metadata()
    meta['title'] = title

    suggested_out = str(Path(os.getcwd()) / f'{sanitize(title)}.m4b')
    out_in = _ask_optional(f'Output path  [{suggested_out}]')
    output_path = Path(out_in or suggested_out)

    with tempfile.TemporaryDirectory() as _tmp:
        tmp                  = Path(_tmp)
        audio_files          = []
        chapter_titles_final = []
        failed               = []

        _sep('DOWNLOADING')
        cover_path = None
        if cv.lower() == 'c':
            cover_path = _ask_path('Path to cover image (JPG or PNG)')
        else:
            thumb_entry = entries[-1] if cv.lower() == 'l' else entries[0]
            _info('Downloading cover…')
            cover_path = _download_thumbnail(_vurl(thumb_entry), tmp, 'cover')
            if cover_path: _ok('Cover saved')

        total = len(entries)
        print()
        overall = tqdm(total=total, unit='track', ncols=W + 4,
                       bar_format='  Overall  {bar}| {n_fmt}/{total_fmt} tracks  [{elapsed}<{remaining}]')

        for i, entry in enumerate(entries, 1):
            url    = _vurl(entry)
            vtitle = entry.get('title', f'Video {i}')
            label  = vtitle[:28] + '…' if len(vtitle) > 30 else vtitle
            stem   = _dl_cache_stem(url, i)
            af     = _download_audio(url, tmp / f'{i:04d}', label=f'[{i}/{total}] {label}',
                                     cache_stem=stem)
            if af and af.exists():
                audio_files.append(af)
                chapter_titles_final.append(chapter_titles[i - 1])
            else:
                failed.append(i)
            overall.update(1)
        overall.close()

        if failed:
            _warn(f'{len(failed)} download(s) failed and were skipped: {failed}')
        if not audio_files:
            print('\n  No audio downloaded. Aborting.\n'); sys.exit(1)

        build_m4b(audio_files, chapter_titles_final,
                  output_path, cover_path, meta, encoder_info)


def _run_per_video(entries, encoder_info):
    """Encode each entry as its own M4B — thumbnail cover, shared metadata."""
    _sep('AUDIOBOOK DETAILS  (applied to all books)')
    meta = ask_metadata()

    total = len(entries)
    for i, entry in enumerate(entries, 1):
        url    = _vurl(entry)
        vtitle = entry.get('title', f'Video {i}')
        _sep(f'VIDEO {i}/{total}')

        title = ask_title(vtitle)
        suggested_out = str(Path(os.getcwd()) / f'{sanitize(title)}.m4b')
        out_in = _ask_optional(f'Output path  [{suggested_out}]')
        output_path = Path(out_in or suggested_out)

        this_meta = dict(meta)
        this_meta['title'] = title

        with tempfile.TemporaryDirectory() as _tmp:
            tmp  = Path(_tmp)
            stem = sanitize(vtitle)

            _sep('DOWNLOADING')
            _info('Downloading thumbnail…')
            cover_path = _download_thumbnail(url, tmp, stem + '_thumb')
            if cover_path:
                _ok('Thumbnail saved')

            audio_file = _download_audio(url, tmp / stem, label=f'[{i}/{total}] Audio')
            if not audio_file or not audio_file.exists():
                _warn(f'Download failed for "{vtitle}" — skipping.')
                continue
            _ok('Audio downloaded')

            yt_chaps = entry.get('chapters') or []
            use_yt_chapters = None
            if yt_chaps:
                ans = _ask(f'Use the {len(yt_chaps)} YouTube chapters? (Y/N)',
                           default='Y', valid=['Y', 'y', 'N', 'n'])
                use_yt_chapters = yt_chaps if ans.upper() == 'Y' else None

            build_m4b([audio_file], [title], output_path, cover_path,
                      this_meta, encoder_info, yt_chapters=use_yt_chapters)


def _pick_videos_from_playlist(entries, pl_title, encoder_info):
    """Let the user select one or more videos from an already-displayed list."""
    raw    = input('  Enter video numbers (e.g. 1,3,5-8): ').strip()
    chosen = _parse_selection(raw, len(entries))
    if not chosen:
        print('\n  No valid selection. Aborting.\n'); sys.exit(1)

    selected = [entries[i - 1] for i in chosen]
    _ok(f'{len(selected)} video(s) selected')

    if len(selected) == 1:
        _handle_single_video_entry(selected[0], encoder_info)
        return

    _handle_playlist(selected, pl_title, encoder_info, detected=False)


def _handle_single_video_entry(entry, encoder_info):
    """Full single-video flow for an entry dict (from a playlist or direct URL)."""
    sv_url = _vurl(entry)
    _info('Fetching video info…')
    info      = _ytdlp_info(sv_url)
    yt_title  = info.get('title', entry.get('title', 'YouTube_Video'))
    duration  = info.get('duration', 0)
    yt_chaps  = info.get('chapters') or []
    box_lines = [f'Title    : {yt_title}', f'Duration : {ms_to_hms(duration * 1000)}']
    if yt_chaps:
        box_lines.append(f'Chapters : {len(yt_chaps)} found in video')
    _box(box_lines)

    use_yt_chapters = None
    if yt_chaps:
        ans = _ask(f'Use the {len(yt_chaps)} YouTube chapters as audiobook chapters? (Y/N)',
                   default='Y', valid=['Y', 'y', 'N', 'n'])
        use_yt_chapters = yt_chaps if ans.upper() == 'Y' else None

    _sep('AUDIOBOOK DETAILS')
    title      = ask_title(yt_title)
    cv         = _ask('Cover — YouTube thumbnail (T) or your own image (C)?',
                      default='T', valid=['T', 't', 'C', 'c'])
    cover_path = _ask_path('Path to cover image (JPG or PNG)') if cv.lower() == 'c' else None
    meta       = ask_metadata()
    meta['title'] = title
    suggested_out = str(Path(os.getcwd()) / f'{sanitize(title)}.m4b')
    out_in     = _ask_optional(f'Output path  [{suggested_out}]')
    output_path = Path(out_in or suggested_out)

    with tempfile.TemporaryDirectory() as _tmp:
        tmp  = Path(_tmp)
        stem = sanitize(yt_title)
        _sep('DOWNLOADING')
        if cover_path is None:
            _info('Downloading thumbnail…')
            cover_path = _download_thumbnail(sv_url, tmp, stem + '_thumb')
            if cover_path: _ok('Thumbnail saved')
        audio_file = _download_audio(sv_url, tmp / stem, label='Audio')
        if not audio_file or not audio_file.exists():
            print('\n  Download failed.\n'); sys.exit(1)
        _ok('Audio downloaded')
        build_m4b([audio_file], [title], output_path, cover_path, meta, encoder_info,
                  yt_chapters=use_yt_chapters)


def _handle_playlist(entries, pl_title, encoder_info, detected=False):
    """
    Unified playlist handler used by both mode 2 (detected playlist URL) and
    mode 3 (explicit playlist input).  `detected=True` changes the opening
    wording to indicate the playlist was auto-detected rather than chosen.
    """
    # Always show the video list first so the user can make an informed choice.
    _sep('PLAYLIST VIDEOS')
    for i, e in enumerate(entries, 1):
        print(f'  {i:>3}.  {e.get("title", f"Video {i}")}')
    print()

    if detected:
        _warn('Playlist URL detected — this is not a single video.')
        print('  How would you like to proceed?\n')
    else:
        print('  How would you like to process this playlist?\n')

    print('  P.  One combined audiobook  (each video = one chapter)')
    print('  V.  Choose one or more videos from the list')
    print('  E.  One separate audiobook per video')
    print()
    route = _ask('Choice', default='P', valid=['P', 'p', 'V', 'v', 'E', 'e'])

    if route.upper() == 'P':
        _run_playlist_to_one(entries, pl_title, encoder_info)
    elif route.upper() == 'V':
        _pick_videos_from_playlist(entries, pl_title, encoder_info)
    else:
        _run_per_video(entries, encoder_info)


# ═════════════════════════════════════════════════════════════════════════════
#  MODE 1: YOUTUBE VIDEO → AUDIOBOOK
# ═════════════════════════════════════════════════════════════════════════════

def mode_single_video(encoder_info):
    _sep('YOUTUBE VIDEO → AUDIOBOOK')
    url = _ask('YouTube video URL')

    if _is_playlist_url(url):
        _warn('Playlist URL detected — this is not a single video.')
        print()
        print('  P.  Proceed with the whole playlist')
        print('  V.  Pick specific video(s) from the playlist')
        print()
        route = _ask('Choice', default='P', valid=['P', 'p', 'V', 'v'])

        if route.upper() == 'P':
            entries, pl_title = _fetch_playlist(url)
            _box([f'Playlist : {pl_title}', f'Videos   : {len(entries)}'])
            _handle_playlist(entries, pl_title, encoder_info, detected=False)
        else:
            entries, pl_title = _fetch_playlist(url)
            _box([f'Playlist : {pl_title}', f'Videos   : {len(entries)}'])
            _sep('PLAYLIST VIDEOS')
            for i, e in enumerate(entries, 1):
                print(f'  {i:>3}.  {e.get("title", f"Video {i}")}')
            print()
            _pick_videos_from_playlist(entries, pl_title, encoder_info)
        return

    _info('Fetching video info…')
    info       = _ytdlp_info(url)
    yt_title   = info.get('title', 'YouTube_Video')
    duration   = info.get('duration', 0)
    yt_chaps   = info.get('chapters') or []
    box_lines  = [f'Title    : {yt_title}', f'Duration : {ms_to_hms(duration * 1000)}']
    if yt_chaps:
        box_lines.append(f'Chapters : {len(yt_chaps)} found in video')
    _box(box_lines)

    use_yt_chapters = None
    if yt_chaps:
        ans = _ask(f'Use the {len(yt_chaps)} YouTube chapters as audiobook chapters? (Y/N)',
                   default='Y', valid=['Y', 'y', 'N', 'n'])
        use_yt_chapters = yt_chaps if ans.upper() == 'Y' else None

    _sep('AUDIOBOOK DETAILS')
    title = ask_title(yt_title)

    cv = _ask('Cover — YouTube thumbnail (T) or your own image (C)?',
              default='T', valid=['T', 't', 'C', 'c'])
    cover_path = _ask_path('Path to cover image (JPG or PNG)') if cv.lower() == 'c' else None

    meta = ask_metadata()
    meta['title'] = title

    suggested_out = str(Path(os.getcwd()) / f'{sanitize(title)}.m4b')
    out_in = _ask_optional(f'Output path  [{suggested_out}]')
    output_path = Path(out_in or suggested_out)

    with tempfile.TemporaryDirectory() as _tmp:
        tmp  = Path(_tmp)
        stem = sanitize(yt_title)

        _sep('DOWNLOADING')
        if cover_path is None:
            _info('Downloading thumbnail…')
            cover_path = _download_thumbnail(url, tmp, stem + '_thumb')
            if cover_path:
                _ok('Thumbnail saved')

        audio_file = _download_audio(url, tmp / stem, label='Audio')
        if not audio_file or not audio_file.exists():
            print('\n  Download failed.\n'); sys.exit(1)
        _ok('Audio downloaded')

        build_m4b([audio_file], [title], output_path, cover_path, meta, encoder_info,
                  yt_chapters=use_yt_chapters)

# ═════════════════════════════════════════════════════════════════════════════
#  MODE 2: YOUTUBE PLAYLIST → AUDIOBOOK
# ═════════════════════════════════════════════════════════════════════════════

def mode_playlist(encoder_info):
    _sep('YOUTUBE PLAYLIST → AUDIOBOOK')
    url             = _ask('YouTube playlist URL')
    entries, pl_title = _fetch_playlist(url)
    _box([f'Playlist : {pl_title}', f'Videos   : {len(entries)}'])

    all_choice = _ask('Process all videos (A) or a subset (S)?',
                      default='A', valid=['A', 'a', 'S', 's'])
    if all_choice.lower() == 's':
        start   = int(_ask('Starting video number (1-based)'))
        end     = int(_ask(f'Ending video number (max {len(entries)})'))
        entries = entries[start - 1:end]
        _ok(f'Selected {len(entries)} video(s)')

    _handle_playlist(entries, pl_title, encoder_info, detected=False)

# ═════════════════════════════════════════════════════════════════════════════
#  MODE 5: SPREADSHEET (Excel / CSV) → ONE AUDIOBOOK
# ═════════════════════════════════════════════════════════════════════════════

def mode_spreadsheet(encoder_info):
    _sep('SPREADSHEET → ONE AUDIOBOOK')
    _box([
        'Your file must have exactly these two column headers',
        'in the first row (case-insensitive, extra spaces OK):',
        '',
        '   title   |   link',
        '  ─────────┼──────────────────────────────────',
        '  Intro     │  https://youtube.com/watch?v=...',
        '  Chapter 1 │  https://youtube.com/watch?v=...',
        '  Chapter 2 │  https://youtube.com/watch?v=...',
        '',
        'Each row = one chapter.  Blank rows are skipped.',
    ], title='EXPECTED FORMAT')

    sheet_path = _ask_path('Path to Excel (.xlsx) or CSV file')
    ext = Path(sheet_path).suffix.lower()
    try:
        df = pd.read_excel(sheet_path) if ext in ('.xls', '.xlsx') else pd.read_csv(sheet_path)
    except Exception as e:
        print(f'\n  Could not read file: {e}\n'); sys.exit(1)

    # Normalise headers: strip whitespace, lowercase — so "Title ", "TITLE", " title" all match
    df.columns = [str(c).strip().lower() for c in df.columns]

    if 'title' not in df.columns or 'link' not in df.columns:
        _warn(f'Could not find "title" and "link" columns.  Found: {list(df.columns)}')
        print('  Make sure the first row of your file has exactly the headers "title" and "link".\n')
        sys.exit(1)

    # Strip whitespace from every cell in both columns
    df['title'] = df['title'].astype(str).str.strip()
    df['link']  = df['link'].astype(str).str.strip()

    # Drop rows where either cell is empty / NaN / literal "nan"
    df = df[df['title'].notna() & df['link'].notna()]
    df = df[~df['title'].isin(['', 'nan']) & ~df['link'].isin(['', 'nan'])]

    rows       = df[['title', 'link']].values.tolist()
    raw_titles = [r[0] for r in rows]
    urls       = [r[1] for r in rows]

    _box([f'Rows loaded : {len(rows)}',
          '',
          'Preview:',
          *[f'  {t[:30]:<32} {u[:30]}' for t, u in zip(raw_titles[:3], urls[:3])]],
         title='File loaded')

    if not rows:
        print('\n  No valid rows found after cleaning.  Check the file.\n'); sys.exit(1)

    _ok(f'{len(rows)} chapter(s) ready')

    _sep('AUDIOBOOK DETAILS')
    namer          = ask_chapter_naming(raw_titles)
    chapter_titles = [namer(i + 1, t) for i, t in enumerate(raw_titles)]

    title = ask_title(Path(sheet_path).stem)

    cv = _ask('Cover — first video thumbnail (F), last video thumbnail (L), or your own (C)?',
              default='F', valid=['F', 'f', 'L', 'l', 'C', 'c'])

    meta = ask_metadata()
    meta['title'] = title

    suggested_out = str(Path(os.getcwd()) / f'{sanitize(title)}.m4b')
    out_in = _ask_optional(f'Output path  [{suggested_out}]')
    output_path = Path(out_in or suggested_out)

    with tempfile.TemporaryDirectory() as _tmp:
        tmp                  = Path(_tmp)
        audio_files          = []
        chapter_titles_final = []
        failed               = []

        # Cover
        _sep('DOWNLOADING')
        cover_path = None
        if cv.lower() == 'c':
            cover_path = _ask_path('Path to cover image (JPG or PNG)')
        else:
            thumb_url  = urls[-1] if cv.lower() == 'l' else urls[0]
            _info('Downloading cover…')
            cover_path = _download_thumbnail(thumb_url, tmp, 'cover')
            if cover_path: _ok('Cover saved')

        # Audio tracks
        total      = len(rows)
        downloaded = 0
        print()

        for i, (raw_title, url) in enumerate(zip(raw_titles, urls), 1):
            # Header line for this track
            label = raw_title[:40] + '…' if len(raw_title) > 42 else raw_title
            print(f'\n  ┌─ [{i}/{total}]  {label}')
            stem  = _dl_cache_stem(url, i)
            af    = _download_audio(url, tmp / f'{i:04d}', label='  │  Downloading',
                                    cache_stem=stem)
            if af and af.exists():
                audio_files.append(af)
                chapter_titles_final.append(chapter_titles[i - 1])
                downloaded += 1
                print(f'  └─ ✓  {downloaded}/{total} audio downloaded')
            else:
                failed.append(i)
                print(f'  └─ ✗  Download failed  ({downloaded}/{total} so far)')

        print()
        if failed:
            _warn(f'{len(failed)} download(s) failed and were skipped — tracks: {failed}')
        _ok(f'{downloaded} of {total} audio tracks downloaded successfully')
        if not audio_files:
            print('\n  No audio downloaded. Aborting.\n'); sys.exit(1)

        build_m4b(audio_files, chapter_titles_final,
                  output_path, cover_path, meta, encoder_info)

# ═════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═════════════════════════════════════════════════════════════════════════════

def main():
    t0 = time.time()

    print()
    print('  ╔' + '═' * W + '╗')
    print('  ║' + '  A U D I O B O O K   M A K E R   P R O'.center(W) + '║')
    print('  ║' + 'YT Video · YT Playlist · Folder · Spreadsheet'.center(W) + '║')
    print('  ╚' + '═' * W + '╝')

    # Encoder detection
    print()
    _info('Detecting best AAC encoder…')
    encoder_info = detect_encoder()
    _, is_hw, enc_desc = encoder_info
    mode_str = 'Hardware · CBR' if is_hw else 'Software · VBR'
    _ok(f'{enc_desc}  [{mode_str}]')

    # Mode selection
    _sep('WHAT WOULD YOU LIKE TO DO?')
    _box([
        '1.  YouTube video            →  audiobook',
        '2.  YouTube playlist         →  combined / pick videos / one per video',
        '3.  Folder of audio files    →  audiobook',
        '4.  Parent folder            →  one audiobook per subfolder  (batch)',
        '5.  Spreadsheet              →  ONE audiobook  (each row = chapter)',
        '    Excel / CSV with title + YouTube URL columns',
    ])

    choice = _ask('Choice', valid=['1', '2', '3', '4', '5'])

    # Browser cookie prompt — only relevant for YouTube modes
    global _COOKIE_BROWSER
    if choice in ('1', '2', '5'):
        _sep('YOUTUBE AUTHENTICATION')
        _box([
            'Use browser cookies to access private / Premium videos.',
            'Skip (press Enter) for public videos only.',
            '',
            'macOS: Chrome/Brave/Edge require Full Disk Access for Terminal.',
            '  System Settings → Privacy & Security → Full Disk Access → add Terminal',
            'If you see: "Could not copy Chrome cookie database" — that is why.',
        ])
        BROWSERS = ['safari', 'chrome', 'firefox', 'edge', 'brave', 'opera']
        for i, b in enumerate(BROWSERS, 1):
            print(f'  {i}.  {b.capitalize()}')
        print('  0.  No cookies  (public only)')
        bchoice = input('\n  Browser [0]: ').strip() or '0'
        if bchoice in [str(i) for i in range(1, len(BROWSERS) + 1)]:
            _COOKIE_BROWSER = BROWSERS[int(bchoice) - 1]
            _ok(f'Using cookies from {_COOKIE_BROWSER.capitalize()}')
        else:
            _COOKIE_BROWSER = None
            _info('No cookies — public videos only')

    dispatch = {
        '1': mode_single_video,
        '2': mode_playlist,
        '3': mode_folder,
        '4': mode_batch_folder,
        '5': mode_spreadsheet,
    }
    while True:
        dispatch[choice](encoder_info)

        elapsed = time.time() - t0
        print(f'\n  Total time : {str(timedelta(seconds=int(elapsed)))}')
        print()
        notify()

        again = _ask('\n  Make another audiobook?  (y/n)', default='no',
                     valid=['yes', 'no', 'y', 'n'])
        if again not in ('yes', 'y'):
            print('\n  Done. Goodbye!\n')
            break

        # Re-show mode menu for the next run
        t0 = time.time()
        _sep('WHAT WOULD YOU LIKE TO DO?')
        _box([
            '1.  YouTube video            →  audiobook',
            '2.  YouTube playlist         →  combined / pick videos / one per video',
            '3.  Folder of audio files    →  audiobook',
            '4.  Parent folder            →  one audiobook per subfolder  (batch)',
            '5.  Spreadsheet              →  ONE audiobook  (each row = chapter)',
            '    Excel / CSV with title + YouTube URL columns',
        ])
        choice = _ask('Choice', valid=['1', '2', '3', '4', '5'])

        if choice in ('1', '2', '5'):
            _sep('YOUTUBE AUTHENTICATION')
            _box([
                'Use browser cookies to access private / Premium videos.',
                'Skip (press Enter) for public videos only.',
                '',
                'macOS: Chrome/Brave/Edge require Full Disk Access for Terminal.',
                '  System Settings → Privacy & Security → Full Disk Access → add Terminal',
                'If you see: "Could not copy Chrome cookie database" — that is why.',
            ])
            BROWSERS = ['safari', 'chrome', 'firefox', 'edge', 'brave', 'opera']
            for i, b in enumerate(BROWSERS, 1):
                print(f'  {i}.  {b.capitalize()}')
            print('  0.  No cookies  (public only)')
            bchoice = input('\n  Browser [0]: ').strip() or '0'
            if bchoice in [str(i) for i in range(1, len(BROWSERS) + 1)]:
                _COOKIE_BROWSER = BROWSERS[int(bchoice) - 1]
                _ok(f'Using cookies from {_COOKIE_BROWSER.capitalize()}')
            else:
                _COOKIE_BROWSER = None
                _info('No cookies — public videos only')


if __name__ == '__main__':
    main()