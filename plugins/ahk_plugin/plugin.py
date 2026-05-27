"""
cc_autohotkey / plugin.py
=========================
AutoHotkey 2.0 interoperability plugin for Command Center.

Purpose
-------
This plugin is NOT an AutoHotkey replacement or editor — it improves
*interoperability* between your existing AHK 2.0 scripts and Command Center:

  • Automatically locates your AutoHotkey 2.0 installation
  • Registers .ahk script files you care about
  • Parses each script and extracts its hotkey definitions
  • Detects conflicts between those hotkeys and:
      - Command Center's built-in keyboard shortcuts
      - Any other active CC plugin hotkeys
  • Runs and stops registered scripts as background subprocesses
  • Shows live status (running / stopped / error) for each script

Open the dashboard via:
  • Footer button  → "AHK"
  • Hotkey         → Ctrl+Shift+A
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
import os
from pathlib import Path
from typing import Optional

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QWidget, QLabel, QPushButton,
    QFrame, QScrollArea, QTableWidget, QTableWidgetItem, QHeaderView,
    QAbstractItemView, QLineEdit, QFileDialog, QSizePolicy, QTabWidget,
    QTextEdit, QMessageBox, QCheckBox
)
from PySide6.QtCore import Qt, QTimer, Signal, QThread, QObject
from PySide6.QtGui import QFont, QColor


# ---------------------------------------------------------------------------
# Module-level state
# ---------------------------------------------------------------------------

_api        = None          # CommandCenterAPI — set in activate()
_footer_btn = None          # QPushButton in the footer bar
_dashboard  = None          # Open QDialog (if any)
_poll_timer = None          # QTimer for polling process status

# Script registry: list of {"path": str, "auto_start": bool}
_scripts: list[dict] = []

# Running subprocesses: path → subprocess.Popen
_procs: dict[str, subprocess.Popen] = {}

# Cached AHK executable path
_ahk_exe: Optional[Path] = None


# ---------------------------------------------------------------------------
# Settings keys
# ---------------------------------------------------------------------------

_KEY_AHK_EXE   = "ahk_exe_path"
_KEY_SCRIPTS   = "scripts_json"
_KEY_CONFLICTS = "check_conflicts_on_open"


# ---------------------------------------------------------------------------
# Known AutoHotkey 2.0 installation paths (searched in order)
# ---------------------------------------------------------------------------

_AHK_SEARCH_PATHS: list[Path] = [
    Path(r"C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe"),
    Path(r"C:\Program Files\AutoHotkey\v2\AutoHotkey32.exe"),
    Path(r"C:\Program Files\AutoHotkey\v2\AutoHotkeyU64.exe"),
    Path(r"C:\Program Files\AutoHotkey\AutoHotkey64.exe"),
    Path(r"C:\Program Files\AutoHotkey\AutoHotkey.exe"),
    Path(r"C:\Program Files (x86)\AutoHotkey\AutoHotkey.exe"),
    Path(r"C:\Program Files (x86)\AutoHotkey\v2\AutoHotkey.exe"),
]

# AHK modifier symbols → Qt modifier names
_AHK_MODS: dict[str, str] = {
    "^": "Ctrl",
    "!": "Alt",
    "+": "Shift",
    "#": "Meta",
}

# AHK special key names → Qt key names (case-insensitive lookup)
_AHK_KEY_MAP: dict[str, str] = {
    "space": "Space",       "enter": "Return",      "return": "Return",
    "tab": "Tab",           "backspace": "Backspace","delete": "Delete",
    "del": "Delete",        "insert": "Insert",      "ins": "Insert",
    "home": "Home",         "end": "End",
    "pgup": "PgUp",         "pgdn": "PgDown",
    "prior": "PgUp",        "next": "PgDown",
    "up": "Up",             "down": "Down",
    "left": "Left",         "right": "Right",
    "esc": "Escape",        "escape": "Escape",
    "capslock": "CapsLock", "scrolllock": "ScrollLock",
    "numlock": "NumLock",   "pause": "Pause",        "break": "Break",
    "printscreen": "Print", "appskey": "Menu",
    "f1": "F1",  "f2": "F2",  "f3": "F3",  "f4": "F4",
    "f5": "F5",  "f6": "F6",  "f7": "F7",  "f8": "F8",
    "f9": "F9",  "f10": "F10","f11": "F11","f12": "F12",
    "f13": "F13","f14": "F14","f15": "F15","f16": "F16",
    "f17": "F17","f18": "F18","f19": "F19","f20": "F20",
    # Numpad
    "numpad0": "Num+0",   "numpad1": "Num+1",   "numpad2": "Num+2",
    "numpad3": "Num+3",   "numpad4": "Num+4",   "numpad5": "Num+5",
    "numpad6": "Num+6",   "numpad7": "Num+7",   "numpad8": "Num+8",
    "numpad9": "Num+9",   "numpadadd": "Num++", "numpadmult": "Num+*",
    "numpadsub": "Num+-", "numpaddiv": "Num+/", "numpaddot": "Num+.",
    "numpadenter": "Enter",
    # Media
    "volume_up": "VolumeUp",    "volume_down": "VolumeDown",
    "volume_mute": "VolumeMute",
    "media_play_pause": "MediaTogglePlayPause",
    "media_next": "MediaNext",  "media_prev": "MediaPrevious",
    "media_stop": "MediaStop",
}

# CC built-in hotkeys for conflict detection (duplicated here as a fallback
# in case the API's list_all() is unavailable on older CC versions)
_CC_BUILTIN_FALLBACK: dict[str, str] = {
    "Ctrl+F":       "Toggle search bar",
    "Escape":       "Close search bar",
    "Ctrl+N":       "New node wizard",
    "Ctrl+Shift+N": "New folder",
    "Ctrl+,":       "Open settings",
    "F1":           "Help / FAQ",
    "Ctrl+Shift+T": "Toggle Time Tracker",
    "Ctrl+Shift+B": "Open Notebook",
    "Ctrl+Shift+M": "Open Media Library",
    "Ctrl+B":       "Notebook: bold",
    "Ctrl+I":       "Notebook: italic",
    "Ctrl+U":       "Notebook: underline",
}


# ---------------------------------------------------------------------------
# Plugin entry points
# ---------------------------------------------------------------------------

def activate(api):
    global _api, _footer_btn, _ahk_exe, _poll_timer

    _api = api
    _load_settings()

    _ahk_exe = _find_ahk_exe()
    _api.log(f"[AHK Bridge] AHK executable: {_ahk_exe or 'NOT FOUND'}")

    # Footer button
    _footer_btn = api.ui.add_footer_button("AHK", _open_dashboard)

    # Global hotkey (Ctrl+Shift+A — unlikely to conflict)
    api.hotkeys.register("Ctrl+Shift+A", _open_dashboard, "AHK Bridge: open dashboard")

    # Poll running processes every 3 s to refresh status
    _poll_timer = api.timers.create(3000, _poll_processes, single_shot=False)

    # Auto-start scripts marked for auto-start
    _autostart_scripts()

    if _ahk_exe is None:
        api.toast(
            "AutoHotkey not found. Open AHK Bridge (footer) to set the path.",
            "warn",
        )
    else:
        api.log(f"[AHK Bridge] activated — {len(_scripts)} script(s) registered.")


def deactivate():
    global _footer_btn, _dashboard, _poll_timer

    # Stop all running scripts
    _stop_all_scripts()

    # Remove footer button
    if _footer_btn is not None:
        _api.ui.remove_footer_button(_footer_btn)
        _footer_btn = None

    # Unregister hotkey
    try:
        _api.hotkeys.unregister("Ctrl+Shift+A")
    except Exception:
        pass

    # Close dashboard
    if _dashboard is not None:
        try:
            if _dashboard.isVisible():
                _dashboard.close()
        except Exception:
            pass
    _dashboard = None

    _api.log("[AHK Bridge] deactivated.")


# ---------------------------------------------------------------------------
# AutoHotkey finder
# ---------------------------------------------------------------------------

def _find_ahk_exe() -> Optional[Path]:
    """Locate the AutoHotkey executable (prefers v2).  Returns Path or None."""

    # 1. User-configured path takes priority
    stored = _api.settings.value(_KEY_AHK_EXE, "").strip()
    if stored:
        p = Path(stored)
        if p.is_file():
            return p
        _api.log(f"[AHK Bridge] Configured AHK path does not exist: {stored}")

    # 2. Windows registry
    try:
        import winreg
        keys = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\AutoHotkey"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\AutoHotkey"),
            (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\AutoHotkey"),
        ]
        candidates_rel = [
            "AutoHotkey64.exe", "AutoHotkey.exe",
            r"v2\AutoHotkey64.exe", r"v2\AutoHotkey.exe",
        ]
        for root, subkey in keys:
            try:
                with winreg.OpenKey(root, subkey) as k:
                    install_dir = winreg.QueryValueEx(k, "InstallDir")[0]
                    for rel in candidates_rel:
                        p = Path(install_dir) / rel
                        if p.is_file():
                            return p
            except OSError:
                pass
    except ImportError:
        pass
    except Exception as exc:
        _api.log(f"[AHK Bridge] Registry search error: {exc}")

    # 3. Well-known install paths
    for p in _AHK_SEARCH_PATHS:
        if p.is_file():
            return p

    # 4. PATH (where autohotkey / where ahk)
    for cmd in ("autohotkey", "autohotkey64", "ahk"):
        try:
            result = subprocess.run(
                ["where", cmd],
                capture_output=True, text=True, timeout=4,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            if result.returncode == 0:
                for line in result.stdout.strip().splitlines():
                    p = Path(line.strip())
                    if p.is_file():
                        return p
        except (OSError, subprocess.TimeoutExpired):
            pass

    return None


def _get_ahk_version(exe: Path) -> str:
    """Return a human-readable version string for the given AHK executable."""
    # Try Windows file version info
    try:
        import ctypes
        import ctypes.wintypes
        size = ctypes.windll.version.GetFileVersionInfoSizeW(str(exe), None)
        if size > 0:
            buf = ctypes.create_string_buffer(size)
            ctypes.windll.version.GetFileVersionInfoW(str(exe), None, size, buf)
            ver_ptr = ctypes.c_void_p()
            ver_len = ctypes.c_uint()
            ctypes.windll.version.VerQueryValueW(
                buf, "\\", ctypes.byref(ver_ptr), ctypes.byref(ver_len)
            )
            class VS_FIXEDFILEINFO(ctypes.Structure):
                _fields_ = [
                    ("dwSignature",        ctypes.c_uint32),
                    ("dwStrucVersion",     ctypes.c_uint32),
                    ("dwFileVersionMS",    ctypes.c_uint32),
                    ("dwFileVersionLS",    ctypes.c_uint32),
                    ("dwProductVersionMS", ctypes.c_uint32),
                    ("dwProductVersionLS", ctypes.c_uint32),
                    ("dwFileFlagsMask",    ctypes.c_uint32),
                    ("dwFileFlags",        ctypes.c_uint32),
                    ("dwFileOS",           ctypes.c_uint32),
                    ("dwFileType",         ctypes.c_uint32),
                    ("dwFileSubtype",      ctypes.c_uint32),
                    ("dwFileDateMS",       ctypes.c_uint32),
                    ("dwFileDateLS",       ctypes.c_uint32),
                ]
            info = ctypes.cast(ver_ptr, ctypes.POINTER(VS_FIXEDFILEINFO)).contents
            ms = info.dwFileVersionMS
            ls = info.dwFileVersionLS
            return f"{ms >> 16}.{ms & 0xFFFF}.{ls >> 16}.{ls & 0xFFFF}"
    except Exception:
        pass

    # Fallback: infer from path
    parts_lower = [p.lower() for p in exe.parts]
    if "v2" in parts_lower:
        return "2.x (path hint)"
    return "Unknown"


# ---------------------------------------------------------------------------
# .ahk script parser
# ---------------------------------------------------------------------------

# Hotkey label pattern:  [modifiers][key]::
# Handles: ^+t::  !F1::  #!a::  ~^+Del::
_RE_HOTKEY_LABEL = re.compile(
    r'^[ \t]*'
    r'([#!^+~\*<>]{0,8})'                # optional modifier chars
    r'('
    r'[Ff]\d{1,2}'                        # F1–F20
    r'|[A-Za-z0-9]'                       # single alphanumeric
    r'|(?:numpad\w+|volume_\w+|media_\w+' # special names
    r'|space|enter|return|tab|backspace'
    r'|delete|del|insert|ins|home|end'
    r'|pgup|pgdn|prior|next|esc|escape'
    r'|capslock|scrolllock|numlock|pause|break'
    r'|printscreen|appskey|lbutton|rbutton|mbutton)'
    r')'
    r'[ \t]*::',
    re.IGNORECASE | re.MULTILINE,
)

# Hotkey("seq", func) / Hotkey 'seq', func
_RE_HOTKEY_FN = re.compile(
    r'\bHotkey\s*[(\s]\s*["\']'
    r'([#!^+~\*<>]{0,8}[A-Za-z0-9_]{1,40})'
    r'["\']',
    re.IGNORECASE,
)

# Hotstring  ::abbrev::replacement  — these are NOT hotkeys, skip them
_RE_HOTSTRING = re.compile(r'^[ \t]*:([^:]*):')


def parse_ahk_hotkeys(script_path: str) -> tuple[list[dict], list[str]]:
    """Parse an .ahk v2 script and extract hotkey definitions.

    Returns (hotkeys, warnings) where:
      hotkeys  — list of {"raw", "normalized", "line", "source"}
      warnings — list of human-readable parse problems
    """
    hotkeys: list[dict] = []
    warnings: list[str] = []

    try:
        text = Path(script_path).read_text(encoding="utf-8-sig", errors="replace")
    except OSError as exc:
        return [], [f"Cannot read file: {exc}"]

    lines = text.splitlines()
    for lineno, line in enumerate(lines, 1):
        stripped = line.strip()

        # Skip blank lines and full-line comments
        if not stripped or stripped.startswith(";"):
            continue

        # Strip inline comment
        code = re.sub(r'\s*;.*$', '', stripped)
        if not code:
            continue

        # Skip hotstrings (:options:abbrev::replacement)
        if _RE_HOTSTRING.match(code):
            continue

        # ── Pattern 1: bare hotkey label ──────────────────────────────────
        m = _RE_HOTKEY_LABEL.match(line)
        if m:
            mods_raw = m.group(1)
            key_raw  = m.group(2)
            norm, warn = _normalize_ahk_hotkey(mods_raw, key_raw)
            if warn:
                warnings.append(f"Line {lineno}: {warn}")
            if norm:
                hotkeys.append({
                    "raw":        f"{mods_raw}{key_raw}",
                    "normalized": norm,
                    "line":       lineno,
                    "source":     str(script_path),
                    "is_label":   True,
                })
            continue

        # ── Pattern 2: Hotkey() function call ────────────────────────────
        for m2 in _RE_HOTKEY_FN.finditer(code):
            seq = m2.group(1)
            # Split off leading modifiers
            mod_chars = ""
            key_part  = seq
            for i, ch in enumerate(seq):
                if ch in "#!^+~*<>":
                    mod_chars += ch
                else:
                    key_part = seq[i:]
                    break
            norm, warn = _normalize_ahk_hotkey(mod_chars, key_part)
            if warn:
                warnings.append(f"Line {lineno}: {warn}")
            if norm:
                hotkeys.append({
                    "raw":        seq,
                    "normalized": norm,
                    "line":       lineno,
                    "source":     str(script_path),
                    "is_label":   False,
                })

    return hotkeys, warnings


def _normalize_ahk_hotkey(mods_raw: str, key_raw: str) -> tuple[str, str]:
    """Convert AHK modifier+key to Qt-style "Ctrl+Shift+A" format.

    Returns (normalized_string, warning_string).
    normalized_string is empty if mapping failed.
    """
    parts: list[str] = []
    warn = ""

    # Modifiers (canonical order: Ctrl, Alt, Shift, Meta/Win)
    if "^" in mods_raw:
        parts.append("Ctrl")
    if "!" in mods_raw:
        parts.append("Alt")
    if "+" in mods_raw:
        parts.append("Shift")
    if "#" in mods_raw:
        parts.append("Meta")

    key_lower = key_raw.lower().strip()

    mapped = _AHK_KEY_MAP.get(key_lower)
    if mapped:
        parts.append(mapped)
    elif len(key_raw) == 1 and key_raw.isalpha():
        # Single letter — uppercase for display
        parts.append(key_raw.upper())
    elif len(key_raw) == 1 and (key_raw.isdigit() or key_raw in r"`-=[]\;',./ "):
        parts.append(key_raw)
    elif key_lower.startswith("sc") and len(key_lower) <= 6:
        # AHK scan code (e.g. sc013) — include raw, can't map to Qt name
        parts.append(key_raw.upper())
        warn = f"Scan code '{key_raw}' cannot be reliably mapped to a Qt key name."
    elif key_lower.startswith("vk") and len(key_lower) <= 6:
        # Virtual key code — include raw
        parts.append(key_raw.upper())
        warn = f"Virtual key code '{key_raw}' cannot be reliably mapped to a Qt key name."
    else:
        # Unknown key — include raw, flag it
        parts.append(key_raw)
        warn = f"Unknown key name '{key_raw}' — using raw value."

    return ("+".join(parts) if parts else "", warn)


# ---------------------------------------------------------------------------
# Conflict checker
# ---------------------------------------------------------------------------

def _get_all_cc_hotkeys() -> dict[str, str]:
    """Return {sequence: source_description} for all active CC hotkeys.

    Reads live user-configured hotkeys from the API so custom remaps are
    reflected.  Falls back to the static built-in list only when the API
    is unavailable.
    """
    try:
        all_hk = _api.hotkeys.list_all()
        result: dict[str, str] = {}
        for seq, info in all_hk.items():
            result[seq] = (
                f"{info.get('source', 'Command Center')}: "
                f"{info.get('description', '')}"
            )
        if result:
            return result
        # Empty response — fall through to static fallback
    except AttributeError:
        pass  # CC version without list_all()
    except Exception as exc:
        _api.log(f"[AHK Bridge] conflict check warning: {exc}")
    return dict(_CC_BUILTIN_FALLBACK)


def check_conflicts(hotkeys: list[dict]) -> list[dict]:
    """Annotate each hotkey dict with a 'conflicts' list."""
    registry = _get_all_cc_hotkeys()
    for entry in hotkeys:
        norm = entry.get("normalized", "")
        conflicts: list[dict] = []
        if norm and norm in registry:
            conflicts.append({"sequence": norm, "cc_source": registry[norm]})
        entry["conflicts"] = conflicts
    return hotkeys


# ---------------------------------------------------------------------------
# AHK hotkey validator & in-file remapper
# ---------------------------------------------------------------------------

def _validate_ahk_raw(seq: str) -> tuple[bool, str]:
    """Return (True, '') if seq is a valid AHK 2.0 hotkey, else (False, reason)."""
    if not seq or not seq.strip():
        return False, "Sequence cannot be empty."
    seq = seq.strip()
    mod_chars = ""
    key_part = seq
    for i, ch in enumerate(seq):
        if ch in "#!^+~*<>":
            mod_chars += ch
        else:
            key_part = seq[i:]
            break
    if not key_part:
        return False, "No key specified (only modifier characters found)."
    key_lower = key_part.lower()
    is_valid = (
        (len(key_part) == 1 and (key_part.isalnum() or key_part in r"`-=[]\;',./"))
        or key_lower in _AHK_KEY_MAP
        or (key_lower.startswith("sc") and len(key_lower) <= 6)
        or (key_lower.startswith("vk") and len(key_lower) <= 6)
    )
    if not is_valid:
        return False, (
            f"Unrecognized key name: '{key_part}'.\n"
            "Use a letter, digit, F-key (F1–F20), or a named key such as\n"
            "Space, Tab, Enter, Delete, Home, End, PgUp, PgDn, Esc, etc."
        )
    return True, ""


def _remap_ahk_hotkey(
    script_path: str,
    line_num: int,
    is_label: bool,
    old_raw: str,
    new_raw: str,
) -> tuple[bool, str]:
    """Rewrite line_num in script_path to use new_raw as the hotkey.

    For label hotkeys (is_label=True):  ^+t::action  →  ^+r::action
    For Hotkey() calls (is_label=False): Hotkey("^+t", fn)  →  Hotkey("^+r", fn)

    Returns (ok, message).
    """
    try:
        p = Path(script_path)
        text = p.read_text(encoding="utf-8-sig", errors="replace")
        lines = text.splitlines(keepends=True)
        if line_num < 1 or line_num > len(lines):
            return False, (
                f"Line {line_num} is out of range "
                f"(file has {len(lines)} lines).\n"
                "Please re-scan and try again."
            )
        idx = line_num - 1
        original = lines[idx]

        if is_label:
            # Replace ^+t:: → ^+r:: at the start of the line
            new_line = re.sub(
                r'^([ \t]*)' + re.escape(old_raw) + r'([ \t]*:[:$]?)',
                lambda m: m.group(1) + new_raw + m.group(2),
                original,
            )
        else:
            # Replace quoted old_raw inside a Hotkey() call
            new_line = re.sub(
                r'(["\'])' + re.escape(old_raw) + r'\1',
                f'"{new_raw}"',
                original,
                count=1,
            )

        if new_line == original:
            return False, (
                f"Could not locate '{old_raw}' on line {line_num}.\n\n"
                "The file may have been modified since the last scan.\n"
                "Please re-scan and try again."
            )

        lines[idx] = new_line
        p.write_text("".join(lines), encoding="utf-8")
        _api.log(
            f"[AHK Bridge] Remapped '{old_raw}' → '{new_raw}' "
            f"in {p.name} line {line_num}"
        )
        return True, f"Remapped '{old_raw}' → '{new_raw}' on line {line_num}."
    except OSError as exc:
        return False, f"Cannot write to file:\n{exc}"
    except Exception as exc:
        return False, f"Unexpected error:\n{exc}"


# ---------------------------------------------------------------------------
# Process manager
# ---------------------------------------------------------------------------

def _run_script(path: str) -> tuple[bool, str]:
    """Start an AHK script as a background subprocess.  Returns (ok, message)."""
    global _ahk_exe

    if _ahk_exe is None:
        _ahk_exe = _find_ahk_exe()
    if _ahk_exe is None:
        return False, (
            "AutoHotkey executable not found.\n"
            "Set the path in AHK Bridge → Config tab."
        )

    p = Path(path)
    if not p.is_file():
        return False, f"Script file not found:\n{path}"

    # Check already running
    proc = _procs.get(path)
    if proc is not None and proc.poll() is None:
        return False, f"'{p.name}' is already running (PID {proc.pid})."

    try:
        proc = subprocess.Popen(
            [str(_ahk_exe), str(p)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        _procs[path] = proc
        _api.log(f"[AHK Bridge] Started '{p.name}' (PID {proc.pid})")
        return True, f"Started '{p.name}'  (PID {proc.pid})"
    except OSError as exc:
        return False, f"Failed to start script: {exc}"
    except Exception as exc:
        return False, f"Unexpected error starting script: {exc}"


def _stop_script(path: str) -> tuple[bool, str]:
    """Terminate a running AHK script.  Returns (ok, message)."""
    proc = _procs.get(path)
    if proc is None:
        return False, "Script is not running."

    ret = proc.poll()
    if ret is not None:
        _procs.pop(path, None)
        return False, f"Script had already exited (code {ret})."

    try:
        proc.terminate()
        try:
            proc.wait(timeout=4)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait(timeout=2)
    except Exception as exc:
        return False, f"Error stopping script: {exc}"

    _procs.pop(path, None)
    _api.log(f"[AHK Bridge] Stopped '{Path(path).name}'")
    return True, f"Stopped '{Path(path).name}'"


def _stop_all_scripts():
    for path in list(_procs.keys()):
        try:
            _stop_script(path)
        except Exception:
            pass


def _is_running(path: str) -> bool:
    proc = _procs.get(path)
    return proc is not None and proc.poll() is None


def _poll_processes():
    """Clean up stale proc entries; called by the repeating timer."""
    for path in list(_procs.keys()):
        proc = _procs[path]
        if proc.poll() is not None:
            _procs.pop(path, None)
            _api.log(f"[AHK Bridge] Script exited: {Path(path).name}")


# ---------------------------------------------------------------------------
# Script auto-start
# ---------------------------------------------------------------------------

def _autostart_scripts():
    for rec in _scripts:
        if rec.get("auto_start") and rec.get("path"):
            ok, msg = _run_script(rec["path"])
            if ok:
                _api.log(f"[AHK Bridge] Auto-started: {Path(rec['path']).name}")
            else:
                _api.log(f"[AHK Bridge] Auto-start failed ({Path(rec['path']).name}): {msg}")


# ---------------------------------------------------------------------------
# Settings I/O
# ---------------------------------------------------------------------------

def _load_settings():
    global _scripts
    raw = _api.settings.value(_KEY_SCRIPTS, "").strip()
    if raw:
        try:
            loaded = json.loads(raw)
            if isinstance(loaded, list):
                # Validate entries
                clean = []
                for item in loaded:
                    if isinstance(item, dict) and "path" in item:
                        clean.append({
                            "path":       str(item["path"]),
                            "auto_start": bool(item.get("auto_start", False)),
                        })
                _scripts = clean
                return
        except (json.JSONDecodeError, TypeError, KeyError) as exc:
            _api.log(f"[AHK Bridge] Settings parse error: {exc}")
    _scripts = []


def _save_settings():
    try:
        _api.settings.set(_KEY_SCRIPTS, json.dumps(_scripts))
    except Exception as exc:
        _api.log(f"[AHK Bridge] Settings save error: {exc}")


# ---------------------------------------------------------------------------
# Dashboard — main entry
# ---------------------------------------------------------------------------

def _open_dashboard():
    global _dashboard
    if _dashboard is not None:
        try:
            if _dashboard.isVisible():
                _dashboard.raise_()
                _dashboard.activateWindow()
                return
        except RuntimeError:
            pass  # underlying C++ object deleted
    _dashboard = _AHKDashboard()
    _api.ui.show_dialog(_dashboard)


# ---------------------------------------------------------------------------
# Dashboard widget
# ---------------------------------------------------------------------------

class _AHKDashboard(QDialog):
    def __init__(self):
        parent = _api.ui.main_window
        super().__init__(parent)
        self.setWindowTitle("AutoHotkey 2.0 Bridge")
        self.setMinimumSize(720, 560)
        self.resize(860, 640)
        self._build_ui()

        # Refresh status every 2 s while the dialog is open
        self._refresh_timer = QTimer(self)
        self._refresh_timer.setInterval(2000)
        self._refresh_timer.timeout.connect(self._refresh_status)
        self._refresh_timer.start()

    def closeEvent(self, event):
        self._refresh_timer.stop()
        super().closeEvent(event)

    # ── UI construction ───────────────────────────────────────────────────

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._make_header())

        tabs = QTabWidget()
        tabs.setStyleSheet(_tab_style())
        tabs.addTab(self._make_scripts_tab(), "  Scripts  ")
        tabs.addTab(self._make_hotkeys_tab(), "  Hotkeys & Conflicts  ")
        tabs.addTab(self._make_config_tab(),  "  Config  ")
        root.addWidget(tabs, stretch=1)

        root.addWidget(self._make_footer())

    def _make_header(self) -> QWidget:
        hdr = QWidget()
        hdr.setFixedHeight(54)
        hdr.setStyleSheet("background:#0d1117; border-bottom:1px solid #30363d;")
        hl = QHBoxLayout(hdr)
        hl.setContentsMargins(18, 0, 18, 0)

        title = QLabel("⚙  AutoHotkey 2.0 Bridge")
        title.setFont(QFont("Segoe UI", 11, QFont.Weight.DemiBold))
        title.setStyleSheet("color:#58a6ff; background:transparent;")
        hl.addWidget(title)
        hl.addStretch()

        # AHK status pill
        self._ahk_status_lbl = QLabel()
        self._ahk_status_lbl.setFont(QFont("Segoe UI", 8))
        self._ahk_status_lbl.setStyleSheet(
            "background:#161b22; border:1px solid #30363d; border-radius:10px;"
            "padding:2px 10px; color:#8b949e;"
        )
        self._update_ahk_pill()
        hl.addWidget(self._ahk_status_lbl)
        return hdr

    def _update_ahk_pill(self):
        global _ahk_exe
        if _ahk_exe and _ahk_exe.is_file():
            ver = _get_ahk_version(_ahk_exe)
            self._ahk_status_lbl.setText(f"✔  AHK {ver}  —  {_ahk_exe.name}")
            self._ahk_status_lbl.setStyleSheet(
                "background:#0d2818; border:1px solid #238636; border-radius:10px;"
                "padding:2px 10px; color:#3fb950;"
            )
        else:
            self._ahk_status_lbl.setText("✘  AutoHotkey not found")
            self._ahk_status_lbl.setStyleSheet(
                "background:#2d0f0f; border:1px solid #6e1717; border-radius:10px;"
                "padding:2px 10px; color:#f85149;"
            )

    # ── Scripts tab ───────────────────────────────────────────────────────

    def _make_scripts_tab(self) -> QWidget:
        w = QWidget()
        w.setStyleSheet("background:#0d1117;")
        vl = QVBoxLayout(w)
        vl.setContentsMargins(14, 12, 14, 10)
        vl.setSpacing(8)

        # Toolbar
        tb = QHBoxLayout()
        add_btn = _mk_btn("＋  Add Script", "#1f6feb", "#388bfd")
        add_btn.clicked.connect(self._on_add_script)
        tb.addWidget(add_btn)

        run_all = _mk_btn("▶  Run All", "#1a3a1a", "#3fb950")
        run_all.clicked.connect(self._on_run_all)
        tb.addWidget(run_all)

        stop_all = _mk_btn("■  Stop All", "#3a1515", "#f85149")
        stop_all.clicked.connect(self._on_stop_all)
        tb.addWidget(stop_all)

        tb.addStretch()
        vl.addLayout(tb)

        # Script list area (scrollable)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea{background:#0d1117;border:none;}")
        self._scripts_inner = QWidget()
        self._scripts_inner.setStyleSheet("background:#0d1117;")
        self._scripts_layout = QVBoxLayout(self._scripts_inner)
        self._scripts_layout.setContentsMargins(0, 0, 0, 0)
        self._scripts_layout.setSpacing(6)
        self._scripts_layout.addStretch()
        scroll.setWidget(self._scripts_inner)
        vl.addWidget(scroll, stretch=1)

        self._script_rows: list[_ScriptRow] = []
        self._rebuild_script_list()
        return w

    def _rebuild_script_list(self):
        # Remove old rows
        for row in self._script_rows:
            row.setParent(None)
            row.deleteLater()
        self._script_rows.clear()

        # Remove stretch at end
        item = self._scripts_layout.takeAt(self._scripts_layout.count() - 1)
        del item

        for rec in _scripts:
            row = _ScriptRow(rec, self)
            self._scripts_layout.addWidget(row)
            self._script_rows.append(row)

        self._scripts_layout.addStretch()

    def _refresh_status(self):
        _poll_processes()
        for row in self._script_rows:
            row.refresh_status()

    def _on_add_script(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select AutoHotkey Script",
            str(Path.home()), "AHK Scripts (*.ahk);;All Files (*.*)"
        )
        if not path:
            return
        # Avoid duplicates
        if any(r["path"] == path for r in _scripts):
            _api.toast(f"'{Path(path).name}' is already in the list.", "warning")
            return
        _scripts.append({"path": path, "auto_start": False})
        _save_settings()
        self._rebuild_script_list()
        _api.toast(f"Added '{Path(path).name}'", "info")

    def remove_script(self, path: str):
        _stop_script(path)
        _scripts[:] = [r for r in _scripts if r["path"] != path]
        _save_settings()
        self._rebuild_script_list()

    def toggle_auto_start(self, path: str, enabled: bool):
        for rec in _scripts:
            if rec["path"] == path:
                rec["auto_start"] = enabled
                break
        _save_settings()

    def _on_run_all(self):
        ok_count = err_count = 0
        for rec in _scripts:
            ok, msg = _run_script(rec["path"])
            if ok:
                ok_count += 1
            else:
                _api.log(f"[AHK Bridge] Run-all error ({Path(rec['path']).name}): {msg}")
                err_count += 1
        self._refresh_status()
        _api.toast(
            f"Started {ok_count} script(s)."
            + (f"  {err_count} failed — see log." if err_count else ""),
            "info" if not err_count else "warning",
        )

    def _on_stop_all(self):
        _stop_all_scripts()
        self._refresh_status()
        _api.toast("All scripts stopped.", "info")

    # ── Hotkeys tab ───────────────────────────────────────────────────────

    def _make_hotkeys_tab(self) -> QWidget:
        w = QWidget()
        w.setStyleSheet("background:#0d1117;")
        vl = QVBoxLayout(w)
        vl.setContentsMargins(14, 12, 14, 10)
        vl.setSpacing(8)

        tb = QHBoxLayout()
        scan_btn = _mk_btn("⟳  Scan All Scripts", "#1f3a5f", "#58a6ff")
        scan_btn.clicked.connect(self._on_scan_hotkeys)
        tb.addWidget(scan_btn)

        self._conflict_summary = QLabel("")
        self._conflict_summary.setFont(QFont("Segoe UI", 8))
        self._conflict_summary.setStyleSheet("color:#8b949e; background:transparent;")
        tb.addStretch()
        tb.addWidget(self._conflict_summary)
        vl.addLayout(tb)

        # Syntax hint
        hint_lbl = QLabel(
            "\U0001f4a1  AHK modifiers:  ^ = Ctrl   ! = Alt   + = Shift   # = Win"
            "     e.g.  ^r   ^!r   ^+r   ^+F5   !F4   #d"
        )
        hint_lbl.setFont(QFont("Consolas", 8))
        hint_lbl.setStyleSheet("color:#8b949e; background:transparent;")
        vl.addWidget(hint_lbl)

        # Table
        self._hk_table = QTableWidget(0, 6)
        self._hk_table.setHorizontalHeaderLabels(
            ["AHK Hotkey", "Normalized", "Script", "Line", "Conflict", "Remap"]
        )
        self._hk_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self._hk_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self._hk_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self._hk_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self._hk_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self._hk_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Fixed)
        self._hk_table.setColumnWidth(5, 210)
        self._hk_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._hk_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._hk_table.setAlternatingRowColors(True)
        self._hk_table.setStyleSheet(_table_style())
        self._hk_table.verticalHeader().setDefaultSectionSize(34)
        vl.addWidget(self._hk_table, stretch=1)

        # Warnings area
        self._warn_box = QTextEdit()
        self._warn_box.setReadOnly(True)
        self._warn_box.setFixedHeight(80)
        self._warn_box.setPlaceholderText("Parse warnings will appear here…")
        self._warn_box.setStyleSheet(
            "background:#161b22; border:1px solid #30363d; color:#8b949e;"
            "font-family:Consolas,monospace; font-size:8pt; border-radius:4px;"
        )
        vl.addWidget(self._warn_box)
        return w

    def _on_scan_hotkeys(self):
        all_hotkeys: list[dict] = []
        all_warnings: list[str] = []

        if not _scripts:
            _api.toast("No scripts registered. Add scripts on the Scripts tab.", "warning")
            return

        for rec in _scripts:
            hks, warns = parse_ahk_hotkeys(rec["path"])
            all_hotkeys.extend(hks)
            all_warnings.extend([f"[{Path(rec['path']).name}] {w}" for w in warns])

        check_conflicts(all_hotkeys)
        self._populate_hotkeys_table(all_hotkeys)

        if all_warnings:
            self._warn_box.setPlainText("\n".join(all_warnings))
        else:
            self._warn_box.setPlainText("")

        conflict_count = sum(1 for h in all_hotkeys if h.get("conflicts"))
        if conflict_count:
            self._conflict_summary.setText(
                f"⚠  {conflict_count} conflict(s) detected — highlighted in red"
            )
            self._conflict_summary.setStyleSheet("color:#f85149; background:transparent;")
        else:
            self._conflict_summary.setText(
                f"✔  {len(all_hotkeys)} hotkey(s) scanned — no conflicts"
            )
            self._conflict_summary.setStyleSheet("color:#3fb950; background:transparent;")

    def _populate_hotkeys_table(self, hotkeys: list[dict]):
        self._hk_table.setRowCount(0)
        for row_idx, hk in enumerate(hotkeys):
            self._hk_table.insertRow(row_idx)
            conflicts = hk.get("conflicts", [])
            has_conflict = bool(conflicts)
            conflict_text = (
                "\n".join(c["cc_source"] for c in conflicts)
                if has_conflict else "—"
            )
            cells = [
                hk.get("raw", ""),
                hk.get("normalized", ""),
                Path(hk.get("source", "")).name,
                str(hk.get("line", "")),
                conflict_text,
            ]
            fg = QColor("#f85149") if has_conflict else QColor("#c9d1d9")
            for col, text in enumerate(cells):
                item = QTableWidgetItem(text)
                item.setForeground(fg)
                self._hk_table.setItem(row_idx, col, item)
            # Remap column: inline QLineEdit + Apply button
            remap_w = self._make_remap_widget(hk, row_idx)
            self._hk_table.setCellWidget(row_idx, 5, remap_w)

    def _make_remap_widget(self, hk: dict, row_idx: int) -> QWidget:
        """Build the inline remap widget (QLineEdit + Apply button) for a table row."""
        w = QWidget()
        w.setStyleSheet("background:transparent;")
        hl = QHBoxLayout(w)
        hl.setContentsMargins(4, 2, 4, 2)
        hl.setSpacing(4)

        le = QLineEdit()
        le.setPlaceholderText("e.g. ^!r")
        le.setFont(QFont("Consolas", 9))
        le.setStyleSheet(
            "QLineEdit{background:#0d1117; color:#c9d1d9;"
            "border:1px solid #30363d; border-radius:3px; padding:2px 6px;}"
            "QLineEdit:focus{border-color:#58a6ff;}"
        )
        le.setToolTip(
            "Enter a new AHK 2.0 hotkey sequence, then click Apply.\n"
            "Modifiers:  ^ = Ctrl   ! = Alt   + = Shift   # = Win\n"
            "Examples:  ^r   ^!r   ^+F5   !F4   #d   ^+Del"
        )
        hl.addWidget(le, stretch=1)

        apply_btn = QPushButton("Apply")
        apply_btn.setFixedSize(52, 24)
        apply_btn.setCursor(Qt.PointingHandCursor)
        apply_btn.setFont(QFont("Segoe UI", 8))
        apply_btn.setStyleSheet(
            "QPushButton{background:#1f3a1f; color:#3fb950;"
            "border:1px solid #3fb95044; border-radius:3px; padding:0;}"
            "QPushButton:hover{background:#3fb95033; border-color:#3fb950;}"
        )
        apply_btn.setToolTip("Write hotkey change to the .ahk file")
        apply_btn.clicked.connect(lambda: self._apply_remap(hk, le, row_idx))
        hl.addWidget(apply_btn)
        return w

    def _apply_remap(self, hk: dict, le: QLineEdit, row_idx: int):
        """Validate and write the new hotkey to the .ahk file, then update the row."""
        new_raw = le.text().strip()
        if not new_raw:
            QMessageBox.warning(self, "Empty Sequence",
                                "Please enter a new hotkey sequence first.")
            return

        old_raw = hk.get("raw", "")
        if new_raw == old_raw:
            QMessageBox.information(self, "No Change",
                                    "The new sequence is the same as the current one.")
            return

        ok, err = _validate_ahk_raw(new_raw)
        if not ok:
            QMessageBox.warning(
                self, "Invalid AHK Syntax",
                f"'{new_raw}' is not a valid AHK 2.0 hotkey sequence.\n\n"
                f"{err}\n\n"
                "Modifier symbols:  ^ = Ctrl   ! = Alt   + = Shift   # = Win\n"
                "Examples:  ^r   ^!r   ^+F5   !F4",
            )
            return

        script_name = Path(hk.get("source", "")).name
        line_num    = hk.get("line", 0)
        is_label    = hk.get("is_label", True)

        if not old_raw:
            QMessageBox.warning(self, "No Original Hotkey",
                                "The original hotkey is missing.\n"
                                "Please re-scan scripts and try again.")
            return

        reply = QMessageBox.question(
            self, "Confirm Remap",
            f"Remap hotkey in  {script_name}  (line {line_num})?\n\n"
            f"  Current:  {old_raw}\n"
            f"  New:      {new_raw}\n\n"
            "This will modify the .ahk file on disk.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        write_ok, msg = _remap_ahk_hotkey(
            hk.get("source", ""), line_num, is_label, old_raw, new_raw
        )
        if write_ok:
            hk["raw"] = new_raw
            cell = self._hk_table.item(row_idx, 0)
            if cell:
                cell.setText(new_raw)
            le.clear()
            le.setPlaceholderText(f"\u2713 was: {old_raw}")
            _api.toast(
                f"Remapped {old_raw} \u2192 {new_raw} in {script_name}.  "
                "Stop and restart the script in the Scripts tab to apply.",
                "info",
            )
        else:
            QMessageBox.critical(self, "Remap Failed", msg)

    # ── Config tab ────────────────────────────────────────────────────────

    def _make_config_tab(self) -> QWidget:
        w = QWidget()
        w.setStyleSheet("background:#0d1117;")
        vl = QVBoxLayout(w)
        vl.setContentsMargins(18, 16, 18, 16)
        vl.setSpacing(14)

        # AHK exe path
        vl.addWidget(_section_label("AutoHotkey Executable"))
        path_row = QHBoxLayout()
        self._ahk_path_edit = QLineEdit()
        self._ahk_path_edit.setPlaceholderText("Leave blank to auto-detect…")
        self._ahk_path_edit.setText(
            str(_ahk_exe) if _ahk_exe and _ahk_exe.is_file() else
            _api.settings.value(_KEY_AHK_EXE, "")
        )
        self._ahk_path_edit.setStyleSheet(_input_style())
        path_row.addWidget(self._ahk_path_edit, stretch=1)
        browse_btn = _mk_btn("Browse…", "#1f3a1f", "#3fb950")
        browse_btn.setFixedWidth(90)
        browse_btn.clicked.connect(self._browse_ahk_exe)
        path_row.addWidget(browse_btn)
        re_detect = _mk_btn("Auto-detect", "#1f1f3a", "#58a6ff")
        re_detect.setFixedWidth(100)
        re_detect.clicked.connect(self._on_re_detect)
        path_row.addWidget(re_detect)
        vl.addLayout(path_row)

        note = QLabel(
            "AHK v2 is required.  Typical location: "
            "C:\\Program Files\\AutoHotkey\\v2\\AutoHotkey64.exe"
        )
        note.setWordWrap(True)
        note.setFont(QFont("Segoe UI", 8))
        note.setStyleSheet("color:#8b949e; background:transparent;")
        vl.addWidget(note)

        _sep(vl)

        save_btn = _mk_btn("💾  Save Config", "#1f3a5f", "#388bfd")
        save_btn.clicked.connect(self._save_config)
        vl.addWidget(save_btn)

        vl.addStretch()

        # CC hotkeys reference
        vl.addWidget(_section_label("Command Center Hotkeys  (for reference)"))
        cc_table = QTableWidget(0, 2)
        cc_table.setHorizontalHeaderLabels(["Hotkey", "Action"])
        cc_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        cc_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        cc_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        cc_table.setMaximumHeight(180)
        cc_table.setStyleSheet(_table_style())
        cc_table.setAlternatingRowColors(True)
        registry = _get_all_cc_hotkeys()
        for i, (seq, desc) in enumerate(sorted(registry.items())):
            cc_table.insertRow(i)
            cc_table.setItem(i, 0, QTableWidgetItem(seq))
            cc_table.setItem(i, 1, QTableWidgetItem(desc))
        vl.addWidget(cc_table)

        return w

    def _browse_ahk_exe(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Locate AutoHotkey Executable",
            r"C:\Program Files\AutoHotkey",
            "Executables (*.exe);;All Files (*.*)"
        )
        if path:
            self._ahk_path_edit.setText(path)

    def _on_re_detect(self):
        global _ahk_exe
        _api.settings.remove(_KEY_AHK_EXE)
        _ahk_exe = _find_ahk_exe()
        if _ahk_exe:
            self._ahk_path_edit.setText(str(_ahk_exe))
            _api.toast(f"Found: {_ahk_exe}", "info")
        else:
            self._ahk_path_edit.clear()
            _api.toast("AutoHotkey not found on this system.", "warning")
        self._update_ahk_pill()

    def _save_config(self):
        global _ahk_exe
        raw = self._ahk_path_edit.text().strip()
        if raw:
            p = Path(raw)
            if not p.is_file():
                QMessageBox.warning(
                    self, "Invalid Path",
                    f"The specified file does not exist:\n{raw}\n\n"
                    "Please browse to the AutoHotkey64.exe file."
                )
                return
            _api.settings.set(_KEY_AHK_EXE, raw)
            _ahk_exe = p
        else:
            _api.settings.remove(_KEY_AHK_EXE)
            _ahk_exe = _find_ahk_exe()
        self._update_ahk_pill()
        _api.toast("Config saved.", "info")

    # ── Footer bar ────────────────────────────────────────────────────────

    def _make_footer(self) -> QWidget:
        foot = QWidget()
        foot.setFixedHeight(50)
        foot.setStyleSheet("background:#161b22; border-top:1px solid #30363d;")
        fl = QHBoxLayout(foot)
        fl.setContentsMargins(14, 0, 14, 0)
        fl.setSpacing(10)

        help_btn = _mk_btn("?  Help", "#1a1a2e", "#8b949e")
        help_btn.clicked.connect(self._show_help)
        fl.addWidget(help_btn)

        fl.addStretch()

        close_btn = _mk_btn("Close", "#1a1a1a", "#8b949e")
        close_btn.clicked.connect(self.close)
        fl.addWidget(close_btn)
        return foot

    def _show_help(self):
        _api.ui.show_message(
            "AHK Bridge \u2014 Help",
            "Scripts tab:\n"
            "  Add .ahk script files using '+ Add Script'.\n"
            "  \u25b6 / \u25a0 buttons run or stop individual scripts.\n"
            "  \u270e Edit \u2014 opens the script in your default text editor.\n"
            "  Auto-start \u2014 launches the script automatically when the\n"
            "  AHK Bridge plugin is activated.\n\n"
            "Hotkeys & Conflicts tab:\n"
            "  Click '\u27f3 Scan All Scripts' to parse all registered scripts.\n"
            "  Hotkeys that conflict with Command Center shortcuts\n"
            "  (including user-customized hotkeys) are highlighted in red.\n"
            "  In the Remap column, type a new AHK 2.0 sequence and click\n"
            "  Apply to edit the hotkey directly in the .ahk file.\n"
            "  After remapping, stop then restart the script (Scripts tab).\n"
            "  AHK syntax:  ^ = Ctrl   ! = Alt   + = Shift   # = Win\n"
            "  Examples:  ^r   ^!r   ^+F5   !F4   #d   ^+Del\n\n"
            "Config tab:\n"
            "  Set or auto-detect the AutoHotkey v2 executable.\n"
            "  All active Command Center hotkeys are shown for reference.\n\n"
            "Note: AHK scripts run as separate processes managed by\n"
            "AutoHotkey \u2014 this plugin provides the launcher, editor,\n"
            "conflict-detection, and in-place remap layer."
        )


# ---------------------------------------------------------------------------
# Script row widget
# ---------------------------------------------------------------------------

class _ScriptRow(QFrame):
    def __init__(self, rec: dict, dashboard: _AHKDashboard):
        super().__init__()
        self._rec = rec
        self._dash = dashboard
        self._path = rec["path"]
        self.setStyleSheet(
            "QFrame{background:#161b22; border:1px solid #30363d;"
            "border-radius:6px;}"
            "QFrame:hover{border-color:#58a6ff;}"
        )
        self._build()

    def _build(self):
        hl = QHBoxLayout(self)
        hl.setContentsMargins(10, 6, 10, 6)
        hl.setSpacing(8)

        # Status dot
        self._status_dot = QLabel("●")
        self._status_dot.setFixedWidth(14)
        self._status_dot.setFont(QFont("Segoe UI", 10))
        hl.addWidget(self._status_dot)

        # Script name + path
        info = QVBoxLayout()
        info.setSpacing(1)
        name_lbl = QLabel(Path(self._path).name)
        name_lbl.setFont(QFont("Segoe UI", 9, QFont.Weight.DemiBold))
        name_lbl.setStyleSheet("color:#c9d1d9; background:transparent; border:none;")
        info.addWidget(name_lbl)
        path_lbl = QLabel(self._path)
        path_lbl.setFont(QFont("Segoe UI", 7))
        path_lbl.setStyleSheet("color:#8b949e; background:transparent; border:none;")
        path_lbl.setWordWrap(False)
        info.addWidget(path_lbl)
        hl.addLayout(info, stretch=1)

        # Auto-start checkbox
        self._auto_cb = QCheckBox("Auto-start")
        self._auto_cb.setChecked(self._rec.get("auto_start", False))
        self._auto_cb.setStyleSheet("color:#8b949e; background:transparent; border:none;")
        self._auto_cb.stateChanged.connect(
            lambda state: self._dash.toggle_auto_start(
                self._path, bool(state)
            )
        )
        hl.addWidget(self._auto_cb)

        # Run / Stop buttons
        self._run_btn = _mk_btn("▶", "#1a3a1a", "#3fb950")
        self._run_btn.setFixedSize(30, 26)
        self._run_btn.setToolTip("Run script")
        self._run_btn.clicked.connect(self._on_run)
        hl.addWidget(self._run_btn)

        self._stop_btn = _mk_btn("■", "#3a1515", "#f85149")
        self._stop_btn.setFixedSize(30, 26)
        self._stop_btn.setToolTip("Stop script")
        self._stop_btn.clicked.connect(self._on_stop)
        hl.addWidget(self._stop_btn)

        # Edit button
        edit_btn = _mk_btn("✎", "#1a2a3a", "#8b949e")
        edit_btn.setFixedSize(28, 26)
        edit_btn.setToolTip("Open in default text editor")
        edit_btn.clicked.connect(self._on_edit)
        hl.addWidget(edit_btn)

        # Remove button
        rem_btn = _mk_btn("✕", "#1a1a1a", "#8b949e")
        rem_btn.setFixedSize(28, 26)
        rem_btn.setToolTip("Remove from list")
        rem_btn.clicked.connect(self._on_remove)
        hl.addWidget(rem_btn)

        self.refresh_status()

    def refresh_status(self):
        running = _is_running(self._path)
        if running:
            self._status_dot.setStyleSheet("color:#3fb950; background:transparent; border:none;")
            self._status_dot.setToolTip("Running")
        else:
            exists = Path(self._path).is_file()
            if exists:
                self._status_dot.setStyleSheet("color:#8b949e; background:transparent; border:none;")
                self._status_dot.setToolTip("Stopped")
            else:
                self._status_dot.setStyleSheet("color:#f85149; background:transparent; border:none;")
                self._status_dot.setToolTip("File not found")
        self._run_btn.setEnabled(not running)
        self._stop_btn.setEnabled(running)

    def _on_run(self):
        ok, msg = _run_script(self._path)
        self.refresh_status()
        _api.toast(msg, "info" if ok else "error")

    def _on_stop(self):
        ok, msg = _stop_script(self._path)
        self.refresh_status()
        _api.toast(msg, "info" if ok else "warning")

    def _on_edit(self):
        """Open the script in the default text editor via the 'edit' shell verb."""
        import ctypes
        try:
            ret = ctypes.windll.shell32.ShellExecuteW(
                None, "edit", self._path, None, None, 1
            )
            if ret <= 32:
                # 'edit' verb not registered for this file type — fall back to Notepad
                subprocess.Popen(
                    ["notepad.exe", self._path],
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
        except Exception as exc:
            QMessageBox.warning(
                None, "Cannot Open File",
                f"Could not open '{Path(self._path).name}' for editing:\n{exc}",
            )

    def _on_remove(self):
        if _is_running(self._path):
            reply = QMessageBox.question(
                self, "Script Running",
                f"'{Path(self._path).name}' is currently running.\n"
                "Stop it and remove from list?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                return
        self._dash.remove_script(self._path)


# ---------------------------------------------------------------------------
# UI helpers
# ---------------------------------------------------------------------------

def _mk_btn(label: str, bg: str, fg: str) -> QPushButton:
    btn = QPushButton(label)
    btn.setFixedHeight(28)
    btn.setCursor(Qt.PointingHandCursor)
    btn.setFont(QFont("Segoe UI", 8))
    btn.setStyleSheet(f"""
        QPushButton {{
            background:{bg}; color:{fg};
            border:1px solid {fg}44; border-radius:4px;
            padding:0 10px;
        }}
        QPushButton:hover {{
            background:{fg}33; border-color:{fg};
        }}
        QPushButton:disabled {{
            background:#0d1117; color:#484f58; border-color:#30363d;
        }}
    """)
    return btn


def _section_label(text: str) -> QLabel:
    lbl = QLabel(text)
    lbl.setFont(QFont("Segoe UI", 8, QFont.Weight.DemiBold))
    lbl.setStyleSheet("color:#58a6ff; background:transparent; letter-spacing:0.5px;")
    return lbl


def _sep(layout: QVBoxLayout):
    line = QFrame()
    line.setFrameShape(QFrame.HLine)
    line.setStyleSheet("color:#30363d; border:none; border-top:1px solid #30363d; background:transparent;")
    layout.addWidget(line)


def _input_style() -> str:
    return (
        "QLineEdit{"
        "background:#161b22; color:#c9d1d9;"
        "border:1px solid #30363d; border-radius:4px; padding:4px 8px;"
        "font-family:'Segoe UI'; font-size:9pt;}"
        "QLineEdit:focus{border-color:#58a6ff;}"
    )


def _tab_style() -> str:
    return """
        QTabWidget::pane {
            background: #0d1117;
            border: 1px solid #30363d;
            border-top: none;
        }
        QTabBar::tab {
            background: #161b22;
            color: #8b949e;
            border: 1px solid #30363d;
            border-bottom: none;
            padding: 6px 16px;
            font-family: 'Segoe UI';
            font-size: 9pt;
        }
        QTabBar::tab:selected {
            background: #0d1117;
            color: #58a6ff;
            border-bottom: 1px solid #0d1117;
        }
        QTabBar::tab:hover:!selected {
            color: #c9d1d9;
        }
    """


def _table_style() -> str:
    return """
        QTableWidget {
            background: #161b22;
            alternate-background-color: #0d1117;
            color: #c9d1d9;
            gridline-color: #30363d;
            border: 1px solid #30363d;
            font-family: 'Segoe UI';
            font-size: 8.5pt;
            border-radius: 4px;
        }
        QHeaderView::section {
            background: #1c2128;
            color: #8b949e;
            border: none;
            border-bottom: 1px solid #30363d;
            padding: 4px 8px;
            font-weight: bold;
        }
        QTableWidget::item:selected {
            background: #1f3a5f;
            color: #c9d1d9;
        }
    """
