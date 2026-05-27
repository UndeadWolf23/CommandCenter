"""
Command Center - A GUI-based workflow hub built with PySide6.
Provides a customizable node-tile dashboard for launching files, URLs, and notes.
"""

import sys
import os
import json
import subprocess
import webbrowser
import hashlib
import shutil
import math
import random
import ctypes
import ctypes.wintypes
import winreg
import zipfile
import openpyxl
import cryptography
from pathlib import Path
from functools import lru_cache
from datetime import datetime, timedelta
from typing import Optional

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QPushButton, QDialog, QLineEdit, QTextEdit,
    QCheckBox, QComboBox, QFileDialog, QScrollArea, QFrame,
    QSizePolicy, QSpacerItem, QGraphicsDropShadowEffect, QStackedWidget,
    QRadioButton, QButtonGroup, QMenu, QSlider, QFormLayout,
    QDialogButtonBox, QSplitter, QAbstractScrollArea, QListWidget,
    QListWidgetItem, QMessageBox, QTabWidget, QColorDialog, QScrollBar,
    QStyledItemDelegate, QSpinBox, QCalendarWidget, QAbstractSpinBox
)
from PySide6.QtCore import (
    Qt, QPoint, QSize, QRect, QPropertyAnimation, QEasingCurve,
    QTimer, QThread, Signal, QObject, QMimeData, QUrl, QSettings,
    QStandardPaths, QEvent, QRectF, QPointF, Property, QByteArray, QDate
)
from PySide6.QtGui import (
    QColor, QPainter, QPen, QBrush, QLinearGradient, QRadialGradient,
    QFont, QFontMetrics, QPainterPath, QIcon, QPixmap, QCursor,
    QDragEnterEvent, QDropEvent, QMouseEvent, QPaintEvent, QImage,
    QGuiApplication, QResizeEvent, QConicalGradient,
    QSyntaxHighlighter, QTextCharFormat, QAction, QTextCursor,
    QIntValidator
)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

APP_NAME = "Command Center"
APP_VERSION = "1.2.6.1"

_settings_store = QSettings("CommandCenter", "CommandCenter")

# If the stored save_path is inside the Store Python sandbox (Packages\...),
# clear it so the app uses the real LOCALAPPDATA default instead.
_stored_save_path = _settings_store.value("save_path", "")
if _stored_save_path and (
    "\\Packages\\" in _stored_save_path or "/Packages/" in _stored_save_path
):
    _settings_store.setValue("save_path", "")
    _settings_store.sync()
    _stored_save_path = ""

def _get_config_dir() -> Path:
    custom = _settings_store.value("save_path", "")
    if custom and "\\Packages\\" not in custom and Path(custom).is_dir():
        return Path(custom)
    # Store Python (WindowsApps) virtualises all AppData writes via NTFS reparse
    # points, silently redirecting them to the Packages sandbox.  Detect this by
    # checking if the interpreter lives in WindowsApps and, if so, store data in
    # a 'data' subfolder next to the script — that path is never virtualised.
    # When packaged as a PyInstaller exe the check fails and we use the proper
    # LOCALAPPDATA\CommandCenter path.
    exe = os.path.abspath(sys.executable)
    if "WindowsApps" in exe or "windowsapps" in exe.lower():
        return Path(os.path.dirname(os.path.abspath(__file__))) / "data"
    localappdata = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(localappdata) / "CommandCenter"

CONFIG_DIR = _get_config_dir()
CONFIG_FILE = CONFIG_DIR / "nodes.json"
CACHE_DIR = CONFIG_DIR / "cache"
ICON_CACHE_DIR = CACHE_DIR / "icons"

CONFIG_DIR.mkdir(parents=True, exist_ok=True)
ICON_CACHE_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# Theme system  —  all UI colors go through ThemeManager
# ---------------------------------------------------------------------------

_BUILTIN_THEMES = {
    "Deep Space": {
        "bg_dark":        (14, 20, 34),
        "bg_mid":         (20, 28, 46),
        "bg_grad_top":    (16, 22, 38),
        "bg_grad_bot":    (22, 14, 36),
        "glow":           (0, 200, 255),
        "titlebar_bg":    (12, 18, 32, 250),
        "text_primary":   (220, 242, 255),
        "text_secondary": (140, 190, 220),
        "text_dim":       (80, 120, 155),
        "accent_blue":    (30, 150, 255),
        "accent_teal":    (0, 210, 190),
        "accent_amber":   (255, 170, 20),
        "accent_red":     (230, 60, 70),
        "tile_bg_base":   (18, 28, 46),
        "tile_bg_hover":  (26, 40, 64),
        "footer_bg":      (12, 18, 32, 240),
    },
    "Midnight Blue": {
        "bg_dark":        (10, 15, 40),
        "bg_mid":         (16, 24, 60),
        "bg_grad_top":    (12, 18, 50),
        "bg_grad_bot":    (8, 12, 36),
        "glow":           (80, 180, 255),
        "titlebar_bg":    (10, 14, 38, 250),
        "text_primary":   (220, 235, 255),
        "text_secondary": (140, 175, 220),
        "text_dim":       (80, 110, 160),
        "accent_blue":    (60, 160, 255),
        "accent_teal":    (40, 210, 200),
        "accent_amber":   (255, 180, 40),
        "accent_red":     (230, 60, 70),
        "tile_bg_base":   (16, 24, 56),
        "tile_bg_hover":  (24, 36, 80),
        "footer_bg":      (10, 14, 38, 240),
    },
    "Slate Light": {
        "bg_dark":        (38, 44, 58),
        "bg_mid":         (48, 56, 72),
        "bg_grad_top":    (44, 50, 66),
        "bg_grad_bot":    (32, 38, 52),
        "glow":           (100, 200, 255),
        "titlebar_bg":    (34, 40, 54, 250),
        "text_primary":   (230, 240, 255),
        "text_secondary": (175, 200, 230),
        "text_dim":       (120, 150, 185),
        "accent_blue":    (80, 170, 255),
        "accent_teal":    (60, 220, 200),
        "accent_amber":   (255, 185, 50),
        "accent_red":     (235, 80, 85),
        "tile_bg_base":   (46, 54, 70),
        "tile_bg_hover":  (58, 68, 88),
        "footer_bg":      (34, 40, 54, 240),
    },
    "Forest Night": {
        "bg_dark":        (10, 22, 18),
        "bg_mid":         (14, 32, 26),
        "bg_grad_top":    (12, 26, 20),
        "bg_grad_bot":    (8, 18, 14),
        "glow":           (60, 230, 160),
        "titlebar_bg":    (10, 20, 16, 250),
        "text_primary":   (215, 245, 230),
        "text_secondary": (130, 195, 165),
        "text_dim":       (70, 130, 100),
        "accent_blue":    (40, 180, 220),
        "accent_teal":    (60, 230, 160),
        "accent_amber":   (255, 195, 60),
        "accent_red":     (225, 70, 75),
        "tile_bg_base":   (14, 30, 22),
        "tile_bg_hover":  (20, 44, 32),
        "footer_bg":      (10, 20, 16, 240),
    },
    "Crimson Dark": {
        "bg_dark":        (24, 10, 14),
        "bg_mid":         (36, 14, 20),
        "bg_grad_top":    (28, 12, 16),
        "bg_grad_bot":    (20, 8, 12),
        "glow":           (255, 100, 140),
        "titlebar_bg":    (22, 10, 14, 250),
        "text_primary":   (255, 235, 240),
        "text_secondary": (210, 165, 180),
        "text_dim":       (140, 90, 110),
        "accent_blue":    (100, 160, 255),
        "accent_teal":    (80, 215, 185),
        "accent_amber":   (255, 185, 50),
        "accent_red":     (255, 80, 100),
        "tile_bg_base":   (32, 14, 20),
        "tile_bg_hover":  (46, 20, 28),
        "footer_bg":      (22, 10, 14, 240),
    },
    # ── Dark Knight: Batman-inspired — pure black/charcoal, gold bat-signal glow ──
    "Dark Knight": {
        "bg_dark":        (4, 4, 6),       # near-pure black
        "bg_mid":         (10, 10, 12),    # dark charcoal
        "bg_grad_top":    (8, 8, 10),
        "bg_grad_bot":    (2, 2, 4),
        "glow":           (190, 158, 22),   # Batman gold (bat symbol)
        "titlebar_bg":    (4, 4, 6, 250),
        "text_primary":   (232, 232, 235),  # near-white
        "text_secondary": (160, 160, 168),  # cool grey
        "text_dim":       (88, 88, 96),     # dark grey
        "accent_blue":    (70, 110, 165),   # muted grey-blue (Batman suit)
        "accent_teal":    (55, 145, 130),   # dark teal
        "accent_amber":   (190, 158, 22),   # gold
        "accent_red":     (195, 48, 58),    # red
        "tile_bg_base":   (12, 12, 15),
        "tile_bg_hover":  (20, 20, 24),
        "footer_bg":      (4, 4, 6, 240),
    },
    # ── Spooky: Halloween — deep purple BG, burnt/dark orange glow ───────────
    "Spooky": {
        "bg_dark":        (14, 4, 28),    # deep saturated purple-black
        "bg_mid":         (26, 8, 52),    # rich dark purple
        "bg_grad_top":    (20, 6, 42),
        "bg_grad_bot":    (10, 2, 22),
        "glow":           (200, 80, 10),  # dark burnt pumpkin orange
        "titlebar_bg":    (12, 2, 24, 250),
        "text_primary":   (250, 228, 200),  # warm cream
        "text_secondary": (185, 130, 80),   # dark pumpkin-tan
        "text_dim":       (110, 68, 42),    # dim burnt orange
        "accent_blue":    (110, 30, 180),   # deep eerie purple
        "accent_teal":    (160, 20, 200),   # dark ghostly violet
        "accent_amber":   (210, 100, 12),   # dark halloween orange
        "accent_red":     (180, 30, 22),    # deep blood red
        "tile_bg_base":   (20, 6, 40),
        "tile_bg_hover":  (34, 12, 62),
        "footer_bg":      (12, 2, 24, 240),
    },
    # ── Noir: Film noir — near-black BG, off-white text, silver/grey tones ────
    "Noir": {
        "bg_dark":        (8, 8, 8),      # almost pure black
        "bg_mid":         (16, 16, 16),   # very dark charcoal
        "bg_grad_top":    (20, 20, 20),
        "bg_grad_bot":    (6, 6, 6),
        "glow":           (195, 195, 195),  # cool silver — the classic spotlight
        "titlebar_bg":    (6, 6, 6, 250),
        "text_primary":   (230, 228, 222),  # warm off-white newsprint
        "text_secondary": (160, 158, 152),  # aged grey
        "text_dim":       (88, 86, 82),     # dim smoke grey
        "accent_blue":    (130, 128, 138),  # grey-blue steel
        "accent_teal":    (105, 145, 130),  # muted teal shadow
        "accent_amber":   (175, 148, 80),   # sepia gold
        "accent_red":     (175, 42, 42),    # dark crimson (cigarette cherry)
        "tile_bg_base":   (20, 20, 20),
        "tile_bg_hover":  (30, 30, 30),
        "footer_bg":      (6, 6, 6, 240),
    },
    "Custom": {},   # merged from saved QSettings keys
}


class ThemeManager:
    """Central color store — all widgets read from here."""
    _instance: "ThemeManager | None" = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._init()
        return cls._instance

    def _init(self):
        self._listeners: list[callable] = []
        self._load()

    def _load(self):
        name = _settings_store.value("theme_name", "Deep Space")
        if name not in _BUILTIN_THEMES:
            name = "Deep Space"
        self._name = name
        base = dict(_BUILTIN_THEMES.get(name, _BUILTIN_THEMES["Deep Space"]))
        # Merge custom overrides
        for key in list(base.keys()):
            stored = _settings_store.value(f"theme_custom/{key}", None)
            if stored:
                try:
                    parts = [int(x) for x in stored.split(",")]
                    base[key] = tuple(parts)
                except Exception as exc:
                    print(f"[CommandCenter] Theme custom color ignored ({key}): {exc}",
                          file=sys.stderr)
        self._colors = base
        # Apply brightness scaling before building QColor objects
        brightness = (int(_settings_store.value("ui_brightness", 100)) + 100) / 100.0
        if brightness != 2.0:
            scaled = {}
            for key, val in self._colors.items():
                r = min(255, int(val[0] * brightness))
                g = min(255, int(val[1] * brightness))
                b = min(255, int(val[2] * brightness))
                scaled[key] = (r, g, b) + val[3:]
            self._colors = scaled
        self._build()

    def _build(self):
        c = self._colors
        def qc(key, a=255):
            v = c[key]
            if len(v) == 4:
                return QColor(*v)
            return QColor(v[0], v[1], v[2], a)
        self.BG_DARK        = qc("bg_dark")
        self.BG_MID         = qc("bg_mid")
        self.BG_GRAD_TOP    = qc("bg_grad_top")
        self.BG_GRAD_BOT    = qc("bg_grad_bot")
        self.GLOW           = qc("glow")
        self.GLOW_DIM       = QColor(self.GLOW.red(), self.GLOW.green(), self.GLOW.blue(), 100)
        self.TITLEBAR_BG    = qc("titlebar_bg") if len(c["titlebar_bg"])==4 else QColor(c["titlebar_bg"][0],c["titlebar_bg"][1],c["titlebar_bg"][2],250)
        self.TEXT_PRIMARY   = qc("text_primary")
        self.TEXT_SECONDARY = qc("text_secondary")
        self.TEXT_DIM       = qc("text_dim")
        self.ACCENT_BLUE    = qc("accent_blue")
        self.ACCENT_TEAL    = qc("accent_teal")
        self.ACCENT_AMBER   = qc("accent_amber")
        self.ACCENT_RED     = qc("accent_red")
        self.TILE_BG_BASE   = qc("tile_bg_base")
        self.TILE_BG_HOVER  = qc("tile_bg_hover")
        self.FOOTER_BG      = qc("footer_bg") if len(c["footer_bg"])==4 else QColor(c["footer_bg"][0],c["footer_bg"][1],c["footer_bg"][2],240)
        self.BTN_CLOSE      = QColor(210, 55, 65)
        self.BTN_MINIMIZE   = QColor(220, 150, 10)
        self.BTN_MAXIMIZE   = QColor(30, 185, 110)

    @property
    def theme_name(self) -> str:
        return self._name

    def apply_theme(self, name: str):
        if name not in _BUILTIN_THEMES:
            return
        _settings_store.setValue("theme_name", name)
        self._name = name
        self._load()
        self._notify()

    def set_custom_color(self, key: str, color: QColor):
        _settings_store.setValue(f"theme_custom/{key}", f"{color.red()},{color.green()},{color.blue()},{color.alpha()}")
        _settings_store.sync()
        self._load()
        self._notify()

    def set_brightness(self, value: int):
        _settings_store.setValue("ui_brightness", value)
        _settings_store.sync()
        self._load()
        self._notify()

    def color_tuple(self, key: str) -> tuple:
        return self._colors.get(key, (128, 128, 128))

    def register(self, fn: callable):
        if fn not in self._listeners:
            self._listeners.append(fn)

    def unregister(self, fn: callable):
        try:
            self._listeners.remove(fn)
        except ValueError:
            pass

    def _notify(self):
        dead = []
        for fn in list(self._listeners):
            try:
                fn()
            except Exception as exc:
                print(f"[CommandCenter] Theme listener error: {exc}", file=sys.stderr)
                dead.append(fn)
        for fn in dead:
            try:
                self._listeners.remove(fn)
            except ValueError:
                pass


_theme = ThemeManager()

# Convenience aliases so existing code works unchanged
def _T(): return _theme   # shorthand for inline use

# Static aliases — kept as module-level vars that point into theme for
# read-only convenience; critical paths should use _theme directly
COLOR_BG_DARK        = _theme.BG_DARK
COLOR_BG_MID         = _theme.BG_MID
COLOR_BG_GRAD_TOP    = _theme.BG_GRAD_TOP
COLOR_BG_GRAD_BOT    = _theme.BG_GRAD_BOT
COLOR_GLOW           = _theme.GLOW
COLOR_GLOW_DIM       = _theme.GLOW_DIM
COLOR_TITLEBAR_BG    = _theme.TITLEBAR_BG
COLOR_TEXT_PRIMARY   = _theme.TEXT_PRIMARY
COLOR_TEXT_SECONDARY = _theme.TEXT_SECONDARY
COLOR_TEXT_DIM       = _theme.TEXT_DIM
COLOR_ACCENT_BLUE    = _theme.ACCENT_BLUE
COLOR_ACCENT_TEAL    = _theme.ACCENT_TEAL
COLOR_ACCENT_AMBER   = _theme.ACCENT_AMBER
COLOR_ACCENT_RED     = _theme.ACCENT_RED
COLOR_BTN_CLOSE      = _theme.BTN_CLOSE
COLOR_BTN_MINIMIZE   = _theme.BTN_MINIMIZE
COLOR_BTN_MAXIMIZE   = _theme.BTN_MAXIMIZE
COLOR_FOOTER_BG      = _theme.FOOTER_BG

BORDER_RADIUS   = 10
TITLEBAR_HEIGHT = 50
FOOTER_HEIGHT   = 36
GLOW_WIDTH      = 2
RESIZE_MARGIN   = 7

FONT_TITLE  = QFont("Segoe UI", 11, QFont.Weight.Bold)
FONT_TITLE.setLetterSpacing(QFont.AbsoluteSpacing, 0.7)
FONT_LABEL  = QFont("Segoe UI", 9)
FONT_SMALL  = QFont("Segoe UI", 8)
FONT_MONO   = QFont("Consolas", 9)

# NODE_SIZES: (grid_cols, grid_rows)  — 2x4 is 4 columns wide, 2 rows tall (horizontal)
NODE_SIZES = {
    "1x1": (1, 1),
    "2x2": (2, 2),
    "2x4": (4, 2),   # 4 grid cols wide × 2 grid rows tall  → landscape
}
TILE_BASE_SIZE = 140
TILE_GAP       = 12

# Grid constants – the canvas works on an invisible cell grid
GRID_COLS      = 6    # default columns; recalculated on canvas resize
GRID_CELL      = TILE_BASE_SIZE + TILE_GAP   # pixels per grid cell

NODE_TYPE_FILE   = "file"
NODE_TYPE_URL    = "url"
NODE_TYPE_NOTE   = "note"   # Notebook-linked note node
NODE_TYPE_FOLDER = "folder"

NODE_FILE_EXT = ".node"

# Drag state IDs
_DRAG_NONE   = 0
_DRAG_MOVE   = 1
_DRAG_HOLD_MS = 300   # ms hold before drag starts


# ---------------------------------------------------------------------------
# Windows startup registry helpers
# ---------------------------------------------------------------------------

_STARTUP_KEY_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
_STARTUP_VALUE    = "CommandCenter"


def _set_windows_startup(enabled: bool) -> bool:
    """Write or delete the HKCU Run registry key. Returns True on success."""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, _STARTUP_KEY_PATH, 0, winreg.KEY_SET_VALUE)
        if enabled:
            cmd = f'"{sys.executable}" "{os.path.abspath(__file__)}"'
            winreg.SetValueEx(key, _STARTUP_VALUE, 0, winreg.REG_SZ, cmd)
        else:
            try:
                winreg.DeleteValue(key, _STARTUP_VALUE)
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
        return True
    except OSError as exc:
        print(f"[CommandCenter] Startup registry error: {exc}", file=sys.stderr)
        return False


def _get_windows_startup() -> bool:
    """Return True if the startup Run entry currently exists."""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, _STARTUP_KEY_PATH, 0, winreg.KEY_READ)
        winreg.QueryValueEx(key, _STARTUP_VALUE)
        winreg.CloseKey(key)
        return True
    except (FileNotFoundError, OSError):
        return False


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def cache_icon(source_path: str) -> Optional[str]:
    if not source_path or not os.path.isfile(source_path):
        return source_path
    ext = Path(source_path).suffix.lower()
    h = hashlib.md5(source_path.encode()).hexdigest()
    dest = ICON_CACHE_DIR / f"{h}{ext}"
    if not dest.exists():
        shutil.copy2(source_path, dest)
    return str(dest)


@lru_cache(maxsize=64)
def load_pixmap_cached(path: str, size: int = 64) -> QPixmap:
    pm = QPixmap(path)
    if pm.isNull():
        return pm
    return pm.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation)


def make_shadow(parent: QWidget, radius: float = 18,
                color: QColor = None) -> QGraphicsDropShadowEffect:
    shadow = QGraphicsDropShadowEffect(parent)
    shadow.setBlurRadius(radius)
    shadow.setOffset(0, 0)
    shadow.setColor(color or QColor(0, 0, 0, 160))
    return shadow


def launch_path(target: str, params: str = "") -> None:
    """
    Smart launcher.

    - .py              → runs with the current Python interpreter
    - .ps1             → runs with PowerShell (bypass execution policy)
    - .bat / .cmd      → runs through cmd.exe
    - .exe / .com / "" → Popen directly (supports params)
    - everything else  → os.startfile() which uses the Windows Shell and
                         respects registered file associations, so .txt opens
                         in Notepad, .pdf in Acrobat, images in Photos, etc.
                         params are ignored for shell-opened files.
    Raises OSError on failure.
    """
    if not target:
        raise OSError("No target path specified.")
    ext = Path(target).suffix.lower()
    # Script types — explicit interpreter
    if ext == ".py":
        cmd = [sys.executable, target]
        if params:
            cmd += params.split()
        subprocess.Popen(cmd, shell=False)
    elif ext == ".ps1":
        cmd = ["powershell", "-noexit"]
        if _settings_store.value("ps1_bypass_execution_policy", "true") == "true":
            cmd += ["-ExecutionPolicy", "Bypass"]
        cmd += ["-File", target]
        if params:
            cmd += params.split()
        subprocess.Popen(cmd, shell=False)
    elif ext in (".bat", ".cmd"):
        cmd = ["cmd", "/c", target]
        if params:
            cmd += params.split()
        subprocess.Popen(cmd, shell=False)
    elif ext in (".exe", ".com", ""):
        # Native executable or extensionless binary / folder shortcut.
        # CREATE_NEW_CONSOLE ensures console apps (cmd, powershell, etc.) get
        # a proper interactive window even when launched from a GUI parent.
        # For GUI subsystem exes (Chrome, etc.) the flag is a no-op.
        cmd = [target]
        if params:
            cmd += params.split()
        subprocess.Popen(cmd, shell=False,
                         creationflags=subprocess.CREATE_NEW_CONSOLE)
    else:
        # All other types (documents, images, text files, PDFs, …)
        # Use ShellExecute via os.startfile so Windows picks the right app.
        os.startfile(target)


# ---------------------------------------------------------------------------
# Node store (persistence)
# ---------------------------------------------------------------------------

class NodeStore:
    def __init__(self):
        self._data: list[dict] = []
        self._config_file = self._resolve_config_file()
        self.load()
        self.seed_defaults_if_needed()
        self._migrate_defaults_v2()
        self._migrate_defaults_v3()

    def _resolve_config_file(self) -> Path:
        # Always derive from CONFIG_DIR so there is exactly one path-resolution
        # code path in the whole app.
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        return CONFIG_DIR / "nodes.json"

    def reload_path(self):
        self._config_file = self._resolve_config_file()
        self.load()

    def load(self):
        if self._config_file.exists():
            try:
                with open(self._config_file, "r", encoding="utf-8") as f:
                    self._data = json.load(f)
            except (json.JSONDecodeError, OSError) as exc:
                print(f"[CommandCenter] Nodes load error: {exc}", file=sys.stderr)
                self._data = []
        else:
            self._data = []

    def save(self):
        try:
            self._config_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self._config_file, "w", encoding="utf-8") as f:
                json.dump(self._data, f, indent=2, ensure_ascii=False)
        except OSError as exc:
            print(f"[CommandCenter] Save error: {exc}", file=sys.stderr)

    def all_nodes(self) -> list[dict]:
        return [n for n in self._data
                if not n.get("archived", False) and n.get("type") != NODE_TYPE_FOLDER]

    def all_folders(self) -> list[dict]:
        return [n for n in self._data
                if not n.get("archived", False) and n.get("type") == NODE_TYPE_FOLDER]

    def all_items(self) -> list[dict]:
        """Root-level non-archived items (nodes + folders), sorted by grid_order.
        Nodes that belong to a folder (folder_id is set) are excluded from the
        main canvas — they are only visible inside their folder's popup.
        """
        items = [n for n in self._data
                 if not n.get("archived", False)
                 and not n.get("folder_id")]
        items.sort(key=lambda n: n.get("grid_order", 9999))
        return items

    def archived_nodes(self) -> list[dict]:
        return [n for n in self._data if n.get("archived", False)]

    def add_node(self, node: dict) -> dict:
        node.setdefault("id", self._new_id())
        # place at end of grid by default
        active = [n for n in self._data if not n.get("archived", False)]
        node.setdefault("grid_order", len(active))
        self._data.append(node)
        self.save()
        return node

    def reorder_items(self, ordered_ids: list[int]):
        """Persist a new display order for grid items."""
        id_to_order = {nid: idx for idx, nid in enumerate(ordered_ids)}
        for n in self._data:
            if n.get("id") in id_to_order:
                n["grid_order"] = id_to_order[n["id"]]
        self.save()

    def add_folder(self, name: str) -> dict:
        active = [n for n in self._data if not n.get("archived", False)]
        folder = {
            "id": self._new_id(),
            "type": NODE_TYPE_FOLDER,
            "name": name,
            "archived": False,
            "grid_order": len(active),
            "children": [],   # list of node ids
        }
        self._data.append(folder)
        self.save()
        return folder

    def move_node_to_folder(self, node_id: int, folder_id: Optional[int]):
        """Set node's folder_id. Pass None to move to root."""
        for n in self._data:
            if n.get("id") == node_id:
                n["folder_id"] = folder_id
                break
        # keep folder's children list in sync
        for n in self._data:
            if n.get("type") == NODE_TYPE_FOLDER:
                children = n.get("children", [])
                if n.get("id") == folder_id:
                    if node_id not in children:
                        children.append(node_id)
                else:
                    if node_id in children:
                        children.remove(node_id)
                n["children"] = children
        self.save()

    def update_node(self, node_id: int, updates: dict):
        for n in self._data:
            if n.get("id") == node_id:
                n.update(updates)
                break
        self.save()

    def remove_node(self, node_id: int):
        self._data = [n for n in self._data if n.get("id") != node_id]
        self.save()

    def archive_node(self, node_id: int):
        self.update_node(node_id, {"archived": True})

    def unarchive_node(self, node_id: int):
        self.update_node(node_id, {"archived": False})

    def seed_defaults_if_needed(self):
        """Create the 'Default Nodes' folder on first ever launch only.
        Uses a sentinel file inside the data directory so the check travels
        with the data folder rather than living in the registry (QSettings).
        A legacy QSettings flag is also checked for backward compatibility.
        """
        _seed_flag = CONFIG_DIR / ".seeded"
        _legacy_flag = _settings_store.value("defaults_seeded_v1", "false") == "true"
        if _seed_flag.exists() or _legacy_flag:
            # Write the file-based flag if we only had the legacy registry one
            if _legacy_flag and not _seed_flag.exists():
                try:
                    _seed_flag.touch()
                except OSError:
                    pass
            return
        # Mark as seeded immediately before writing so a crash mid-seed
        # doesn't cause a partial re-seed on next launch.
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            _seed_flag.touch()
        except OSError:
            pass
        _settings_store.setValue("defaults_seeded_v1", "true")
        _settings_store.sync()

        folder = self.add_folder("Default Nodes")
        fid = folder["id"]

        lp  = os.path.expandvars   # expand %LOCALAPPDATA% and similar vars
        _sd = os.path.dirname(os.path.abspath(__file__))  # script directory

        def _img(filename):
            """Return absolute path to a bundled image, empty string if missing."""
            p = os.path.join(_sd, filename)
            return p if os.path.isfile(p) else ""

        defaults = [
            # name,               type,             target,                                                            accent_color,  icon_file
            # ── Office 365 (new Store-based apps) ──────────────────────────────────────
            ("MS Outlook",         NODE_TYPE_FILE,  lp(r"%LOCALAPPDATA%\Microsoft\WindowsApps\olk.exe"),              "#0078D4", "outlook.png"),
            ("MS Teams",           NODE_TYPE_FILE,  lp(r"%LOCALAPPDATA%\Microsoft\WindowsApps\ms-teams.exe"),         "#6264A7", "teams.png"),
            ("MS OneNote",         NODE_TYPE_FILE,  r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote.lnk",   "#7719AA", "onenote.png"),
            # ── Browser ────────────────────────────────────────────────────────────────
            ("Google Chrome",      NODE_TYPE_FILE,  r"C:\Program Files\Google\Chrome\Application\chrome.exe",         "#DB4437", "chrome.png"),
            # ── Work URLs ──────────────────────────────────────────────────────────────
            ("SmartTrak",          NODE_TYPE_URL,   "https://stp.boit.us/displaycases",                               "#00B4D8", ""),
            ("ScreenConnect",      NODE_TYPE_URL,   "https://remote.bankonitusa.com/Host#Access",                      "#F77F00", "ScreenConnect.png"),
            ("BoitDOCS",           NODE_TYPE_URL,   "https://boitdocs.boit.us/",                                      "#2DC653", ""),
            ("ChatGPT",            NODE_TYPE_URL,   "https://chatgpt.com/",                                           "#10A37F", "GPT.png"),
            ("SentinelOne",        NODE_TYPE_URL,   "https://usea1-navanta.sentinelone.net/login",                     "#7B2D8B", "SentinelOne.png"),
            ("Lunch Schedule",     NODE_TYPE_URL,   "https://boitdocs.boit.us/tiki-index.php?page=OKC+Lunch+Schedule","#F4A261", ""),
            ("ADP",                NODE_TYPE_URL,   "https://workforcenow.adp.com/theme/index.html#/home",            "#E63946", "ADP.png"),
            ("SafeSend",           NODE_TYPE_URL,   "https://mail.secureyouremail.com/Default.aspx?ReturnUrl=%2f",    "#457B9D", ""),
            # ── System tools ───────────────────────────────────────────────────────────
            ("Command Prompt",     NODE_TYPE_FILE,  r"C:\Windows\System32\cmd.exe",                                   "#4A9E5C", "cmd.png"),
            ("PowerShell ISE",      NODE_TYPE_FILE,  r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell_ise.exe",     "#012456", "powershell.png"),
        ]

        for name, ntype, target, color, icon_file in defaults:
            node = {
                "name":         name,
                "type":         ntype,
                "target":       target,
                "description":  "",
                "auto_launch":  False,
                "accent_color": color,
                "archived":     False,
                "tags":         [],
                "open_behavior":"normal",
                "icon":         _img(icon_file) if icon_file else "",
            }
            saved = self.add_node(node)
            # Places node inside the folder (sets folder_id + updates children list)
            self.move_node_to_folder(saved["id"], fid)

    def _migrate_defaults_v2(self):
        """One-time migration: rename PowerShell node to PowerShell ISE and
        update its target from powershell.exe → powershell_ise.exe.
        """
        if _settings_store.value("defaults_migrated_v2", "false") == "true":
            return
        _settings_store.setValue("defaults_migrated_v2", "true")
        old_target = r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
        new_target = r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell_ise.exe"
        changed = False
        for n in self._data:
            if n.get("target") == old_target and n.get("name") == "PowerShell":
                n["target"] = new_target
                n["name"]   = "PowerShell ISE"
                changed = True
        if changed:
            self.save()

    def _migrate_defaults_v3(self):
        """One-time migration: update OneNote target to use the Start Menu .lnk."""
        if _settings_store.value("defaults_migrated_v3", "false") == "true":
            return
        _settings_store.setValue("defaults_migrated_v3", "true")
        old_targets = [
            r"C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE",
            r"%LOCALAPPDATA%\Microsoft\WindowsApps\ONENOTE.EXE",
        ]
        new_target = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote.lnk"
        changed = False
        for n in self._data:
            if n.get("name") == "MS OneNote" and n.get("target") in old_targets:
                n["target"] = new_target
                changed = True
        if changed:
            self.save()

    def _new_id(self) -> int:
        if not self._data:
            return 1
        return max(n.get("id", 0) for n in self._data) + 1


    @staticmethod
    def export_node(node: dict, path: str):
        data = {k: v for k, v in node.items() if k != "id"}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    @staticmethod
    def import_node(path: str) -> dict:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)


# ---------------------------------------------------------------------------
# Custom cursor support
# ---------------------------------------------------------------------------

# (name, filename) — empty string means "use system default"
_CURSOR_OPTIONS = [
    ("Standard",     "DefaultCursor.png"),
    ("Cosmic",       "CosmicCursor.png"),
    ("High Fantasy", "HighFantasyCursor.png"),
    ("High Tech",    "HighTechCursor.png"),
    ("Medieval",     "MedievalCursor.png"),
    ("Noir",         "NoirCursor.png"),
    ("Reactor",      "ReactorCursor.png"),
    ("Spooky",       "SpookyCursor.png"),
]

_cursor_override_active = False


def _apply_app_cursor():
    """Read the saved cursor preference and apply it application-wide."""
    global _cursor_override_active
    saved = _settings_store.value("cursor_name", "Standard")
    app = QApplication.instance()
    if app is None:
        return
    if saved == "Standard":
        if _cursor_override_active:
            app.restoreOverrideCursor()
            _cursor_override_active = False
        return
    # Resolve cursor file path — built-in presets vs. user-uploaded custom cursor
    if saved == "Custom":
        path = _settings_store.value("cursor_custom_path", "")
        if not path or not os.path.isfile(path):
            # No custom file saved yet; fall back gracefully to system cursor
            if _cursor_override_active:
                app.restoreOverrideCursor()
                _cursor_override_active = False
            return
    else:
        filename = next((f for n, f in _CURSOR_OPTIONS if n == saved and f), None)
        if not filename:
            if _cursor_override_active:
                app.restoreOverrideCursor()
                _cursor_override_active = False
            return
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
        if not os.path.isfile(path):
            return
    px = QPixmap(path)
    if px.isNull():
        return
    # Scale to user-chosen cursor size (default 40×40)
    sz = int(_settings_store.value("cursor_size", 40))
    sz = max(16, min(128, sz))   # clamp to sane range
    px = px.scaled(sz, sz, Qt.KeepAspectRatio, Qt.SmoothTransformation)
    # hotspot (0, 0) = click point at top-left corner of the image
    cursor = QCursor(px, 0, 0)
    if _cursor_override_active:
        app.changeOverrideCursor(cursor)
    else:
        app.setOverrideCursor(cursor)
        _cursor_override_active = True


# ---------------------------------------------------------------------------
# Animated glow border overlay
# ---------------------------------------------------------------------------

class GlowBorderWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self._phase = 0.0
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(40)

    def _tick(self):
        self._phase = (self._phase + 0.035) % (2 * math.pi)
        self.update()

    def paintEvent(self, event: QPaintEvent):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        pulse = 0.6 + 0.4 * math.sin(self._phase)
        alpha = int(140 + 100 * pulse)
        w, h = self.width(), self.height()
        r = BORDER_RADIUS + 1

        glow = QColor(t.GLOW)
        glow.setAlpha(int(alpha * 0.18))
        p.setPen(QPen(glow, GLOW_WIDTH + 8))
        p.setBrush(Qt.NoBrush)
        p.drawRoundedRect(2, 2, w - 4, h - 4, r + 4, r + 4)

        glow2 = QColor(t.GLOW)
        glow2.setAlpha(int(alpha * 0.38))
        p.setPen(QPen(glow2, GLOW_WIDTH + 3))
        p.drawRoundedRect(3, 3, w - 6, h - 6, r + 2, r + 2)

        sharp = QColor(t.GLOW)
        sharp.setAlpha(min(alpha, 220))
        p.setPen(QPen(sharp, 1.5))
        p.drawRoundedRect(4, 4, w - 8, h - 8, r, r)

        # Corner accent sparks
        spark_a = int(80 * pulse)
        sc = QColor(t.GLOW); sc.setAlpha(spark_a)
        p.setPen(QPen(sc, 2))
        for cx, cy in [(6, 6), (w-6, 6), (6, h-6), (w-6, h-6)]:
            p.drawLine(cx - 5, cy, cx + 5, cy)
            p.drawLine(cx, cy - 5, cx, cy + 5)
        p.end()


# ---------------------------------------------------------------------------
# Title bar button
# ---------------------------------------------------------------------------

class TitleBarButton(QWidget):
    clicked = Signal()

    def __init__(self, color: QColor, symbol: str, parent=None):
        super().__init__(parent)
        self._color = color
        self._symbol = symbol
        self._hovered = False
        self.setFixedSize(14, 14)
        self.setCursor(Qt.PointingHandCursor)

    def enterEvent(self, e):  self._hovered = True;  self.update()
    def leaveEvent(self, e):  self._hovered = False; self.update()

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.clicked.emit()

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        c = self._color.lighter(140) if self._hovered else self._color
        # Outer glow when hovered
        if self._hovered:
            glow = QColor(c); glow.setAlpha(60)
            p.setPen(Qt.NoPen); p.setBrush(glow)
            p.drawEllipse(-3, -3, 20, 20)
        p.setBrush(QBrush(c))
        p.setPen(Qt.NoPen)
        p.drawEllipse(0, 0, 14, 14)
        if self._hovered:
            p.setPen(QPen(QColor(0, 0, 0, 200), 1.2))
            p.setFont(QFont("Segoe UI", 6, QFont.Weight.Bold))
            p.drawText(QRect(0, 0, 14, 14), Qt.AlignCenter, self._symbol)
        p.end()


# ---------------------------------------------------------------------------
# Animated title label (pulses subtly between TEXT_PRIMARY and GLOW)
# ---------------------------------------------------------------------------

class _AnimatedTitleLabel(QLabel):
    """Title label with two layered animations:

    1. Sinusoidal colour pulse — smoothly blends TEXT_PRIMARY → GLOW on a
       ~6.5 s cycle (unchanged from before).
    2. Decode burst — every 15 s each character rolls through random techno
       glyphs left-to-right, snapping to its real value before moving on.
       After all characters resolve the 15 s countdown restarts.

    Uses a custom paintEvent rather than setStyleSheet so the Qt style
    engine is never re-evaluated on every animation tick, eliminating
    unnecessary layout / style recalculations.
    """

    _DECODE_CHARSET = "!@#$%^&*<>?/|0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    _DECODE_FLASHES = 3   # ticks per character before locking (× 50 ms = 250 ms/char)

    def __init__(self, text: str, parent=None):
        super().__init__(text, parent)
        self.setFont(FONT_TITLE)
        self.setAutoFillBackground(False)   # let parent title-bar gradient show through
        self._phase = 0.0
        self._color = _theme.TEXT_PRIMARY

        # ── decode animation state ───────────────────────────────────────────
        self._permanent_text = text   # the real title — decode always targets this
        self._display_text   = text   # what paintEvent actually renders
        self._decode_active  = False
        self._decode_char    = 0     # index of character currently being decoded
        self._decode_flashes = 0     # flash frames consumed for current char
        self._transient_active = False   # True while a status message is showing

        # ── pulse timer (50 ms / 20 fps) ────────────────────────────────────
        self._timer = QTimer(self)
        self._timer.setInterval(50)
        self._timer.timeout.connect(self._tick)
        self._timer.start()

        # ── trigger the first decode 15 s after startup ──────────────────────
        self._decode_trigger = QTimer(self)
        self._decode_trigger.setSingleShot(True)
        self._decode_trigger.timeout.connect(self._start_decode)
        self._decode_trigger.start(30_000)
        # Respect the "disable title bar animation" setting on startup
        if _settings_store.value("disable_titlebar_anim", "false") == "true":
            self._timer.stop()
            self._decode_trigger.stop()

    # ── decode helpers ────────────────────────────────────────────────────────

    def _start_decode(self):
        """Kick off a new left-to-right decode animation."""
        if self._transient_active:
            # Status message is showing — postpone decode until title is restored
            self._decode_trigger.start(15_000)
            return
        actual = self._permanent_text
        self._decode_char = 0
        # Skip any leading spaces before the first real character
        while self._decode_char < len(actual) and actual[self._decode_char] == ' ':
            self._decode_char += 1
        self._decode_flashes = 0
        self._decode_active  = True

    def _decode_step(self):
        """Advance the decode by one 50 ms tick."""
        actual = self._permanent_text
        done   = len(actual)

        if self._decode_char >= done:
            # All characters have been resolved
            self._decode_active = False
            self._display_text  = actual
            self._decode_trigger.start(15_000)
            return

        # Advance flash counter, or lock current char and move to next
        if self._decode_flashes < self._DECODE_FLASHES:
            self._decode_flashes += 1
        else:
            self._decode_char   += 1
            # Skip past consecutive spaces automatically
            while self._decode_char < done and actual[self._decode_char] == ' ':
                self._decode_char += 1
            self._decode_flashes = 0
            if self._decode_char >= done:
                self._decode_active = False
                self._display_text  = actual
                self._decode_trigger.start(15_000)
                return

        # Build display string: resolved chars on the left, scrambled on the right
        idx = self._decode_char
        result = []
        for i, c in enumerate(actual):
            if c == ' ':
                result.append(' ')
            elif i < idx:
                result.append(c)                            # already locked
            else:
                result.append(random.choice(self._DECODE_CHARSET))  # scrambled
        self._display_text = "".join(result)

    # ── animation tick (50 ms) ────────────────────────────────────────────────

    def _tick(self):
        # ~6.5 s full cycle  (2π / 0.048 × 0.05 s/step)
        self._phase = (self._phase + 0.048) % (2.0 * math.pi)
        t = _theme
        base, glow = t.TEXT_PRIMARY, t.GLOW
        # Blend 0–65 % toward the theme accent at the sine peak
        a = (math.sin(self._phase) * 0.5 + 0.5) * 0.65
        self._color = QColor(
            int(base.red()   + (glow.red()   - base.red())   * a),
            int(base.green() + (glow.green() - base.green()) * a),
            int(base.blue()  + (glow.blue()  - base.blue())  * a),
        )
        if self._decode_active:
            self._decode_step()
        self.update()

    def paintEvent(self, _e: QPaintEvent):
        """Draw the label text directly over the title-bar background."""
        p = QPainter(self)
        # Fill background first — without this, each frame draws ON TOP of the
        # previous frame's text pixels and the colour change is invisible.
        bg = _theme.TITLEBAR_BG
        p.fillRect(self.rect(), QColor(bg.red(), bg.green(), bg.blue()))
        p.setFont(self.font())
        p.setPen(self._color)
        p.drawText(self.rect(), Qt.AlignLeft | Qt.AlignVCenter, self._display_text)
        p.end()

    def set_animation_enabled(self, enabled: bool):
        """Start or stop the pulse and decode animations."""
        if enabled:
            if not self._timer.isActive():
                self._timer.start()
            if not self._decode_trigger.isActive() and not self._decode_active:
                self._decode_trigger.start(30_000)
        else:
            self._timer.stop()
            self._decode_trigger.stop()
            self._decode_active = False
            self._display_text  = self._permanent_text
            self._color         = _theme.TEXT_PRIMARY
            self.update()

    def set_permanent(self, text: str):
        """Update the permanent title and clear any transient override."""
        was_transient = self._transient_active
        self._permanent_text   = text
        self._transient_active = False
        # Abort any in-progress decode so it doesn't finish onto stale text
        if self._decode_active:
            self._decode_active = False
        # Always re-arm the decode trigger when restoring from a status message,
        # since set_transient may have stopped it while no decode was running yet.
        if (was_transient or not self._decode_trigger.isActive()) and self._timer.isActive():
            self._decode_trigger.start(15_000)
        self._display_text = text
        super().setText(text)
        self.update()

    def set_transient(self, text: str):
        """Show a temporary status string without touching the permanent title."""
        self._transient_active = True
        # Abort any running decode — it must not continue on the status text
        if self._decode_active:
            self._decode_active = False
            self._decode_trigger.stop()   # will be restarted when title is restored
        self._display_text = text
        super().setText(text)             # keep QLabel text in sync for size hints
        self.update()


# ---------------------------------------------------------------------------
# Custom title bar
# ---------------------------------------------------------------------------

class CustomTitleBar(QWidget):
    help_requested          = Signal()
    search_requested        = Signal()
    quick_connect_requested = Signal()

    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.setFixedHeight(TITLEBAR_HEIGHT)
        self._drag_pos: Optional[QPoint] = None
        self._main_win = None
        self._build_ui(title)

    def set_main_window(self, win):
        self._main_win = win

    def _build_ui(self, title: str):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 0, 8, 0)
        layout.setSpacing(0)

        _icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "CommandCenter.png")
        logo_lbl = QLabel()
        logo_lbl.setStyleSheet("background: transparent;")
        if os.path.isfile(_icon_path):
            logo_px = QPixmap(_icon_path).scaled(
                44, 44, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_lbl.setPixmap(logo_px)
        else:
            logo_lbl.setText("CC")
            logo_lbl.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
            logo_lbl.setStyleSheet(f"color: {_theme.GLOW.name()}; letter-spacing: 3px;")
        layout.addWidget(logo_lbl)
        layout.addSpacing(10)

        self._title_label = _AnimatedTitleLabel(title)
        layout.addWidget(self._title_label)
        layout.addStretch()

        qc_btn = QPushButton("  Quick Connect")
        qc_btn.setCursor(Qt.PointingHandCursor)
        qc_btn.setToolTip("Quick Connect  (Ctrl+Q)")
        qc_btn.setFont(QFont("Segoe UI", 8, QFont.Weight.DemiBold))
        qc_btn.setFixedHeight(26)
        _sc_path = str(Path(__file__).parent / "ScreenConnect.png")
        qc_btn.setIcon(QIcon(_sc_path))
        qc_btn.setIconSize(QSize(28, 28))
        qc_btn.clicked.connect(self.quick_connect_requested)
        layout.addWidget(qc_btn)
        layout.addSpacing(6)
        self._qc_btn = qc_btn

        search_btn = QPushButton("\U0001f50d")
        search_btn.setFixedSize(22, 22)
        search_btn.setCursor(Qt.PointingHandCursor)
        search_btn.setFont(QFont("Segoe UI", 8))
        search_btn.setToolTip("Search nodes  (Ctrl+F)")
        search_btn.clicked.connect(self.search_requested)
        layout.addWidget(search_btn)
        layout.addSpacing(6)
        self._search_btn = search_btn

        help_btn = QPushButton("?")
        help_btn.setFixedSize(22, 22)
        help_btn.setCursor(Qt.PointingHandCursor)
        help_btn.setFont(QFont("Segoe UI", 8, QFont.Weight.Bold))
        help_btn.setToolTip("Help & FAQ")
        help_btn.clicked.connect(self.help_requested)
        layout.addWidget(help_btn)
        layout.addSpacing(10)
        self._help_btn = help_btn
        self._refresh_icon_btn_styles()

        self._btn_min   = TitleBarButton(_theme.BTN_MINIMIZE, "-")
        self._btn_max   = TitleBarButton(_theme.BTN_MAXIMIZE, "+")
        self._btn_close = TitleBarButton(_theme.BTN_CLOSE,    "x")
        for btn in (self._btn_min, self._btn_max, self._btn_close):
            layout.addWidget(btn); layout.addSpacing(5)

        self._btn_min.clicked.connect(
            lambda: self._main_win and self._main_win.showMinimized())
        self._btn_max.clicked.connect(self._toggle_max)
        self._btn_close.clicked.connect(
            lambda: self._main_win and self._main_win.close())

    def _refresh_icon_btn_styles(self):
        g = _theme.GLOW
        ts = _theme.TEXT_SECONDARY
        ss = f"""
            QPushButton {{
                background:rgba({g.red()},{g.green()},{g.blue()},40);
                border:1px solid rgba({g.red()},{g.green()},{g.blue()},100);
                border-radius:11px;
                color:{ts.name()};
            }}
            QPushButton:hover {{
                background:rgba({g.red()},{g.green()},{g.blue()},90);
                color:{g.name()};
                border-color:rgba({g.red()},{g.green()},{g.blue()},180);
            }}
        """
        if hasattr(self, "_search_btn"):
            self._search_btn.setStyleSheet(ss)
        if hasattr(self, "_help_btn"):
            self._help_btn.setStyleSheet(ss)
        if hasattr(self, "_qc_btn"):
            self._qc_btn.setStyleSheet(
                f"""QPushButton {{
                        background:rgba({g.red()},{g.green()},{g.blue()},35);
                        border:1px solid rgba({g.red()},{g.green()},{g.blue()},110);
                        border-radius:10px;
                        color:{ts.name()};
                        padding:0 10px 0 8px;
                    }}
                    QPushButton:hover {{
                        background:rgba({g.red()},{g.green()},{g.blue()},85);
                        color:{g.name()};
                        border-color:rgba({g.red()},{g.green()},{g.blue()},200);
                    }}
                    QPushButton:pressed {{
                        background:rgba({g.red()},{g.green()},{g.blue()},130);
                    }}"""
            )

    def refresh_theme(self):
        self._refresh_icon_btn_styles()
        self.update()

    def set_title(self, t: str):
        self._title_label.set_permanent(t)

    def _toggle_max(self):
        if not self._main_win: return
        if self._main_win.isMaximized():
            self._main_win.showNormal()
        else:
            self._main_win.showMaximized()

    def mousePressEvent(self, e: QMouseEvent):
        if e.button() == Qt.LeftButton:
            self._drag_pos = e.globalPosition().toPoint()

    def mouseMoveEvent(self, e: QMouseEvent):
        if self._drag_pos and self._main_win and not self._main_win.isMaximized():
            delta = e.globalPosition().toPoint() - self._drag_pos
            self._main_win.move(self._main_win.pos() + delta)
            self._drag_pos = e.globalPosition().toPoint()

    def mouseReleaseEvent(self, e: QMouseEvent): self._drag_pos = None
    def mouseDoubleClickEvent(self, e: QMouseEvent): self._toggle_max()

    def paintEvent(self, e: QPaintEvent):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        grad = QLinearGradient(0, 0, self.width(), 0)
        grad.setColorAt(0, t.TITLEBAR_BG)
        grad.setColorAt(0.5, QColor(t.TITLEBAR_BG.red()+4, t.TITLEBAR_BG.green()+4, t.TITLEBAR_BG.blue()+8, 250))
        grad.setColorAt(1, QColor(t.BG_GRAD_BOT.red(), t.BG_GRAD_BOT.green()+2, t.BG_GRAD_BOT.blue()+8, 250))
        p.fillRect(self.rect(), grad)

        # Bottom separator with gradient glow
        sep_grad = QLinearGradient(0, 0, self.width(), 0)
        sep_grad.setColorAt(0,   QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 0))
        sep_grad.setColorAt(0.2, QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 55))
        sep_grad.setColorAt(0.8, QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 55))
        sep_grad.setColorAt(1,   QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 0))
        p.setPen(QPen(QBrush(sep_grad), 1))
        p.drawLine(0, self.height() - 1, self.width(), self.height() - 1)
        p.end()


# ---------------------------------------------------------------------------
# Plugin-buttons scroll area (embedded in footer)
# ---------------------------------------------------------------------------

class _PluginScrollArea(QWidget):
    """Horizontal scrollable strip that holds plugin-injected footer buttons.

    Built-in footer buttons are fixed on the left; this widget expands to
    fill the remaining space and lets plugin buttons scroll horizontally so
    the footer never overflows no matter how many plugins are installed.
    A right-edge chevron fades in whenever there is hidden content to the right.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.setFixedHeight(FOOTER_HEIGHT)

        self._inner = QWidget()
        self._inner.setStyleSheet("background: transparent;")
        self._inner_layout = QHBoxLayout(self._inner)
        self._inner_layout.setContentsMargins(4, 0, 4, 0)
        self._inner_layout.setSpacing(4)
        self._inner_layout.addStretch()

        self._scroll = QScrollArea()
        self._scroll.setWidget(self._inner)
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setStyleSheet(
            "QScrollArea { background: transparent; border: none; }"
            "QWidget      { background: transparent; }"
        )

        # Right-edge scroll indicator — shown when content overflows right
        self._indicator = QLabel("\u276f", self)  # ❯ chevron
        self._indicator.setFont(QFont("Segoe UI", 8, QFont.Weight.Bold))
        self._indicator.setFixedSize(14, FOOTER_HEIGHT)
        self._indicator.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._indicator.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self._indicator.hide()

        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        lay.addWidget(self._scroll)

        # Connect scrollbar changes to indicator update
        self._scroll.horizontalScrollBar().rangeChanged.connect(self._update_indicator)
        self._scroll.horizontalScrollBar().valueChanged.connect(self._update_indicator)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # Keep indicator pinned to the right edge
        self._indicator.move(self.width() - self._indicator.width(), 0)
        self._update_indicator()

    def _update_indicator(self):
        """Show or hide the scroll chevron based on overflow."""
        sb = self._scroll.horizontalScrollBar()
        can_scroll_right = sb.value() < sb.maximum()
        t = _theme
        self._indicator.setStyleSheet(
            f"color: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},180);"
            "background: transparent;"
        )
        if can_scroll_right:
            self._indicator.show()
            self._indicator.raise_()
        else:
            self._indicator.hide()

    def refresh_theme(self):
        self._update_indicator()

    # ── scroll with mouse wheel (horizontal) ──────────────────────────────
    def wheelEvent(self, event):
        sb = self._scroll.horizontalScrollBar()
        delta = -event.angleDelta().y() // 2
        sb.setValue(sb.value() + delta)
        event.accept()

    # ── button management ─────────────────────────────────────────────────
    def add_button(self, btn: QPushButton) -> None:
        """Insert btn before the trailing stretch."""
        idx = max(0, self._inner_layout.count() - 1)
        self._inner_layout.insertWidget(idx, btn)
        self._inner.adjustSize()
        self._update_indicator()

    def remove_button(self, btn: QPushButton) -> None:
        self._inner_layout.removeWidget(btn)
        self._inner.adjustSize()
        self._update_indicator()

    def has_buttons(self) -> bool:
        """True when at least one plugin button is present."""
        return self._inner_layout.count() > 1


# ---------------------------------------------------------------------------
# Footer toolbar
# ---------------------------------------------------------------------------

class FooterToolBar(QWidget):
    settings_requested       = Signal()
    clipboard_requested      = Signal()
    plugins_requested        = Signal()
    new_folder_requested     = Signal()
    time_tracker_requested   = Signal()
    notebook_requested       = Signal()
    reminders_requested      = Signal()
    media_library_requested  = Signal()
    lock_screen_requested    = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(FOOTER_HEIGHT)
        self._build_ui()

    def _build_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 0, 8, 0)
        layout.setSpacing(4)

        def _sep():
            s = QFrame(); s.setFrameShape(QFrame.VLine)
            s.setStyleSheet("color: rgba(0,180,200,40);")
            layout.addSpacing(4); layout.addWidget(s); layout.addSpacing(4)

        # Primary actions — left side
        layout.addWidget(self._make_btn("Settings",      self.settings_requested.emit))
        _sep()
        layout.addWidget(self._make_btn("📋 Clipboard",  self.clipboard_requested.emit))
        _sep()
        layout.addWidget(self._make_btn("Plugins",       self.plugins_requested.emit))
        _sep()
        layout.addWidget(self._make_btn("New Folder",    self.new_folder_requested.emit))
        _sep()
        layout.addWidget(self._make_btn("Time Tracker",  self.time_tracker_requested.emit))
        _sep()
        layout.addWidget(self._make_btn("Notebook",      self.notebook_requested.emit))
        _sep()
        layout.addWidget(self._make_btn("Reminders",   self.reminders_requested.emit))
        _sep()
        layout.addWidget(self._make_btn("Media Library", self.media_library_requested.emit))
        _sep()
        layout.addWidget(self._make_btn("File Explorer", self._open_file_explorer))
        _sep()
        layout.addWidget(self._make_btn("Calculator",    self._open_calculator))
        _sep()
        layout.addWidget(self._make_btn("Lock Screen",   self.lock_screen_requested.emit))

        # Plugin-injected buttons live here — horizontally scrollable so the
        # footer never overflows regardless of how many plugins are installed.
        _sep()
        self._plugin_area = _PluginScrollArea()
        layout.addWidget(self._plugin_area)

        # Secondary action + version — right side
        layout.addWidget(self._make_btn("Feedback",      self._open_feedback))
        layout.addSpacing(12)
        self._ver_lbl = QLabel(f"v{APP_VERSION}")
        self._ver_lbl.setFont(FONT_SMALL)
        self._ver_lbl.setStyleSheet(f"color: {_theme.TEXT_DIM.name()};")
        layout.addWidget(self._ver_lbl)

    def _make_btn(self, label: str, slot) -> QPushButton:
        t = _theme
        btn = QPushButton(label)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setFont(QFont("Segoe UI", 8))
        btn.setFixedHeight(24)
        btn.clicked.connect(slot)
        self._apply_btn_style(btn)
        if not hasattr(self, "_managed_btns"):
            self._managed_btns: list[QPushButton] = []
        self._managed_btns.append(btn)
        return btn

    def _apply_btn_style(self, btn: QPushButton):
        t = _theme
        btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: 1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
                border-radius: 5px;
                color: {t.TEXT_SECONDARY.name()};
                padding: 0 10px;
            }}
            QPushButton:hover {{
                background: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},30);
                color: {t.GLOW.name()};
                border-color: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},130);
            }}
            QPushButton:pressed {{
                background: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);
            }}
        """)

    def refresh_theme(self):
        for btn in getattr(self, "_managed_btns", []):
            self._apply_btn_style(btn)
        if hasattr(self, "_ver_lbl"):
            self._ver_lbl.setStyleSheet(f"color: {_theme.TEXT_DIM.name()};")
        if hasattr(self, "_plugin_area"):
            self._plugin_area.refresh_theme()
        self.update()

    @staticmethod
    def _open_calculator():
        try:
            subprocess.Popen(["calc.exe"], shell=False)
        except OSError:
            subprocess.Popen(["calc"], shell=True)

    @staticmethod
    def _open_file_explorer():
        try:
            subprocess.Popen(["explorer.exe"], shell=False)
        except OSError:
            subprocess.Popen(["explorer"], shell=True)

    @staticmethod
    def _open_feedback():
        webbrowser.open(
            "mailto:caleb.chaney@navanta.com"
            "?subject=Command%20Center%20Feedback"
        )

    def paintEvent(self, e: QPaintEvent):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        grad = QLinearGradient(0, 0, self.width(), 0)
        grad.setColorAt(0, t.FOOTER_BG)
        grad.setColorAt(1, QColor(t.BG_GRAD_BOT.red(), t.BG_GRAD_BOT.green(), t.BG_GRAD_BOT.blue()+4, 240))
        p.fillRect(self.rect(), grad)
        # Top separator glow line
        sep_grad = QLinearGradient(0, 0, self.width(), 0)
        sep_grad.setColorAt(0,   QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 0))
        sep_grad.setColorAt(0.2, QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 45))
        sep_grad.setColorAt(0.8, QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 45))
        sep_grad.setColorAt(1,   QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 0))
        p.setPen(QPen(QBrush(sep_grad), 1))
        p.drawLine(0, 0, self.width(), 0)
        p.end()


# ---------------------------------------------------------------------------
# Footer toolbar (separator helper)
# ---------------------------------------------------------------------------

def _footer_sep() -> QFrame:
    t = _theme
    sep = QFrame(); sep.setFrameShape(QFrame.VLine)
    sep.setStyleSheet(f"color: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},45);")
    return sep


# ---------------------------------------------------------------------------
# Node tooltip
# ---------------------------------------------------------------------------

class NodeToolTip(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent, Qt.ToolTip | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self._name = self._desc = self._type = ""
        self._hide_timer = QTimer(self)
        self._hide_timer.setSingleShot(True)
        self._hide_timer.timeout.connect(self.hide)

    def show_for(self, name: str, desc: str, node_type: str, gpos: QPoint):
        self._name, self._desc, self._type = name, desc, node_type
        self._hide_timer.stop()
        self.adjustSize()
        self.move(gpos + QPoint(16, 12))
        self.show(); self.update()

    def hide_delayed(self, ms: int = 4000):
        self._hide_timer.start(ms)

    def sizeHint(self):
        fm_n = QFontMetrics(QFont("Segoe UI", 10, QFont.Weight.DemiBold))
        fm_d = QFontMetrics(QFont("Segoe UI", 9))
        w = max(fm_n.horizontalAdvance(self._name),
                fm_d.horizontalAdvance(self._desc[:60])) + 28
        lines = max(1, len(self._desc) // 40 + 1) if self._desc else 1
        return QSize(min(w, 320), 52 + lines * 18)

    def paintEvent(self, e: QPaintEvent):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        rect = QRectF(1, 1, self.width() - 2, self.height() - 2)
        path = QPainterPath(); path.addRoundedRect(rect, 8, 8)
        grad = QLinearGradient(0, 0, 0, self.height())
        bg = t.BG_MID
        grad.setColorAt(0, QColor(bg.red()+4, bg.green()+4, bg.blue()+6, 240))
        grad.setColorAt(1, QColor(t.BG_DARK.red(), t.BG_DARK.green(), t.BG_DARK.blue(), 240))
        p.fillPath(path, grad)
        border = QColor(t.GLOW); border.setAlpha(110)
        p.setPen(QPen(border, 1)); p.drawPath(path)

        # Type badge
        badge = {NODE_TYPE_FILE: t.ACCENT_BLUE,
                 NODE_TYPE_URL:  t.ACCENT_TEAL,
                 NODE_TYPE_NOTE: t.ACCENT_AMBER}.get(self._type, t.ACCENT_BLUE)
        p.setBrush(QBrush(badge)); p.setPen(Qt.NoPen)
        p.drawRoundedRect(10, 11, 6, 6, 3, 3)
        p.setPen(QColor(t.TEXT_PRIMARY))
        p.setFont(QFont("Segoe UI", 10, QFont.Weight.DemiBold))
        p.drawText(QRect(22, 5, self.width() - 28, 22),
                   Qt.AlignLeft | Qt.AlignVCenter, self._name)
        if self._desc:
            p.setPen(QColor(t.TEXT_SECONDARY))
            p.setFont(QFont("Segoe UI", 9))
            p.drawText(QRect(10, 28, self.width() - 20, self.height() - 34),
                       Qt.TextWordWrap | Qt.AlignLeft | Qt.AlignTop, self._desc)
        p.end()


# ---------------------------------------------------------------------------
# Note tile hover-preview card
# ---------------------------------------------------------------------------

_NOTE_PREVIEW_CARD: Optional["_NotePreviewCard"] = None


def _get_note_card() -> "_NotePreviewCard":
    global _NOTE_PREVIEW_CARD
    if _NOTE_PREVIEW_CARD is None:
        _NOTE_PREVIEW_CARD = _NotePreviewCard()
    return _NOTE_PREVIEW_CARD


class _NotePreviewCard(QWidget):
    """Floating scrollable preview of a notebook note, shown on tile hover."""

    def __init__(self):
        super().__init__(None, Qt.ToolTip | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setFixedWidth(330)
        self._current_id: Optional[str] = None
        self._hide_timer = QTimer(self)
        self._hide_timer.setSingleShot(True)
        self._hide_timer.setInterval(120)
        self._hide_timer.timeout.connect(self.hide)
        self._build_ui()

    def _build_ui(self):
        from PySide6.QtWidgets import QTextBrowser
        lay = QVBoxLayout(self); lay.setContentsMargins(0, 0, 0, 0)
        self._card = QWidget(); self._card.setObjectName("npc_card")
        cl = QVBoxLayout(self._card)
        cl.setContentsMargins(12, 8, 12, 10); cl.setSpacing(4)
        # Title row
        hrow = QHBoxLayout(); hrow.setSpacing(6)
        self._icon_lbl = QLabel("📝"); self._icon_lbl.setFont(QFont("Segoe UI", 9))
        self._icon_lbl.setStyleSheet("background:transparent;")
        self._title_lbl = QLabel("")
        self._title_lbl.setFont(QFont("Segoe UI", 9, QFont.Weight.DemiBold))
        hrow.addWidget(self._icon_lbl); hrow.addWidget(self._title_lbl); hrow.addStretch()
        cl.addLayout(hrow)
        self._sep = QFrame(); self._sep.setFrameShape(QFrame.HLine)
        cl.addWidget(self._sep)
        # Scrollable content
        self._browser = QTextBrowser()
        self._browser.setMaximumHeight(240)
        self._browser.setMinimumHeight(60)
        self._browser.setFrameShape(QFrame.NoFrame)
        self._browser.setOpenExternalLinks(False)
        self._browser.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        cl.addWidget(self._browser)
        # Hint label
        self._hint_lbl = QLabel("Scroll here to read  ·  Click tile to open")
        self._hint_lbl.setFont(QFont("Segoe UI", 7)); self._hint_lbl.setAlignment(Qt.AlignCenter)
        cl.addWidget(self._hint_lbl)
        lay.addWidget(self._card)
        self._restyle()

    def _restyle(self):
        t = _theme
        g = t.GLOW
        self._card.setStyleSheet(f"""
            QWidget#npc_card {{
                background:rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},245);
                border:1px solid rgba({g.red()},{g.green()},{g.blue()},100);
                border-radius:10px;
            }}
        """)
        self._title_lbl.setStyleSheet(
            f"color:{t.TEXT_PRIMARY.name()}; background:transparent;")
        self._sep.setStyleSheet(
            f"color:rgba({g.red()},{g.green()},{g.blue()},40);")
        self._browser.setStyleSheet(f"""
            QTextBrowser {{
                background:transparent; border:none;
                color:{t.TEXT_SECONDARY.name()};
                font-family:'Segoe UI'; font-size:8.5pt;
                padding:0;
            }}
            QScrollBar:vertical {{
                background:transparent; width:4px; border-radius:2px;
            }}
            QScrollBar::handle:vertical {{
                background:rgba({g.red()},{g.green()},{g.blue()},80);
                border-radius:2px;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height:0; }}
        """)
        self._hint_lbl.setStyleSheet(
            f"color:{t.TEXT_DIM.name()}; background:transparent;")

    def show_for(self, note_id: str, global_pos: QPoint):
        self._hide_timer.stop()
        self._restyle()   # always sync to current theme
        if note_id != self._current_id or not self.isVisible():
            self._current_id = note_id
            data = NotebookStore.load_note(note_id)
            self._title_lbl.setText(data.get("title", "Untitled"))
            html = data.get("content", "")
            if html:
                self._browser.setHtml(html)
            else:
                self._browser.clear()
                self._browser.setPlaceholderText("(empty note)")
        self.adjustSize()
        screen = QGuiApplication.primaryScreen().availableGeometry()
        x = min(global_pos.x(), screen.right() - self.width() - 12)
        y = global_pos.y() - self.height() - 10
        if y < screen.top() + 10:
            y = global_pos.y() + 10
        self.move(x, y)
        self.show(); self.raise_()

    def hide_delayed(self):
        self._hide_timer.start()

    def scroll_by(self, delta: int):
        """Forward a wheel scroll from the hovering tile into the browser."""
        sb = self._browser.verticalScrollBar()
        sb.setValue(sb.value() - delta)

    def paintEvent(self, e):
        pass  # card widget handles its own background


# ---------------------------------------------------------------------------
# Node tile
# ---------------------------------------------------------------------------

class NodeTile(QWidget):
    launch_requested    = Signal(dict)
    edit_requested      = Signal(dict)
    delete_requested    = Signal(dict)
    archive_requested   = Signal(dict)
    export_requested    = Signal(dict)
    duplicate_requested = Signal(dict)
    remove_from_folder_requested = Signal(dict)
    # Drag is handled entirely by NodeCanvas via eventFilter

    def __init__(self, node: dict, tooltip_widget: NodeToolTip, parent=None):
        super().__init__(parent)
        self._node = node
        self._tooltip = tooltip_widget
        self._hovered = False
        self._pressed = False
        self._dragging_visual = False   # set by canvas during drag
        self._anim = 0.0
        self._setup_size()
        self.setMouseTracking(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setGraphicsEffect(make_shadow(self, 20, QColor(0, 0, 0, 140)))
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)

    def _setup_size(self):
        cols, rows = NODE_SIZES.get(self._node.get("size", "1x1"), (1, 1))
        # pixel size = number_of_cells * cell_size - one trailing gap
        w = cols * TILE_BASE_SIZE + (cols - 1) * TILE_GAP
        h = rows * TILE_BASE_SIZE + (rows - 1) * TILE_GAP
        self.setFixedSize(w, h)

    def node_data(self) -> dict: return self._node

    def update_node(self, node: dict):
        self._node = node; self._setup_size(); self.update()

    def enterEvent(self, e):
        self._hovered = True; self._timer.start(16)
        nt = self._node.get("type", NODE_TYPE_FILE)
        if nt == NODE_TYPE_NOTE:
            note_id = self._node.get("target", "")
            if note_id:
                _get_note_card().show_for(
                    note_id,
                    self.mapToGlobal(QPoint(self.width() + 6, self.height() // 2)))
        elif self._node.get("description"):
            self._tooltip.show_for(self._node.get("name", ""),
                                   self._node.get("description", ""),
                                   nt,
                                   QCursor.pos())
            self._tooltip.hide_delayed()

    def leaveEvent(self, e):
        self._hovered = False; self._timer.start(16)
        self._tooltip.hide()
        _get_note_card().hide_delayed()

    def wheelEvent(self, e):
        card = _get_note_card()
        if card.isVisible() and self._node.get("type") == NODE_TYPE_NOTE:
            # Forward wheel scroll into the preview browser instead of the canvas
            delta = e.angleDelta().y() // 2
            card.scroll_by(delta)
            e.accept()
        else:
            e.ignore()

    def _tick(self):
        target = 1.0 if self._hovered else 0.0
        if abs(self._anim - target) < 0.08:
            self._anim = target; self._timer.stop()
        else:
            self._anim += 0.08 if target > self._anim else -0.08
        self.update()

    def set_dragging(self, on: bool):
        """Canvas calls this to toggle the drag-lifted visual state."""
        self._dragging_visual = on
        self.setCursor(Qt.ClosedHandCursor if on else Qt.PointingHandCursor)
        self.update()

    def mousePressEvent(self, e: QMouseEvent):
        if e.button() == Qt.LeftButton:
            self._pressed = True; self.update()

    def mouseReleaseEvent(self, e: QMouseEvent):
        if e.button() == Qt.LeftButton:
            self._pressed = False; self.update()
            # Only emit launch if canvas did NOT consume this as a drag
            if not self._dragging_visual and self.rect().contains(e.position().toPoint()):
                if _settings_store.value("launch_on_single_click", "true") == "true":
                    self.launch_requested.emit(self._node)

    def mouseDoubleClickEvent(self, e: QMouseEvent):
        if e.button() == Qt.LeftButton:
            if _settings_store.value("launch_on_single_click", "true") == "false":
                self.launch_requested.emit(self._node)

    def contextMenuEvent(self, e):
        menu = QMenu(self)
        menu.setStyleSheet(self._menu_css())
        in_folder = bool(self._node.get("folder_id"))
        if in_folder:
            act_remove_from_folder = menu.addAction("⬅  Remove From Folder")
            menu.addSeparator()
        act_run     = menu.addAction("Run")
        act_edit    = menu.addAction("Edit")
        act_export  = menu.addAction("Export  (.node file)")
        act_duplicate = menu.addAction("Duplicate")
        menu.addSeparator()
        # Actions that depend on node type / target availability
        target  = self._node.get("target", "")
        nt      = self._node.get("type", NODE_TYPE_FILE)
        is_file = nt == NODE_TYPE_FILE and bool(target)
        is_exe  = is_file and Path(target).suffix.lower() in {".exe", ".com", ".bat", ".cmd", ".ps1", ".py"}
        act_open_location = menu.addAction("Open File Location")
        act_copy_path     = menu.addAction("Copy Path")
        act_run_admin     = menu.addAction("Run as Administrator")
        act_open_location.setEnabled(is_file and os.path.isfile(target))
        act_copy_path.setEnabled(bool(target))
        act_run_admin.setEnabled(is_exe)
        menu.addSeparator()
        act_copy_file     = menu.addAction("Copy File to Clipboard")
        act_copy_contents = menu.addAction("Copy Contents to Clipboard")
        act_copy_file.setEnabled(is_file and os.path.isfile(target))
        act_copy_contents.setEnabled(is_file and os.path.isfile(target))
        menu.addSeparator()
        act_archive = menu.addAction("Archive")
        act_delete  = menu.addAction("Delete")
        chosen = menu.exec(e.globalPos())
        if in_folder and chosen == act_remove_from_folder:
            self.remove_from_folder_requested.emit(self._node)
        elif chosen == act_run:             self.launch_requested.emit(self._node)
        elif chosen == act_edit:          self.edit_requested.emit(self._node)
        elif chosen == act_export:        self.export_requested.emit(self._node)
        elif chosen == act_duplicate:     self.duplicate_requested.emit(self._node)
        elif chosen == act_open_location: self._open_file_location(target)
        elif chosen == act_copy_path:     self._copy_path_to_clipboard(target)
        elif chosen == act_run_admin:     self._run_as_admin(target, self._node.get("params", ""))
        elif chosen == act_copy_file:     self._copy_file_to_clipboard(target)
        elif chosen == act_copy_contents: self._copy_contents_to_clipboard(target)
        elif chosen == act_archive:       self.archive_requested.emit(self._node)
        elif chosen == act_delete:        self.delete_requested.emit(self._node)

    def _open_file_location(self, path: str):
        """Open Windows Explorer with the file highlighted."""
        try:
            abs_path = os.path.abspath(path)
            subprocess.Popen(["explorer", "/select,", abs_path])
        except Exception as exc:
            print(f"[CommandCenter] Open file location error: {exc}", file=sys.stderr)

    def _copy_path_to_clipboard(self, path: str):
        """Copy the raw target path/URL string to the clipboard."""
        try:
            QApplication.clipboard().setText(path)
        except Exception as exc:
            print(f"[CommandCenter] Copy path error: {exc}", file=sys.stderr)

    def _run_as_admin(self, target: str, params: str = ""):
        """Launch the target with the runas verb (triggers UAC prompt)."""
        try:
            ret = ctypes.windll.shell32.ShellExecuteW(
                None, "runas", target, params if params else None, None, 1)
            if ret <= 32:
                # User likely cancelled the UAC prompt or access was denied
                print(
                    f"[CommandCenter] Run as admin: ShellExecute returned {ret} (user may have cancelled UAC)",
                    file=sys.stderr)
        except OSError as exc:
            print(f"[CommandCenter] Run as admin error: {exc}", file=sys.stderr)

    def _copy_file_to_clipboard(self, path: str):
        try:
            mime = QMimeData()
            url = QUrl.fromLocalFile(os.path.abspath(path))
            mime.setUrls([url])
            QApplication.clipboard().setMimeData(mime)
        except Exception as exc:
            print(f"[CommandCenter] Clipboard copy error: {exc}", file=sys.stderr)

    def _copy_contents_to_clipboard(self, path: str):
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                text = f.read()
            QApplication.clipboard().setText(text)
        except Exception as exc:
            print(f"[CommandCenter] Clipboard copy-contents error: {exc}", file=sys.stderr)

    def _menu_css(self):
        t = _theme
        return f"""
        QMenu {{
            background: rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},245);
            border: 1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},110);
            border-radius: 8px;
            padding: 4px 0;
            color: {t.TEXT_PRIMARY.name()};
            font-family:'Segoe UI'; font-size:9pt;
        }}
        QMenu::item {{ padding:6px 18px; border-radius:4px; }}
        QMenu::item:selected {{
            background: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);
            color: {t.GLOW.name()};
        }}
        QMenu::separator {{
            height:1px; background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},45); margin:3px 8px;
        }}
        """

    def paintEvent(self, e: QPaintEvent):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        w, h = self.width(), self.height()
        rect = QRectF(1, 1, w - 2, h - 2)
        path = QPainterPath(); path.addRoundedRect(rect, BORDER_RADIUS, BORDER_RADIUS)

        # --- Background glass card ---
        grad = QLinearGradient(0, 0, w * 0.6, h)
        ba = int(215 + 30 * self._anim)
        base = t.TILE_BG_BASE
        hov  = t.TILE_BG_HOVER
        if self._pressed:
            r = max(base.red()   - 6, 0)
            g = max(base.green() - 6, 0)
            b = max(base.blue()  - 6, 0)
            grad.setColorAt(0, QColor(r, g, b, ba))
            grad.setColorAt(1, QColor(r, g, b, ba - 20))
        elif self._hovered:
            grad.setColorAt(0, QColor(hov.red(), hov.green(), hov.blue(), ba))
            grad.setColorAt(0.6, QColor(hov.red()+6, hov.green()+6, hov.blue()+10, ba))
            grad.setColorAt(1, QColor(hov.red()-4, hov.green()-4, hov.blue()+2, ba - 20))
        else:
            grad.setColorAt(0, QColor(base.red(), base.green(), base.blue(), 210))
            grad.setColorAt(1, QColor(base.red()-4, base.green()-4, base.blue()+2, 200))
        p.fillPath(path, grad)

        # --- Glass highlight streak (top edge shimmer) ---
        if self._hovered or self._anim > 0:
            shimmer = QLinearGradient(0, 0, w, 0)
            sa = int(18 * self._anim)
            shimmer.setColorAt(0,    QColor(255, 255, 255, 0))
            shimmer.setColorAt(0.3,  QColor(255, 255, 255, sa))
            shimmer.setColorAt(0.7,  QColor(255, 255, 255, sa // 2))
            shimmer.setColorAt(1.0,  QColor(255, 255, 255, 0))
            p.fillRect(QRectF(1, 1, w - 2, 2), QBrush(shimmer))

        # --- Border (per-node accent color takes priority over theme glow) ---
        _ac = self._node.get("accent_color", "")
        gc = QColor(_ac) if _ac and QColor(_ac).isValid() else t.GLOW
        if self._dragging_visual:
            pen = QPen(QColor(gc.red(), gc.green(), gc.blue(), 255), 2.5)
        elif self._hovered:
            pen = QPen(QColor(gc.red(), gc.green(), gc.blue(),
                              int(200 + 55 * self._anim)), 2.0)
        else:
            pen = QPen(QColor(gc.red(), gc.green(), gc.blue(), 150), 1.5)
        p.setPen(pen); p.setBrush(Qt.NoBrush); p.drawPath(path)

        # --- Drag lifted outer glow ---
        if self._dragging_visual:
            pg = QColor(gc); pg.setAlpha(50)
            p.setPen(QPen(pg, 8)); p.drawPath(path)

        # --- Icon ---
        icon_path = self._node.get("icon", "")
        icon_size = min(w, h) // 3 + 4
        icon_drawn = False
        if icon_path and os.path.isfile(icon_path):
            pm = load_pixmap_cached(icon_path, icon_size)
            if not pm.isNull():
                ix = (w - pm.width()) // 2
                iy = (h - pm.height()) // 2 - 14
                p.drawPixmap(ix, iy, pm); icon_drawn = True

        if not icon_drawn:
            glyph = {"file": "F", "url": "W", "note": "📝"}.get(
                self._node.get("type", "file"), "?")
            if glyph == "📝":
                fs = max(icon_size // 2 + 4, 20)
                p.setFont(QFont("Segoe UI Emoji", fs))
            else:
                fs = max(icon_size // 2, 16)
                p.setFont(QFont("Segoe UI", fs, QFont.Weight.Bold))
            gc = QColor(t.ACCENT_TEAL)
            gc.setAlpha(int(80 + 110 * self._anim))
            p.setPen(gc)
            p.drawText(QRect(0, 0, w, h - 24), Qt.AlignCenter, glyph)

        # --- Name label with subtle backing ---
        name = self._node.get("name", "")
        if name:
            label_rect = QRect(4, h - 28, w - 8, 22)
            # Frosted label backing
            lbg = QColor(t.BG_DARK); lbg.setAlpha(140)
            p.setBrush(lbg); p.setPen(Qt.NoPen)
            p.drawRoundedRect(label_rect, 4, 4)
            p.setFont(QFont("Segoe UI", 8, QFont.Weight.DemiBold))
            lc = QColor(t.TEXT_PRIMARY); lc.setAlpha(220 + int(35 * self._anim))
            p.setPen(lc)
            fm = QFontMetrics(p.font())
            p.drawText(label_rect, Qt.AlignHCenter | Qt.AlignVCenter,
                       fm.elidedText(name, Qt.ElideRight, label_rect.width() - 6))

        # --- Schedule clock badge ---
        sched = self._node.get("schedule")
        if sched and isinstance(sched, dict) and sched.get("enabled"):
            badge_size = 14
            bx, by = w - badge_size - 4, 4
            badge_bg = QColor(t.ACCENT_AMBER); badge_bg.setAlpha(210)
            p.setBrush(badge_bg); p.setPen(Qt.NoPen)
            p.drawEllipse(bx, by, badge_size, badge_size)
            p.setPen(QColor(0, 0, 0, 200))
            p.setFont(QFont("Segoe UI Emoji", 7))
            p.drawText(QRect(bx, by, badge_size, badge_size), Qt.AlignCenter, "⏰")

        p.end()


# ---------------------------------------------------------------------------
# Add-node tile (+)
# ---------------------------------------------------------------------------

class AddNodeTile(QWidget):
    add_requested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(TILE_BASE_SIZE, TILE_BASE_SIZE)
        self._hovered = False; self._anim = 0.0
        self.setCursor(Qt.PointingHandCursor)
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)

    def enterEvent(self, e): self._hovered = True;  self._timer.start(16)
    def leaveEvent(self, e): self._hovered = False; self._timer.start(16)

    def _tick(self):
        target = 1.0 if self._hovered else 0.0
        if abs(self._anim - target) < 0.08:
            self._anim = target; self._timer.stop()
        else:
            self._anim += 0.08 if target > self._anim else -0.08
        self.update()

    def mousePressEvent(self, e: QMouseEvent):
        if e.button() == Qt.LeftButton: self.add_requested.emit()

    def paintEvent(self, e: QPaintEvent):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        rect = QRectF(1, 1, self.width() - 2, self.height() - 2)
        path = QPainterPath(); path.addRoundedRect(rect, BORDER_RADIUS, BORDER_RADIUS)
        grad = QLinearGradient(0, 0, self.width(), self.height())
        a = int(40 + 70 * self._anim)
        base = t.TILE_BG_BASE
        grad.setColorAt(0, QColor(base.red(), base.green(), base.blue(), a))
        grad.setColorAt(1, QColor(base.red()-4, base.green()-4, base.blue()+2, max(a-20, 0)))
        p.fillPath(path, grad)
        # Dashed border
        glow_r, glow_g, glow_b = t.GLOW.red(), t.GLOW.green(), t.GLOW.blue()
        dc = QColor(glow_r, glow_g, glow_b, int(80 + 120 * self._anim))
        p.setPen(QPen(dc, 1.4, Qt.DashLine)); p.setBrush(Qt.NoBrush); p.drawPath(path)
        cx, cy = self.width() // 2, self.height() // 2
        arm = 15 + int(5 * self._anim)
        pc = QColor(t.GLOW if self._hovered else t.ACCENT_TEAL)
        pc.setAlpha(int(130 + 100 * self._anim))
        p.setPen(QPen(pc, 2.4, Qt.SolidLine, Qt.RoundCap))
        p.drawLine(cx - arm, cy, cx + arm, cy)
        p.drawLine(cx, cy - arm, cx, cy + arm)
        if self._hovered or self._anim > 0.3:
            p.setFont(QFont("Segoe UI", 8))
            lc = QColor(t.TEXT_SECONDARY); lc.setAlpha(int(180 * self._anim))
            p.setPen(lc)
            p.drawText(QRect(0, cy + 22, self.width(), 18), Qt.AlignCenter, "Add Node")
        p.end()


# ---------------------------------------------------------------------------
# Icon drop zone
# ---------------------------------------------------------------------------

class IconDropZone(QWidget):
    icon_changed = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(80, 80)
        self.setAcceptDrops(True)
        self._path = ""; self._hovered = False
        self.setCursor(Qt.PointingHandCursor)

    def current_path(self) -> str: return self._path

    def set_path(self, path: str): self._path = path; self.update()

    def mousePressEvent(self, e: QMouseEvent):
        if e.button() == Qt.LeftButton:
            path, _ = QFileDialog.getOpenFileName(
                self, "Select Icon", "",
                "Images (*.png *.jpg *.jpeg *.bmp *.ico *.svg *.webp)")
            if path:
                self._path = cache_icon(path)
                self.icon_changed.emit(self._path); self.update()

    def dragEnterEvent(self, e: QDragEnterEvent):
        if e.mimeData().hasUrls():
            url = e.mimeData().urls()[0]
            if url.isLocalFile():
                ext = Path(url.toLocalFile()).suffix.lower()
                if ext in (".png",".jpg",".jpeg",".bmp",".ico",".svg",".webp"):
                    e.acceptProposedAction()
                    self._hovered = True; self.update(); return
        e.ignore()

    def dragLeaveEvent(self, e): self._hovered = False; self.update()

    def dropEvent(self, e: QDropEvent):
        self._hovered = False
        self._path = cache_icon(e.mimeData().urls()[0].toLocalFile())
        self.icon_changed.emit(self._path); self.update()

    def paintEvent(self, e):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        rect = QRectF(1, 1, self.width() - 2, self.height() - 2)
        path = QPainterPath(); path.addRoundedRect(rect, 7, 7)
        bg = t.TILE_BG_BASE
        p.fillPath(path, QColor(bg.red(), bg.green(), bg.blue(), 200))
        pc = QColor(t.GLOW if self._hovered else t.ACCENT_TEAL)
        p.setPen(QPen(pc, 1.4, Qt.SolidLine if self._path else Qt.DashLine))
        p.setBrush(Qt.NoBrush); p.drawPath(path)
        if self._path and os.path.isfile(self._path):
            pm = load_pixmap_cached(self._path, 64)
            if not pm.isNull():
                p.drawPixmap((self.width()-pm.width())//2,
                             (self.height()-pm.height())//2, pm)
        else:
            p.setPen(QColor(t.TEXT_DIM)); p.setFont(QFont("Segoe UI", 8))
            p.drawText(self.rect(), Qt.AlignCenter, "Icon\n(drop or click)")
        p.end()


# ---------------------------------------------------------------------------
# Node Creation / Edit Wizard
# ---------------------------------------------------------------------------

class NodeWizard(QDialog):
    node_saved = Signal(dict)

    def __init__(self, existing_node: Optional[dict] = None, parent=None):
        super().__init__(parent, Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self._existing = existing_node
        self._drag_pos: Optional[QPoint] = None
        self._build_ui()
        if existing_node: self._populate(existing_node)
        self.setMinimumSize(520, 660)

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(10, 10, 10, 10); outer.setSpacing(0)
        self._card = QWidget()
        cl = QVBoxLayout(self._card)
        cl.setContentsMargins(0, 0, 0, 0); cl.setSpacing(0)
        self._card.setGraphicsEffect(make_shadow(self._card, 30, QColor(0,0,0,200)))

        tbar = QWidget(); tbar.setFixedHeight(44)
        tbar.setStyleSheet("background:transparent;")
        tbl = QHBoxLayout(tbar); tbl.setContentsMargins(16,0,12,0)
        ttl = QLabel("Node Creation Wizard" if not self._existing else "Edit Node")
        ttl.setFont(QFont("Segoe UI",11,QFont.Weight.DemiBold))
        ttl.setStyleSheet(f"color:{_theme.TEXT_PRIMARY.name()};")
        tbl.addWidget(ttl); tbl.addStretch()

        imp_btn = QPushButton("Import .node")
        imp_btn.setCursor(Qt.PointingHandCursor)
        imp_btn.setFont(QFont("Segoe UI",8)); imp_btn.setFixedHeight(24)
        imp_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({_theme.TILE_BG_BASE.red()},{_theme.TILE_BG_BASE.green()},{_theme.TILE_BG_BASE.blue()},160);
                border:1px solid rgba({_theme.GLOW.red()},{_theme.GLOW.green()},{_theme.GLOW.blue()},80);
                border-radius:5px;
                color:{_theme.TEXT_SECONDARY.name()};
                padding:0 10px;
            }}
            QPushButton:hover {{
                background:rgba({_theme.GLOW.red()},{_theme.GLOW.green()},{_theme.GLOW.blue()},40);
                color:{_theme.GLOW.name()};
            }}
        """)
        imp_btn.clicked.connect(self._import_node_file)
        tbl.addWidget(imp_btn); tbl.addSpacing(8)

        cb = TitleBarButton(COLOR_BTN_CLOSE, "x"); cb.clicked.connect(self.reject)
        tbl.addWidget(cb)
        tbar.mousePressEvent = lambda ev: (
            setattr(self, "_drag_pos", ev.globalPosition().toPoint())
            if ev.button() == Qt.LeftButton else None)
        tbar.mouseMoveEvent = self._drag_move
        tbar.mouseReleaseEvent = lambda ev: setattr(self, "_drag_pos", None)
        cl.addWidget(tbar)

        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"color:rgba({_theme.GLOW.red()},{_theme.GLOW.green()},{_theme.GLOW.blue()},55);")
        cl.addWidget(sep)

        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("background:transparent;")
        fw = QWidget(); fw.setStyleSheet("background:transparent;")
        form = QVBoxLayout(fw)
        form.setContentsMargins(20,16,20,16); form.setSpacing(12)

        form.addWidget(self._sec("Node Size"))
        sr = QHBoxLayout()
        self._size_group = QButtonGroup(self)
        for sk, sl in [("1x1","Small (1x1)"),("2x2","Medium (2x2)"),("2x4","Large (2x4)")]:
            rb = QRadioButton(sl); rb.setProperty("size_key", sk)
            self._style_rb(rb)
            if sk == "1x1": rb.setChecked(True)
            self._size_group.addButton(rb); sr.addWidget(rb)
        form.addLayout(sr)

        form.addWidget(self._sec("Node Name"))
        self._name_edit = self._le("Enter a name for this node")
        self._name_edit.setMaxLength(80)
        form.addWidget(self._name_edit)

        form.addWidget(self._sec("Node Type"))
        tr = QHBoxLayout()
        self._type_group = QButtonGroup(self)
        for tk, tl in [(NODE_TYPE_FILE,"File / App"),(NODE_TYPE_URL,"URL"),(NODE_TYPE_NOTE,"Notebook Note")]:
            rb = QRadioButton(tl); rb.setProperty("type_key", tk)
            self._style_rb(rb)
            if tk == NODE_TYPE_FILE: rb.setChecked(True)
            self._type_group.addButton(rb); tr.addWidget(rb)
        self._type_group.buttonClicked.connect(self._on_type_changed)
        form.addLayout(tr)

        self._target_lbl = QLabel("File Path / URL")
        self._target_lbl.setFont(FONT_LABEL)
        self._target_lbl.setStyleSheet(f"color:{_theme.TEXT_SECONDARY.name()};")
        form.addWidget(self._target_lbl)
        trow = QHBoxLayout()
        self._target_edit = self._le("Path to file or URL")
        br_btn = QPushButton("Browse"); self._style_btn(br_btn, secondary=True, small=True)
        br_btn.clicked.connect(self._browse_target)
        trow.addWidget(self._target_edit); trow.addWidget(br_btn)
        form.addLayout(trow)

        self._note_lbl = QLabel("Select Notebook Note")
        self._note_lbl.setFont(FONT_LABEL)
        self._note_lbl.setStyleSheet(f"color:{_theme.TEXT_SECONDARY.name()};")
        self._note_lbl.setVisible(False); form.addWidget(self._note_lbl)
        self._note_combo = QComboBox()
        self._note_combo.setFont(FONT_LABEL)
        t = _theme
        self._note_combo.setStyleSheet(f"""
            QComboBox {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:5px; color:{t.TEXT_PRIMARY.name()};
                padding:5px 9px; font-family:'Segoe UI'; font-size:9pt;
            }}
            QComboBox::drop-down {{ border:none; width:22px; }}
            QComboBox QAbstractItemView {{
                background:rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},240);
                color:{t.TEXT_PRIMARY.name()};
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},80);
                selection-background-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},55);
            }}
        """)
        self._note_combo.setVisible(False); form.addWidget(self._note_combo)
        self._new_note_btn = QPushButton("+ Create new note")
        self._style_btn(self._new_note_btn, secondary=True, small=True)
        self._new_note_btn.setVisible(False)
        self._new_note_btn.clicked.connect(self._create_new_linked_note)
        form.addWidget(self._new_note_btn)
        # Placeholder for old note_edit (unused, kept for _populate compat)
        self._note_edit = QTextEdit(); self._note_edit.setVisible(False)
        form.addWidget(self._note_edit)

        form.addWidget(self._sec("Launch Parameters (optional)"))
        self._params_edit = self._le("e.g.  --flag value")
        form.addWidget(self._params_edit)

        # ── Open Behavior (file nodes only) ──────────────────────────────
        self._behavior_lbl = self._sec("Open Behavior")
        form.addWidget(self._behavior_lbl)
        self._behavior_combo = QComboBox()
        self._behavior_combo.setFont(FONT_LABEL)
        for label, data in [
            ("Launch normally",               "normal"),
            ("Open containing folder",        "open_folder"),
            ("Run as administrator",          "run_admin"),
            ("Copy file to clipboard",        "copy_file"),
            ("Copy file contents to clipboard", "copy_contents"),
        ]:
            self._behavior_combo.addItem(label, data)
        t = _theme
        self._behavior_combo.setStyleSheet(f"""
            QComboBox {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:5px; color:{t.TEXT_PRIMARY.name()};
                padding:5px 9px; font-family:'Segoe UI'; font-size:9pt;
            }}
            QComboBox::drop-down {{ border:none; width:22px; }}
            QComboBox QAbstractItemView {{
                background:rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},240);
                color:{t.TEXT_PRIMARY.name()};
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},80);
                selection-background-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},55);
            }}
        """)
        form.addWidget(self._behavior_combo)

        form.addWidget(self._sec("Description"))
        self._desc_edit = QTextEdit()
        self._desc_edit.setPlaceholderText("Shown in tooltip when hovering the node...")
        self._desc_edit.setFixedHeight(60); self._style_te(self._desc_edit)
        form.addWidget(self._desc_edit)

        form.addWidget(self._sec("Tags  (comma-separated, optional)"))
        self._tags_edit = self._le("e.g.  work, tools, python")
        form.addWidget(self._tags_edit)

        # ── Tile Accent Color ────────────────────────────────────────────
        form.addWidget(self._sec("Tile Accent Color  (optional)"))
        self._accent_color_val = ""
        accent_row = QHBoxLayout(); accent_row.setSpacing(6)
        self._accent_btn = QPushButton("Default  (use theme color)")
        self._accent_btn.setFixedHeight(28)
        self._accent_btn.setCursor(Qt.PointingHandCursor)
        self._accent_btn.setFont(QFont("Segoe UI", 8))
        self._style_btn(self._accent_btn, secondary=True, small=True)
        self._accent_btn.clicked.connect(self._pick_accent_color)
        self._accent_clear_btn = QPushButton("Clear")
        self._accent_clear_btn.setFixedHeight(28)
        self._accent_clear_btn.setCursor(Qt.PointingHandCursor)
        self._accent_clear_btn.setFont(QFont("Segoe UI", 8))
        self._style_btn(self._accent_clear_btn, secondary=True, small=True)
        self._accent_clear_btn.clicked.connect(self._clear_accent_color)
        accent_row.addWidget(self._accent_btn)
        accent_row.addWidget(self._accent_clear_btn)
        accent_row.addStretch()
        form.addLayout(accent_row)
        accent_note = QLabel("Per-node color overrides the global theme accent on this tile.")
        accent_note.setFont(FONT_SMALL)
        accent_note.setStyleSheet(f"color:{_theme.TEXT_DIM.name()};")
        form.addWidget(accent_note)

        form.addWidget(self._sec("Icon"))
        irow = QHBoxLayout()
        self._icon_zone = IconDropZone()
        irow.addWidget(self._icon_zone)
        inf = QLabel("Drag an image here or click to browse.\nPNG, JPG, ICO, SVG, WEBP.")
        inf.setFont(FONT_SMALL); inf.setStyleSheet(f"color:{_theme.TEXT_DIM.name()};")
        irow.addWidget(inf); irow.addStretch(); form.addLayout(irow)

        self._auto_launch = QCheckBox("Auto-launch when Command Center opens")
        self._style_cb(self._auto_launch); form.addWidget(self._auto_launch)

        # ── Schedule ─────────────────────────────────────────────────────
        form.addWidget(self._sec("Schedule  (optional)"))
        self._sched_enabled = QCheckBox("Enable scheduled execution")
        self._style_cb(self._sched_enabled)
        form.addWidget(self._sched_enabled)

        # type row
        sched_type_row = QHBoxLayout()
        sched_type_lbl = QLabel("Type:")
        sched_type_lbl.setFont(FONT_LABEL)
        sched_type_lbl.setStyleSheet(f"color:{_theme.TEXT_SECONDARY.name()};")
        self._sched_type_combo = QComboBox()
        self._sched_type_combo.setFont(FONT_LABEL)
        for key, label in SCHED_TYPES:
            self._sched_type_combo.addItem(label, key)
        self._sched_combo_style()
        sched_type_row.addWidget(sched_type_lbl)
        sched_type_row.addWidget(self._sched_type_combo)
        sched_type_row.addStretch()
        form.addLayout(sched_type_row)

        # time row  (HH:MM)
        self._sched_time_row = QWidget(); sched_time_lay = QHBoxLayout(self._sched_time_row)
        sched_time_lay.setContentsMargins(0, 0, 0, 0)
        sched_time_lbl = QLabel("Time  (HH:MM):")
        sched_time_lbl.setFont(FONT_LABEL)
        sched_time_lbl.setStyleSheet(f"color:{_theme.TEXT_SECONDARY.name()};")
        self._sched_time_edit = self._le("e.g.  09:00")
        self._sched_time_edit.setMaximumWidth(100)
        sched_time_lay.addWidget(sched_time_lbl); sched_time_lay.addWidget(self._sched_time_edit)
        sched_time_lay.addStretch()
        form.addWidget(self._sched_time_row)

        # date row  (YYYY-MM-DD)  — for ONCE
        self._sched_date_row = QWidget(); sched_date_lay = QHBoxLayout(self._sched_date_row)
        sched_date_lay.setContentsMargins(0, 0, 0, 0)
        sched_date_lbl = QLabel("Date  (YYYY-MM-DD):")
        sched_date_lbl.setFont(FONT_LABEL)
        sched_date_lbl.setStyleSheet(f"color:{_theme.TEXT_SECONDARY.name()};")
        self._sched_date_edit = self._le("e.g.  2026-12-31")
        self._sched_date_edit.setMaximumWidth(140)
        sched_date_lay.addWidget(sched_date_lbl); sched_date_lay.addWidget(self._sched_date_edit)
        sched_date_lay.addStretch()
        form.addWidget(self._sched_date_row)

        # days-of-week checkboxes — for WEEKLY
        self._sched_days_row = QWidget(); sched_days_lay = QHBoxLayout(self._sched_days_row)
        sched_days_lay.setContentsMargins(0, 0, 0, 0)
        _day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        self._sched_day_cbs: list[QCheckBox] = []
        for i, dn in enumerate(_day_names):
            cb = QCheckBox(dn); self._style_cb(cb)
            if i < 5: cb.setChecked(True)   # default: Mon-Fri
            self._sched_day_cbs.append(cb); sched_days_lay.addWidget(cb)
        sched_days_lay.addStretch()
        form.addWidget(self._sched_days_row)

        # day-of-month row — for MONTHLY
        self._sched_dom_row = QWidget(); sched_dom_lay = QHBoxLayout(self._sched_dom_row)
        sched_dom_lay.setContentsMargins(0, 0, 0, 0)
        sched_dom_lbl = QLabel("Day of month  (1–31):")
        sched_dom_lbl.setFont(FONT_LABEL)
        sched_dom_lbl.setStyleSheet(f"color:{_theme.TEXT_SECONDARY.name()};")
        self._sched_dom_edit = self._le("1")
        self._sched_dom_edit.setMaximumWidth(70)
        sched_dom_lay.addWidget(sched_dom_lbl); sched_dom_lay.addWidget(self._sched_dom_edit)
        sched_dom_lay.addStretch()
        form.addWidget(self._sched_dom_row)

        # interval row — for INTERVAL
        self._sched_interval_row = QWidget(); sched_int_lay = QHBoxLayout(self._sched_interval_row)
        sched_int_lay.setContentsMargins(0, 0, 0, 0)
        sched_int_lbl = QLabel("Every:")
        sched_int_lbl.setFont(FONT_LABEL)
        sched_int_lbl.setStyleSheet(f"color:{_theme.TEXT_SECONDARY.name()};")
        self._sched_interval_val = self._le("30")
        self._sched_interval_val.setMaximumWidth(70)
        self._sched_interval_unit = QComboBox()
        self._sched_interval_unit.setFont(FONT_LABEL)
        self._sched_interval_unit.addItem("minutes", "minutes")
        self._sched_interval_unit.addItem("hours", "hours")
        self._sched_combo_style(self._sched_interval_unit)
        sched_int_lay.addWidget(sched_int_lbl)
        sched_int_lay.addWidget(self._sched_interval_val)
        sched_int_lay.addWidget(self._sched_interval_unit)
        sched_int_lay.addStretch()
        form.addWidget(self._sched_interval_row)

        self._sched_note = QLabel(
            "⚠  Scheduled nodes launch silently in the background even if Command Center is minimized.")
        self._sched_note.setFont(FONT_SMALL)
        self._sched_note.setWordWrap(True)
        self._sched_note.setStyleSheet(f"color:{_theme.TEXT_DIM.name()};")
        form.addWidget(self._sched_note)

        # connect signals to update visibility
        self._sched_enabled.toggled.connect(self._refresh_sched_ui)
        self._sched_type_combo.currentIndexChanged.connect(self._refresh_sched_ui)
        self._refresh_sched_ui()

        form.addStretch(); scroll.setWidget(fw); cl.addWidget(scroll)

        brow = QHBoxLayout(); brow.setContentsMargins(20,8,20,14); brow.addStretch()
        cancel = QPushButton("Cancel"); self._style_btn(cancel, secondary=True)
        cancel.clicked.connect(self.reject)
        save = QPushButton("Save Node"); self._style_btn(save)
        save.clicked.connect(self._on_save)
        brow.addWidget(cancel); brow.addWidget(save); cl.addLayout(brow)
        outer.addWidget(self._card)

    def _drag_move(self, e: QMouseEvent):
        if self._drag_pos:
            delta = e.globalPosition().toPoint() - self._drag_pos
            self.move(self.pos() + delta)
            self._drag_pos = e.globalPosition().toPoint()

    def _sec(self, txt):
        t = _theme
        l = QLabel(txt); l.setFont(QFont("Segoe UI",8,QFont.Weight.DemiBold))
        l.setStyleSheet(f"color:{t.GLOW.name()}; letter-spacing:1px;"); return l

    def _sched_combo_style(self, combo: "QComboBox" = None):
        """Apply standard combo-box styling. If combo is None, style _sched_type_combo."""
        target = combo if combo is not None else self._sched_type_combo
        t = _theme
        target.setStyleSheet(f"""
            QComboBox {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:5px; color:{t.TEXT_PRIMARY.name()};
                padding:5px 9px; font-family:'Segoe UI'; font-size:9pt;
            }}
            QComboBox::drop-down {{ border:none; width:22px; }}
            QComboBox QAbstractItemView {{
                background:rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},240);
                color:{t.TEXT_PRIMARY.name()};
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},80);
                selection-background-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},55);
            }}
        """)

    def _refresh_sched_ui(self):
        """Show/hide schedule sub-widgets based on type selection & enabled state."""
        enabled = self._sched_enabled.isChecked()
        stype = self._sched_type_combo.currentData() or SCHED_DAILY
        self._sched_type_combo.setEnabled(enabled)
        has_time     = stype in (SCHED_DAILY, SCHED_WORKDAYS, SCHED_WEEKENDS,
                                 SCHED_WEEKLY, SCHED_MONTHLY, SCHED_ONCE)
        has_date     = stype == SCHED_ONCE
        has_days     = stype == SCHED_WEEKLY
        has_dom      = stype == SCHED_MONTHLY
        has_interval = stype == SCHED_INTERVAL
        self._sched_time_row.setVisible(enabled and has_time)
        self._sched_date_row.setVisible(enabled and has_date)
        self._sched_days_row.setVisible(enabled and has_days)
        self._sched_dom_row.setVisible(enabled and has_dom)
        self._sched_interval_row.setVisible(enabled and has_interval)
        self._sched_note.setVisible(enabled)

    def _get_schedule_dict(self) -> dict:
        """Build and validate the schedule dict from wizard fields."""
        if not self._sched_enabled.isChecked():
            return {"enabled": False}
        stype = self._sched_type_combo.currentData() or SCHED_DAILY
        sched: dict = {"enabled": True, "type": stype}
        if stype != SCHED_INTERVAL:
            raw_time = self._sched_time_edit.text().strip()
            # validate HH:MM
            try:
                h, m = (int(x) for x in raw_time.split(":"))
                if not (0 <= h <= 23 and 0 <= m <= 59):
                    raise ValueError
                sched["time"] = f"{h:02d}:{m:02d}"
            except (ValueError, TypeError):
                sched["time"] = "09:00"
        if stype == SCHED_ONCE:
            raw_date = self._sched_date_edit.text().strip()
            try:
                y, mo, d = (int(x) for x in raw_date.split("-"))
                datetime(y, mo, d)   # validate
                sched["date"] = f"{y:04d}-{mo:02d}-{d:02d}"
            except (ValueError, TypeError):
                sched["date"] = datetime.now().strftime("%Y-%m-%d")
        elif stype == SCHED_WEEKLY:
            sched["days"] = [i for i, cb in enumerate(self._sched_day_cbs) if cb.isChecked()]
        elif stype == SCHED_MONTHLY:
            raw_dom = self._sched_dom_edit.text().strip()
            try:
                dom = int(raw_dom)
                sched["day_of_month"] = max(1, min(31, dom))
            except (ValueError, TypeError):
                sched["day_of_month"] = 1
        elif stype == SCHED_INTERVAL:
            raw_val = self._sched_interval_val.text().strip()
            try:
                val = int(raw_val)
                sched["interval_value"] = max(1, val)
            except (ValueError, TypeError):
                sched["interval_value"] = 30
            sched["interval_unit"] = self._sched_interval_unit.currentData() or "minutes"
        # Preserve last_run if editing an existing node
        if self._existing:
            old_sched = self._existing.get("schedule") or {}
            if old_sched.get("last_run"):
                sched["last_run"] = old_sched["last_run"]
        return sched

    def _populate_schedule(self, node: dict):
        """Restore schedule fields from an existing node."""
        sched = node.get("schedule") or {}
        self._sched_enabled.setChecked(bool(sched.get("enabled", False)))
        stype = sched.get("type", SCHED_DAILY)
        for i in range(self._sched_type_combo.count()):
            if self._sched_type_combo.itemData(i) == stype:
                self._sched_type_combo.setCurrentIndex(i); break
        self._sched_time_edit.setText(sched.get("time", "09:00"))
        self._sched_date_edit.setText(sched.get("date", ""))
        saved_days = sched.get("days", list(range(5)))
        for i, cb in enumerate(self._sched_day_cbs):
            cb.setChecked(i in saved_days)
        dom = sched.get("day_of_month", 1)
        self._sched_dom_edit.setText(str(dom))
        self._sched_interval_val.setText(str(sched.get("interval_value", 30)))
        unit = sched.get("interval_unit", "minutes")
        for i in range(self._sched_interval_unit.count()):
            if self._sched_interval_unit.itemData(i) == unit:
                self._sched_interval_unit.setCurrentIndex(i); break
        self._refresh_sched_ui()

    def _le(self, ph):
        t = _theme
        le = QLineEdit(); le.setPlaceholderText(ph); le.setFont(FONT_LABEL)
        le.setStyleSheet(f"""
            QLineEdit {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:5px; color:{t.TEXT_PRIMARY.name()};
                padding:5px 9px; font-family:'Segoe UI'; font-size:9pt;
            }}
            QLineEdit:focus {{ border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},160); }}
        """)
        return le

    def _style_te(self, te):
        t = _theme
        te.setFont(FONT_LABEL)
        te.setStyleSheet(f"""
            QTextEdit {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:5px; color:{t.TEXT_PRIMARY.name()};
                padding:5px 9px; font-family:'Segoe UI'; font-size:9pt;
            }}
            QTextEdit:focus {{ border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},160); }}
        """)

    def _style_rb(self, rb):
        t = _theme
        rb.setFont(FONT_LABEL)
        rb.setStyleSheet(f"""
            QRadioButton {{
                color:{t.TEXT_SECONDARY.name()}; spacing:5px;
                font-family:'Segoe UI'; font-size:9pt;
            }}
            QRadioButton::indicator {{
                width:13px; height:13px; border-radius:7px;
                border:1.5px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},110);
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
            }}
            QRadioButton::indicator:checked {{
                background:{t.GLOW.name()}; border-color:{t.GLOW.name()};
            }}
            QRadioButton:checked {{ color:{t.TEXT_PRIMARY.name()}; }}
        """)

    def _style_cb(self, cb):
        t = _theme
        cb.setFont(FONT_LABEL)
        cb.setStyleSheet(f"""
            QCheckBox {{
                color:{t.TEXT_SECONDARY.name()}; spacing:7px;
                font-family:'Segoe UI'; font-size:9pt;
            }}
            QCheckBox::indicator {{
                width:13px; height:13px; border-radius:3px;
                border:1.5px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},110);
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
            }}
            QCheckBox::indicator:checked {{
                background:{t.GLOW.name()}; border-color:{t.GLOW.name()};
            }}
        """)

    def _style_btn(self, btn, secondary=False, small=False):
        t = _theme
        btn.setCursor(Qt.PointingHandCursor)
        btn.setFont(QFont("Segoe UI",9,QFont.Weight.DemiBold))
        if small: btn.setFixedHeight(28)
        if secondary:
            btn.setStyleSheet(f"""
                QPushButton {{
                    background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},150);
                    border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                    border-radius:6px; color:{t.TEXT_SECONDARY.name()};
                    padding:6px 16px;
                }}
                QPushButton:hover {{
                    background:rgba({t.TILE_BG_HOVER.red()},{t.TILE_BG_HOVER.green()},{t.TILE_BG_HOVER.blue()},180);
                    color:{t.TEXT_PRIMARY.name()};
                    border-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},110);
                }}
            """)
        else:
            g = t.GLOW
            dr, dg, db = int(g.red()*0.65), int(g.green()*0.65), int(g.blue()*0.65)
            btn.setStyleSheet(f"""
                QPushButton {{
                    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                        stop:0 rgba({g.red()},{g.green()},{g.blue()},195),
                        stop:1 rgba({dr},{dg},{db},195));
                    border:none; border-radius:6px;
                    color:{t.TEXT_PRIMARY.name()}; padding:6px 16px;
                }}
                QPushButton:hover {{
                    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                        stop:0 rgba({g.red()},{g.green()},{g.blue()},240),
                        stop:1 rgba({dr},{dg},{db},240));
                }}
                QPushButton:pressed {{
                    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                        stop:0 rgba({int(g.red()*0.8)},{int(g.green()*0.8)},{int(g.blue()*0.8)},220),
                        stop:1 rgba({int(g.red()*0.5)},{int(g.green()*0.5)},{int(g.blue()*0.5)},220));
                }}
            """)

    def _on_type_changed(self, btn):
        nt = btn.property("type_key"); is_note = nt == NODE_TYPE_NOTE; is_file = nt == NODE_TYPE_FILE
        self._target_lbl.setVisible(not is_note)
        self._target_edit.setVisible(not is_note)
        self._note_lbl.setVisible(is_note)
        self._note_combo.setVisible(is_note)
        self._new_note_btn.setVisible(is_note)
        self._behavior_lbl.setVisible(is_file)
        self._behavior_combo.setVisible(is_file)
        if is_note:
            self._refresh_note_combo()

    def _pick_accent_color(self):
        """Open color dialog and store the chosen hex color."""
        initial = QColor(self._accent_color_val) if self._accent_color_val else _theme.GLOW
        dlg = QColorDialog(initial, self)
        t = _theme
        dlg.setStyleSheet(f"background:{t.BG_MID.name()}; color:{t.TEXT_PRIMARY.name()};")
        if dlg.exec() == QDialog.Accepted:
            c = dlg.currentColor()
            self._accent_color_val = c.name()
            self._accent_btn.setText("")
            self._accent_btn.setStyleSheet(
                f"background:{c.name()}; border:none; border-radius:5px; "
                f"color:white; padding:0 10px; font-size:9pt;")

    def _clear_accent_color(self):
        """Remove the per-node accent color, reverting to the theme default."""
        self._accent_color_val = ""
        self._accent_btn.setText("Default  (use theme color)")
        self._style_btn(self._accent_btn, secondary=True, small=True)

    def _refresh_note_combo(self):
        """Reload notebook notes into the combo box."""
        current_id = self._note_combo.currentData()
        self._note_combo.clear()
        notes = NotebookStore.list_notes()
        for n in notes:
            self._note_combo.addItem(n["title"] or "Untitled", n["id"])
        if not notes:
            self._note_combo.addItem("(no notes — click + Create new note)", "")
        # Restore previous selection
        if current_id:
            for i in range(self._note_combo.count()):
                if self._note_combo.itemData(i) == current_id:
                    self._note_combo.setCurrentIndex(i); break

    def _create_new_linked_note(self):
        """Create a new empty notebook note and select it in the combo."""
        note_id = NotebookStore.new_id()
        title = self._name_edit.text().strip() or "Untitled"
        NotebookStore.save_note(note_id, title, "")
        self._refresh_note_combo()
        for i in range(self._note_combo.count()):
            if self._note_combo.itemData(i) == note_id:
                self._note_combo.setCurrentIndex(i); break

    def _browse_target(self):
        if self._current_type() == NODE_TYPE_URL: return
        path, _ = QFileDialog.getOpenFileName(self, "Select File")
        if path: self._target_edit.setText(path)

    def _import_node_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Import Node File", "",
            f"Node Files (*{NODE_FILE_EXT});;All Files (*)")
        if not path: return
        try:
            node = NodeStore.import_node(path); node.pop("id", None)
            self._populate(node)
        except Exception as exc:
            QMessageBox.warning(self, "Import Error", str(exc))

    def _current_type(self):
        for rb in self._type_group.buttons():
            if rb.isChecked(): return rb.property("type_key")
        return NODE_TYPE_FILE

    def _current_size(self):
        for rb in self._size_group.buttons():
            if rb.isChecked(): return rb.property("size_key")
        return "1x1"

    def _on_save(self):
        name = self._name_edit.text().strip()
        if not name: self._name_edit.setFocus(); return
        nt = self._current_type()
        if nt == NODE_TYPE_NOTE:
            note_id = self._note_combo.currentData() or ""
            if not note_id:
                self._new_note_btn.setFocus(); return
            target = note_id
        else:
            target = self._target_edit.text().strip()
        node = {
            "name": name, "type": nt, "size": self._current_size(),
            "target": target,
            "note": "",
            "params": self._params_edit.text().strip(),
            "description": self._desc_edit.toPlainText().strip(),
            "icon": self._icon_zone.current_path(),
            "auto_launch": self._auto_launch.isChecked(),
            "archived": False,
        }
        if self._existing: node["id"] = self._existing["id"]
        # ── New fields ──────────────────────────────────────────────────
        raw_tags = self._tags_edit.text().strip()
        node["tags"] = [t.strip() for t in raw_tags.split(",") if t.strip()] if raw_tags else []
        node["accent_color"] = self._accent_color_val
        node["open_behavior"] = (
            self._behavior_combo.currentData()
            if self._current_type() == NODE_TYPE_FILE else "normal")
        node["schedule"] = self._get_schedule_dict()
        self.node_saved.emit(node); self.accept()

    def _populate(self, node: dict):
        self._name_edit.setText(node.get("name",""))
        self._desc_edit.setPlainText(node.get("description",""))
        self._params_edit.setText(node.get("params",""))
        self._auto_launch.setChecked(node.get("auto_launch",False))
        self._populate_schedule(node)
        if node.get("icon"): self._icon_zone.set_path(node["icon"])
        for rb in self._size_group.buttons():
            if rb.property("size_key") == node.get("size","1x1"):
                rb.setChecked(True); break
        nt = node.get("type", NODE_TYPE_FILE)
        for rb in self._type_group.buttons():
            if rb.property("type_key") == nt:
                rb.setChecked(True); self._on_type_changed(rb); break
        if nt == NODE_TYPE_NOTE:
            self._refresh_note_combo()
            note_id = node.get("target", "")
            for i in range(self._note_combo.count()):
                if self._note_combo.itemData(i) == note_id:
                    self._note_combo.setCurrentIndex(i); break
        else:
            self._target_edit.setText(node.get("target", ""))
        # Restore open behavior
        ob = node.get("open_behavior", "normal")
        for i in range(self._behavior_combo.count()):
            if self._behavior_combo.itemData(i) == ob:
                self._behavior_combo.setCurrentIndex(i); break
        # Restore tags
        tags = node.get("tags", [])
        self._tags_edit.setText(", ".join(tags) if isinstance(tags, list) else str(tags))
        # Restore accent color
        ac = node.get("accent_color", "")
        self._accent_color_val = ac
        if ac and QColor(ac).isValid():
            self._accent_btn.setText("")
            self._accent_btn.setStyleSheet(
                f"background:{ac}; border:none; border-radius:5px; "
                f"color:white; padding:0 10px; font-size:9pt;")
        else:
            self._accent_color_val = ""
            self._accent_btn.setText("Default  (use theme color)")
            self._style_btn(self._accent_btn, secondary=True, small=True)

    def paintEvent(self, e: QPaintEvent):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        rect = QRectF(10,10,self.width()-20,self.height()-20)
        path = QPainterPath(); path.addRoundedRect(rect, BORDER_RADIUS+2, BORDER_RADIUS+2)
        grad = QLinearGradient(0,0,self.width(),self.height())
        grad.setColorAt(0, QColor(t.BG_MID.red()+4, t.BG_MID.green()+4, t.BG_MID.blue()+6, 248))
        grad.setColorAt(1, QColor(t.BG_DARK.red(), t.BG_DARK.green(), t.BG_DARK.blue(), 248))
        p.fillPath(path, grad)
        border = QColor(t.GLOW); border.setAlpha(100)
        p.setPen(QPen(border,1.3)); p.setBrush(Qt.NoBrush); p.drawPath(path)
        p.end()


# ---------------------------------------------------------------------------
# Quick Connect
# ---------------------------------------------------------------------------

_QC_CLIENTS: list[str] = sorted([
    "ABCOLUMBUS", "ALLAMERICA", "ANBSHAWNEE", "ABWAGONER", "ABTTULSA",
    "AEBHENRYETTA", "ACBBLUESPRINGS", "ADBTEXHOMA", "ABCOGCLEVELAND",
    "ABSPRINGFIELD", "AVBBROKENARROW", "BCNAALVA", "BANCPARTNER",
    "BANK7", "BIIRVINE", "BOEEUFAULA", "BOHHYDRO", "BOLLUXEMBURG",
    "BOORRICK", "BOPVPRAIRIEVILLAGE", "BOFHWAMEGO", "BOWTHOMAS",
    "BANKOFVICI", "B19MEMPHIS", "B360CORDELL", "BCCBIRVINE",
    "BBTCHICAGO", "BOOKC", "BCBLITTLECHUTE", "BBBYRON", "CBALTUS",
    "CPOLOKC", "CBPARKER", "CBJOHNSTON", "CSBCHILLICOTHE", "CBAMARILLO",
    "CSBCLEOSPRINGS", "CBTLAJUNTA", "CNBTEXARKANA", "CBTOPEKA", "CNBALVA",
    "CBMRICHMOND", "CBOELWEIN", "CBPPARKERSBURG", "CSBCANTON", "CSBCASHION",
    "CSBCOFFEYVILLE", "CSBGALVA", "CSBROYALCENTER", "CWBFRESNO",
    "CBFORTMADISON", "CBPLATTECITY", "CBWAYNOKA", "CNBHILLSDALE",
    "CFCUEDGEWOOD", "CBKREMLIN", "CUOOOKC", "DCBDALLAS", "DNBCLEARLAKE",
    "DFBSYRACUSE", "ERBROCHESTER", "EBENCINO", "EBSEDAN", "ECUOKC",
    "ESBFEMPORIA", "ENBEVERGREEN", "FMPIEDMONT", "FNBPTOWN", "FSBWEVER",
    "FSBCAMERON", "FSBALLEN", "FFBFRESNO", "FABSTONEWALL", "FAMYUKON",
    "FBTPERRY", "FBBELOIT", "FIRSTBETHANY", "FCBNEWELL", "FFSLADELTA",
    "FFSLSANRAFAEL", "FFSBCHAMPAIGN", "FIRSTLIBERTY", "FNBMCALESTER",
    "FBTWEATHERFORD", "FNBFRANKFORT", "FNBMARLOW", "FNBETGILMER",
    "FNBHUGO", "FNBMKALAMAZOO", "FNBOK", "FNBVINITA", "FPBPRYOR",
    "FSBBELMOND", "FSBNOBLE", "FSBTAHLEQUAH", "FSBWATONGA", "FSBDECATUR",
    "FSBSHANNON", "FBBRICHMOND", "FSBFOWLER", "FBALTUS", "FBROCKRAPIDS",
    "FSBOKC", "GBOAKLAND", "GBLASVEGAS", "GHSFCUBINGHAMTON", "GNBGILMER",
    "GSBGLENWOOD", "GCBMEDFORD", "GCBULYSSES", "GREATNATIONS",
    "GPTCOVERLANDPARK", "GWBVANWERT", "GBMORRIS", "GBWILLIAMSON",
    "GSLAGUNNISON", "GCSBGUTHRIE", "HCBSALIDA", "HPBKEYES", "HSLANORBORNE",
    "HSBJEFFERSON", "HSBROYAL", "HPBMANSFIELD", "HCBEASTTAWAS",
    "IDBGFCUWASHINGTONDC", "IBRATON", "IBTMONUMENT", "ISBCARROLL",
    "JBSANANTONIO", "JOHNSONSTATE", "KCBLAKIN", "KTCKEARNEY",
    "KSBWISCONSINRAPIDS", "LSBOOLOGAH", "LCSBHOBBS", "LSBPELLA",
    "LBLOSANGELES", "LNBLAWTON", "LSBLISLE", "LTKINGFISHER", "MBMALVERN",
    "MSBMCCLAVE", "MCCORMACK", "MBWYANDOTTE", "MBBLOOMFIELDHILLS",
    "MABBALDWIN", "MBDESPLAINES", "MBFRESNO", "NBIRVINE", "NNBSANDIEGO",
    "NBNEWBURYPORT", "NVBMARYVILLE", "NCBHENNEPIN", "NCHOKC", "OBA",
    "OCBTULSA", "OKFCU", "OHBROFF", "OSBBUFFALO", "OBCOLBY", "OBOXFORD",
    "PVBSALINAS", "PCBPERKINS", "PBPENTUCKET", "PFBJOLIET", "PNBCHECOTAH",
    "PSBINDIANOLA", "PNBPERRYTON", "PPNBCOSPR", "PFSLADEERLODGE",
    "PWCBWINDSOR", "PCBMARENGO", "PBROTHVILLE", "PBGUTHRIE", "POBSALEM",
    "RIAFCUBETTENDORF", "RRFCUALTUS", "RGBMONTEVISTA", "RFSBRIVERFALLS",
    "RRBNEWBURYPORT", "REBMARION", "RBTFITCHBURG", "RSBRUSHVILLE",
    "SLVFBALAMOSA", "SBSPRINGFIELD", "SECURITYTULSA", "SSBCHEYENNE",
    "SSBTFREDERICKSBURG", "SSBWEWOKA", "SNBSHATTUCK", "SNBSTOCKTON",
    "SSB", "SWHBSCOTTSDALE", "SNBWichita", "SOFCULAWTON", "SBSHAWNEE",
    "SSBSPIRO", "SVCBSPRINGVALLEY", "SLBTAC", "SBBISMARCK", "SBOCCHERRY",
    "SBMCONCORDIA", "SBREESEVILLE", "SBCOLORADOSPRINGS", "SBPARKCITY",
    "SBBLOOMFIELD", "SASBCOURTLAND", "BALDWINSTBANK", "BANKADVANCE",
    "TBBBENNINGTON", "CSBCCHENEY", "TCNBCOLORADOCITY", "FSBMCPHERSON",
    "FNBELKCITY", "FNBBROKENARROW", "FNBMIAMI", "FNBTREMONT", "FNBAVA",
    "FSBABERNATHY", "FSBBOISE", "FSBPONDCRK", "GBTGUNNISON", "THSBHOPETON",
    "INBIDABEL", "STATEBK", "SEBWOODWARD", "UBCCOLUMBUSGROVE", "TCUBCAMDEN",
    "TCTGLASGOW", "TCBTROACHDALE", "TRUSTOK", "USBGREENFIELD", "USBHAZEN",
    "UBTMARYSVILLE", "UBATMORE", "USBFRESNO", "VBOKC", "VBEDMOND",
    "WCBWALPOLE", "WBOKC", "WIBWESTBEND", "WSBWRAY", "WBTCHEYENNE",
    "ZBTSCOTTSDALE",
], key=str.upper)


class _AutoCompleteEdit(QLineEdit):
    """QLineEdit with inline ghost-text typeahead AND a live-filter dropdown popup.

    - Typing instantly filters a popup list to matching client IDs.
    - Up / Down navigate the popup; Tab or Enter accepts the highlighted item.
    - Ghost suffix (greyed-out text) previews the first match inline; Tab / → accepts it.
    - Escape closes the popup without selecting; clicking outside also closes it.
    """

    def __init__(self, items: list, parent=None):
        super().__init__(parent)
        self._items        = items
        self._ghost        = ""
        self._ghost_full   = ""
        self._popup: Optional[QFrame]       = None
        self._popup_list: Optional[QListWidget] = None
        self.textChanged.connect(self._on_text_changed)

    # ── public API ────────────────────────────────────────────────────────

    def show_all_popup(self):
        """Show the full unfiltered list (called by the ▾ browse button)."""
        self._ensure_popup()
        self._populate_popup(self._items)
        self._reposition_popup()
        self._popup.show()
        self.setFocus()

    # ── ghost suggestion ──────────────────────────────────────────────────

    def _on_text_changed(self, text: str):
        upper = text.upper()
        if text != upper:
            pos = self.cursorPosition()
            self.blockSignals(True)
            self.setText(upper)
            self.setCursorPosition(pos)
            self.blockSignals(False)
            text = upper
        self._refresh_ghost(text)
        self._refresh_popup(text)

    def _refresh_ghost(self, text: str):
        if text:
            tu = text.upper()
            match = next((c for c in self._items if c.upper().startswith(tu)), None)
            if match and len(match) > len(text):
                self._ghost      = match[len(text):]
                self._ghost_full = match
            else:
                self._ghost = self._ghost_full = ""
        else:
            self._ghost = self._ghost_full = ""
        self.update()

    def _accept_ghost(self):
        if not self._ghost:
            return
        full = self._ghost_full or (self.text() + self._ghost)
        self.blockSignals(True)
        self.setText(full)
        self.blockSignals(False)
        self._ghost = self._ghost_full = ""
        self.end(False)
        self._hide_popup()
        self.update()

    # ── popup lifecycle ───────────────────────────────────────────────────

    def _ensure_popup(self):
        if self._popup is not None:
            return
        t = _theme
        self._popup = QFrame(None)
        self._popup.setWindowFlags(
            Qt.FramelessWindowHint | Qt.Tool | Qt.WindowDoesNotAcceptFocus
        )
        self._popup.setAttribute(Qt.WA_ShowWithoutActivating)
        self._popup.setAttribute(Qt.WA_QuitOnClose, False)
        self._popup.setStyleSheet(
            f"""QFrame {{
                background:rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},252);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},150);
                border-radius:6px;
            }}"""
        )
        lay = QVBoxLayout(self._popup)
        lay.setContentsMargins(4, 4, 4, 4)
        lay.setSpacing(0)

        self._popup_list = QListWidget()
        self._popup_list.setFont(QFont("Consolas", 9))
        self._popup_list.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._popup_list.setFrameShape(QFrame.NoFrame)
        self._popup_list.setStyleSheet(
            f"""QListWidget {{
                background:transparent; border:none; outline:none;
                color:{t.TEXT_PRIMARY.name()};
            }}
            QListWidget::item {{ padding:4px 14px; border-radius:3px; }}
            QListWidget::item:hover {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},55);
            }}
            QListWidget::item:selected {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},105);
                color:{t.TEXT_PRIMARY.name()};
            }}
            QScrollBar:vertical {{
                background:rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},120);
                width:6px; border-radius:3px; margin:0;
            }}
            QScrollBar::handle:vertical {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},130);
                border-radius:3px; min-height:20px;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height:0; }}"""
        )
        self._popup_list.itemClicked.connect(self._accept_popup_item)
        lay.addWidget(self._popup_list)

    def _populate_popup(self, items: list):
        lst = self._popup_list
        lst.blockSignals(True)
        lst.clear()
        lst.addItems(items)
        lst.blockSignals(False)
        if items:
            lst.setCurrentRow(0)
        item_h = lst.sizeHintForRow(0) if lst.count() > 0 else 24
        lst.setFixedHeight(min(len(items), 14) * item_h + 2)
        self._popup.adjustSize()

    def _refresh_popup(self, text: str):
        if not text:
            self._hide_popup()
            return
        tu = text.upper()
        matches = [c for c in self._items if c.upper().startswith(tu)]
        if not matches:
            self._hide_popup()
            return
        self._ensure_popup()
        self._populate_popup(matches)
        self._reposition_popup()
        self._popup.show()

    def _reposition_popup(self):
        if not self._popup:
            return
        # Match width of full input row (edit + ▾ button) when parent dialog exposes it
        popup_w = self.width()
        parent_dlg = self.window()
        if parent_dlg and hasattr(parent_dlg, '_dropdown_btn'):
            popup_w = self.width() + parent_dlg._dropdown_btn.width() + 4
        self._popup.setFixedWidth(popup_w)
        popup_h = self._popup.sizeHint().height()

        gp    = self.mapToGlobal(QPoint(0, self.height() + 2))
        scr_o = QGuiApplication.screenAt(gp)
        scr   = scr_o.availableGeometry() if scr_o else QGuiApplication.primaryScreen().availableGeometry()

        x, y = gp.x(), gp.y()
        if y + popup_h > scr.bottom():          # flip above if off-screen
            y = self.mapToGlobal(QPoint(0, -popup_h - 2)).y()
        x = max(scr.left() + 4, min(x, scr.right() - popup_w - 4))
        self._popup.move(x, y)

    def _hide_popup(self):
        if self._popup and self._popup.isVisible():
            self._popup.hide()

    def _accept_popup_item(self, item):
        self.blockSignals(True)
        self.setText(item.text())
        self.blockSignals(False)
        self._ghost = self._ghost_full = ""
        self._hide_popup()
        self.end(False)
        self.setFocus()
        self.update()

    # ── event handling ────────────────────────────────────────────────────

    def event(self, ev):
        """Intercept Tab before Qt's focus-chain so it accepts the suggestion."""
        from PySide6.QtCore import QEvent
        if ev.type() == QEvent.Type.KeyPress and ev.key() == Qt.Key_Tab:
            if self._popup and self._popup.isVisible():
                item = self._popup_list.currentItem()
                if item:
                    self._accept_popup_item(item)
                    return True
            if self._ghost:
                self._accept_ghost()
                return True
        return super().event(ev)

    def keyPressEvent(self, event):
        key = event.key()
        popup_open = bool(self._popup and self._popup.isVisible())

        if popup_open:
            if key == Qt.Key_Down:
                r = self._popup_list.currentRow()
                self._popup_list.setCurrentRow(min(r + 1, self._popup_list.count() - 1))
                event.accept(); return
            if key == Qt.Key_Up:
                r = self._popup_list.currentRow()
                if r <= 0:
                    self._hide_popup()
                else:
                    self._popup_list.setCurrentRow(r - 1)
                event.accept(); return
            if key in (Qt.Key_Return, Qt.Key_Enter):
                item = self._popup_list.currentItem()
                if item:
                    self._accept_popup_item(item)
                    event.accept(); return
            if key == Qt.Key_Escape:
                self._hide_popup()
                event.accept(); return

        # Right-arrow at end of typed text → accept inline ghost
        if (
            key == Qt.Key_Right
            and self.cursorPosition() == len(self.text())
            and self._ghost
        ):
            self._accept_ghost()
            event.accept()
            return

        super().keyPressEvent(event)

    def hideEvent(self, ev):
        super().hideEvent(ev)
        self._hide_popup()

    # ── ghost-text painting ───────────────────────────────────────────────

    def paintEvent(self, event):
        super().paintEvent(event)
        if not self._ghost:
            return

        # Ask Qt's style engine for the ACTUAL inner text rectangle.
        # This correctly accounts for the stylesheet padding (0 10px), border, and
        # any platform-specific insets — unlike contentsRect() which ignores padding.
        from PySide6.QtWidgets import QStyleOptionFrame, QStyle
        opt = QStyleOptionFrame()
        self.initStyleOption(opt)
        inner = self.style().subElementRect(
            QStyle.SubElement.SE_LineEditContents, opt, self
        )
        # Measure how wide the typed text is and add a 4 px visual gap so the
        # ghost is clearly separated from both the last character and the cursor bar.
        x = inner.left() + self.fontMetrics().horizontalAdvance(self.text()) + 4

        p = QPainter(self)
        p.setFont(self.font())
        t = _theme
        p.setPen(QColor(t.TEXT_DIM.red(), t.TEXT_DIM.green(), t.TEXT_DIM.blue(), 155))
        r = QRect(x, inner.top(), max(0, inner.right() - x), inner.height())
        p.drawText(r, Qt.AlignVCenter | Qt.AlignLeft, self._ghost)
        p.end()


class QuickConnectDialog(QDialog):
    """Frameless dialog for launching a remote-access session.

    Builds a URL of the form:
        https://remote.bankonitusa.com/Host#Access/<ClientID>[/<DeviceName>]
    and opens it in the system default browser.
    """

    _BASE_URL = "https://remote.bankonitusa.com/Host#Access/"

    def __init__(self, parent=None):
        super().__init__(parent, Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self._drag_pos: Optional[QPoint] = None
        self._build_ui()
        self.setFixedSize(460, 302)

    # ── drag support ──────────────────────────────────────────────────────

    def _drag_move(self, ev: QMouseEvent):
        if self._drag_pos and (ev.buttons() & Qt.LeftButton):
            delta = ev.globalPosition().toPoint() - self._drag_pos
            self.move(self.pos() + delta)
            self._drag_pos = ev.globalPosition().toPoint()

    # ── UI construction ───────────────────────────────────────────────────

    def _build_ui(self):
        t = _theme
        outer = QVBoxLayout(self)
        outer.setContentsMargins(10, 10, 10, 10)
        outer.setSpacing(0)

        self._card = QWidget()
        self._card.setGraphicsEffect(make_shadow(self._card, 32, QColor(0, 0, 0, 210)))
        cl = QVBoxLayout(self._card)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(0)

        # ── title bar ─────────────────────────────────────────────────────
        tbar = QWidget()
        tbar.setFixedHeight(46)
        tbar.setStyleSheet("background:transparent;")
        tbl = QHBoxLayout(tbar)
        tbl.setContentsMargins(16, 0, 12, 0)
        tbl.setSpacing(0)

        bolt = QLabel("⚡")
        bolt.setFont(QFont("Segoe UI", 12))
        bolt.setStyleSheet(f"color:{t.GLOW.name()}; background:transparent;")
        tbl.addWidget(bolt)
        tbl.addSpacing(7)

        ttl = QLabel("Quick Connect")
        ttl.setFont(QFont("Segoe UI", 11, QFont.Weight.DemiBold))
        ttl.setStyleSheet(f"color:{t.TEXT_PRIMARY.name()}; background:transparent;")
        tbl.addWidget(ttl)
        tbl.addStretch()

        close_btn = TitleBarButton(COLOR_BTN_CLOSE, "x")
        close_btn.clicked.connect(self.reject)
        tbl.addWidget(close_btn)

        tbar.mousePressEvent = lambda ev: (
            setattr(self, "_drag_pos", ev.globalPosition().toPoint())
            if ev.button() == Qt.LeftButton else None
        )
        tbar.mouseMoveEvent = self._drag_move
        tbar.mouseReleaseEvent = lambda ev: setattr(self, "_drag_pos", None)
        cl.addWidget(tbar)

        # ── separator ─────────────────────────────────────────────────────
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(
            f"color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);"
        )
        cl.addWidget(sep)

        # ── form ──────────────────────────────────────────────────────────
        form_w = QWidget()
        form_w.setStyleSheet("background:transparent;")
        fl = QVBoxLayout(form_w)
        fl.setContentsMargins(22, 16, 22, 8)
        fl.setSpacing(7)

        # CLIENT ID
        cid_lbl = QLabel(
            f"<span style='font-family:Segoe UI;font-size:8.5pt;"
            f"font-weight:600;letter-spacing:1.2px;"
            f"color:{t.TEXT_SECONDARY.name()};'>CLIENT ID</span>"
            f"<span style='font-size:7.5pt;color:{t.ACCENT_RED.name()};'>"
            f"  required</span>"
        )
        fl.addWidget(cid_lbl)

        cid_row = QHBoxLayout()
        cid_row.setSpacing(4)

        self._client_edit = _AutoCompleteEdit(_QC_CLIENTS)
        self._client_edit.setPlaceholderText(
            "Type or select a Client ID…  (Tab to complete)"
        )
        self._client_edit.setFont(QFont("Segoe UI", 9))
        self._client_edit.setFixedHeight(32)
        self._style_input(self._client_edit)
        self._client_edit.returnPressed.connect(self._do_connect)
        cid_row.addWidget(self._client_edit, 1)

        self._dropdown_btn = QPushButton("▾")
        self._dropdown_btn.setFixedSize(32, 32)
        self._dropdown_btn.setCursor(Qt.PointingHandCursor)
        self._dropdown_btn.setFont(QFont("Segoe UI", 11))
        self._dropdown_btn.setToolTip("Browse all cached client IDs")
        self._dropdown_btn.setStyleSheet(
            f"""QPushButton {{
                    background:rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},200);
                    border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},80);
                    border-radius:7px;
                    color:{t.TEXT_SECONDARY.name()};
                }}
                QPushButton:hover {{
                    background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);
                    color:{t.GLOW.name()};
                    border-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},160);
                }}
                QPushButton:pressed {{
                    background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},80);
                }}"""
        )
        self._dropdown_btn.clicked.connect(self._show_client_menu)
        cid_row.addWidget(self._dropdown_btn)

        fl.addLayout(cid_row)
        fl.addSpacing(4)

        # DEVICE NAME
        dev_lbl = QLabel(
            f"<span style='font-family:Segoe UI;font-size:8.5pt;"
            f"font-weight:600;letter-spacing:1.2px;"
            f"color:{t.TEXT_SECONDARY.name()};'>DEVICE NAME</span>"
            f"<span style='font-size:7.5pt;color:{t.TEXT_DIM.name()};'>"
            f"  optional</span>"
        )
        fl.addWidget(dev_lbl)

        self._device_edit = QLineEdit()
        self._device_edit.setPlaceholderText("e.g. FrontDesk01")
        self._device_edit.setFont(QFont("Segoe UI", 9))
        self._device_edit.setFixedHeight(32)
        self._style_input(self._device_edit)
        self._device_edit.returnPressed.connect(self._do_connect)
        fl.addWidget(self._device_edit)

        cl.addWidget(form_w)
        cl.addStretch()

        # ── button row ────────────────────────────────────────────────────
        btn_w = QWidget()
        btn_w.setStyleSheet("background:transparent;")
        bl = QHBoxLayout(btn_w)
        bl.setContentsMargins(22, 2, 22, 16)
        bl.setSpacing(8)

        self._err_lbl = QLabel("")
        self._err_lbl.setFont(QFont("Segoe UI", 7, QFont.Weight.Normal))
        self._err_lbl.setStyleSheet(
            f"color:{t.ACCENT_RED.name()}; background:transparent;"
        )
        bl.addWidget(self._err_lbl, 1)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setFixedHeight(32)
        cancel_btn.setFont(QFont("Segoe UI", 9))
        cancel_btn.setCursor(Qt.PointingHandCursor)
        cancel_btn.setStyleSheet(
            f"""QPushButton {{
                    background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},180);
                    border:1px solid rgba({t.TEXT_DIM.red()},{t.TEXT_DIM.green()},{t.TEXT_DIM.blue()},80);
                    border-radius:7px;
                    color:{t.TEXT_SECONDARY.name()};
                    padding:0 16px;
                }}
                QPushButton:hover {{
                    background:rgba({t.TILE_BG_HOVER.red()},{t.TILE_BG_HOVER.green()},{t.TILE_BG_HOVER.blue()},220);
                    color:{t.TEXT_PRIMARY.name()};
                    border-color:rgba({t.TEXT_SECONDARY.red()},{t.TEXT_SECONDARY.green()},{t.TEXT_SECONDARY.blue()},120);
                }}"""
        )
        cancel_btn.clicked.connect(self.reject)
        bl.addWidget(cancel_btn)

        connect_btn = QPushButton("⚡  Connect")
        connect_btn.setFixedHeight(32)
        connect_btn.setFont(QFont("Segoe UI", 9, QFont.Weight.DemiBold))
        connect_btn.setCursor(Qt.PointingHandCursor)
        connect_btn.setDefault(True)
        connect_btn.setStyleSheet(
            f"""QPushButton {{
                    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                        stop:0 rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},210),
                        stop:1 rgba({t.ACCENT_TEAL.red()},{t.ACCENT_TEAL.green()},{t.ACCENT_TEAL.blue()},190));
                    border:none;
                    border-radius:7px;
                    color:white;
                    padding:0 20px;
                }}
                QPushButton:hover {{
                    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                        stop:0 rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},255),
                        stop:1 rgba({t.ACCENT_TEAL.red()},{t.ACCENT_TEAL.green()},{t.ACCENT_TEAL.blue()},240));
                }}
                QPushButton:pressed {{
                    background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},160);
                }}"""
        )
        connect_btn.clicked.connect(self._do_connect)
        bl.addWidget(connect_btn)

        cl.addWidget(btn_w)
        outer.addWidget(self._card)

    # ── helpers ───────────────────────────────────────────────────────────

    def _style_input(self, edit: QLineEdit):
        t = _theme
        edit.setStyleSheet(
            f"""QLineEdit {{
                    background:rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},200);
                    border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                    border-radius:7px;
                    color:{t.TEXT_PRIMARY.name()};
                    padding:0 10px;
                    selection-background-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},100);
                }}
                QLineEdit:focus {{
                    border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},180);
                    background:rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},200);
                }}"""
        )

    def _show_client_menu(self):
        """Browse all client IDs via the live-filter popup (shows full list)."""
        self._client_edit.show_all_popup()

    # ── connect action ────────────────────────────────────────────────────

    def _do_connect(self):
        client = self._client_edit.text().strip()
        device  = self._device_edit.text().strip()
        if not client:
            self._err_lbl.setText("⚠  Client ID is required")
            self._client_edit.setFocus()
            return
        self._err_lbl.setText("")
        url = self._BASE_URL + client
        if device:
            url += "/" + device
        try:
            import webbrowser
            webbrowser.open(url)
        except Exception as exc:
            self._err_lbl.setText(f"⚠  Could not open browser: {exc}")
            return
        self.accept()

    # ── key handling ──────────────────────────────────────────────────────

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.reject()
        else:
            super().keyPressEvent(event)

    # ── painting ──────────────────────────────────────────────────────────

    def paintEvent(self, e: QPaintEvent):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        rect = QRectF(self._card.geometry()).adjusted(0.5, 0.5, -0.5, -0.5)
        path = QPainterPath()
        path.addRoundedRect(rect, BORDER_RADIUS + 2, BORDER_RADIUS + 2)
        grad = QLinearGradient(0, 0, self.width(), self.height())
        grad.setColorAt(0, QColor(
            t.BG_MID.red() + 4, t.BG_MID.green() + 4, t.BG_MID.blue() + 6, 248))
        grad.setColorAt(1, QColor(
            t.BG_DARK.red(), t.BG_DARK.green(), t.BG_DARK.blue(), 248))
        p.fillPath(path, grad)
        border = QColor(t.GLOW)
        border.setAlpha(100)
        p.setPen(QPen(border, 1.3))
        p.setBrush(Qt.NoBrush)
        p.drawPath(path)
        p.end()


# ---------------------------------------------------------------------------
# Clipboard Manager
# ---------------------------------------------------------------------------

import re as _re
import base64 as _base64
import io as _io

_CLIP_CACHE_DIR = CACHE_DIR / "clipboard"
_CLIP_CACHE_DIR.mkdir(parents=True, exist_ok=True)
_CLIP_HISTORY_FILE = CONFIG_DIR / "clipboard_history.json"
_CLIP_DEFAULT_MAX  = 200
_CLIP_MAX_IMAGES   = 50   # cap stored image entries to avoid runaway disk use


def _cb_max_size() -> int:
    try:
        return int(_settings_store.value("clipboard_max", _CLIP_DEFAULT_MAX))
    except Exception:
        return _CLIP_DEFAULT_MAX


def _cb_auto_close() -> bool:
    return _settings_store.value("clipboard_auto_close", "true") == "true"


def _cb_ignore_dupes() -> bool:
    return _settings_store.value("clipboard_ignore_dupes", "true") == "true"


def _cb_capture_images() -> bool:
    return _settings_store.value("clipboard_capture_images", "true") == "true"


def _cb_capture_files() -> bool:
    return _settings_store.value("clipboard_capture_files", "true") == "true"


def _cb_click_copy_close() -> bool:
    return _settings_store.value("clipboard_click_copy_close", "false") == "true"


def _fetch_windows_clipboard_history() -> list[str]:
    """Return plain-text items from Windows clipboard history (Win+V), newest first.

    Requires the `winsdk` package (pip install winsdk-Windows-ApplicationModel-DataTransfer).
    Falls back to an empty list silently when unavailable or when the user has
    Windows clipboard history disabled.
    """
    try:
        try:
            import winsdk.windows.applicationmodel.datatransfer as _wadt
        except ImportError:
            try:
                import winrt.windows.applicationmodel.datatransfer as _wadt  # type: ignore
            except ImportError:
                return []

        import asyncio

        async def _fetch():
            result = await _wadt.Clipboard.get_history_items_async()
            # ClipboardHistoryItemsResultStatus: Success=0, AccessDenied=1, Disabled=2
            if result.status != 0:
                return []
            texts: list[str] = []
            for item in result.items:
                try:
                    dp = item.content
                    if dp.contains(_wadt.StandardDataFormats.TEXT):
                        text = await dp.get_text_async()
                        if text and text.strip():
                            texts.append(text)
                except Exception:
                    continue
            return texts

        loop = asyncio.new_event_loop()
        try:
            return loop.run_until_complete(_fetch())
        finally:
            loop.close()
    except Exception:
        return []


# ---------------------------------------------------------------------------
# Password pattern detection
# ---------------------------------------------------------------------------
import re as _re_pwd

# Matches generated-password style: Word(caps) + 3-6 digits + Word(caps) + symbol(s)
# Examples: Delivered9898Flower$   Mountain42Dog!   River2024Tree#
_PWD_PATTERN = _re_pwd.compile(
    r'^[A-Z][a-z]{2,14}\d{3,8}[A-Z][a-z]{2,14}[^A-Za-z0-9\s]{1,4}$'
)

def _looks_like_password(text: str) -> bool:
    """Return True if the text looks like a generated password."""
    t = text.strip()
    # Must be a single token (no spaces, no newlines)
    if not t or ' ' in t or '\n' in t or '\t' in t:
        return False
    # Reasonable password length range
    if not (8 <= len(t) <= 40):
        return False
    return bool(_PWD_PATTERN.match(t))


class ClipboardStore(QObject):
    """Monitors the system clipboard and maintains a typed history.

    History entries are dicts:
        id        – unique hex string
        type      – "text" | "html" | "image" | "files" | "rich"
        label     – short display string (shown in list)
        text      – plain-text representation (used for search & preview)
        html      – full HTML if available (None otherwise)
        image_file – filename (relative to _CLIP_CACHE_DIR) or None
        files     – list[str] of file paths (for "files" type)
        ts        – ISO timestamp string
        pinned    – bool
        mime_raw  – {format: base64_str} for exact restoration of extra data
    """

    history_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._history: list[dict] = []
        self._ignoring = False          # guard against our own clipboard writes
        self._load()
        # Connect to clipboard dataChanged – event-driven, no polling overhead
        cb = QGuiApplication.clipboard()
        cb.dataChanged.connect(self._on_clipboard_changed)
        # Capture the current clipboard content on startup
        QTimer.singleShot(500, self._on_clipboard_changed)
        # Pull in anything copied while Command Center was closed
        QTimer.singleShot(800, self._sync_from_windows_history)

    # ── public read API ────────────────────────────────────────────────────

    @property
    def history(self) -> list[dict]:
        return self._history

    def get_entry(self, entry_id: str) -> Optional[dict]:
        return next((e for e in self._history if e["id"] == entry_id), None)

    # ── clipboard restoration ──────────────────────────────────────────────

    def restore_to_clipboard(self, entry: dict) -> bool:
        """Re-apply a history entry to the system clipboard. Returns True on success."""
        try:
            self._ignoring = True
            md = QMimeData()
            etype = entry.get("type", "text")

            if etype == "image":
                img_file = entry.get("image_file")
                if img_file:
                    px = QPixmap(_str(_CLIP_CACHE_DIR / img_file))
                    if not px.isNull():
                        md.setImageData(px.toImage())
                # also set text label so pasting in text contexts works
                if entry.get("text"):
                    md.setText(entry["text"])

            elif etype == "files":
                urls = [QUrl.fromLocalFile(f) for f in entry.get("files", [])]
                md.setUrls(urls)
                if entry.get("text"):
                    md.setText(entry["text"])

            else:
                # text / html / rich — restore all captured MIME formats
                if entry.get("html"):
                    md.setHtml(entry["html"])
                if entry.get("text"):
                    md.setText(entry["text"])

            # Restore any extra raw MIME payloads (e.g. CF_RTFTEXT, x-special/*)
            for fmt, b64 in entry.get("mime_raw", {}).items():
                try:
                    md.setData(fmt, QByteArray.fromBase64(b64.encode()))
                except Exception:
                    pass

            QGuiApplication.clipboard().setMimeData(md)
            # Move to top of history (without adding a duplicate)
            self._promote(entry["id"])
            return True
        except Exception as exc:
            print(f"[ClipboardStore] restore error: {exc}", file=sys.stderr)
            return False
        finally:
            self._ignoring = False

    # ── mutation ───────────────────────────────────────────────────────────

    def toggle_pin(self, entry_id: str):
        e = self.get_entry(entry_id)
        if e is not None:
            e["pinned"] = not e["pinned"]
            self._save()
            self.history_changed.emit()

    def delete_entry(self, entry_id: str):
        e = self.get_entry(entry_id)
        if e is None:
            return
        # Delete cached image file
        img = e.get("image_file")
        if img:
            try:
                (_CLIP_CACHE_DIR / img).unlink(missing_ok=True)
            except Exception:
                pass
        self._history = [x for x in self._history if x["id"] != entry_id]
        self._save()
        self.history_changed.emit()

    def clear_unpinned(self):
        # Delete image files for unpinned image entries
        for e in self._history:
            if not e["pinned"] and e.get("image_file"):
                try:
                    (_CLIP_CACHE_DIR / e["image_file"]).unlink(missing_ok=True)
                except Exception:
                    pass
        self._history = [e for e in self._history if e["pinned"]]
        self._save()
        self.history_changed.emit()

    # ── clipboard monitoring ───────────────────────────────────────────────

    def _on_clipboard_changed(self):
        if self._ignoring:
            return
        try:
            cb = QGuiApplication.clipboard()

            # Capture image data IMMEDIATELY before anything can change clipboard state.
            # Chrome/Windows fire two consecutive clipboard events when copying an
            # image: first with PNG data, then HTML-only.  By the time _build_entry
            # would call cb.image(), the clipboard is already in the HTML-only state.
            eager_img: Optional[QImage] = None
            if _cb_capture_images():
                try:
                    _img = cb.image()
                    if _img and not _img.isNull():
                        eager_img = _img.copy()  # deep copy – clipboard may change
                except Exception:
                    pass

            md = cb.mimeData()
            if md is None:
                return
            entry = self._build_entry(md, eager_img)
            if entry is None:
                return
            self._add(entry)
        except Exception as exc:
            print(f"[ClipboardStore] capture error: {exc}", file=sys.stderr)

    # ── image helpers ──────────────────────────────────────────────────────

    @staticmethod
    def _extract_img_from_html(html: str, entry_id: str) -> Optional[str]:
        """Find the first <img> in an HTML clipboard fragment and save a PNG
        preview to _CLIP_CACHE_DIR.  Returns the filename or None.

        Handles:
          - data:image/... base64 URIs
          - file:///... local file paths
        Web http(s) URLs are skipped (would require network access).
        """
        if not html:
            return None
        try:
            import re as _re_h
            m = _re_h.search(r'src=["\']([^"\']{10,})["\']', html, _re_h.IGNORECASE)
            if not m:
                return None
            src = m.group(1).strip()

            img = QImage()

            # data URI
            dm = _re_h.match(r'data:image/[^;]+;base64,(.+)', src, _re_h.DOTALL)
            if dm:
                import base64 as _b64h
                raw = _b64h.b64decode(dm.group(1).strip() + '==')
                if not img.loadFromData(raw):
                    return None

            # local file URI or absolute path
            elif src.startswith(('file:///', 'file://', '/', 'C:', 'D:', 'E:')):
                from urllib.parse import unquote as _unq
                path = src
                if path.startswith('file:///'):
                    path = path[8:]
                elif path.startswith('file://'):
                    path = path[7:]
                path = _unq(path)
                img = QImage(path)
                if img.isNull():
                    return None

            else:
                return None  # http/https – skip

            fname     = f"{entry_id}_prev.png"
            save_path = _CLIP_CACHE_DIR / fname
            if img.save(_str(save_path), "PNG"):
                return fname
        except Exception:
            pass
        return None

    @staticmethod
    def _extract_qimage(md: QMimeData) -> Optional[QImage]:
        """Fallback image extraction from a QMimeData snapshot.
        Only called when the eager cb.image() capture in _on_clipboard_changed
        already failed (e.g. no image on clipboard at that moment).
        Tries every available raw-bytes path so we never lose an image.
        """
        if not md.hasImage():
            return None

        # a) imageData() — PySide6 should return QImage directly
        try:
            data_obj = md.imageData()
            if data_obj is not None:
                if type(data_obj).__name__ == "QImage":
                    img: QImage = data_obj
                    if not img.isNull():
                        return img.copy()
                # Try generic QImage construction from whatever was returned
                try:
                    img = QImage(data_obj)
                    if not img.isNull():
                        return img
                except Exception:
                    pass
        except Exception:
            pass

        # b) Raw bytes from every format Qt reports, plus hard-coded fallbacks
        checked: set = set()
        all_fmts = list(md.formats()) + [
            'application/x-qt-windows-mime;value="PNG"',
            "image/png", "image/jpeg", "image/jpg",
            "image/gif", "image/webp", "image/bmp", "image/x-png",
        ]
        for fmt in all_fmts:
            fl = fmt.lower()
            if fl in checked:
                continue
            if not (fl.startswith("image/") or "png" in fl or "jpeg" in fl
                    or "jfif" in fl or "bmp" in fl or "gif" in fl or "webp" in fl):
                continue
            checked.add(fl)
            try:
                raw = md.data(fmt)
                if raw and len(raw) > 16:
                    img = QImage()
                    if img.loadFromData(bytes(raw)):
                        return img
            except Exception:
                pass

        return None

    def _build_entry(self, md: QMimeData,
                     eager_img: Optional[QImage] = None) -> Optional[dict]:
        """Parse QMimeData into a history entry dict, or return None to skip."""
        # ── Image ──────────────────────────────────────────────────────────
        if _cb_capture_images():
            # Prefer the eagerly-captured image (grabbed before clipboard state
            # could change); fall back to extraction from md.
            img = eager_img if (eager_img and not eager_img.isNull()) else self._extract_qimage(md)
            if img and not img.isNull():
                try:
                    # Check dupe by approximate size signature
                    sig = f"img_{img.width()}_{img.height()}_{img.byteCount()}"
                    if _cb_ignore_dupes():
                        for e in self._history[:20]:
                            if e.get("_sig") == sig:
                                return None
                    # Count existing image entries; evict oldest if over limit
                    img_entries = [e for e in self._history if e["type"] == "image" and not e.get("pinned")]
                    while len(img_entries) >= _CLIP_MAX_IMAGES and img_entries:
                        oldest = img_entries.pop()
                        if oldest.get("image_file"):
                            try:
                                (_CLIP_CACHE_DIR / oldest["image_file"]).unlink(missing_ok=True)
                            except Exception:
                                pass
                        self._history = [x for x in self._history if x["id"] != oldest["id"]]
                    entry_id  = hashlib.md5(sig.encode()).hexdigest()[:12]
                    fname     = f"{entry_id}.png"
                    save_path = _CLIP_CACHE_DIR / fname
                    if not img.save(_str(save_path), "PNG"):
                        # fall through to HTML/text capture
                        pass
                    else:
                        w, h = img.width(), img.height()
                        return {
                            "id": entry_id, "type": "image",
                            "label": f"Image  {w}\u00d7{h}",
                            "text":  f"[Image {w}\u00d7{h}]",
                            "html":  None, "image_file": fname, "files": [],
                            "ts": datetime.now().isoformat(timespec="seconds"),
                            "pinned": False, "mime_raw": {}, "_sig": sig,
                        }
                except Exception:
                    pass

        # ── File list ──────────────────────────────────────────────────────
        if md.hasUrls() and _cb_capture_files():
            try:
                urls = md.urls()
                file_paths = [u.toLocalFile() for u in urls if u.isLocalFile()]
                if file_paths:
                    key = "|".join(sorted(file_paths))
                    if _cb_ignore_dupes() and self._history:
                        if self._history[0].get("files") == file_paths:
                            return None
                    entry_id = hashlib.md5(key.encode()).hexdigest()[:12]
                    label    = (Path(file_paths[0]).name if len(file_paths) == 1
                                else f"{len(file_paths)} files")
                    return {
                        "id": entry_id, "type": "files",
                        "label": label,
                        "text":  "\n".join(file_paths),
                        "html":  None, "image_file": None, "files": file_paths,
                        "ts": datetime.now().isoformat(timespec="seconds"),
                        "pinned": False, "mime_raw": {},
                    }
            except Exception:
                pass

        # ── HTML / rich text ───────────────────────────────────────────────
        if md.hasHtml():
            try:
                html = md.html()
                text = md.text() if md.hasText() else ""
                if not (html or text):
                    return None
                content_key = (text or html).strip()
                if not content_key:
                    return None
                if _cb_ignore_dupes() and self._history:
                    top = self._history[0]
                    if top.get("type") in ("html", "rich", "text") and top.get("text", "").strip() == content_key:
                        return None
                entry_id = hashlib.md5(content_key.encode("utf-8", errors="replace")).hexdigest()[:12]
                # Collect extra MIME formats (e.g. RTF)
                mime_raw = {}
                for fmt in md.formats():
                    if fmt not in ("text/plain", "text/html", "application/x-qt-windows-mime;value=\"HTML Format\""):
                        try:
                            data = md.data(fmt)
                            if data and len(data) < 128 * 1024:  # skip huge blobs
                                mime_raw[fmt] = _base64.b64encode(bytes(data)).decode()
                        except Exception:
                            pass
                etype = "html" if html else "text"
                label = (text or _strip_html(html))[:120].replace("\n", " ").strip()

                # Try to extract an image preview from <img> src (data URI or local file).
                img_file = self._extract_img_from_html(html or "", entry_id)

                return {
                    "id": entry_id, "type": etype,
                    "label": label or "(empty)",
                    "text":  text or _strip_html(html),
                    "html":  html or None,
                    "image_file": img_file, "files": [],
                    "ts": datetime.now().isoformat(timespec="seconds"),
                    "pinned": False, "mime_raw": mime_raw,
                }
            except Exception:
                pass

        # ── Plain text ─────────────────────────────────────────────────────
        if md.hasText():
            try:
                text = md.text()
                if not text or not text.strip():
                    return None
                if _cb_ignore_dupes() and self._history:
                    if self._history[0].get("text", "") == text:
                        return None
                entry_id = hashlib.md5(text.encode("utf-8", errors="replace")).hexdigest()[:12]
                label    = text[:120].replace("\n", " ").strip()
                # Capture extra MIME (e.g. CF_RTFTEXT from Word)
                mime_raw = {}
                for fmt in md.formats():
                    if fmt not in ("text/plain",):
                        try:
                            data = md.data(fmt)
                            if data and len(data) < 128 * 1024:
                                mime_raw[fmt] = _base64.b64encode(bytes(data)).decode()
                        except Exception:
                            pass
                etype = "rich" if mime_raw else "text"
                if etype in ("text", "rich") and _looks_like_password(text):
                    etype = "pwd"
                return {
                    "id": entry_id, "type": etype,
                    "label": label or "(empty)",
                    "text":  text,
                    "html":  None, "image_file": None, "files": [],
                    "ts": datetime.now().isoformat(timespec="seconds"),
                    "pinned": False, "mime_raw": mime_raw,
                }
            except Exception:
                pass

        return None

    # ── internal helpers ───────────────────────────────────────────────────

    def _sync_from_windows_history(self):
        """Import text items from Windows clipboard history that aren't already
        in our store.  Items are appended at the END (older) so they don't push
        out things the user copied recently; pinned items are never displaced.
        Runs in a background thread to avoid blocking the UI."""
        import threading

        def _worker():
            try:
                texts = _fetch_windows_clipboard_history()
                if not texts:
                    return
                changed = False
                # texts[0] is newest – already captured by _on_clipboard_changed.
                # Walk oldest-first so insertion order is chronological.
                for text in reversed(texts):
                    if not text or not text.strip():
                        continue
                    eid = hashlib.md5(text.encode("utf-8", errors="replace")).hexdigest()[:12]
                    if any(e.get("id") == eid for e in self._history):
                        continue
                    label = text[:120].replace("\n", " ").strip()
                    etype = "rich" if False else "text"
                    if _looks_like_password(text):
                        etype = "pwd"
                    entry = {
                        "id": eid, "type": etype,
                        "label": label or "(empty)",
                        "text": text,
                        "html": None, "image_file": None, "files": [],
                        "ts": datetime.now().isoformat(timespec="seconds"),
                        "pinned": False, "mime_raw": {},
                    }
                    self._history.append(entry)
                    changed = True
                if changed:
                    self._trim()
                    self._save()
                    self.history_changed.emit()
            except Exception as exc:
                print(f"[ClipboardStore] win-history sync error: {exc}", file=sys.stderr)

        threading.Thread(target=_worker, daemon=True).start()

    def _add(self, entry: dict):
        existing_id = next((e["id"] for e in self._history if e["id"] == entry["id"]), None)
        if existing_id:
            self._promote(existing_id)
            return
        self._history.insert(0, entry)
        self._trim()
        self._save()
        self.history_changed.emit()

    def _promote(self, entry_id: str):
        idx = next((i for i, e in enumerate(self._history) if e["id"] == entry_id), None)
        if idx is None or idx == 0:
            return
        entry = self._history.pop(idx)
        entry["ts"] = datetime.now().isoformat(timespec="seconds")
        self._history.insert(0, entry)
        self._save()
        self.history_changed.emit()

    def _trim(self):
        max_size = _cb_max_size()
        pinned   = [e for e in self._history if e.get("pinned")]
        unpinned = [e for e in self._history if not e.get("pinned")]
        evicted  = unpinned[max_size:]
        for e in evicted:
            if e.get("image_file"):
                try:
                    (_CLIP_CACHE_DIR / e["image_file"]).unlink(missing_ok=True)
                except Exception:
                    pass
        self._history = pinned + unpinned[:max_size]

    def _save(self):
        try:
            # Don't write _sig to disk (it's a runtime key only)
            clean = [{k: v for k, v in e.items() if k != "_sig"} for e in self._history]
            _CLIP_HISTORY_FILE.write_text(
                json.dumps(clean, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        except Exception as exc:
            print(f"[ClipboardStore] save error: {exc}", file=sys.stderr)

    def _load(self):
        try:
            if _CLIP_HISTORY_FILE.exists():
                data = json.loads(_CLIP_HISTORY_FILE.read_text(encoding="utf-8"))
                if isinstance(data, list):
                    self._history = data
        except Exception as exc:
            print(f"[ClipboardStore] load error: {exc}", file=sys.stderr)
            self._history = []


def _strip_html(html: str) -> str:
    """Very fast HTML tag stripper for label generation."""
    try:
        import re as _r
        return _r.sub(r"<[^>]+>", "", html or "").strip()
    except Exception:
        return html or ""


def _str(p) -> str:
    return str(p)


# Global singleton — created once MainWindow starts
_clip_store: Optional[ClipboardStore] = None


# ── Clipboard Manager Window ───────────────────────────────────────────────

class _ClipItemDelegate(QStyledItemDelegate):
    """Paints each clipboard row with type badge + inline preview (Ditto-style).

    Image entries get a taller row with an actual thumbnail.
    File entries show a folder icon with full file paths.
    """

    ROW_H     = 64   # standard row height
    ROW_H_IMG = 84   # taller row for image thumbnails
    THUMB_W   = 72   # thumbnail width for image rows

    def __init__(self, parent=None):
        super().__init__(parent)
        self._px_cache: dict = {}   # entry_id -> QPixmap (pre-scaled)

    def row_height(self, entry: dict) -> int:
        if not isinstance(entry, dict):
            return self.ROW_H
        if entry.get("type") == "image":
            return self.ROW_H_IMG
        if entry.get("image_file") and entry.get("type") in ("html", "rich"):
            return self.ROW_H_IMG
        return self.ROW_H

    def sizeHint(self, option, index):
        entry = index.data(Qt.UserRole)
        return QSize(option.rect.width(), self.row_height(entry))

    def paint(self, painter, option, index):
        from PySide6.QtWidgets import QStyle
        entry = index.data(Qt.UserRole)
        if not isinstance(entry, dict):
            return

        t = _theme
        etype = entry.get("type", "text")
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.SmoothPixmapTransform)

        rect = QRectF(option.rect).adjusted(5, 2, -5, -2)

        # ── background ──────────────────────────────────────────────────
        is_sel = bool(option.state & QStyle.StateFlag.State_Selected)
        is_hov = bool(option.state & QStyle.StateFlag.State_MouseOver)
        if is_sel:
            bg = QColor(t.GLOW); bg.setAlpha(88)
            bp = QPainterPath(); bp.addRoundedRect(rect, 6, 6)
            painter.fillPath(bp, bg)
        elif is_hov:
            bg = QColor(t.GLOW); bg.setAlpha(28)
            bp = QPainterPath(); bp.addRoundedRect(rect, 6, 6)
            painter.fillPath(bp, bg)

        # ── pin strip ────────────────────────────────────────────────────
        if entry.get("pinned"):
            sr = QRectF(rect.left(), rect.top() + 6, 3, rect.height() - 12)
            sp = QPainterPath(); sp.addRoundedRect(sr, 1.5, 1.5)
            painter.fillPath(sp, QColor(t.GLOW))

        # ── row number ───────────────────────────────────────────────────
        painter.setFont(QFont("Segoe UI", 7))
        painter.setPen(QColor(t.TEXT_DIM))
        painter.drawText(QRectF(rect.left() + 5, rect.top() + 6, 20, 14), Qt.AlignCenter, str(index.row() + 1))

        # ── type badge ───────────────────────────────────────────────────
        _type_lbl = {
            "text":  "TXT",
            "html":  "HTM",
            "rich":  "RTF",
            "image": "IMG",
            "files": "FIL",
            "pwd":   "PWD",
        }
        # Fixed per-type colors so badges are immediately recognisable
        # regardless of the active theme.
        _type_col = {
            "text":  QColor("#7cb9e8"),   # sky blue
            "html":  QColor("#f4a261"),   # amber
            "rich":  QColor("#a8d8a8"),   # sage green
            "image": QColor("#c77dff"),   # purple
            "files": QColor("#56cfe1"),   # cyan
            "pwd":   QColor("#ff6b6b"),   # coral-red  (stands out as sensitive)
        }
        badge_clr = _type_col.get(etype, t.TEXT_SECONDARY)
        bx = rect.left() + 30
        bw, bh = 32, 17
        br = QRectF(bx, rect.top() + 8, bw, bh)
        bp2 = QPainterPath(); bp2.addRoundedRect(br, 3, 3)
        bf = QColor(badge_clr); bf.setAlpha(40)
        painter.fillPath(bp2, bf)
        painter.setPen(QColor(badge_clr))
        painter.setFont(QFont("Segoe UI", 6, QFont.Weight.Bold))
        painter.drawText(br, Qt.AlignCenter, _type_lbl.get(etype, "?"))

        ts_col_w = 46

        # ──────────────────────────────────────────────────────────────────
        # image_file may be set on html/rich entries too (extracted from <img>)
        if etype == "image" or (entry.get("image_file") and etype in ("html", "rich")):
            self._paint_image_row(painter, entry, rect, bx, bw, ts_col_w, badge_clr, t)
        elif etype == "files":
            self._paint_files_row(painter, entry, rect, bx, bw, ts_col_w, t)
        elif etype == "pwd":
            self._paint_pwd_row(painter, entry, rect, bx, bw, ts_col_w, badge_clr, t)
        else:
            self._paint_text_row(painter, entry, rect, bx, bw, ts_col_w, t)

        # ── timestamp ────────────────────────────────────────────────────
        ts = entry.get("ts", "")
        if ts:
            painter.setFont(QFont("Segoe UI", 7))
            painter.setPen(QColor(t.TEXT_DIM))
            painter.drawText(
                QRectF(rect.right() - ts_col_w, rect.top() + 8, ts_col_w - 4, 16),
                Qt.AlignRight | Qt.AlignVCenter, self._fmt_ts(ts)
            )

        # ── pin icon ─────────────────────────────────────────────────────
        if entry.get("pinned"):
            painter.setFont(QFont("Segoe UI", 8))
            painter.setPen(QColor(t.GLOW))
            painter.drawText(
                QRectF(rect.right() - ts_col_w, rect.top() + 28, ts_col_w - 4, 16),
                Qt.AlignRight | Qt.AlignVCenter, "📌"
            )

        painter.restore()

    # ── type-specific row painters ────────────────────────────────────────

    def _paint_image_row(self, painter, entry, rect, bx, bw, ts_col_w, badge_clr, t):
        """Draw thumbnail on the left, label + dimensions on the right."""
        thumb_x  = bx + bw + 8
        thumb_h  = rect.height() - 12
        thumb_w  = self.THUMB_W
        thumb_y  = rect.top() + 6

        px = self._get_thumbnail(entry, thumb_w, int(thumb_h))
        if px and not px.isNull():
            # clip thumbnail to rounded rect
            clip = QPainterPath()
            clip.addRoundedRect(QRectF(thumb_x, thumb_y, thumb_w, thumb_h), 4, 4)
            painter.save()
            painter.setClipPath(clip)
            scaled = px.scaled(int(thumb_w), int(thumb_h), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            dx = thumb_x + (thumb_w - scaled.width()) / 2
            dy = thumb_y + (thumb_h - scaled.height()) / 2
            painter.drawPixmap(int(dx), int(dy), scaled)
            painter.restore()
            # border
            border = QColor(badge_clr); border.setAlpha(90)
            painter.setPen(QPen(border, 1.0))
            painter.setBrush(Qt.NoBrush)
            painter.drawRoundedRect(QRectF(thumb_x, thumb_y, thumb_w, thumb_h), 4, 4)
        else:
            # placeholder
            ph = QPainterPath()
            ph.addRoundedRect(QRectF(thumb_x, thumb_y, thumb_w, thumb_h), 4, 4)
            pf = QColor(badge_clr); pf.setAlpha(20)
            painter.fillPath(ph, pf)
            painter.setFont(QFont("Segoe UI", 16))
            painter.setPen(QColor(badge_clr))
            painter.drawText(QRectF(thumb_x, thumb_y, thumb_w, thumb_h), Qt.AlignCenter, "🖼")

        cx = thumb_x + thumb_w + 10
        cw = rect.right() - cx - ts_col_w - 4

        fl = QFont("Segoe UI", 8.5)
        if entry.get("pinned"): fl.setWeight(QFont.Weight.DemiBold)
        painter.setFont(fl)
        painter.setPen(QColor(t.TEXT_PRIMARY))
        lbl = QFontMetrics(fl).elidedText(entry.get("label", ""), Qt.ElideRight, int(cw))
        painter.drawText(QRectF(cx, rect.top() + 10, cw, 18), Qt.AlignLeft | Qt.AlignVCenter, lbl)

        # dimensions
        dims = self._get_img_dims(entry)
        if dims:
            painter.setFont(QFont("Segoe UI", 7.5))
            painter.setPen(QColor(t.TEXT_DIM))
            painter.drawText(QRectF(cx, rect.top() + 32, cw, 16), Qt.AlignLeft | Qt.AlignVCenter, dims)

    def _paint_files_row(self, painter, entry, rect, bx, bw, ts_col_w, t):
        """Folder icon + full path(s)."""
        # folder icon in a box after the badge
        icon_x = bx + bw + 6
        icon_r = QRectF(icon_x, rect.top() + (self.ROW_H - 30) / 2, 28, 30)
        painter.setFont(QFont("Segoe UI Emoji", 16))
        painter.setPen(QColor(t.TEXT_SECONDARY))
        painter.drawText(icon_r, Qt.AlignCenter, "📁")

        cx = icon_x + 34
        cw = rect.right() - cx - ts_col_w - 4
        files = entry.get("files") or []
        first_path = files[0] if files else entry.get("label", "")

        # top line: full path of first file (elide middle so both ends visible)
        fl = QFont("Segoe UI", 8.5)
        if entry.get("pinned"): fl.setWeight(QFont.Weight.DemiBold)
        painter.setFont(fl)
        painter.setPen(QColor(t.TEXT_PRIMARY))
        lbl = QFontMetrics(fl).elidedText(first_path, Qt.ElideMiddle, int(cw))
        painter.drawText(QRectF(cx, rect.top() + 8, cw, 18), Qt.AlignLeft | Qt.AlignVCenter, lbl)

        # second line
        fp = QFont("Segoe UI", 7.5)
        painter.setFont(fp)
        painter.setPen(QColor(t.TEXT_DIM))
        if len(files) > 1:
            extra_txt = f"+{len(files) - 1} more file{'s' if len(files) - 1 > 1 else ''}"
            painter.drawText(QRectF(cx, rect.top() + 30, cw, 16), Qt.AlignLeft | Qt.AlignVCenter, extra_txt)
        elif first_path:
            fname = os.path.basename(first_path)
            prev  = QFontMetrics(fp).elidedText(fname, Qt.ElideRight, int(cw))
            painter.drawText(QRectF(cx, rect.top() + 30, cw, 16), Qt.AlignLeft | Qt.AlignVCenter, prev)

    def _paint_pwd_row(self, painter, entry, rect, bx, bw, ts_col_w, badge_clr, t):
        """Password row: lock icon + plain-text label + preview."""
        cx = bx + bw + 8
        cw = rect.right() - cx - ts_col_w - 4

        # lock icon
        painter.setFont(QFont("Segoe UI Emoji", 13))
        painter.setPen(QColor(badge_clr))
        painter.drawText(QRectF(cx, rect.top() + 8, 20, 20), Qt.AlignCenter, "🔑")

        cx2 = cx + 26

        # top line: the password in the badge accent colour so it stands out
        fl = QFont("Segoe UI", 8.5)
        if entry.get("pinned"):
            fl.setWeight(QFont.Weight.DemiBold)
        painter.setFont(fl)
        painter.setPen(QColor(badge_clr))
        lbl = QFontMetrics(fl).elidedText(entry.get("label", ""), Qt.ElideRight, int(cw - 26))
        painter.drawText(QRectF(cx2, rect.top() + 8, cw - 26, 18), Qt.AlignLeft | Qt.AlignVCenter, lbl)

        # second line: length hint in dim colour
        pwd = entry.get("text", "")
        fp = QFont("Segoe UI", 7.5)
        painter.setFont(fp)
        painter.setPen(QColor(t.TEXT_DIM))
        painter.drawText(QRectF(cx2, rect.top() + 30, cw - 26, 16),
                         Qt.AlignLeft | Qt.AlignVCenter, f"{len(pwd)}-char password")

    def _paint_text_row(self, painter, entry, rect, bx, bw, ts_col_w, t):
        """Standard text/html/rich row: label + preview line."""
        cx = bx + bw + 8
        cw = rect.right() - cx - ts_col_w - 4

        fl = QFont("Segoe UI", 8.5)
        if entry.get("pinned"): fl.setWeight(QFont.Weight.DemiBold)
        painter.setFont(fl)
        painter.setPen(QColor(t.TEXT_PRIMARY))
        lbl = QFontMetrics(fl).elidedText(entry.get("label", ""), Qt.ElideRight, int(cw))
        painter.drawText(QRectF(cx, rect.top() + 8, cw, 18), Qt.AlignLeft | Qt.AlignVCenter, lbl)

        preview = ""
        for line in (entry.get("text") or "").splitlines():
            line = line.strip()
            if line:
                preview = line
                break
        if preview:
            fp = QFont("Segoe UI", 7.5)
            painter.setFont(fp)
            painter.setPen(QColor(t.TEXT_DIM))
            prev = QFontMetrics(fp).elidedText(preview, Qt.ElideRight, int(cw))
            painter.drawText(QRectF(cx, rect.top() + 30, cw, 16), Qt.AlignLeft | Qt.AlignVCenter, prev)

    # ── helpers ───────────────────────────────────────────────────────────

    def _get_thumbnail(self, entry: dict, max_w: int, max_h: int) -> Optional[QPixmap]:
        eid = entry.get("id", "")
        if eid in self._px_cache:
            return self._px_cache[eid]
        img_file = entry.get("image_file")
        if not img_file:
            return None
        px = QPixmap(_str(_CLIP_CACHE_DIR / img_file))
        if px.isNull():
            return None
        # cache a 2x-res version so it looks sharp on high-DPI
        cached = px.scaled(max_w * 2, max_h * 2, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self._px_cache[eid] = cached
        return cached

    def _get_img_dims(self, entry: dict) -> str:
        eid = entry.get("id", "")
        if eid in self._px_cache:
            px = self._px_cache[eid]
            if not px.isNull():
                # cached pixmap may be pre-scaled; use original file for real dims
                img_file = entry.get("image_file")
                if img_file:
                    orig = QPixmap(_str(_CLIP_CACHE_DIR / img_file))
                    if not orig.isNull():
                        return f"{orig.width()} × {orig.height()} px"
        return ""

    def invalidate_cache(self, entry_id: str = ""):
        """Remove one or all cached thumbnails (call when entries are deleted)."""
        if entry_id:
            self._px_cache.pop(entry_id, None)
        else:
            self._px_cache.clear()

    def _fmt_ts(self, ts: str) -> str:
        try:
            from datetime import datetime as _dtt
            dt = _dtt.fromisoformat(ts)
            now = _dtt.now()
            return dt.strftime("%H:%M") if dt.date() == now.date() else dt.strftime("%m/%d")
        except Exception:
            return ""


class ClipboardManagerWindow(QWidget):
    """Standalone, frameless clipboard manager — Ditto-style single-column list.

    Top-level Qt.Window so it stays visible even when the main window is
    minimized.  Theme-aware via _theme.register().
    """

    _TYPE_FILTERS = [
        ("All",    None),
        ("Text",   ("text", "rich")),
        ("HTML",   ("html",)),
        ("Images", ("image",)),
        ("Files",  ("files",)),
        ("PWD",    ("pwd",)),
    ]

    def __init__(self):
        super().__init__(None)
        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_QuitOnClose, False)
        self._drag_pos: Optional[QPoint] = None
        self._active_type_filter = None
        self._selected_id: Optional[str] = None
        self._delegate = _ClipItemDelegate(self)
        self._build_ui()
        self.setMinimumSize(400, 460)
        self.resize(460, 600)
        _theme.register(self._refresh_theme)
        self._refresh_theme()
        if _clip_store:
            _clip_store.history_changed.connect(self._refresh_list)

    # ── UI construction ────────────────────────────────────────────────────

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(10, 10, 10, 10)
        outer.setSpacing(0)

        self._card = QWidget()
        self._card.setObjectName("clip_card")
        self._card.setGraphicsEffect(make_shadow(self._card, 32, QColor(0, 0, 0, 210)))
        cl = QVBoxLayout(self._card)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(0)

        # ── title bar ─────────────────────────────────────────────────────
        tbar = QWidget()
        tbar.setFixedHeight(46)
        tbar.setStyleSheet("background:transparent;")
        tbl = QHBoxLayout(tbar)
        tbl.setContentsMargins(16, 0, 12, 0)
        tbl.setSpacing(0)

        ico = QLabel("📋")
        ico.setFont(QFont("Segoe UI", 11))
        ico.setStyleSheet("background:transparent;")
        tbl.addWidget(ico)
        tbl.addSpacing(8)

        ttl = QLabel("Clipboard Manager")
        ttl.setFont(QFont("Segoe UI", 11, QFont.Weight.DemiBold))
        ttl.setObjectName("clip_title")
        ttl.setStyleSheet("background:transparent;")
        tbl.addWidget(ttl)
        tbl.addStretch()

        self._count_lbl = QLabel("")
        self._count_lbl.setObjectName("clip_count")
        self._count_lbl.setStyleSheet("background:transparent;")
        tbl.addWidget(self._count_lbl)
        tbl.addSpacing(10)

        self._settings_btn = QPushButton("⚙")
        self._settings_btn.setFixedSize(26, 26)
        self._settings_btn.setCursor(Qt.PointingHandCursor)
        self._settings_btn.setToolTip("Clipboard Manager settings")
        self._settings_btn.setObjectName("clip_icon_btn")
        self._settings_btn.clicked.connect(self._open_settings_panel)
        tbl.addWidget(self._settings_btn)
        tbl.addSpacing(6)

        close_btn = TitleBarButton(COLOR_BTN_CLOSE, "x")
        close_btn.clicked.connect(self.hide)
        tbl.addWidget(close_btn)

        tbar.mousePressEvent   = lambda ev: (setattr(self, "_drag_pos", ev.globalPosition().toPoint()) if ev.button() == Qt.LeftButton else None)
        tbar.mouseMoveEvent    = self._drag_move
        tbar.mouseReleaseEvent = lambda ev: setattr(self, "_drag_pos", None)
        cl.addWidget(tbar)

        # ── separator ─────────────────────────────────────────────────────
        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        sep.setObjectName("clip_sep")
        cl.addWidget(sep)

        # ── search bar ────────────────────────────────────────────────────
        search_w = QWidget()
        search_w.setStyleSheet("background:transparent;")
        search_w.setFixedHeight(42)
        sl = QHBoxLayout(search_w)
        sl.setContentsMargins(12, 7, 12, 0)
        sl.setSpacing(0)

        self._search = QLineEdit()
        self._search.setPlaceholderText("  Search clipboard history…")
        self._search.setFixedHeight(28)
        self._search.setClearButtonEnabled(True)
        self._search.textChanged.connect(self._refresh_list)
        self._search.setObjectName("clip_search")
        sl.addWidget(self._search)
        cl.addWidget(search_w)

        # ── type filter chips ─────────────────────────────────────────────
        chip_bar = QWidget()
        chip_bar.setStyleSheet("background:transparent;")
        chip_bar.setFixedHeight(32)
        cbl = QHBoxLayout(chip_bar)
        cbl.setContentsMargins(12, 0, 12, 0)
        cbl.setSpacing(6)

        self._chip_btns: list = []
        for label, fval in self._TYPE_FILTERS:
            btn = QPushButton(label)
            btn.setFixedHeight(22)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setCheckable(True)
            btn.setChecked(fval is None)
            btn.setObjectName("clip_chip")
            btn.clicked.connect(lambda checked, v=fval: self._set_type_filter(v))
            cbl.addWidget(btn)
            self._chip_btns.append((fval, btn))
        cbl.addStretch()
        cl.addWidget(chip_bar)

        # ── list ──────────────────────────────────────────────────────────
        list_w = QWidget()
        list_w.setStyleSheet("background:transparent;")
        ll = QVBoxLayout(list_w)
        ll.setContentsMargins(10, 4, 10, 4)
        ll.setSpacing(0)

        self._list = QListWidget()
        self._list.setObjectName("clip_list")
        self._list.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._list.setSelectionMode(QListWidget.SingleSelection)
        self._list.setMouseTracking(True)
        self._list.viewport().setMouseTracking(True)
        self._list.setItemDelegate(self._delegate)
        self._list.currentItemChanged.connect(self._on_item_changed)
        self._list.itemActivated.connect(self._copy_and_close)
        self._list.itemClicked.connect(self._on_item_clicked)
        ll.addWidget(self._list)
        cl.addWidget(list_w, 1)

        # ── action bar ────────────────────────────────────────────────────
        action_w = QWidget()
        action_w.setStyleSheet("background:transparent;")
        action_w.setFixedHeight(44)
        al = QHBoxLayout(action_w)
        al.setContentsMargins(12, 6, 12, 6)
        al.setSpacing(6)

        self._copy_close_btn = QPushButton("⎘  Copy & Close")
        self._copy_close_btn.setFixedHeight(28)
        self._copy_close_btn.setCursor(Qt.PointingHandCursor)
        self._copy_close_btn.setObjectName("clip_action_btn")
        self._copy_close_btn.clicked.connect(self._copy_and_close)
        al.addWidget(self._copy_close_btn)

        self._copy_btn = QPushButton("Copy")
        self._copy_btn.setFixedHeight(28)
        self._copy_btn.setCursor(Qt.PointingHandCursor)
        self._copy_btn.setObjectName("clip_action_btn")
        self._copy_btn.clicked.connect(self._copy_only)
        al.addWidget(self._copy_btn)

        self._pin_btn = QPushButton("📌 Pin")
        self._pin_btn.setFixedHeight(28)
        self._pin_btn.setCursor(Qt.PointingHandCursor)
        self._pin_btn.setObjectName("clip_action_btn")
        self._pin_btn.clicked.connect(self._toggle_pin)
        al.addWidget(self._pin_btn)

        self._del_btn = QPushButton("✕ Delete")
        self._del_btn.setFixedHeight(28)
        self._del_btn.setCursor(Qt.PointingHandCursor)
        self._del_btn.setObjectName("clip_danger_btn")
        self._del_btn.clicked.connect(self._delete_selected)
        al.addWidget(self._del_btn)

        al.addStretch()

        self._clear_btn = QPushButton("Clear All")
        self._clear_btn.setFixedHeight(28)
        self._clear_btn.setCursor(Qt.PointingHandCursor)
        self._clear_btn.setObjectName("clip_danger_btn")
        self._clear_btn.clicked.connect(self._clear_unpinned)
        al.addWidget(self._clear_btn)

        cl.addWidget(action_w)
        outer.addWidget(self._card)
        self._update_action_buttons()

    # ── settings ───────────────────────────────────────────────────────────

    def _open_settings_panel(self):
        dlg = _ClipSettingsDialog(self)
        dlg.exec()
        if _clip_store:
            _clip_store.history_changed.emit()

    # ── list management ────────────────────────────────────────────────────

    def _filtered_history(self) -> list[dict]:
        if _clip_store is None:
            return []
        query = self._search.text().lower().strip()
        tfilter = self._active_type_filter
        result = []
        for e in _clip_store.history:
            if tfilter and e.get("type") not in tfilter:
                continue
            if query and query not in (e.get("text") or "").lower() and query not in (e.get("label") or "").lower():
                continue
            result.append(e)
        return result

    def _refresh_list(self):
        entries = self._filtered_history()
        prev_id = self._selected_id
        self._list.blockSignals(True)
        self._list.clear()
        for e in entries:
            item = QListWidgetItem()
            item.setData(Qt.UserRole, e)
            item.setSizeHint(QSize(100, self._delegate.row_height(e)))
            self._list.addItem(item)
        self._list.blockSignals(False)
        # Restore selection
        if prev_id:
            for i in range(self._list.count()):
                e = self._list.item(i).data(Qt.UserRole)
                if isinstance(e, dict) and e.get("id") == prev_id:
                    self._list.setCurrentRow(i)
                    break
        total = len(_clip_store.history) if _clip_store else 0
        shown = len(entries)
        self._count_lbl.setText(f"{shown}/{total}")
        self._update_action_buttons()

    # ── selection ─────────────────────────────────────────────────────────

    def _on_item_changed(self, current, _prev):
        entry = current.data(Qt.UserRole) if current else None
        self._selected_id = entry.get("id") if isinstance(entry, dict) else None
        self._update_action_buttons()

    def _on_item_clicked(self, item):
        """Single-click copy-and-close when the setting is enabled."""
        if _cb_click_copy_close():
            entry = item.data(Qt.UserRole) if item else None
            if isinstance(entry, dict) and _clip_store:
                if _clip_store.restore_to_clipboard(entry):
                    self.hide()

    def _update_action_buttons(self):
        entry = self._current_entry()
        has = entry is not None
        pinned = entry.get("pinned", False) if has else False
        self._copy_close_btn.setEnabled(has)
        self._copy_btn.setEnabled(has)
        self._pin_btn.setEnabled(has)
        self._del_btn.setEnabled(has)
        self._pin_btn.setText("📌 Unpin" if pinned else "📌 Pin")

    # ── actions ────────────────────────────────────────────────────────────

    def _current_entry(self) -> Optional[dict]:
        item = self._list.currentItem()
        if item is None:
            return None
        e = item.data(Qt.UserRole)
        return e if isinstance(e, dict) else None

    def _copy_and_close(self, *_):
        entry = self._current_entry()
        if entry and _clip_store:
            if _clip_store.restore_to_clipboard(entry):
                self.hide()

    def _copy_only(self, *_):
        entry = self._current_entry()
        if entry and _clip_store:
            _clip_store.restore_to_clipboard(entry)

    def _toggle_pin(self):
        entry = self._current_entry()
        if entry and _clip_store:
            _clip_store.toggle_pin(entry["id"])

    def _delete_selected(self):
        entry = self._current_entry()
        if entry and _clip_store:
            self._delegate.invalidate_cache(entry["id"])
            _clip_store.delete_entry(entry["id"])
            self._selected_id = None
            self._update_action_buttons()

    def _clear_unpinned(self):
        if _clip_store is None:
            return
        total_unpinned = sum(1 for e in _clip_store.history if not e.get("pinned"))
        if total_unpinned == 0:
            return
        from PySide6.QtWidgets import QMessageBox
        mb = QMessageBox(self)
        mb.setWindowFlags(mb.windowFlags() | Qt.FramelessWindowHint)
        mb.setIcon(QMessageBox.Warning)
        mb.setWindowTitle("Clear History")
        mb.setText(f"Delete {total_unpinned} unpinned entries?")
        mb.setStandardButtons(QMessageBox.Yes | QMessageBox.Cancel)
        mb.setDefaultButton(QMessageBox.Cancel)
        if mb.exec() == QMessageBox.Yes:
            self._delegate.invalidate_cache()
            _clip_store.clear_unpinned()

    def _set_type_filter(self, fval):
        self._active_type_filter = fval
        for v, btn in self._chip_btns:
            btn.setChecked(v == fval)
        self._refresh_list()

    # ── drag support ───────────────────────────────────────────────────────

    def _drag_move(self, ev: QMouseEvent):
        if self._drag_pos and (ev.buttons() & Qt.LeftButton):
            delta = ev.globalPosition().toPoint() - self._drag_pos
            self.move(self.pos() + delta)
            self._drag_pos = ev.globalPosition().toPoint()

    # ── keyboard ───────────────────────────────────────────────────────────

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.hide()
        elif event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self._copy_and_close()
        elif event.key() == Qt.Key_Delete:
            self._delete_selected()
        else:
            super().keyPressEvent(event)

    # ── painting ───────────────────────────────────────────────────────────

    def paintEvent(self, e: QPaintEvent):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        rect = QRectF(self._card.geometry()).adjusted(0.5, 0.5, -0.5, -0.5)
        path = QPainterPath()
        path.addRoundedRect(rect, BORDER_RADIUS + 2, BORDER_RADIUS + 2)
        grad = QLinearGradient(0, 0, self.width(), self.height())
        grad.setColorAt(0, QColor(t.BG_MID.red() + 4, t.BG_MID.green() + 4, t.BG_MID.blue() + 6, 248))
        grad.setColorAt(1, QColor(t.BG_DARK.red(), t.BG_DARK.green(), t.BG_DARK.blue(), 248))
        p.fillPath(path, grad)
        border = QColor(t.GLOW); border.setAlpha(100)
        p.setPen(QPen(border, 1.3)); p.setBrush(Qt.NoBrush); p.drawPath(path)
        p.end()

    def showEvent(self, ev):
        super().showEvent(ev)
        self._refresh_list()
        self._search.clear()
        self._search.setFocus()

    def hideEvent(self, ev):
        super().hideEvent(ev)

    # ── theming ────────────────────────────────────────────────────────────

    def _refresh_theme(self):
        t = _theme
        g = t.GLOW
        bg_d = t.BG_DARK; bg_m = t.BG_MID
        tp = t.TEXT_PRIMARY; ts_ = t.TEXT_SECONDARY
        ar = t.ACCENT_RED

        for child in self.findChildren(QLabel, "clip_title"):
            child.setStyleSheet(f"color:{tp.name()}; background:transparent;")
        for child in self.findChildren(QLabel, "clip_count"):
            child.setStyleSheet(f"color:{t.TEXT_DIM.name()}; background:transparent; font-size:8pt;")
        for child in self.findChildren(QFrame, "clip_sep"):
            child.setStyleSheet(f"color:rgba({g.red()},{g.green()},{g.blue()},60);")

        for child in self.findChildren(QLineEdit, "clip_search"):
            child.setStyleSheet(f"""
                QLineEdit {{
                    background:rgba({bg_d.red()},{bg_d.green()},{bg_d.blue()},200);
                    border:1px solid rgba({g.red()},{g.green()},{g.blue()},70);
                    border-radius:6px; color:{tp.name()}; padding:0 8px;
                    selection-background-color:rgba({g.red()},{g.green()},{g.blue()},100);
                }}
                QLineEdit:focus {{
                    border:1px solid rgba({g.red()},{g.green()},{g.blue()},180);
                }}
            """)

        # List — delegate owns item drawing; stylesheet handles background + scrollbar only
        for child in self.findChildren(QListWidget, "clip_list"):
            child.setStyleSheet(f"""
                QListWidget {{
                    background:rgba({bg_d.red()},{bg_d.green()},{bg_d.blue()},190);
                    border:1px solid rgba({g.red()},{g.green()},{g.blue()},55);
                    border-radius:6px; outline:none; color:{tp.name()};
                }}
                QListWidget::item {{ background:transparent; border:none; }}
                QListWidget::item:selected {{ background:transparent; }}
                QListWidget::item:hover {{ background:transparent; }}
                QScrollBar:vertical {{
                    background:rgba({bg_d.red()},{bg_d.green()},{bg_d.blue()},120);
                    width:6px; border-radius:3px; margin:0;
                }}
                QScrollBar::handle:vertical {{
                    background:rgba({g.red()},{g.green()},{g.blue()},130);
                    border-radius:3px; min-height:20px;
                }}
                QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height:0; }}
            """)
        self._list.viewport().update()

        icon_ss = f"""
            QPushButton {{
                background:rgba({g.red()},{g.green()},{g.blue()},30);
                border:1px solid rgba({g.red()},{g.green()},{g.blue()},80);
                border-radius:6px; color:{ts_.name()};
            }}
            QPushButton:hover {{
                background:rgba({g.red()},{g.green()},{g.blue()},80);
                color:{g.name()};
                border-color:rgba({g.red()},{g.green()},{g.blue()},180);
            }}
        """
        for child in self.findChildren(QPushButton, "clip_icon_btn"):
            child.setStyleSheet(icon_ss)

        action_ss = f"""
            QPushButton {{
                background:rgba({bg_m.red()},{bg_m.green()},{bg_m.blue()},180);
                border:1px solid rgba({g.red()},{g.green()},{g.blue()},80);
                border-radius:6px; color:{ts_.name()};
                padding:0 12px; font-size:8.5pt;
            }}
            QPushButton:hover {{
                background:rgba({g.red()},{g.green()},{g.blue()},70);
                color:{g.name()};
                border-color:rgba({g.red()},{g.green()},{g.blue()},180);
            }}
            QPushButton:pressed {{ background:rgba({g.red()},{g.green()},{g.blue()},110); }}
            QPushButton:disabled {{
                background:rgba({bg_m.red()},{bg_m.green()},{bg_m.blue()},80);
                border-color:rgba({g.red()},{g.green()},{g.blue()},30);
                color:rgba({ts_.red()},{ts_.green()},{ts_.blue()},80);
            }}
        """
        for child in self.findChildren(QPushButton, "clip_action_btn"):
            child.setStyleSheet(action_ss)

        danger_ss = f"""
            QPushButton {{
                background:rgba({bg_m.red()},{bg_m.green()},{bg_m.blue()},180);
                border:1px solid rgba({ar.red()},{ar.green()},{ar.blue()},70);
                border-radius:6px; color:rgba({ar.red()},{ar.green()},{ar.blue()},200);
                padding:0 12px; font-size:8.5pt;
            }}
            QPushButton:hover {{
                background:rgba({ar.red()},{ar.green()},{ar.blue()},60);
                color:{tp.name()};
                border-color:rgba({ar.red()},{ar.green()},{ar.blue()},200);
            }}
        """
        for child in self.findChildren(QPushButton, "clip_danger_btn"):
            child.setStyleSheet(danger_ss)

        chip_active = f"""
            QPushButton {{
                background:rgba({g.red()},{g.green()},{g.blue()},110);
                border:1px solid rgba({g.red()},{g.green()},{g.blue()},200);
                border-radius:10px; color:{tp.name()};
                padding:0 10px; font-size:8pt;
            }}
        """
        chip_inactive = f"""
            QPushButton {{
                background:rgba({bg_m.red()},{bg_m.green()},{bg_m.blue()},140);
                border:1px solid rgba({g.red()},{g.green()},{g.blue()},50);
                border-radius:10px; color:{ts_.name()};
                padding:0 10px; font-size:8pt;
            }}
            QPushButton:hover {{
                background:rgba({g.red()},{g.green()},{g.blue()},55);
                color:{tp.name()};
            }}
        """
        for _fval, btn in self._chip_btns:
            btn.setStyleSheet(chip_active if btn.isChecked() else chip_inactive)


class _ClipSettingsDialog(QDialog):
    """Compact settings dialog for the Clipboard Manager."""

    def __init__(self, parent=None):
        super().__init__(parent, Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self._drag_pos: Optional[QPoint] = None
        self._build_ui()
        self.setFixedSize(400, 410)

    def _build_ui(self):
        t = _theme
        outer = QVBoxLayout(self)
        outer.setContentsMargins(10, 10, 10, 10)
        outer.setSpacing(0)

        card = QWidget()
        card.setGraphicsEffect(make_shadow(card, 28, QColor(0, 0, 0, 200)))
        cl = QVBoxLayout(card)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(0)

        tbar = QWidget(); tbar.setFixedHeight(44)
        tbar.setStyleSheet("background:transparent;")
        tbl = QHBoxLayout(tbar); tbl.setContentsMargins(16, 0, 12, 0)
        ttl = QLabel("⚙  Clipboard Settings")
        ttl.setFont(QFont("Segoe UI", 10, QFont.Weight.DemiBold))
        ttl.setStyleSheet(f"color:{t.TEXT_PRIMARY.name()}; background:transparent;")
        tbl.addWidget(ttl); tbl.addStretch()
        cb2 = TitleBarButton(COLOR_BTN_CLOSE, "x"); cb2.clicked.connect(self.accept)
        tbl.addWidget(cb2)
        tbar.mousePressEvent  = lambda ev: (setattr(self, "_drag_pos", ev.globalPosition().toPoint()) if ev.button() == Qt.LeftButton else None)
        tbar.mouseMoveEvent   = self._drag_move
        tbar.mouseReleaseEvent = lambda ev: setattr(self, "_drag_pos", None)
        cl.addWidget(tbar)

        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);")
        cl.addWidget(sep)

        form_w = QWidget(); form_w.setStyleSheet("background:transparent;")
        fl = QVBoxLayout(form_w)
        fl.setContentsMargins(22, 16, 22, 16)
        fl.setSpacing(10)

        def row_label(txt):
            lbl = QLabel(txt)
            lbl.setFont(QFont("Segoe UI", 8, QFont.Weight.DemiBold))
            lbl.setStyleSheet(f"color:{t.TEXT_SECONDARY.name()}; background:transparent; letter-spacing:0.8px;")
            return lbl

        input_ss = f"""
            QComboBox, QSpinBox {{
                background:rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},200);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},80);
                border-radius:6px; color:{t.TEXT_PRIMARY.name()};
                padding:2px 8px; min-height:26px;
            }}
        """

        # Max history size
        fl.addWidget(row_label("MAX HISTORY SIZE"))
        self._max_combo = QComboBox()
        self._max_combo.setStyleSheet(input_ss)
        for v in (50, 100, 200, 500, 1000):
            self._max_combo.addItem(str(v), v)
        current_max = _cb_max_size()
        idx = self._max_combo.findData(current_max)
        self._max_combo.setCurrentIndex(idx if idx >= 0 else 2)
        fl.addWidget(self._max_combo)

        fl.addSpacing(4)

        def make_toggle(label, key, default_true=True):
            chk = QCheckBox(label)
            chk.setFont(QFont("Segoe UI", 9))
            chk.setStyleSheet(f"color:{t.TEXT_PRIMARY.name()}; background:transparent;")
            chk.setChecked(_settings_store.value(key, "true" if default_true else "false") == "true")
            return chk

        self._auto_close_chk       = make_toggle("Close manager after copying",         "clipboard_auto_close")
        self._click_copy_close_chk = make_toggle("Auto Copy and Close on Click",         "clipboard_click_copy_close", default_true=False)
        self._dedup_chk            = make_toggle("Ignore duplicate entries",              "clipboard_ignore_dupes")
        self._img_chk              = make_toggle("Capture image copies",                 "clipboard_capture_images")
        self._file_chk             = make_toggle("Capture file copies",                  "clipboard_capture_files")

        for chk in (self._auto_close_chk, self._click_copy_close_chk, self._dedup_chk, self._img_chk, self._file_chk):
            fl.addWidget(chk)

        fl.addStretch()
        cl.addWidget(form_w, 1)

        # ── save button ───────────────────────────────────────────────────
        btn_w = QWidget(); btn_w.setStyleSheet("background:transparent;")
        bl = QHBoxLayout(btn_w); bl.setContentsMargins(22, 0, 22, 16)
        bl.addStretch()
        save_btn = QPushButton("Save")
        save_btn.setFixedHeight(30)
        save_btn.setFont(QFont("Segoe UI", 9, QFont.Weight.DemiBold))
        save_btn.setCursor(Qt.PointingHandCursor)
        save_btn.setStyleSheet(f"""
            QPushButton {{
                background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},210),
                    stop:1 rgba({t.ACCENT_TEAL.red()},{t.ACCENT_TEAL.green()},{t.ACCENT_TEAL.blue()},190));
                border:none; border-radius:7px; color:white; padding:0 20px;
            }}
            QPushButton:hover {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},255);
            }}
        """)
        save_btn.clicked.connect(self._save_and_close)
        bl.addWidget(save_btn)
        cl.addWidget(btn_w)

        outer.addWidget(card)

    def _save_and_close(self):
        max_val = self._max_combo.currentData()
        if max_val:
            _settings_store.setValue("clipboard_max", str(max_val))
        _settings_store.setValue("clipboard_auto_close",         "true" if self._auto_close_chk.isChecked()       else "false")
        _settings_store.setValue("clipboard_click_copy_close",  "true" if self._click_copy_close_chk.isChecked() else "false")
        _settings_store.setValue("clipboard_ignore_dupes",      "true" if self._dedup_chk.isChecked()           else "false")
        _settings_store.setValue("clipboard_capture_images",    "true" if self._img_chk.isChecked()             else "false")
        _settings_store.setValue("clipboard_capture_files",     "true" if self._file_chk.isChecked()            else "false")
        _settings_store.sync()
        self.accept()

    def _drag_move(self, ev: QMouseEvent):
        if self._drag_pos:
            delta = ev.globalPosition().toPoint() - self._drag_pos
            self.move(self.pos() + delta)
            self._drag_pos = ev.globalPosition().toPoint()

    def paintEvent(self, e: QPaintEvent):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        rect = QRectF(10, 10, self.width() - 20, self.height() - 20)
        path = QPainterPath(); path.addRoundedRect(rect, BORDER_RADIUS + 2, BORDER_RADIUS + 2)
        grad = QLinearGradient(0, 0, self.width(), self.height())
        grad.setColorAt(0, QColor(t.BG_MID.red() + 4, t.BG_MID.green() + 4, t.BG_MID.blue() + 6, 248))
        grad.setColorAt(1, QColor(t.BG_DARK.red(), t.BG_DARK.green(), t.BG_DARK.blue(), 248))
        p.fillPath(path, grad)
        border = QColor(t.GLOW); border.setAlpha(90)
        p.setPen(QPen(border, 1.3)); p.setBrush(Qt.NoBrush); p.drawPath(path)
        p.end()

    def keyPressEvent(self, ev):
        if ev.key() == Qt.Key_Escape:
            self.reject()
        else:
            super().keyPressEvent(ev)


# ---------------------------------------------------------------------------
# Notebook system
# ---------------------------------------------------------------------------

import base64 as _base64

NOTEBOOK_DIR = CONFIG_DIR / "notebooks"
NOTEBOOK_DIR.mkdir(parents=True, exist_ok=True)

MEDIA_LIBRARY_DIR  = CONFIG_DIR / "media_library"
MEDIA_LIBRARY_GIFS = MEDIA_LIBRARY_DIR / "gifs"
MEDIA_LIBRARY_PICS = MEDIA_LIBRARY_DIR / "pictures"
for _d in (MEDIA_LIBRARY_DIR, MEDIA_LIBRARY_GIFS, MEDIA_LIBRARY_PICS):
    _d.mkdir(parents=True, exist_ok=True)
MEDIA_LIBRARY_JSON = MEDIA_LIBRARY_DIR / "library.json"

_DEFAULT_EMOTICONS = [
    "¯\\_(ツ)_/¯",
    "(╯°□°)╯︵ ┻━┻",
    "( ͡° ͜ʖ ͡°)",
    "ヽ(°〇°)ﾉ",
    "(ง'̀-'́)ง",
    "┬─┬ノ( º _ ºノ)",
    "ಠ_ಠ",
    "(づ｡◕‿‿◕｡)づ",
    "乁( ・ω・ )ㄏ",
    "(ﾉ◕ヮ◕)ﾉ*:･ﾟ✧",
]


class NotebookStore:
    """Manages .ccnote files on disk (JSON with rich HTML content)."""

    @staticmethod
    def list_notes() -> list[dict]:
        """Return all notes sorted: pinned first, then by modified date descending."""
        notes = []
        for f in NOTEBOOK_DIR.glob("*.ccnote"):
            try:
                data = json.loads(f.read_text(encoding="utf-8"))
                notes.append({
                    "id":       f.stem,
                    "title":    data.get("title", f.stem),
                    "modified": data.get("modified", ""),
                    "color":    data.get("color", ""),
                    "pinned":   bool(data.get("pinned", False)),
                    "path":     f,
                })
            except Exception:
                pass
        # Sort: pinned first, then by modified descending (newest on top)
        notes.sort(key=lambda n: (
            not n["pinned"],
            n["modified"] if n["modified"] else ""
        ), reverse=False)
        # Secondary: within same pinned-group, newest modified first
        pinned   = [n for n in notes if     n["pinned"]]
        unpinned = [n for n in notes if not n["pinned"]]
        pinned.sort(  key=lambda n: n["modified"], reverse=True)
        unpinned.sort(key=lambda n: n["modified"], reverse=True)
        return pinned + unpinned

    @staticmethod
    def load_note(note_id: str) -> dict:
        path = NOTEBOOK_DIR / f"{note_id}.ccnote"
        if path.exists():
            try:
                return json.loads(path.read_text(encoding="utf-8"))
            except Exception:
                pass
        return {"title": "Untitled", "content": "", "modified": "", "color": "", "pinned": False}

    @staticmethod
    def save_note(note_id: str, title: str, html: str) -> None:
        """Save note content/title, preserving any existing metadata (color, pinned, etc.)."""
        path = NOTEBOOK_DIR / f"{note_id}.ccnote"
        # Preserve existing metadata fields
        existing: dict = {}
        if path.exists():
            try:
                existing = json.loads(path.read_text(encoding="utf-8"))
            except Exception:
                pass
        data = {
            **existing,
            "title":    title,
            "content":  html,
            "modified": datetime.now().isoformat(timespec="seconds"),
        }
        try:
            NOTEBOOK_DIR.mkdir(parents=True, exist_ok=True)
            path.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                            encoding="utf-8")
        except OSError as exc:
            print(f"[CommandCenter] Note save error ({note_id}): {exc}", file=sys.stderr)

    @staticmethod
    def update_note_meta(note_id: str, **kwargs) -> None:
        """Update metadata fields (color, pinned, …) without touching content/modified."""
        path = NOTEBOOK_DIR / f"{note_id}.ccnote"
        data: dict = {}
        if path.exists():
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
            except Exception:
                pass
        data.update(kwargs)
        try:
            NOTEBOOK_DIR.mkdir(parents=True, exist_ok=True)
            path.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                            encoding="utf-8")
        except OSError as exc:
            print(f"[CommandCenter] Note meta save error ({note_id}): {exc}", file=sys.stderr)

    @staticmethod
    def delete_note(note_id: str) -> None:
        path = NOTEBOOK_DIR / f"{note_id}.ccnote"
        if path.exists():
            try:
                path.unlink()
            except OSError as exc:
                print(f"[CommandCenter] Note delete error ({note_id}): {exc}", file=sys.stderr)

    @staticmethod
    def new_id() -> str:
        return datetime.now().strftime("%Y%m%d_%H%M%S_%f")


# ---------------------------------------------------------------------------
# Spell-check engine (optional — degrades gracefully if no library found)
# ---------------------------------------------------------------------------

import weakref as _weakref
import re     as _re_spell


class _SpellBackend:
    """Unified wrapper around enchant or pyspellchecker.

    Provides a consistent interface regardless of which back-end is installed:
      available()  → bool
      check(word)  → bool  (True = correctly spelled or in ignore list)
      suggest(word)→ list[str]  (up to 8 suggestions)
      ignore(word) → adds to session-level ignore set and rehighlights all
                     registered highlighters
    """

    _WORD_RE = _re_spell.compile(r"\b[A-Za-z']+\b")

    def __init__(self):
        self._checker   = None
        self._backend   = None          # "enchant" | "pyspellchecker"
        self._ignored: set = set()      # session-level ignore list (lowercase)
        self._highlighters: list = []   # weakrefs to active _SpellCheckHighlighter instances
        self._init()

    # ── initialisation ───────────────────────────────────────────────────────

    def _init(self):
        # Priority 1: pyenchant
        try:
            import enchant as _enc
            self._checker = _enc.Dict("en_US")
            self._backend = "enchant"
            return
        except Exception:
            pass
        # Priority 2: pyspellchecker
        try:
            from spellchecker import SpellChecker as _SC
            self._checker = _SC()
            self._backend = "pyspellchecker"
        except Exception:
            pass

    # ── public API ───────────────────────────────────────────────────────────

    def available(self) -> bool:
        return self._checker is not None

    def check(self, word: str) -> bool:
        """Return True if the word is correctly spelled (or ignored/unavailable)."""
        if not word:
            return True
        if word.lower() in self._ignored:
            return True
        if self._checker is None:
            return True
        try:
            if self._backend == "enchant":
                return bool(self._checker.check(word))
            else:
                return not bool(self._checker.unknown([word]))
        except Exception:
            return True

    def suggest(self, word: str) -> list:
        """Return up to 8 spelling suggestions, ordered by likelihood."""
        if self._checker is None or not word:
            return []
        try:
            if self._backend == "enchant":
                return list(self._checker.suggest(word))[:8]
            else:
                candidates = self._checker.candidates(word)
                return sorted(candidates)[:8] if candidates else []
        except Exception:
            return []

    def ignore(self, word: str):
        """Add *word* to the session ignore list and rehighlight all documents."""
        self._ignored.add(word.lower())
        dead = []
        for ref in self._highlighters:
            h = ref()
            if h is None:
                dead.append(ref)
                continue
            try:
                h.rehighlight()
            except Exception:
                pass
        for ref in dead:
            self._highlighters.remove(ref)

    def register_highlighter(self, highlighter):
        """Track a highlighter so ignore() can trigger rehighlight on it."""
        self._highlighters.append(_weakref.ref(highlighter))


_spell_backend = _SpellBackend()


class _SpellCheckHighlighter(QSyntaxHighlighter):
    """Highlights misspelled words with a red spell-check underline.

    Enabled/disabled via set_enabled(). When disabled the highlighter is
    still attached to the document but does nothing, so re-enabling it is
    instant (just call rehighlight()).
    """

    _WORD_RE = _re_spell.compile(r"\b[A-Za-z']+\b")

    def __init__(self, document):
        super().__init__(document)
        self._enabled = False
        self._fmt = QTextCharFormat()
        self._fmt.setUnderlineColor(QColor(220, 50, 50))
        self._fmt.setUnderlineStyle(QTextCharFormat.UnderlineStyle.SpellCheckUnderline)
        _spell_backend.register_highlighter(self)

    def set_enabled(self, enabled: bool):
        self._enabled = enabled and _spell_backend.available()
        self.rehighlight()

    def is_available(self) -> bool:
        return _spell_backend.available()

    def highlightBlock(self, text: str):
        if not self._enabled or not _spell_backend.available():
            return
        for m in self._WORD_RE.finditer(text):
            word = m.group()
            # Skip short words and ALL-CAPS acronyms
            if len(word) <= 2 or word.isupper():
                continue
            # Strip possessives ("word's" → check "word")
            if word.endswith("'s"):
                stripped = word[:-2]
            elif word.endswith("'"):
                stripped = word[:-1]
            else:
                stripped = word
            try:
                if not _spell_backend.check(stripped):
                    self.setFormat(m.start(), m.end() - m.start(), self._fmt)
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Note color presets
# ---------------------------------------------------------------------------

_NOTE_COLOR_PRESETS: list[tuple[str, str]] = [
    ("Sky Blue",  "#2196F3"),
    ("Teal",      "#00BCD4"),
    ("Green",     "#4CAF50"),
    ("Amber",     "#FFC107"),
    ("Orange",    "#FF9800"),
    ("Red",       "#F44336"),
    ("Pink",      "#E91E63"),
    ("Purple",    "#9C27B0"),
    ("Indigo",    "#3F51B5"),
    ("Brown",     "#795548"),
]


class _NoteListItem(QWidget):
    """Custom list row: title + modified date, optional color strip + pin badge."""
    def __init__(self, note_id: str, title: str, modified: str,
                 color: str = "", pinned: bool = False, parent=None):
        super().__init__(parent)
        self.note_id  = note_id
        self._hovered = False
        self._color: Optional[QColor] = QColor(color) if color and QColor(color).isValid() else None
        self._pinned = pinned
        self.setFixedHeight(52)
        self.setCursor(Qt.PointingHandCursor)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 6, 12, 6)   # extra left margin for color strip
        lay.setSpacing(2)
        self._title_lbl = QLabel(title)
        self._title_lbl.setFont(QFont("Segoe UI", 9, QFont.Weight.DemiBold))
        self._title_lbl.setStyleSheet(f"color:{_theme.TEXT_PRIMARY.name()}; background:transparent;")
        mod_str = ""
        if modified:
            try:
                dt = datetime.fromisoformat(modified)
                mod_str = dt.strftime("%b %d, %Y  %H:%M")
            except Exception:
                mod_str = modified
        self._mod_lbl = QLabel(mod_str)
        self._mod_lbl.setFont(QFont("Segoe UI", 7))
        self._mod_lbl.setStyleSheet(f"color:{_theme.TEXT_DIM.name()}; background:transparent;")
        lay.addWidget(self._title_lbl)
        lay.addWidget(self._mod_lbl)
        self.setAttribute(Qt.WA_StyledBackground, True)

    def set_title(self, title: str):
        self._title_lbl.setText(title)

    def set_modified(self, modified: str):
        try:
            dt = datetime.fromisoformat(modified)
            self._mod_lbl.setText(dt.strftime("%b %d, %Y  %H:%M"))
        except Exception:
            self._mod_lbl.setText(modified)

    def set_color(self, color_hex: str):
        """Set the left-border accent color. Empty string clears it."""
        self._color = QColor(color_hex) if color_hex and QColor(color_hex).isValid() else None
        self.update()

    def set_pinned(self, pinned: bool):
        self._pinned = pinned
        self.update()

    def enterEvent(self, e): self._hovered = True;  self.update()
    def leaveEvent(self, e): self._hovered = False; self.update()

    def paintEvent(self, e):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        r = QRectF(2, 2, self.width()-4, self.height()-4)
        path = QPainterPath(); path.addRoundedRect(r, 6, 6)
        if self._hovered:
            p.fillPath(path, QBrush(QColor(
                t.TILE_BG_HOVER.red(), t.TILE_BG_HOVER.green(), t.TILE_BG_HOVER.blue(), 180)))

        # Left color strip (4 px wide, rounded left side)
        if self._color is not None:
            strip_c = QColor(self._color)
            strip_c.setAlpha(220)
            p.setPen(Qt.NoPen)
            p.setBrush(strip_c)
            p.drawRoundedRect(QRectF(2, 4, 4, self.height() - 8), 2, 2)

        # Pin badge (top-right corner)
        if self._pinned:
            p.setPen(Qt.NoPen)
            badge = QColor(t.GLOW); badge.setAlpha(180)
            p.setBrush(badge)
            p.drawEllipse(QRectF(self.width() - 14, 4, 8, 8))

        p.end()
        super().paintEvent(e)


class _NoteSidePanel(QWidget):
    """Left panel: note list + toolbar."""
    note_selected = Signal(str)   # note_id
    note_new      = Signal()
    note_delete   = Signal(str)   # note_id

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumWidth(200)
        self.setMaximumWidth(300)
        self._selected_id: Optional[str] = None
        self._items: list[_NoteListItem] = []
        self._build_ui()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # Header row
        hdr = QWidget()
        hdr.setFixedHeight(44)
        hdr.setStyleSheet("background:transparent;")
        hl = QHBoxLayout(hdr)
        hl.setContentsMargins(12, 0, 8, 0)
        lbl = QLabel("Notes")
        lbl.setFont(QFont("Segoe UI", 10, QFont.Weight.DemiBold))
        lbl.setStyleSheet(f"color:{_theme.TEXT_PRIMARY.name()};")
        hl.addWidget(lbl); hl.addStretch()

        self._new_btn = QPushButton("+")
        self._new_btn.setFixedSize(26, 26)
        self._new_btn.setCursor(Qt.PointingHandCursor)
        self._new_btn.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        self._new_btn.setToolTip("New Note")
        t = _theme
        self._new_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},120);
                border-radius:13px; color:{t.GLOW.name()};
            }}
            QPushButton:hover {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},100);
            }}
        """)
        self._new_btn.clicked.connect(self.note_new)
        hl.addWidget(self._new_btn)
        lay.addWidget(hdr)

        # Separator
        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"color:rgba({_theme.GLOW.red()},{_theme.GLOW.green()},{_theme.GLOW.blue()},40);")
        lay.addWidget(sep)

        # Scrollable note list
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        t = _theme
        self._scroll.setStyleSheet(f"""
            QScrollArea {{ background:transparent; border:none; }}
            QScrollBar:vertical {{
                background:rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},120);
                width:5px; border-radius:2px;
            }}
            QScrollBar::handle:vertical {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},100);
                border-radius:2px; min-height:16px;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height:0; }}
        """)
        self._list_container = QWidget()
        self._list_container.setStyleSheet("background:transparent;")
        self._list_layout = QVBoxLayout(self._list_container)
        self._list_layout.setContentsMargins(4, 4, 4, 4)
        self._list_layout.setSpacing(2)
        self._list_layout.addStretch()
        self._scroll.setWidget(self._list_container)
        lay.addWidget(self._scroll)

    def load_notes(self, notes: list[dict]):
        # Clear
        for item in self._items:
            item.deleteLater()
        self._items.clear()
        # Remove stretch
        item_c = self._list_layout.count()
        if item_c > 0:
            stretch = self._list_layout.itemAt(item_c - 1)
            if stretch and stretch.spacerItem():
                self._list_layout.removeItem(stretch)
        for note in notes:
            self._add_item(note["id"], note["title"], note["modified"],
                           note.get("color", ""), bool(note.get("pinned", False)))
        self._list_layout.addStretch()

    def _add_item(self, note_id: str, title: str, modified: str,
                  color: str = "", pinned: bool = False):
        item = _NoteListItem(note_id, title, modified, color, pinned,
                             self._list_container)
        item.mousePressEvent = lambda e, nid=note_id: self._on_item_clicked(nid, e)
        # Right-click context menu
        item.setContextMenuPolicy(Qt.CustomContextMenu)
        item.customContextMenuRequested.connect(
            lambda pos, nid=note_id: self._on_item_ctx(nid, pos, item))
        self._list_layout.insertWidget(self._list_layout.count() - 1, item)
        self._items.append(item)
        return item

    def add_note(self, note_id: str, title: str, modified: str,
                 color: str = "", pinned: bool = False):
        item = self._add_item(note_id, title, modified, color, pinned)
        self._select_item(note_id)
        return item

    def update_item(self, note_id: str, title: str, modified: str):
        for item in self._items:
            if item.note_id == note_id:
                item.set_title(title)
                item.set_modified(modified)
                item.update()
                break

    def remove_item(self, note_id: str):
        for item in self._items[:]:
            if item.note_id == note_id:
                self._items.remove(item)
                item.deleteLater()
                break

    def _on_item_clicked(self, note_id: str, e: QMouseEvent):
        if e.button() == Qt.LeftButton:
            self._select_item(note_id)
            self.note_selected.emit(note_id)

    def _on_item_ctx(self, note_id: str, pos, item: _NoteListItem):
        t = _theme
        menu = QMenu(self)
        menu.setStyleSheet(f"""
            QMenu {{
                background:rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},245);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},80);
                border-radius:8px; padding:4px 0;
                color:{t.TEXT_PRIMARY.name()};
                font-family:'Segoe UI'; font-size:9pt;
            }}
            QMenu::item {{ padding:6px 18px; border-radius:4px; }}
            QMenu::item:selected {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);
                color:{t.GLOW.name()};
            }}
            QMenu::separator {{ height:1px;
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},40);
                margin:3px 8px;
            }}
        """)
        # ── Pin ──────────────────────────────────────────────────────────────
        is_pinned = item._pinned
        act_pin = menu.addAction("📌  Unpin from Top" if is_pinned else "📌  Pin to Top")
        menu.addSeparator()

        # ── Color submenu ────────────────────────────────────────────────────
        color_menu = menu.addMenu("🎨  Set Color")
        color_menu.setStyleSheet(menu.styleSheet())
        color_actions = {}
        for cname, chex in _NOTE_COLOR_PRESETS:
            act = color_menu.addAction(f"● {cname}")
            act.setData(chex)
            # Colorize the bullet with the preset color using rich text trick via
            # stylesheet on the action icon area — we use the action data instead
            color_actions[act] = chex
        color_menu.addSeparator()
        act_custom_color = color_menu.addAction("✏  Custom Color…")
        act_clear_color  = color_menu.addAction("✕  Clear Color")
        menu.addSeparator()

        # ── Delete ───────────────────────────────────────────────────────────
        act_del = menu.addAction("🗑  Delete Note")

        chosen = menu.exec(item.mapToGlobal(pos))
        if chosen is None:
            return

        if chosen == act_pin:
            new_pinned = not is_pinned
            NotebookStore.update_note_meta(note_id, pinned=new_pinned)
            item.set_pinned(new_pinned)
            # Re-sort the list to float pinned items to top
            self._resort_items()

        elif chosen in color_actions:
            chex = color_actions[chosen]
            NotebookStore.update_note_meta(note_id, color=chex)
            item.set_color(chex)

        elif chosen == act_custom_color:
            current_hex = ""
            if item._color is not None:
                current_hex = item._color.name()
            dlg = QColorDialog(QColor(current_hex) if current_hex else t.GLOW, self)
            dlg.setWindowTitle("Choose Note Color")
            if dlg.exec() == QDialog.Accepted:
                chosen_color = dlg.currentColor()
                chex = chosen_color.name()
                NotebookStore.update_note_meta(note_id, color=chex)
                item.set_color(chex)

        elif chosen == act_clear_color:
            NotebookStore.update_note_meta(note_id, color="")
            item.set_color("")

        elif chosen == act_del:
            self.note_delete.emit(note_id)

    def _resort_items(self):
        """Re-read all notes from disk and reload the list in sorted order,
        preserving the current selection."""
        selected = self._selected_id
        notes = NotebookStore.list_notes()
        self.load_notes(notes)
        if selected:
            self._select_item(selected)

    def _select_item(self, note_id: str):
        self._selected_id = note_id
        t = _theme
        for item in self._items:
            is_sel = item.note_id == note_id
            item._title_lbl.setStyleSheet(
                f"color:{t.GLOW.name() if is_sel else t.TEXT_PRIMARY.name()};"
                f" background:transparent;")
            item.setStyleSheet(
                f"background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},25);"
                f" border-radius:6px;"
                if is_sel else "background:transparent;")

    def selected_id(self) -> Optional[str]:
        return self._selected_id

    def paintEvent(self, e):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        r = QRectF(0, 0, self.width(), self.height())
        p.fillRect(r, QBrush(QColor(
            t.BG_DARK.red(), t.BG_DARK.green(), t.BG_DARK.blue(), 220)))
        # Right border glow line
        g = QLinearGradient(0, 0, 0, self.height())
        g.setColorAt(0, QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 0))
        g.setColorAt(0.3, QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 50))
        g.setColorAt(0.7, QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 50))
        g.setColorAt(1, QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 0))
        p.setPen(QPen(QBrush(g), 1))
        p.drawLine(self.width()-1, 0, self.width()-1, self.height())
        p.end()


class _NoteEditor(QWidget):
    """Right panel: title bar + formatting toolbar + QTextEdit."""
    content_changed = Signal()   # fired on every edit

    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_id: Optional[str] = None
        self._block_signals = False
        # Spell-check toggle state — persist between sessions
        self._spell_check_on: bool = (
            _settings_store.value("notebook_spellcheck", "false") == "true")
        self._build_ui()

    def _build_ui(self):
        from PySide6.QtWidgets import QTextBrowser
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # Title input row
        title_bar = QWidget(); title_bar.setFixedHeight(48)
        title_bar.setStyleSheet("background:transparent;")
        tbl = QHBoxLayout(title_bar)
        tbl.setContentsMargins(20, 8, 20, 8)
        t = _theme
        self._title_edit = QLineEdit()
        self._title_edit.setPlaceholderText("Note title…")
        self._title_edit.setFont(QFont("Segoe UI", 12, QFont.Weight.DemiBold))
        self._title_edit.setStyleSheet(f"""
            QLineEdit {{
                background:transparent; border:none; border-bottom:1px solid
                    rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
                color:{t.TEXT_PRIMARY.name()};
                padding:2px 4px;
            }}
            QLineEdit:focus {{ border-bottom:1px solid
                rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},160); }}
        """)
        self._title_edit.textChanged.connect(self._on_title_changed)
        tbl.addWidget(self._title_edit)
        lay.addWidget(title_bar)

        # Formatting toolbar
        fbar = QWidget(); fbar.setFixedHeight(36)
        fbar.setStyleSheet("background:transparent;")
        fbl = QHBoxLayout(fbar)
        fbl.setContentsMargins(14, 4, 14, 4)
        fbl.setSpacing(2)

        def _fmt_btn(lbl: str, tip: str) -> QPushButton:
            b = QPushButton(lbl)
            b.setFixedSize(28, 26)
            b.setCursor(Qt.PointingHandCursor)
            b.setFont(QFont("Segoe UI", 8, QFont.Weight.DemiBold))
            b.setToolTip(tip)
            t = _theme
            b.setStyleSheet(f"""
                QPushButton {{
                    background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},160);
                    border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);
                    border-radius:4px; color:{t.TEXT_SECONDARY.name()};
                }}
                QPushButton:hover {{
                    background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
                    color:{t.GLOW.name()};
                    border-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},130);
                }}
                QPushButton:pressed {{
                    background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},90);
                }}
            """)
            return b

        def _fmt_sep():
            s = QFrame(); s.setFrameShape(QFrame.VLine)
            s.setFixedHeight(20)
            s.setStyleSheet(f"color:rgba({_theme.GLOW.red()},{_theme.GLOW.green()},{_theme.GLOW.blue()},35);")
            return s

        self._btn_bold      = _fmt_btn("B",  "Bold (Ctrl+B)")
        self._btn_italic    = _fmt_btn("I",  "Italic (Ctrl+I)")
        self._btn_underline = _fmt_btn("U",  "Underline (Ctrl+U)")
        self._btn_strike    = _fmt_btn("S̶",  "Strikethrough")
        self._btn_bold.setFont(QFont("Segoe UI", 8, QFont.Weight.ExtraBold))
        self._btn_italic.setFont(QFont("Segoe UI", 8, QFont.Weight.Normal))
        font_i = self._btn_italic.font(); font_i.setItalic(True); self._btn_italic.setFont(font_i)

        fbl.addWidget(self._btn_bold)
        fbl.addWidget(self._btn_italic)
        fbl.addWidget(self._btn_underline)
        fbl.addWidget(self._btn_strike)
        fbl.addSpacing(4); fbl.addWidget(_fmt_sep()); fbl.addSpacing(4)

        self._btn_h1 = _fmt_btn("H1", "Heading 1")
        self._btn_h2 = _fmt_btn("H2", "Heading 2")
        self._btn_h3 = _fmt_btn("H3", "Heading 3")
        fbl.addWidget(self._btn_h1)
        fbl.addWidget(self._btn_h2)
        fbl.addWidget(self._btn_h3)
        fbl.addSpacing(4); fbl.addWidget(_fmt_sep()); fbl.addSpacing(4)

        self._btn_ul = _fmt_btn("•—", "Bulleted List")
        self._btn_ol = _fmt_btn("1.", "Numbered List")
        fbl.addWidget(self._btn_ul)
        fbl.addWidget(self._btn_ol)
        fbl.addSpacing(4); fbl.addWidget(_fmt_sep()); fbl.addSpacing(4)

        self._btn_link  = _fmt_btn("🔗", "Insert Link")
        self._btn_img   = _fmt_btn("🖼", "Insert Image")
        self._btn_code  = _fmt_btn("</>", "Inline Code")
        self._btn_quote = _fmt_btn("❝",  "Block Quote")
        self._btn_hr    = _fmt_btn("—",  "Horizontal Rule")
        fbl.addWidget(self._btn_link)
        fbl.addWidget(self._btn_img)
        fbl.addWidget(self._btn_code)
        fbl.addWidget(self._btn_quote)
        fbl.addWidget(self._btn_hr)
        fbl.addSpacing(4); fbl.addWidget(_fmt_sep()); fbl.addSpacing(4)

        self._btn_color = _fmt_btn("A", "Text Color")
        self._btn_color.setStyleSheet(
            self._btn_color.styleSheet() +
            f"\nQPushButton {{ color:{_theme.ACCENT_AMBER.name()}; }}")
        fbl.addWidget(self._btn_color)

        fbl.addStretch()

        # Font size combo
        self._size_combo = QComboBox()
        for s in ["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "32", "36", "48"]:
            self._size_combo.addItem(s)
        self._size_combo.setCurrentText("10")
        self._size_combo.setFixedWidth(52)
        self._size_combo.setFixedHeight(26)
        t = _theme
        self._size_combo.setStyleSheet(f"""
            QComboBox {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},160);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);
                border-radius:4px; color:{t.TEXT_SECONDARY.name()};
                padding:0 4px; font-size:8pt;
            }}
            QComboBox::drop-down {{ border:none; }}
            QComboBox QAbstractItemView {{
                background:rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},240);
                color:{t.TEXT_PRIMARY.name()};
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
                selection-background-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
            }}
        """)
        self._size_combo.currentTextChanged.connect(self._on_font_size)
        fbl.addWidget(self._size_combo)

        # Spellcheck toggle button
        fbl.addSpacing(4)
        self._btn_spell = QPushButton("ABC✓")
        self._btn_spell.setFixedSize(46, 26)
        self._btn_spell.setCursor(Qt.PointingHandCursor)
        self._btn_spell.setFont(QFont("Segoe UI", 7, QFont.Weight.DemiBold))
        self._btn_spell.setCheckable(True)
        self._btn_spell.setChecked(self._spell_check_on)
        if not _spell_backend.available():
            self._btn_spell.setToolTip(
                "Spellcheck unavailable.\n"
                "Install a spell-check library:\n"
                "  pip install pyspellchecker\n"
                "or:  pip install pyenchant")
            self._btn_spell.setEnabled(False)
        else:
            self._btn_spell.setToolTip("Toggle Spell Check")
        self._btn_spell.clicked.connect(self._on_spell_toggle)
        fbl.addWidget(self._btn_spell)
        lay.addWidget(fbar)

        # Separator
        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(
            f"color:rgba({_theme.GLOW.red()},{_theme.GLOW.green()},{_theme.GLOW.blue()},30);")
        lay.addWidget(sep)

        # The actual editor
        self._editor = _RichTextEditor()
        self._editor.setStyleSheet(f"""
            QTextEdit {{
                background:transparent; border:none;
                color:{_theme.TEXT_PRIMARY.name()};
                font-family:'Segoe UI'; font-size:10pt;
                padding:16px 20px;
                selection-background-color:rgba({_theme.GLOW.red()},{_theme.GLOW.green()},{_theme.GLOW.blue()},60);
            }}
            QScrollBar:vertical {{
                background:rgba({_theme.BG_DARK.red()},{_theme.BG_DARK.green()},{_theme.BG_DARK.blue()},120);
                width:6px; border-radius:3px;
            }}
            QScrollBar::handle:vertical {{
                background:rgba({_theme.GLOW.red()},{_theme.GLOW.green()},{_theme.GLOW.blue()},100);
                border-radius:3px; min-height:20px;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height:0; }}
        """)
        self._editor.document().contentsChanged.connect(self._on_content_changed)
        self._editor.paste_plain_status.connect(self._on_paste_plain_status)
        lay.addWidget(self._editor)

        # Connect toolbar buttons
        self._btn_bold.clicked.connect(self._fmt_bold)
        self._btn_italic.clicked.connect(self._fmt_italic)
        self._btn_underline.clicked.connect(self._fmt_underline)
        self._btn_strike.clicked.connect(self._fmt_strike)
        self._btn_h1.clicked.connect(lambda: self._fmt_heading(1))
        self._btn_h2.clicked.connect(lambda: self._fmt_heading(2))
        self._btn_h3.clicked.connect(lambda: self._fmt_heading(3))
        self._btn_ul.clicked.connect(self._fmt_ul)
        self._btn_ol.clicked.connect(self._fmt_ol)
        self._btn_link.clicked.connect(self._fmt_link)
        self._btn_img.clicked.connect(self._fmt_img)
        self._btn_code.clicked.connect(self._fmt_code)
        self._btn_quote.clicked.connect(self._fmt_quote)
        self._btn_hr.clicked.connect(self._fmt_hr)
        self._btn_color.clicked.connect(self._fmt_color)

        # Status bar
        self._status_lbl = QLabel("")
        self._status_lbl.setFont(QFont("Segoe UI", 7))
        self._status_lbl.setStyleSheet(
            f"color:{_theme.TEXT_DIM.name()}; padding:2px 20px;")
        self._status_lbl.setAlignment(Qt.AlignRight)
        self._status_lbl.setFixedHeight(18)
        lay.addWidget(self._status_lbl)

    # ── load / clear ─────────────────────────────────────────────────────────

    def load_note(self, note_id: str, data: dict):
        self._block_signals = True
        self._current_id = note_id
        self._title_edit.setText(data.get("title", ""))
        html = data.get("content", "")
        if html:
            self._editor.setHtml(html)
        else:
            self._editor.clear()
        self._block_signals = False
        self._editor.setFocus()
        # Apply current spell-check state to the freshly loaded document
        self._editor.set_spell_check(self._spell_check_on)
        self._update_spell_btn_style()
        self._update_status()

    # ── spell check ──────────────────────────────────────────────────────────

    def _on_spell_toggle(self):
        self._spell_check_on = self._btn_spell.isChecked()
        _settings_store.setValue("notebook_spellcheck",
                                 "true" if self._spell_check_on else "false")
        self._editor.set_spell_check(self._spell_check_on)
        self._update_spell_btn_style()

    def _update_spell_btn_style(self):
        t = _theme
        on = self._spell_check_on and (self._btn_spell.isEnabled())
        if on:
            self._btn_spell.setStyleSheet(f"""
                QPushButton {{
                    background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},80);
                    border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},160);
                    border-radius:4px; color:{t.GLOW.name()};
                }}
                QPushButton:hover {{
                    background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},110);
                }}
            """)
        else:
            self._btn_spell.setStyleSheet(f"""
                QPushButton {{
                    background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},160);
                    border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);
                    border-radius:4px; color:{t.TEXT_DIM.name()};
                }}
                QPushButton:hover {{
                    background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},40);
                    color:{t.TEXT_SECONDARY.name()};
                }}
                QPushButton:disabled {{
                    color:rgba({t.TEXT_DIM.red()},{t.TEXT_DIM.green()},{t.TEXT_DIM.blue()},80);
                    border-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},25);
                }}
            """)

    def clear(self):
        self._block_signals = True
        self._current_id = None
        self._title_edit.clear()
        self._editor.clear()
        self._block_signals = False

    def current_id(self) -> Optional[str]:
        return self._current_id

    def current_title(self) -> str:
        return self._title_edit.text().strip() or "Untitled"

    def current_html(self) -> str:
        return self._editor.toHtml()

    # ── callbacks ─────────────────────────────────────────────────────────────

    def _on_title_changed(self):
        if not self._block_signals and self._current_id:
            self.content_changed.emit()

    def _on_content_changed(self):
        if not self._block_signals and self._current_id:
            self._update_status()
            self.content_changed.emit()

    def _update_status(self):
        doc = self._editor.document()
        words = len(doc.toPlainText().split())
        chars = len(doc.toPlainText())
        self._status_lbl.setText(f"{words} words  ·  {chars} characters")

    def _on_paste_plain_status(self, msg: str):
        """Show a brief paste-plain error in the status bar, then restore word count."""
        self._status_lbl.setText(msg)
        QTimer.singleShot(3000, self._update_status)

    # ── formatting helpers ────────────────────────────────────────────────────

    def _cur(self):
        return self._editor.textCursor()

    def _fmt_bold(self):
        from PySide6.QtGui import QTextCharFormat
        fmt = QTextCharFormat()
        cur = self._cur()
        bold = cur.charFormat().fontWeight() != QFont.Bold
        fmt.setFontWeight(QFont.Bold if bold else QFont.Normal)
        cur.mergeCharFormat(fmt)
        self._editor.setTextCursor(cur)

    def _fmt_italic(self):
        from PySide6.QtGui import QTextCharFormat
        cur = self._cur()
        fmt = QTextCharFormat()
        fmt.setFontItalic(not cur.charFormat().fontItalic())
        cur.mergeCharFormat(fmt)
        self._editor.setTextCursor(cur)

    def _fmt_underline(self):
        from PySide6.QtGui import QTextCharFormat
        cur = self._cur()
        fmt = QTextCharFormat()
        fmt.setFontUnderline(not cur.charFormat().fontUnderline())
        cur.mergeCharFormat(fmt)
        self._editor.setTextCursor(cur)

    def _fmt_strike(self):
        from PySide6.QtGui import QTextCharFormat
        cur = self._cur()
        fmt = QTextCharFormat()
        fmt.setFontStrikeOut(not cur.charFormat().fontStrikeOut())
        cur.mergeCharFormat(fmt)
        self._editor.setTextCursor(cur)

    def _fmt_heading(self, level: int):
        from PySide6.QtGui import QTextBlockFormat, QTextCharFormat, QTextCursor
        sizes = {1: 20, 2: 16, 3: 13}
        cur = self._cur()
        bfmt = QTextBlockFormat()
        bfmt.setTopMargin(8); bfmt.setBottomMargin(4)
        cfmt = QTextCharFormat()
        cfmt.setFontWeight(QFont.Bold)
        cfmt.setFontPointSize(sizes[level])
        cur.beginEditBlock()
        cur.select(QTextCursor.SelectionType.BlockUnderCursor)
        cur.setBlockFormat(bfmt)
        cur.mergeCharFormat(cfmt)
        cur.endEditBlock()
        self._editor.setTextCursor(cur)

    def _fmt_ul(self):
        from PySide6.QtGui import QTextListFormat
        cur = self._cur()
        fmt = QTextListFormat()
        fmt.setStyle(QTextListFormat.ListDisc)
        fmt.setIndent(1)
        cur.createList(fmt)
        self._editor.setTextCursor(cur)

    def _fmt_ol(self):
        from PySide6.QtGui import QTextListFormat
        cur = self._cur()
        fmt = QTextListFormat()
        fmt.setStyle(QTextListFormat.ListDecimal)
        fmt.setIndent(1)
        cur.createList(fmt)
        self._editor.setTextCursor(cur)

    def _fmt_link(self):
        from PySide6.QtGui import QTextCharFormat
        cur = self._cur()
        url, ok = QLineEdit(), False
        # Tiny inline dialog
        t = _theme
        dlg = QDialog(self, Qt.FramelessWindowHint)
        dlg.setAttribute(Qt.WA_TranslucentBackground)
        dlg.setFixedSize(360, 130)
        ll = QVBoxLayout(dlg); ll.setContentsMargins(18, 14, 18, 14)
        ll.addWidget(QLabel("URL:", parent=dlg))
        url_edit = QLineEdit(dlg)
        url_edit.setPlaceholderText("https://…")
        url_edit.setStyleSheet(f"""
            QLineEdit {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},80);
                border-radius:5px; color:{t.TEXT_PRIMARY.name()};
                padding:5px 9px;
            }}
        """)
        ll.addWidget(url_edit)
        btn_row = QHBoxLayout(); btn_row.addStretch()
        ok_btn = QPushButton("Insert", dlg)
        ok_btn.setCursor(Qt.PointingHandCursor)
        ok_btn.clicked.connect(dlg.accept)
        btn_row.addWidget(ok_btn)
        ll.addLayout(btn_row)
        if dlg.exec() == QDialog.Accepted and url_edit.text().strip():
            href = url_edit.text().strip()
            display = cur.selectedText() or href
            fmt = QTextCharFormat()
            fmt.setAnchor(True)
            fmt.setAnchorHref(href)
            fmt.setForeground(QBrush(_theme.ACCENT_BLUE))
            fmt.setFontUnderline(True)
            cur.removeSelectedText()
            cur.insertText(display, fmt)
            self._editor.setTextCursor(cur)

    def _fmt_img(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Insert Image", "",
            "Images (*.png *.jpg *.jpeg *.bmp *.gif *.webp);;All Files (*)")
        if not path:
            return
        self._insert_image_file(path)

    def _insert_image_file(self, path: str):
        img = QImage(path)
        if img.isNull():
            return
        w = _pick_image_width(self, img.width())
        if w is None:
            return
        if img.width() > w:
            img = img.scaledToWidth(w, Qt.SmoothTransformation)
        buf = QByteArray()
        from PySide6.QtCore import QBuffer
        buf_io = QBuffer(buf)
        buf_io.open(QBuffer.WriteOnly)
        img.save(buf_io, "PNG")
        buf_io.close()
        b64 = _base64.b64encode(bytes(buf)).decode()
        cur = self._cur()
        cur.insertHtml(f'<img src="data:image/png;base64,{b64}" width="{w}"/>')
        self._editor.setTextCursor(cur)

    def _fmt_code(self):
        from PySide6.QtGui import QTextCharFormat
        cur = self._cur()
        fmt = QTextCharFormat()
        fmt.setFontFamilies(["Consolas", "Courier New", "monospace"])
        fmt.setFontPointSize(9)
        t = _theme
        fmt.setBackground(QBrush(QColor(
            t.TILE_BG_BASE.red(), t.TILE_BG_BASE.green(), t.TILE_BG_BASE.blue(), 200)))
        fmt.setForeground(QBrush(t.ACCENT_TEAL))
        cur.mergeCharFormat(fmt)
        self._editor.setTextCursor(cur)

    def _fmt_quote(self):
        from PySide6.QtGui import QTextBlockFormat, QTextCharFormat, QTextCursor
        cur = self._cur()
        bfmt = QTextBlockFormat()
        bfmt.setLeftMargin(20)
        bfmt.setTopMargin(4); bfmt.setBottomMargin(4)
        t = _theme
        cfmt = QTextCharFormat()
        cfmt.setFontItalic(True)
        cfmt.setForeground(QBrush(t.TEXT_SECONDARY))
        cur.beginEditBlock()
        cur.select(QTextCursor.SelectionType.BlockUnderCursor)
        cur.setBlockFormat(bfmt)
        cur.mergeCharFormat(cfmt)
        cur.endEditBlock()
        self._editor.setTextCursor(cur)

    def _fmt_hr(self):
        cur = self._cur()
        cur.insertHtml("<hr/>")
        self._editor.setTextCursor(cur)

    def _fmt_color(self):
        from PySide6.QtGui import QTextCharFormat
        dlg = QColorDialog(_theme.TEXT_PRIMARY, self)
        if dlg.exec() == QDialog.Accepted:
            color = dlg.currentColor()
            cur = self._cur()
            fmt = QTextCharFormat()
            fmt.setForeground(QBrush(color))
            cur.mergeCharFormat(fmt)
            self._editor.setTextCursor(cur)

    def _on_font_size(self, size_str: str):
        try:
            size = float(size_str)
        except ValueError:
            return
        from PySide6.QtGui import QTextCharFormat
        cur = self._cur()
        fmt = QTextCharFormat()
        fmt.setFontPointSize(size)
        cur.mergeCharFormat(fmt)
        self._editor.setTextCursor(cur)

    def paintEvent(self, e):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.fillRect(self.rect(), QBrush(QColor(
            t.BG_MID.red(), t.BG_MID.green(), t.BG_MID.blue(), 200)))
        p.end()


def _pick_image_width(parent_widget: QWidget, natural_w: int) -> Optional[int]:
    """Show a small themed dialog to pick the display width of a pasted image.
    Returns the chosen width in pixels, or None if the user cancelled."""
    t = _theme
    default_w = min(natural_w, 480)
    dlg = QDialog(parent_widget, Qt.FramelessWindowHint)
    dlg.setAttribute(Qt.WA_TranslucentBackground)
    dlg.setFixedSize(340, 140)
    outer = QWidget(dlg)
    outer.setGeometry(0, 0, 340, 140)
    outer.setStyleSheet(f"""
        QWidget {{
            background:rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},242);
            border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},90);
            border-radius:10px;
        }}
    """)
    ll = QVBoxLayout(outer)
    ll.setContentsMargins(18, 14, 18, 14)
    ll.setSpacing(10)
    hdr = QLabel("Image display width")
    hdr.setFont(QFont("Segoe UI", 9, QFont.Weight.DemiBold))
    hdr.setStyleSheet(f"color:{t.TEXT_PRIMARY.name()}; background:transparent; border:none;")
    ll.addWidget(hdr)
    row = QHBoxLayout(); row.setSpacing(8)
    slider = QSlider(Qt.Horizontal)
    slider.setRange(80, max(natural_w, 800))
    slider.setValue(default_w)
    slider.setStyleSheet(f"""
        QSlider::groove:horizontal {{
            height:4px;
            background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},40);
            border-radius:2px;
        }}
        QSlider::handle:horizontal {{
            width:14px; height:14px; margin:-5px 0;
            background:{t.GLOW.name()}; border-radius:7px;
        }}
        QSlider::sub-page:horizontal {{
            background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},130);
            border-radius:2px;
        }}
    """)
    val_lbl = QLabel(f"{default_w} px")
    val_lbl.setFont(QFont("Segoe UI", 8))
    val_lbl.setFixedWidth(56)
    val_lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
    val_lbl.setStyleSheet(f"color:{t.TEXT_SECONDARY.name()}; background:transparent; border:none;")
    slider.valueChanged.connect(lambda v: val_lbl.setText(f"{v} px"))
    row.addWidget(slider); row.addWidget(val_lbl)
    ll.addLayout(row)
    btn_row = QHBoxLayout(); btn_row.addStretch()
    ok_btn  = QPushButton("Insert", outer)
    can_btn = QPushButton("Cancel", outer)
    for b in (ok_btn, can_btn):
        b.setFixedHeight(26)
        b.setCursor(Qt.PointingHandCursor)
        b.setStyleSheet(f"""
            QPushButton {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},180);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:5px; color:{t.TEXT_SECONDARY.name()};
                padding:0 12px; font-size:8pt;
            }}
            QPushButton:hover {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                color:{t.GLOW.name()};
            }}
        """)
    ok_btn.clicked.connect(dlg.accept)
    can_btn.clicked.connect(dlg.reject)
    btn_row.addWidget(can_btn); btn_row.addSpacing(6); btn_row.addWidget(ok_btn)
    ll.addLayout(btn_row)
    if dlg.exec() == QDialog.Accepted:
        return slider.value()
    return None


class _RichTextEditor(QTextEdit):
    """QTextEdit subclass that supports paste-image from clipboard."""
    paste_plain_status = Signal(str)   # emitted when paste-plain cannot proceed

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptRichText(True)
        self.setAcceptDrops(True)
        self.setLineWrapMode(QTextEdit.WidgetWidth)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setPlaceholderText("Start writing your note here…")
        # Spell-check highlighter (always attached; enabled/disabled on demand)
        self._spell_highlighter = _SpellCheckHighlighter(self.document())

    def set_spell_check(self, enabled: bool):
        """Enable or disable spell-check underlines."""
        self._spell_highlighter.set_enabled(enabled)

    def _embed_image(self, img: QImage) -> None:
        """Show size picker, then embed img as a base64 data-URI at chosen width."""
        w = _pick_image_width(self, img.width())
        if w is None:
            return
        if img.width() > w:
            img = img.scaledToWidth(w, Qt.SmoothTransformation)
        buf = QByteArray()
        from PySide6.QtCore import QBuffer
        buf_io = QBuffer(buf)
        buf_io.open(QBuffer.WriteOnly)
        img.save(buf_io, "PNG")
        buf_io.close()
        b64 = _base64.b64encode(bytes(buf)).decode()
        cur = self.textCursor()
        cur.insertHtml(f'<img src="data:image/png;base64,{b64}" width="{w}"/>')
        self.setTextCursor(cur)

    def insertFromMimeData(self, source: QMimeData):
        if source.hasImage():
            img = source.imageData()
            if isinstance(img, QImage) and not img.isNull():
                self._embed_image(img)
                return
        # Fall through for image file URLs dragged in
        if source.hasUrls():
            for url in source.urls():
                lf = url.toLocalFile()
                if lf and Path(lf).suffix.lower() in {".png",".jpg",".jpeg",".bmp",".gif",".webp"}:
                    img = QImage(lf)
                    if not img.isNull():
                        self._embed_image(img)
                        return
        super().insertFromMimeData(source)

    def dragEnterEvent(self, e: QDragEnterEvent):
        if e.mimeData().hasImage() or e.mimeData().hasUrls():
            e.acceptProposedAction()
        else:
            super().dragEnterEvent(e)

    def dropEvent(self, e: QDropEvent):
        self.insertFromMimeData(e.mimeData())
        e.acceptProposedAction()

    def keyPressEvent(self, e):
        # Ctrl+B/I/U shortcuts; Ctrl+Shift+V = paste as plain text
        if e.modifiers() & Qt.ControlModifier:
            if e.modifiers() & Qt.ShiftModifier:
                if e.key() == Qt.Key_V:
                    self._paste_plain_text(); return
            else:
                if e.key() == Qt.Key_B:
                    self._toggle_bold(); return
                if e.key() == Qt.Key_I:
                    self._toggle_italic(); return
                if e.key() == Qt.Key_U:
                    self._toggle_underline(); return
        super().keyPressEvent(e)

    # ── paste as plain text ───────────────────────────────────────────────

    def _can_paste_plain(self) -> bool:
        """Return True if the clipboard has something that can become plain text."""
        mime = QApplication.clipboard().mimeData()
        if mime is None:
            return False
        return mime.hasText() or mime.hasHtml()

    def _paste_plain_text(self):
        """Paste clipboard contents as unformatted plain text.

        Handles all common clipboard types:
          - plain text / Excel TSV  → inserted verbatim
          - HTML (web pages, Word)  → converted to plain text via QTextDocument
          - image only              → rejected with status message
          - file URLs only          → rejected with status message
          - empty / unrecognised    → silent no-op
        """
        if self.isReadOnly():
            return
        mime = QApplication.clipboard().mimeData()
        if mime is None or not mime.formats():
            return

        has_text = mime.hasText()
        has_html = mime.hasHtml()
        has_img  = mime.hasImage()
        has_urls = mime.hasUrls()

        # Images with no text representation
        if has_img and not has_text and not has_html:
            self._notify_paste_status("⚠  Images cannot be pasted as plain text.")
            return

        # File drops with no text representation
        if has_urls and not has_text and not has_html:
            self._notify_paste_status("⚠  Files cannot be pasted as plain text.")
            return

        # text/plain covers Excel TSV, terminal output, plain copies, etc.
        if has_text:
            text = mime.text()
            if text:
                cur = self.textCursor()
                cur.insertText(text)
                self.setTextCursor(cur)
                return

        # HTML with no text/plain companion — strip tags via QTextDocument
        if has_html:
            try:
                from PySide6.QtGui import QTextDocument as _QTD
                _doc = _QTD()
                _doc.setHtml(mime.html())
                text = _doc.toPlainText()
                if text:
                    cur = self.textCursor()
                    cur.insertText(text)
                    self.setTextCursor(cur)
                    return
            except Exception:
                pass

        self._notify_paste_status("⚠  Nothing on the clipboard can be pasted as plain text.")

    def _notify_paste_status(self, msg: str):
        """Emit a status message for the parent editor to display briefly."""
        self.paste_plain_status.emit(msg)

    def _toggle_bold(self):
        from PySide6.QtGui import QTextCharFormat
        cur = self.textCursor()
        fmt = QTextCharFormat()
        fmt.setFontWeight(
            QFont.Normal if cur.charFormat().fontWeight() == QFont.Bold else QFont.Bold)
        cur.mergeCharFormat(fmt)
        self.setTextCursor(cur)

    def _toggle_italic(self):
        from PySide6.QtGui import QTextCharFormat
        cur = self.textCursor()
        fmt = QTextCharFormat()
        fmt.setFontItalic(not cur.charFormat().fontItalic())
        cur.mergeCharFormat(fmt)
        self.setTextCursor(cur)

    def _toggle_underline(self):
        from PySide6.QtGui import QTextCharFormat
        cur = self.textCursor()
        fmt = QTextCharFormat()
        fmt.setFontUnderline(not cur.charFormat().fontUnderline())
        cur.mergeCharFormat(fmt)
        self.setTextCursor(cur)

    def contextMenuEvent(self, e):
        """Right-click menu with spell-check suggestions when applicable."""
        # ── Determine word under click ────────────────────────────────────
        word_cursor = self.cursorForPosition(e.pos())
        word_cursor.select(QTextCursor.SelectionType.WordUnderCursor)
        raw_word  = word_cursor.selectedText()
        # Keep only letters and apostrophes
        clean_word  = _re_spell.sub(r"[^A-Za-z']", "", raw_word)
        check_word  = clean_word.rstrip("'s").rstrip("'")  # strip possessives

        spell_active = (
            hasattr(self, "_spell_highlighter") and
            self._spell_highlighter._enabled and
            _spell_backend.available()
        )
        is_misspelled = (
            spell_active and
            len(check_word) > 2 and
            not check_word.isupper() and
            not _spell_backend.check(check_word)
        )

        # ── Build menu ────────────────────────────────────────────────────
        menu = self.createStandardContextMenu()

        # Insert "Paste as Plain Text" directly after the standard Paste action
        paste_plain_act = QAction("Paste as Plain Text\tCtrl+Shift+V", menu)
        paste_plain_act.setEnabled(self._can_paste_plain())
        paste_plain_act.triggered.connect(self._paste_plain_text)
        _actions = menu.actions()
        _inserted = False
        for _i, _act in enumerate(_actions):
            if "paste" in _act.text().lower() and "plain" not in _act.text().lower():
                _ref = _actions[_i + 1] if _i + 1 < len(_actions) else None
                if _ref:
                    menu.insertAction(_ref, paste_plain_act)
                else:
                    menu.addAction(paste_plain_act)
                _inserted = True
                break
        if not _inserted:
            menu.addAction(paste_plain_act)

        t = _theme
        menu.setStyleSheet(f"""
            QMenu {{
                background:rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},245);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},80);
                border-radius:8px; padding:4px 0;
                color:{t.TEXT_PRIMARY.name()};
                font-family:'Segoe UI'; font-size:9pt;
            }}
            QMenu::item {{ padding:6px 18px; border-radius:4px; }}
            QMenu::item:selected {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);
                color:{t.GLOW.name()};
            }}
            QMenu::item:disabled {{
                color:rgba({t.TEXT_DIM.red()},{t.TEXT_DIM.green()},{t.TEXT_DIM.blue()},150);
            }}
            QMenu::separator {{
                height:1px;
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},40);
                margin:3px 8px;
            }}
        """)

        if is_misspelled:
            suggestions = _spell_backend.suggest(check_word)
            first_std   = menu.actions()[0] if menu.actions() else None

            # Bottom separator: divides spell section from standard items
            sep_bottom = menu.insertSeparator(first_std)

            # Ignore action
            ignore_act = QAction(f'Ignore \u201c{clean_word}\u201d', menu)
            menu.insertAction(sep_bottom, ignore_act)

            # Mid separator: divides suggestions from ignore
            sep_mid = menu.insertSeparator(ignore_act)

            # Suggestions (or a disabled placeholder)
            suggestion_acts: list[QAction] = []
            if suggestions:
                bold_font = QFont("Segoe UI", 9, QFont.Weight.DemiBold)
                for suggestion in suggestions:
                    act = QAction(suggestion, menu)
                    act.setFont(bold_font)
                    menu.insertAction(sep_mid, act)
                    suggestion_acts.append(act)
            else:
                no_act = QAction("(No suggestions)", menu)
                no_act.setEnabled(False)
                menu.insertAction(sep_mid, no_act)

            chosen = menu.exec(e.globalPos())

            if chosen == ignore_act:
                _spell_backend.ignore(clean_word)
            elif chosen in suggestion_acts:
                word_cursor.insertText(chosen.text())
                self.setTextCursor(word_cursor)
        else:
            menu.exec(e.globalPos())


class NotebookWindow(QWidget):
    """Standalone resizable notebook window."""
    node_creation_requested = Signal(str, str)  # (note_id, note_title)
    def __init__(self, parent=None):
        super().__init__(parent,
                         Qt.Window | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setMinimumSize(700, 500)
        self.resize(980, 680)
        self._drag_pos:          Optional[QPoint] = None
        self._resize_dir:        Optional[str]    = None
        self._resize_start_geom: Optional[QRect]  = None
        self._resize_start_pos:  Optional[QPoint] = None
        self._current_id:        Optional[str]    = None
        self._save_timer = QTimer(self)
        self._save_timer.setSingleShot(True)
        self._save_timer.setInterval(300)   # 300 ms debounce
        self._save_timer.timeout.connect(self._flush_save)
        self._build_ui()
        self._load_all_notes()
        _theme.register(self._on_theme)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(0)

        self._card = QWidget()
        self._card.setObjectName("nb_card")
        cl = QVBoxLayout(self._card)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(0)
        self._card.setGraphicsEffect(make_shadow(self._card, 32, QColor(0, 0, 0, 210)))

        # Title bar
        tbar = QWidget(); tbar.setFixedHeight(44)
        tbar.setStyleSheet("background:transparent;")
        tbl = QHBoxLayout(tbar)
        tbl.setContentsMargins(16, 0, 12, 0)
        ico = QLabel("📓")
        ico.setFont(QFont("Segoe UI", 12)); ico.setStyleSheet("background:transparent;")
        ttl = QLabel("  Notebook")
        ttl.setFont(QFont("Segoe UI", 11, QFont.Weight.DemiBold))
        ttl.setStyleSheet(f"color:{_theme.TEXT_PRIMARY.name()}; background:transparent;")
        tbl.addWidget(ico); tbl.addWidget(ttl); tbl.addStretch()
        # "Pin as Node" button
        pin_btn = QPushButton("📌")
        pin_btn.setFixedSize(28, 28)
        pin_btn.setCursor(Qt.PointingHandCursor)
        pin_btn.setFont(QFont("Segoe UI Emoji", 11))
        pin_btn.setToolTip("Create a Note node on the main canvas\nlinked to the currently selected note")
        pin_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({_theme.GLOW.red()},{_theme.GLOW.green()},{_theme.GLOW.blue()},40);
                border:1px solid rgba({_theme.GLOW.red()},{_theme.GLOW.green()},{_theme.GLOW.blue()},90);
                border-radius:6px; color:{_theme.GLOW.name()};
            }}
            QPushButton:hover {{
                background:rgba({_theme.GLOW.red()},{_theme.GLOW.green()},{_theme.GLOW.blue()},80);
            }}
        """)
        pin_btn.clicked.connect(self._on_pin_as_node)
        tbl.addWidget(pin_btn); tbl.addSpacing(6)
        mb = TitleBarButton(COLOR_BTN_MINIMIZE, "-")
        mb.setToolTip("Minimize")
        mb.clicked.connect(self.showMinimized)
        tbl.addWidget(mb); tbl.addSpacing(5)
        mx = TitleBarButton(COLOR_BTN_MAXIMIZE, "+")
        mx.setToolTip("Maximize / Restore")
        mx.clicked.connect(self._toggle_maximize)
        tbl.addWidget(mx); tbl.addSpacing(5)
        cb = TitleBarButton(COLOR_BTN_CLOSE, "x")
        cb.clicked.connect(self.close)
        tbl.addWidget(cb)
        tbar.mousePressEvent  = self._tbar_press
        tbar.mouseMoveEvent   = self._tbar_move
        tbar.mouseReleaseEvent = lambda e: setattr(self, "_drag_pos", None)
        cl.addWidget(tbar)

        # Top separator
        topsep = QFrame(); topsep.setFrameShape(QFrame.HLine)
        topsep.setStyleSheet(
            f"color:rgba({_theme.GLOW.red()},{_theme.GLOW.green()},{_theme.GLOW.blue()},50);")
        cl.addWidget(topsep)

        # Body: side panel + editor in a splitter
        self._splitter = QSplitter(Qt.Horizontal)
        self._splitter.setHandleWidth(1)
        self._splitter.setStyleSheet("""
            QSplitter::handle { background: rgba(128,128,128,30); }
        """)

        self._side = _NoteSidePanel()
        self._editor = _NoteEditor()

        self._splitter.addWidget(self._side)
        self._splitter.addWidget(self._editor)
        self._splitter.setSizes([220, 760])
        self._splitter.setCollapsible(0, False)
        self._splitter.setCollapsible(1, False)
        cl.addWidget(self._splitter)

        root.addWidget(self._card)

        # Connect side panel signals
        self._side.note_selected.connect(self._on_note_selected)
        self._side.note_new.connect(self._on_new_note)
        self._side.note_delete.connect(self._on_delete_note)
        self._editor.content_changed.connect(self._on_content_changed)

    # ── note operations ───────────────────────────────────────────────────────

    def _on_pin_as_node(self):
        """Emit signal so MainWindow can create a note node for the current note."""
        nid = self._editor.current_id()
        if not nid:
            return
        title = self._editor.current_title()
        self.node_creation_requested.emit(nid, title)

    def _toggle_maximize(self):
        if self.isMaximized():
            self.showNormal()
        else:
            self.showMaximized()

    def _load_all_notes(self):
        notes = NotebookStore.list_notes()
        self._side.load_notes(notes)
        if notes:
            first = notes[0]
            self._side._select_item(first["id"])
            self._load_note(first["id"])

    def _load_note(self, note_id: str):
        self._current_id = note_id
        data = NotebookStore.load_note(note_id)
        self._editor.load_note(note_id, data)

    def _on_note_selected(self, note_id: str):
        if note_id == self._current_id:
            return
        # Flush current note first
        self._flush_save()
        self._load_note(note_id)

    def _on_new_note(self):
        note_id = NotebookStore.new_id()
        NotebookStore.save_note(note_id, "Untitled", "")
        data = NotebookStore.load_note(note_id)
        self._side.add_note(note_id, "Untitled", data["modified"])
        self._flush_save()
        self._load_note(note_id)

    def _on_delete_note(self, note_id: str):
        NotebookStore.delete_note(note_id)
        self._side.remove_item(note_id)
        if self._current_id == note_id:
            self._current_id = None
            self._editor.clear()
            # Select first remaining
            remaining = NotebookStore.list_notes()
            if remaining:
                self._side._select_item(remaining[0]["id"])
                self._load_note(remaining[0]["id"])

    def _on_content_changed(self):
        """Schedule a save 300 ms after the last keystroke — zero data loss."""
        self._save_timer.start()

    def _flush_save(self):
        """Immediately write current note to disk."""
        self._save_timer.stop()
        nid = self._editor.current_id()
        if not nid:
            return
        title = self._editor.current_title()
        html  = self._editor.current_html()
        NotebookStore.save_note(nid, title, html)
        data = NotebookStore.load_note(nid)
        self._side.update_item(nid, title, data.get("modified", ""))

    # ── theme ─────────────────────────────────────────────────────────────────

    def _on_theme(self):
        self.update()

    # ── window painting (translucent + rounded) ───────────────────────────────

    def paintEvent(self, e: QPaintEvent):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        rect = QRectF(8, 8, self.width()-16, self.height()-16)
        path = QPainterPath()
        path.addRoundedRect(rect, BORDER_RADIUS+2, BORDER_RADIUS+2)
        grad = QLinearGradient(0, 0, 0, self.height())
        grad.setColorAt(0, QColor(t.BG_MID.red()+4, t.BG_MID.green()+4, t.BG_MID.blue()+6, 252))
        grad.setColorAt(1, QColor(t.BG_DARK.red(), t.BG_DARK.green(), t.BG_DARK.blue(), 252))
        p.fillPath(path, grad)
        border = QColor(t.GLOW); border.setAlpha(80)
        p.setPen(QPen(border, 1.3))
        p.setBrush(Qt.NoBrush)
        p.drawPath(path)
        p.end()

    # ── resize / drag ─────────────────────────────────────────────────────────

    _RESIZE_MARGIN = 7
    _CURSOR_MAP = {
        "tl": Qt.SizeFDiagCursor, "br": Qt.SizeFDiagCursor,
        "tr": Qt.SizeBDiagCursor, "bl": Qt.SizeBDiagCursor,
        "l":  Qt.SizeHorCursor,   "r":  Qt.SizeHorCursor,
        "t":  Qt.SizeVerCursor,   "b":  Qt.SizeVerCursor,
    }

    def _resize_edge(self, pos: QPoint) -> Optional[str]:
        m = self._RESIZE_MARGIN
        w, h = self.width(), self.height()
        x, y = pos.x(), pos.y()
        left = x < m; right = x > w - m; top = y < m; bottom = y > h - m
        if top and left:    return "tl"
        if top and right:   return "tr"
        if bottom and left: return "bl"
        if bottom and right:return "br"
        if left:   return "l"
        if right:  return "r"
        if top:    return "t"
        if bottom: return "b"
        return None

    def mousePressEvent(self, e: QMouseEvent):
        if e.button() != Qt.LeftButton: return
        d = self._resize_edge(e.position().toPoint())
        if d:
            self._resize_dir = d
            self._resize_start_geom = self.geometry()
            self._resize_start_pos  = e.globalPosition().toPoint()
        else:
            self._drag_pos = e.globalPosition().toPoint()

    def mouseMoveEvent(self, e: QMouseEvent):
        pos = e.position().toPoint()
        if self._resize_dir and self._resize_start_geom and self._resize_start_pos:
            delta = e.globalPosition().toPoint() - self._resize_start_pos
            g     = QRect(self._resize_start_geom)
            d     = self._resize_dir
            if "l" in d: g.setLeft(g.left() + delta.x())
            if "r" in d: g.setRight(g.right() + delta.x())
            if "t" in d: g.setTop(g.top() + delta.y())
            if "b" in d: g.setBottom(g.bottom() + delta.y())
            if g.width() >= self.minimumWidth() and g.height() >= self.minimumHeight():
                self.setGeometry(g)
        elif self._drag_pos:
            pass   # drag handled by title bar
        else:
            d = self._resize_edge(pos)
            self.setCursor(QCursor(self._CURSOR_MAP[d]) if d else Qt.ArrowCursor)

    def mouseReleaseEvent(self, e: QMouseEvent):
        self._resize_dir = None
        self._resize_start_geom = None
        self._resize_start_pos  = None
        self._drag_pos = None
        self.setCursor(Qt.ArrowCursor)

    def _tbar_press(self, e: QMouseEvent):
        if e.button() == Qt.LeftButton:
            self._drag_pos = e.globalPosition().toPoint()

    def _tbar_move(self, e: QMouseEvent):
        if self._drag_pos and e.buttons() == Qt.LeftButton:
            delta = e.globalPosition().toPoint() - self._drag_pos
            self.move(self.pos() + delta)
            self._drag_pos = e.globalPosition().toPoint()

    def closeEvent(self, e):
        self._flush_save()
        _theme.unregister(self._on_theme)
        super().closeEvent(e)


# ---------------------------------------------------------------------------
# Help dialog
# ---------------------------------------------------------------------------

_HELP_CONTENT = """
<style>
  body  { font-family:'Segoe UI',sans-serif; font-size:9pt; line-height:1.6; margin:0; padding:0 4px; }
  h2    { margin-top:20px; margin-bottom:5px; font-size:11pt; letter-spacing:0.5px; }
  h3    { margin-top:14px; margin-bottom:3px; font-size:9.5pt; }
  p, li { margin:3px 0; }
  ul    { padding-left:20px; margin:4px 0; }
  code  { font-family:'Consolas','Courier New',monospace; font-size:8.5pt;
          background:rgba(128,128,128,0.18); border-radius:3px; padding:1px 4px; }
  hr    { border:none; border-top:1px solid rgba(128,128,128,0.25); margin:14px 0; }
  .tip  { font-style:italic; opacity:0.70; }
  .key  { font-weight:600; }
  .badge { font-size:7.5pt; opacity:0.65; font-style:italic; }
  .highlight-box {
    border-left: 3px solid currentColor;
    padding: 8px 14px;
    margin: 8px 0;
    border-radius: 0 6px 6px 0;
    background: rgba(128,128,128,0.10);
  }
</style>

<h2>&#128187; Welcome to Command Center</h2>
<p>Command Center is a tile-based workflow dashboard. Every shortcut, file, URL,
note, or app you use daily lives on a single, themeable canvas in tiles called <b>nodes</b>. You can drag to reorder, group into folders,
and launch with one click.</p>

<hr>
<h2>&#10024; Why is Command Center Useful?</h2>
<p>Your daily tools are scattered across the taskbar, desktop, File Explorer, and browser bookmarks.
Command Center puts <b>everything on one canvas</b> &mdash; one click opens any file, app, URL, script, or note,
no matter where it actually lives on disk.</p>
<ul>
  <li><b>Scripts to clipboard instantly</b> &mdash; set a node&rsquo;s Open Behavior to <b>&ldquo;Copy file contents to clipboard&rdquo;</b> and your entire PowerShell or batch script is on the clipboard in a single click, ready to paste into any terminal.</li>
  <li><b>Organise by workflow</b> &mdash; group tools into folders (Dev, HR, Reporting&hellip;) regardless of where files live.</li>
  <li><b>Multi-select &amp; batch actions</b> &mdash; rubber-band drag to select multiple tiles, then delete, archive, or group them at once.</li>
  <li><b>Built-in extras</b> &mdash; Notebook, Time Tracker, Media Library, Calculator, and Task Manager are always one footer click away.</li>
</ul>

<hr>
<h2>&#9654; Nodes &mdash; The Building Blocks</h2>
<p>Each item on the canvas is called a <b>node</b>. There are three node types:</p>
<ul>
  <li><b>File</b> &mdash; opens or runs a file or executable with its default application.</li>
  <li><b>URL</b> &mdash; opens a web address in your default browser.</li>
  <li><b>Notebook Note</b> &mdash; links to a note in the built-in Notebook. Hover the tile to preview the note content; click to open it in the Notebook.</li>
</ul>

<h3>Creating a Node</h3>
<ul>
  <li>Click <b>+ Add Node</b> on the canvas, or <b>right-click empty canvas space</b> and choose <b>New Node</b>.</li>
  <li>In the wizard: enter a <b>Name</b>, pick a <b>Type</b>, fill in the target (path / URL / note).</li>
  <li>Choose a <b>Size</b>: <b>Small</b> (1&times;1), <b>Medium</b> (2&times;2), or <b>Large</b> (2&times;4).</li>
  <li>Add an optional <b>Description</b> &mdash; shown as a tooltip on hover.</li>
  <li>Optionally, enable <b>Auto-launch</b> to run the node automatically when Command Center starts.</li>
  <li>Add a custom <b>icon image</b> by clicking the <i>icon preview area</i> in the wizard to open a file browser, or drag an image file directly onto the icon area.</li>
  <li>Click <b>Save Node</b> to place it on the canvas.</li>
</ul>

<h3>Open Behavior (File Nodes)</h3>
<p>When creating or editing a File node, you can choose how clicking it behaves:</p>
<ul>
  <li><b>Launch normally</b> &mdash; opens the file with its default application.</li>
  <li><b>Open containing folder</b> &mdash; opens Windows Explorer to the folder that contains the file.</li>
  <li><b>Run as administrator</b> &mdash; launches the file elevated (triggers a UAC prompt).</li>
  <li><b>Copy file to clipboard</b> &mdash; places the file itself on the clipboard (for pasting into Explorer, email, etc.).</li>
  <li><b>Copy file contents to clipboard</b> &mdash; reads the file as text and copies its full contents to the clipboard.
      <i>Ideal for PowerShell scripts, batch files, SQL queries, and any text you frequently paste.</i></li>
</ul>

<h3>Launching a Node</h3>
<ul>
  <li><b>Single-click</b> a tile to launch it (or trigger its configured Open Behavior) immediately.</li>
  <li>File nodes open with the system default app (unless a custom behavior is set); URL nodes open in your default browser.</li>
  <li>Note nodes open the <b>Notebook</b> window and jump directly to the linked note.</li>
</ul>

<h3>&#128464; Drag &amp; Drop Files onto the Canvas</h3>
<ul>
  <li>Drag any file from Windows Explorer and drop it anywhere on the canvas.</li>
  <li><b>Image files</b> (<code>.png .jpg .jpeg .bmp .gif .webp .ico .svg</code>) &mdash; automatically create a 1&times;1 <b>File node</b> with the image as its icon.</li>
  <li><b>Text files</b> (<code>.txt .md .rst</code>) &mdash; the file&rsquo;s contents are imported into the Notebook as a new note, and a 1&times;1 <b>Note node</b> is created and linked to it.</li>
  <li><b>Other files</b> &mdash; a generic 1&times;1 File node is created that opens the file with its default application.</li>
</ul>

<h3>Editing &amp; Deleting</h3>
<ul>
  <li><b>Right-click</b> any tile to open its context menu. Key actions include:
    <ul>
      <li><b>Run</b> &mdash; launch the node immediately.</li>
      <li><b>Edit</b> &mdash; re-open the Node Wizard to change any field.</li>
      <li><b>Duplicate</b> &mdash; create an identical copy of the node on the canvas.</li>
      <li><b>Export (.node file)</b> &mdash; save the node as a portable <code>.node</code> file.</li>
      <li><b>Open File Location</b> / <b>Copy Path</b> &mdash; File nodes only.</li>
      <li><b>Run as Administrator</b> &mdash; elevated launch for executable File nodes.</li>
      <li><b>Copy File to Clipboard</b> / <b>Copy Contents to Clipboard</b> &mdash; File nodes only.</li>
      <li><b>Archive</b> &mdash; hide the node without deleting it (recoverable from Settings &#8594; Archived Nodes).</li>
      <li><b>Delete</b> &mdash; permanently remove the node (no undo).</li>
    </ul>
  </li>
</ul>

<h3>Importing &amp; Exporting</h3>
<ul>
  <li>Right-click a tile and choose <b>Export (.node file)</b> to save it as a portable <code>.node</code> file.</li>
  <li>In the Node Wizard click <b>Import .node</b> to load a previously exported node &mdash; useful for sharing configurations between machines or backing up individual nodes.</li>
</ul>

<hr>
<h2>&#9200; Scheduled Node Execution</h2>
<p>Any node can be configured to launch automatically on a schedule &mdash; without you having to click it. Open the <b>Node Wizard</b> (create or edit a node) and scroll to the <b>Schedule</b> section.</p>

<h3>Enabling a Schedule</h3>
<ul>
  <li>Check <b>Enable scheduled execution</b> to activate scheduling for that node.</li>
  <li>A small <b>&#9200;</b> clock badge appears on the tile&rsquo;s top-right corner to indicate an active schedule.</li>
  <li>The scheduler checks every <b>30 seconds</b> in the background while Command Center is running.</li>
</ul>

<h3>Schedule Types</h3>
<ul>
  <li><b>Daily</b> &mdash; runs every day at the specified time (HH:MM, 24-hour).</li>
  <li><b>Workdays (Mon&ndash;Fri)</b> &mdash; runs Monday through Friday at the specified time.</li>
  <li><b>Weekends (Sat&ndash;Sun)</b> &mdash; runs Saturday and Sunday at the specified time.</li>
  <li><b>Weekly (specific days)</b> &mdash; tick any combination of Mon&ndash;Sun to build a custom weekly pattern (e.g. Mon + Wed + Fri). Runs at the specified time on each selected day.</li>
  <li><b>Monthly (day of month)</b> &mdash; runs on a specific calendar day (1&ndash;31) each month at the specified time. If the month is shorter than the chosen day (e.g. day 31 in February), it fires on the last valid day.</li>
  <li><b>Interval (every N minutes/hours)</b> &mdash; fires repeatedly every N minutes or hours regardless of clock time. The first fire happens immediately when Command Center starts; subsequent fires are spaced by the interval. Ideal for polling tasks or periodic reminders.</li>
  <li><b>Once (specific date &amp; time)</b> &mdash; fires exactly once on the chosen date (YYYY-MM-DD) at the specified time, then never again unless you re-enable it or change the date.</li>
</ul>

<h3>Time Format</h3>
<p>All time fields use <b>24-hour HH:MM</b> notation (e.g. <code>09:00</code> for 9&nbsp;AM, <code>17:30</code> for 5:30&nbsp;PM).</p>

<h3>Notes &amp; Limitations</h3>
<ul>
  <li>Scheduled execution only works while Command Center is <b>running</b> (it does not use Windows Task Scheduler; there is no wake-from-sleep behaviour).</li>
  <li>The scheduler has a <b>tolerance window of &plusmn;35 seconds</b> around the target time to account for the 30-second poll interval.</li>
  <li>A node will not fire twice within the same minute, even if the scheduler ticks more than once near the boundary.</li>
  <li>The <b>last execution time</b> is stored with the node and survives app restarts, so interval nodes resume correctly after a relaunch.</li>
</ul>

<hr>
<h2>&#128276; Reminder System</h2>
<p>Command Center includes a fully independent reminder system. Click <b>Reminders</b> in the footer toolbar to open the Reminder Manager.</p>

<h3>Reminder Alert Levels</h3>
<ul>
  <li><b>Level 1 — Notify</b> &mdash; One or more of: <em>Windows Toast Notification</em>, <em>in-app popup dialog</em>, or <em>sound playback</em>. Combine any or all. The popup supports a configurable snooze duration.</li>
  <li><b>Level 2 — Full Screen</b> &mdash; Covers <em>every connected monitor</em> with a full-screen coloured overlay displaying the reminder message. Click anywhere or press any key to dismiss instantly. A configurable countdown timer auto-dismisses it (default 10 s). Background and text colours are fully customisable.</li>
  <li><b>Level 3 — Critical</b> &mdash; A flashing red warning dialog that can <em>only</em> be closed by typing the word <b>acknowledge</b> into the input box. The dialog is forced always-on-top and cannot be dismissed by pressing Escape or the window close button. Use for truly critical alerts.</li>
</ul>

<h3>Creating a Reminder</h3>
<ul>
  <li>Click <b>+ New Reminder</b> in the Reminder Manager header.</li>
  <li>Enter a <b>Title</b> (required) and optional <b>Message</b> body and <b>Tags</b>.</li>
  <li>Select an <b>Alert Level</b> (1, 2, or 3) — the options panel below updates to show the relevant settings.</li>
  <li>Set the <b>Schedule</b> using the same schedule types as Node Scheduling (Daily, Workdays, Weekly, Monthly, Interval, Once).</li>
  <li>Optionally set a <b>Priority</b> (Normal / High / Urgent), <b>Tag Colour</b>, and <b>Repeat Limit</b>.</li>
  <li>Click <b>Save Reminder</b>.</li>
</ul>

<h3>Level 1 — Notification Options</h3>
<ul>
  <li><b>Windows Notification</b> &mdash; sends a system toast notification visible in the Windows Action Centre.</li>
  <li><b>App Popup</b> &mdash; shows a themed in-app popup near the top-right of the main window. Includes an optional <b>Snooze</b> button.</li>
  <li><b>Play Sound</b> &mdash; plays either the built-in <em>Default Beep</em> or <em>Critical Stop</em> system sound, or a <b>custom audio file</b> (.wav plays natively; .mp3 and other formats play via the WPF MediaPlayer). Use the Volume slider to control playback volume.</li>
  <li>Select one, two, or all three notification types — all selected types fire simultaneously.</li>
</ul>

<h3>Level 2 — Full Screen Options</h3>
<ul>
  <li><b>Auto-dismiss timeout</b> &mdash; number of seconds before the overlay closes automatically (minimum 1, default 10).</li>
  <li><b>Background colour</b> &mdash; choose any colour for the overlay background (default near-black red).</li>
  <li><b>Text colour</b> &mdash; choose any colour for the title and message text.</li>
  <li>Optionally enable <b>Play sound on trigger</b> — uses the same sound settings as Level 1.</li>
</ul>

<h3>Level 3 — Critical Options</h3>
<ul>
  <li>The acknowledgement dialog flashes its red border to demand attention.</li>
  <li>The user <em>must</em> type exactly <b>acknowledge</b> (case-insensitive) into the text box to enable the dismiss button. Pressing Enter also works once the text matches.</li>
  <li>Optionally enable <b>Play sound on trigger</b>.</li>
</ul>

<h3>Schedule Options for Reminders</h3>
<p>Reminders use the exact same schedule engine as Node Scheduling (see above). All seven schedule types are available: Daily, Workdays, Weekends, Weekly, Monthly, Interval, and Once.</p>

<h3>Additional Options</h3>
<ul>
  <li><b>Snooze duration</b> (Level 1) &mdash; when the user clicks Snooze in a popup, the reminder is silenced for this many minutes before it can fire again.</li>
  <li><b>Repeat limit</b> &mdash; maximum number of times this reminder will fire (0 = unlimited). Useful for &ldquo;remind me 3 times then stop&rdquo; scenarios.</li>
  <li><b>Priority</b> &mdash; Normal / High / Urgent. Urgent reminders sort to the top of the Reminder Manager list and show a 🚨 icon.</li>
  <li><b>Tag colour</b> &mdash; accent colour applied to the reminder title in the list. Helps visually group reminders at a glance.</li>
  <li><b>Enable / Disable toggle</b> &mdash; quickly pause a reminder without deleting it. Disabled reminders are greyed out in the list and never fire.</li>
</ul>

<h3>Testing a Reminder</h3>
<ul>
  <li>Click the <b>▶</b> button on a reminder row (or right-click → <b>Test now</b>) to fire the reminder immediately, regardless of its schedule. Useful for verifying that the alert looks right.</li>
</ul>

<h3>Notes &amp; Limitations</h3>
<ul>
  <li>The reminder scheduler polls every 30 seconds, so there is up to a 30-second delay between a scheduled time and firing.</li>
  <li>Reminders only fire while Command Center is running — they do not use Windows Task Scheduler.</li>
  <li>Custom sounds must be accessible at the stored path each time the reminder fires. If the file is missing, the reminder still fires without sound.</li>
</ul>

<hr>
<h2>&#128269; Search</h2>
<ul>
  <li>Press <code>Ctrl+F</code> or click the <b>&#128269;</b> button in the title bar to open the search bar.</li>
  <li>Type any part of a node&rsquo;s <b>name</b> or <b>tag</b> to filter the canvas instantly &mdash; matching tiles float to the top.</li>
  <li>Search also surfaces nodes that are <b>nested inside folders</b> &mdash; they appear as temporary tiles on the main canvas for direct access.</li>
  <li>Press <code>Escape</code> or click the <b>&times;</b> on the search bar to clear the filter and restore normal order.</li>
  <li>Right-clicking a search result from inside a folder gives a <b>Remove From Folder</b> option to move it permanently to the main canvas.</li>
</ul>

<hr>
<h2>&#128193; Folders</h2>
<ul>
  <li>Click <b>New Folder</b> in the footer toolbar (or right-click empty canvas space and choose <b>New Folder</b>) to create a folder tile.</li>
  <li>Drag any <b>1&times;1</b> node tile <i>onto</i> a folder tile to nest it inside.</li>
  <li>Click a folder to open it and browse its contents.</li>
  <li><b>Right-click</b> a folder for the full context menu:</li>
  <ul>
    <li><b>Open Folder</b> &mdash; browse the folder&rsquo;s contents.</li>
    <li><b>Rename</b> &mdash; rename the folder.</li>
    <li><b>Empty Folder</b> &mdash; return all nested nodes to the main canvas without deleting them.</li>
    <li><b>Delete Folder (keep nodes)</b> &mdash; remove the folder tile and return its nodes to the canvas.</li>
    <li><b>Delete Folder and Contents</b> &mdash; permanently delete the folder <i>and</i> all nodes inside it.</li>
  </ul>
  <li>Inside an open folder, right-click a node and choose <b>Remove From Folder</b> to move it back to the main canvas individually.</li>
  <li>Only <b>1&times;1</b> nodes can be nested &mdash; larger nodes must remain on the main canvas.</li>
</ul>

<hr>
<h2>&#8645; Reordering &amp; Drag-and-Drop</h2>
<ul>
  <li><b>Click and hold</b> a tile to pick it up and drag it.</li>
  <li>Other tiles animate out of the way in real time as you move yours around.</li>
  <li>Drop onto an <b>empty spot</b> to move the tile there, or onto a <b>folder</b> to nest it.</li>
  <li>The canvas layout is automatically saved after every reorder.</li>
</ul>

<hr>
<h2>&#9744; Multi-Select &amp; Batch Actions</h2>
<p>You can select multiple tiles at once using a rubber-band drag, then act on them all in a single step.</p>
<ul>
  <li><b>Click and drag on empty canvas space</b> to draw a selection rectangle &mdash; all tiles within the rectangle are highlighted with a vibrant glow border.</li>
  <li>On releasing the drag, a <b>batch action menu</b> appears showing only the actions relevant to what you selected:</li>
  <ul>
    <li><b>Delete selected nodes</b> &mdash; permanently removes all selected node tiles.</li>
    <li><b>Delete selected folders</b> &mdash; removes selected folder tiles and returns their children to the canvas.</li>
    <li><b>Delete all selected</b> &mdash; removes every selected item (nodes <i>and</i> folder contents) permanently.</li>
    <li><b>Archive selected nodes</b> &mdash; archives all selected nodes (recoverable from Settings).</li>
    <li><b>Empty selected folders</b> &mdash; returns all children of selected folders to the main canvas.</li>
    <li><b>Place selected nodes in new folder</b> &mdash; prompts for a folder name and moves all selected nodes into it.</li>
  </ul>
  <li>Click <b>Clear selection</b> in the menu (or click any tile) to dismiss the selection without taking action.</li>
</ul>

<hr>
<h2>&#128073; Right-Click Canvas Menu</h2>
<p>Right-clicking on any <b>empty area</b> of the main canvas opens a quick-action menu:</p>
<ul>
  <li><b>New Node</b> &mdash; opens the Node Creation Wizard.</li>
  <li><b>New Folder</b> &mdash; creates a new empty folder tile.</li>
  <li><b>Refresh Display</b> &mdash; reloads all tiles from saved data.</li>
  <li><b>Settings</b> &mdash; opens the Settings dialog.</li>
  <li><b>Help</b> &mdash; opens this Help dialog.</li>
</ul>

<hr>
<h2>&#128221; Notebook</h2>
<p>The built-in Notebook is a full rich-text note editor. Notes can be linked to canvas tiles via <b>Notebook Note</b> nodes &mdash; hover a note tile to preview its contents, or click to open the note.</p>
<h3>Opening the Notebook</h3>
<ul>
  <li>Click <b>Notebook</b> in the footer toolbar or press <code>Ctrl+Shift+B</code>. The window is independently resizable and can stay open alongside the main canvas.</li>
</ul>
<h3>&#128204; Creating a Note Node from the Notebook</h3>
<ul>
  <li>While a note is open in the editor, click the <b>&#128204; (pin)</b> button in the Notebook title bar.</li>
  <li>A 1&times;1 <b>Notebook Note</b> tile is instantly created on the main canvas and linked to that note.</li>
  <li>You can also create Notebook Note nodes via the Node Wizard (<b>Add Node</b> &#8594; type = Notebook Note).</li>
</ul>
<h3>&#128065; Hover Previews on Note Tiles</h3>
<ul>
  <li>Hover your mouse over any Notebook Note tile on the canvas.</li>
  <li>A floating preview card appears showing the note title and the first portion of its content.</li>
</ul>
<h3>Managing Notes</h3>
<ul>
  <li>Click <b>+</b> in the left panel to create a new note.</li>
  <li>Right-click a note in the list and choose <b>Delete Note</b> to permanently delete it.</li>
  <li>Notes are stored as individual <code>.ccnote</code> JSON files in <code>%APPDATA%\\CommandCenter\\notebooks\\</code>.</li>
</ul>
<h3>Editing &amp; Formatting</h3>
<ul>
  <li>Use the <b>formatting toolbar</b> to apply: <b>Bold</b>, <i>Italic</i>, <u>Underline</u>, <s>Strikethrough</s>.</li>
  <li>Apply <b>H1 / H2 / H3</b> headings, <b>bullet</b> or <b>numbered</b> lists.</li>
  <li>Insert <b>hyperlinks</b> (&#128279;), <b>inline code</b> (&lt;/&gt;), <b>block quotes</b> (&#10078;), and <b>horizontal rules</b> (&#8212;).</li>
  <li>Keyboard shortcuts: <code>Ctrl+B</code> Bold &nbsp;|&nbsp; <code>Ctrl+I</code> Italic &nbsp;|&nbsp; <code>Ctrl+U</code> Underline.</li>
</ul>
<h3>Auto-Save</h3>
<ul>
  <li>Every keystroke schedules a save <b>300 ms later</b>. Closing the window always flushes any pending save.</li>
</ul>

<hr>
<h2>&#128248; Media Library</h2>
<p>The Media Library is your personal stash of frequently-used visual assets and text snippets &mdash; all in one place, always one click away.
Open it from the footer toolbar with the <b>Media Library</b> button, or press <code>Ctrl+Shift+M</code>.</p>
<p>It has four tabs:</p>
<ul>
  <li><b>GIFs</b> &mdash; store your favorite animated GIFs. They display as static thumbnails and only animate when you hover over them.</li>
  <li><b>Emojis</b> &mdash; save your most-used emoji characters for instant access.</li>
  <li><b>Emoticons</b> &mdash; pre-loaded with classic text emoticons (&#175;\\_(ツ)_/&#175; and friends) and fully customisable.</li>
  <li><b>Pictures</b> &mdash; store reference images, screenshots, diagrams, or any picture you frequently need.</li>
</ul>
<div class="highlight-box">
  <b>&#128203; Clicking any item copies it to your clipboard and closes the library</b> &mdash;
  just open the library, click the item you want, and paste it wherever you need it.
  No dragging, no right-clicking, no copy dialogs.
</div>
<h3>Adding Media</h3>
<ul>
  <li><b>GIFs &amp; Pictures:</b> <b>drag and drop</b> files directly onto the GIFs or Pictures tab from Windows Explorer.
      Accepted formats: <code>.gif</code> for GIFs; <code>.png .jpg .jpeg .bmp .webp</code> for pictures.
      Files are <b>copied into</b> the library folder so originals can be moved freely.</li>
  <li><b>Emojis &amp; Emoticons:</b> type or <b>paste</b> any emoji or text emoticon into the input box at the top of the tab and click <b>Add</b> (or press <code>Enter</code>).</li>
</ul>
<h3>Removing Items</h3>
<ul>
  <li><b>Right-click</b> any item in the library and choose <b>Remove from library</b> to delete it.</li>
  <li>Removing a GIF or picture also deletes the stored copy from the library folder on disk.</li>
</ul>
<h3>Storage Location</h3>
<ul>
  <li>GIFs: <code>%APPDATA%\\CommandCenter\\media_library\\gifs\\</code></li>
  <li>Pictures: <code>%APPDATA%\\CommandCenter\\media_library\\pictures\\</code></li>
  <li>Emojis &amp; emoticons: <code>%APPDATA%\\CommandCenter\\media_library\\library.json</code></li>
</ul>

<hr>
<h2>&#128203; Clipboard Manager</h2>
<p>The Clipboard Manager automatically records everything you copy &mdash; text, rich text, HTML, images, and files &mdash; and lets you paste any of them again at any time.
Open it with <code>Ctrl+`</code> (backtick) or click the <b>Clipboard</b> button in the footer toolbar.
The <code>Ctrl+`</code> hotkey is <b>global</b> &mdash; it works even when Command Center is not the focused window.</p>

<h3>The List</h3>
<ul>
  <li>Each entry shows a colored type badge (<b>TXT</b>, <b>HTM</b>, <b>RTF</b>, <b>IMG</b>, <b>FIL</b>, <b>PWD</b>), a preview of the content, and a timestamp.</li>
  <li>Password entries (<b>PWD</b>) are automatically detected from generated-password patterns (e.g. <code>Delivered9898Flower$</code>) and shown with a masked preview of bullets &mdash; the actual password is copied normally when you select the entry.</li>
  <li><b>Image entries</b> display an inline thumbnail and pixel dimensions.</li>
  <li><b>File entries</b> display the full path and a count of any additional files copied at the same time.</li>
  <li>Pinned items are marked with a teal left-edge bar and always stay at the top of the list regardless of new copies.</li>
  <li>Use the <b>search bar</b> at the top to filter by content, and use the type chip buttons (<b>All / Text / HTML / Images / Files / PWD</b>) to narrow by category.</li>
</ul>

<h3>Copying an Entry</h3>
<ul>
  <li>Select any item and click <b>&#10696; Copy &amp; Close</b> to restore it to the clipboard and dismiss the manager.</li>
  <li>Click <b>Copy</b> to restore it to the clipboard while keeping the manager open.</li>
  <li>If <b>Auto Copy and Close on Click</b> is enabled in settings, a single click on any item immediately copies it and closes the manager &mdash; no button needed.</li>
</ul>

<h3>Other Actions</h3>
<ul>
  <li><b>&#128204; Pin</b> &mdash; pins the selected entry so it stays at the top and is never auto-expired. Click <b>Unpin</b> to release it.</li>
  <li><b>&#10005; Delete</b> &mdash; removes the selected entry from history.</li>
  <li><b>Clear All</b> &mdash; removes all unpinned entries. Pinned entries are unaffected.</li>
</ul>

<h3>Settings <span class="badge">(&#9881; icon in the title bar)</span></h3>
<ul>
  <li><b>Max History Size</b> &mdash; how many entries to keep (50&ndash;1000). Oldest unpinned entries are dropped when the limit is reached.</li>
  <li><b>Close manager after copying</b> &mdash; automatically hides the window after any copy action.</li>
  <li><b>Auto Copy and Close on Click</b> &mdash; a single click on any list item immediately copies it and closes the manager.</li>
  <li><b>Ignore duplicate entries</b> &mdash; prevents recording an item if it is identical to the most recent entry.</li>
  <li><b>Capture image copies</b> &mdash; records image data when you copy an image (e.g. right-click &rarr; Copy image in a browser).</li>
  <li><b>Capture file copies</b> &mdash; records file paths when you copy files in Windows Explorer.</li>
</ul>

<hr>
<h2>&#9201; Time Tracker</h2>
<ul>
  <li>Click <b>Time Tracker</b> in the footer (or press <code>Ctrl+Shift+T</code>) to open the floating HUD. Click again to dismiss it.</li>
  <li><b>Green timer</b> &mdash; counts <i>unlocked</i> (active) time.</li>
  <li><b>Red timer</b> &mdash; counts <i>locked</i> time (e.g., away / screen-locked).</li>
  <li><b>Orange decimal</b> &mdash; a lap counter; click it to reset the lap without touching the main timers.</li>
  <li>Click <b>&#177;</b> to manually add or subtract time from either timer.</li>
  <li>Click <b>&#8801;</b> to open the full time log for the current session.</li>
  <li>Minimize the HUD to a compact pill with the <b>&ndash;</b> button; expand it again with the pill button.</li>
  <li>Closing and reopening the Time Tracker always starts a <b>fresh session</b> from zero by design.</li>
</ul>

<hr>
<h2>&#9881; Settings</h2>
<h3>Appearance</h3>
<ul>
  <li>Choose a <b>preset theme</b>: Deep Space, Midnight Blue, Slate Light, Forest Night, Crimson Dark, Dark Knight, Spooky, Noir, or Custom. Each theme sets the full color palette app-wide.</li>
  <li>Click any <b>color swatch</b> to override individual theme colors (background dark, background mid, glow/accent, text primary, text secondary, text dim, and each accent color).</li>
  <li>Click <b>Reset Custom Colors</b> to discard all overrides and restore the current preset&rsquo;s default palette.</li>
  <li>Use the <b>UI Brightness</b> slider (30%&ndash;200%) to globally scale the luminance of every theme color at once &mdash; useful for bright or dim displays.</li>
  <li>Choose a <b>Canvas Background</b> style: <b>Solid</b>, <b>Dots</b>, <b>Grid</b>, <b>Noise</b>, <b>Gradient</b>, <b>Hexagons</b>, <b>Web</b>, or <b>Custom Image</b>. Each style (except Custom Image) uses your current theme colors.</li>
  <li>Set a custom <b>cursor style</b> (Standard, Cosmic, High Fantasy, High Tech, Medieval, Noir, Reactor, Spooky) or upload your own cursor image with the <b>Custom</b> slot. Adjust the <b>cursor size</b> with the slider. Changes apply instantly and persist across sessions.</li>
  <li>Theme, brightness, and cursor changes apply <b>live</b> &mdash; every open window (including dialogs, Notebook, and plugin windows) updates immediately.</li>
  <li>Toggle <b>Disable title bar animation</b> to stop the pulsing title text and keep it static.</li>
  <li>Toggle <b>Skip startup animation</b> to bypass the splash screen and jump straight to the canvas on launch.</li>
</ul>
<h3>General</h3>
<ul>
  <li><b>Save File Location</b> &mdash; change where nodes, settings, and data are stored. Defaults to <code>%APPDATA%\\CommandCenter</code>. Save-path changes require an <b>app restart</b> to take effect.</li>
</ul>
<p><b>Startup</b></p>
<ul>
  <li><b>Show a tip at startup</b> &mdash; display a helpful tip card when Command Center launches.</li>
  <li><b>Auto Launch Time Tracker</b> &mdash; automatically open the Time Tracker HUD when the app starts.</li>
</ul>
<p><b>Behavior</b></p>
<ul>
  <li><b>Confirm before deleting / archiving a node</b> &mdash; adds an extra confirmation prompt before destructive actions.</li>
  <li><b>Run auto-launch nodes on startup</b> &mdash; globally enables or disables auto-launch for all nodes that have it set.</li>
  <li><b>Single-click to launch nodes</b> &mdash; uncheck to require a double-click to launch a node instead of a single click.</li>
</ul>
<p><b>System</b></p>
<ul>
  <li><b>Launch Command Center on Windows startup</b> &mdash; adds Command Center to your Windows startup registry key so it opens automatically when you log in.</li>
</ul>
<p><b>Scripting</b></p>
<ul>
  <li><b>Bypass Script Execution Policy for PowerShell scripts</b> &mdash; when enabled, <code>.ps1</code> nodes are run with <code>-ExecutionPolicy Bypass</code> so they always execute regardless of system policy.</li>
</ul>
<p><b>Hotkeys (in General tab)</b></p>
<ul>
  <li><b>Disable all keyboard shortcuts</b> &mdash; suppresses all built-in app hotkeys at once. Shortcuts inside the Notebook editor are not affected.</li>
</ul>
<h3>Hotkeys Tab</h3>
<ul>
  <li>View and configure every built-in keyboard shortcut from one place.</li>
  <li>Click the <b>&#9210;</b> record button on any row, then press your desired key combination. The new shortcut takes effect as soon as you click <b>Apply and Close</b>.</li>
  <li>Press <b>Escape</b> while recording to cancel without changing the binding.</li>
  <li>Use the <b>&#8635;</b> reset button on a row to restore just that shortcut to its default.</li>
  <li>Click <b>&#8635; Reset All to Defaults</b> to restore every shortcut at once.</li>
  <li>A <b>&#9888;</b> icon appears next to a binding that uses only a single key &mdash; consider adding a modifier (Ctrl, Shift, Alt) to avoid accidental activation.</li>
  <li>A <b>&#9889;</b> icon appears when two actions share the same key combination &mdash; resolve the conflict before clicking Apply.</li>
</ul>
<h3>Archived Nodes</h3>
<ul>
  <li>View all archived nodes and <b>restore</b> them to the canvas, or <b>permanently delete</b> them from here.</li>
</ul>

<hr>
<h2>&#9776; Footer Toolbar</h2>
<p>The footer bar runs along the bottom of the window. Built-in buttons appear on the left; plugin-injected buttons appear to the right of the built-in buttons.</p>
<ul>
  <li><b>Settings</b> &#8594; Opens the Settings dialog.</li>
  <li><b>Plugins</b> &#8594; Opens the Plugin Manager dialog. Install, enable/disable, reload, and uninstall <code>.ccplug</code> plugins. Also access the full plugin developer guide.</li>
  <li><b>New Folder</b> &#8594; Creates a new empty folder tile on the canvas.</li>
  <li><b>Time Tracker</b> &#8594; Toggles the floating Time Tracker HUD (<code>Ctrl+Shift+T</code>).</li>
  <li><b>Notebook</b> &#8594; Opens the full rich-text Notebook window (<code>Ctrl+Shift+B</code>).</li>
  <li><b>Media Library</b> &#8594; Opens your personal GIF, emoji, emoticon, and picture stash. Click any item to copy it to clipboard (<code>Ctrl+Shift+M</code>).</li>
  <li><b>File Explorer</b> &#8594; Opens Windows File Explorer.</li>
  <li><b>Calculator</b> &#8594; Launches Windows Calculator.</li>
  <li><b>Notepad++</b> &#8594; Opens Notepad++ (must be installed at its default location).</li>
  <li><b>Lock Screen</b> &#8594; Immediately locks the Windows session.</li>
  <li><b>Feedback</b> &#8594; Opens your default mail client with a pre-addressed feedback email.</li>
</ul>

<hr>
<h2>&#128279; Quick Connect</h2>
<p>Launch a remote ScreenConnect session from the title bar icon or <code>Ctrl+Q</code>.</p>
<ul>
  <li>Type a <b>Client ID</b> &mdash; the list filters live. Accept the inline suggestion with <b>Tab</b> or <b>&rarr;</b>, or pick from the dropdown.</li>
  <li>Add an optional <b>Device Name</b> to target a specific machine, then press <b>Enter</b> or click <b>Connect</b>.</li>
</ul>

<hr>
<h2>&#128736; Title Bar</h2>
<ul>
  <li><b>Drag</b> the title bar to move the window anywhere on screen.</li>
  <li><b>Double-click</b> the title bar to maximise / restore the window.</li>
  <li>The <b>&#128269;</b> button opens the search bar (<code>Ctrl+F</code>).</li>
  <li>The <b>?</b> button opens this Help dialog (<code>F1</code>).</li>
  <li>The <b>&#9889; Quick Connect</b> button (or <code>Ctrl+Q</code>) opens the Quick Connect dialog for instant remote access.</li>
  <li><b>&ndash;</b> minimises, <b>+</b> maximises/restores, <b>&times;</b> closes.</li>
</ul>

<hr>
<h2>&#9000; Keyboard Shortcuts</h2>
<p>All built-in shortcuts are fully customisable in <b>Settings &#8594; Hotkeys</b>. The defaults are shown below. Installed plugins may register additional shortcuts &mdash; these are listed in the Plugins dialog next to each plugin.</p>
<table style="border-collapse:collapse; width:100%;">
  <tr><td style="padding:3px 8px; width:38%;"><code>Ctrl+F</code></td><td style="padding:3px 8px;">Open / focus the search bar</td></tr>
  <tr><td style="padding:3px 8px;"><code>Escape</code></td><td style="padding:3px 8px;">Close the search bar</td></tr>
  <tr><td style="padding:3px 8px;"><code>Ctrl+N</code></td><td style="padding:3px 8px;">Open the New Node wizard</td></tr>
  <tr><td style="padding:3px 8px;"><code>Ctrl+Shift+N</code></td><td style="padding:3px 8px;">Create a new folder</td></tr>
  <tr><td style="padding:3px 8px;"><code>Ctrl+,</code></td><td style="padding:3px 8px;">Open Settings</td></tr>
  <tr><td style="padding:3px 8px;"><code>F1</code></td><td style="padding:3px 8px;">Open Help &amp; FAQ (this dialog)</td></tr>
  <tr><td style="padding:3px 8px;"><code>Ctrl+Shift+T</code></td><td style="padding:3px 8px;">Toggle the Time Tracker HUD</td></tr>
  <tr><td style="padding:3px 8px;"><code>Ctrl+Shift+B</code></td><td style="padding:3px 8px;">Open the Notebook</td></tr>
  <tr><td style="padding:3px 8px;"><code>Ctrl+Shift+M</code></td><td style="padding:3px 8px;">Open the Media Library</td></tr>
  <tr><td style="padding:3px 8px;"><code>Ctrl+`</code></td><td style="padding:3px 8px;">Open the Clipboard Manager (global &mdash; works system-wide)</td></tr>
  <tr><td style="padding:3px 8px;"><code>Ctrl+Q</code></td><td style="padding:3px 8px;">Open Quick Connect</td></tr>
</table>
<p class="tip">All built-in shortcuts above can be remapped in <b>Settings &#8594; Hotkeys</b> or disabled in bulk via <b>Settings &#8594; General &#8594; Hotkeys</b>.  Notebook editor shortcuts (<code>Ctrl+B</code> Bold, <code>Ctrl+I</code> Italic, <code>Ctrl+U</code> Underline) are always active inside the Notebook editor and are not user-configurable.  Plugin-registered shortcuts are application-wide and fire even when a dialog is open.</p>

<hr>
<h2>&#128268; Plugins</h2>
<p>Command Center supports <b>.ccplug plugins</b> &mdash; small Python extensions that add new features, buttons, hotkeys, dialogs, and automations directly inside the app.</p>

<h3>Opening the Plugin Manager</h3>
<p>Click <b>Plugins</b> in the footer toolbar. The Plugin Manager dialog lists all installed plugins with their status, description, and version.</p>

<h3>Installing a Plugin</h3>
<ol>
  <li>Click <b>+ Install Plugin</b> in the Plugin Manager.</li>
  <li>Select a <code>.ccplug</code> file (a ZIP archive with a renamed extension).</li>
  <li>The plugin is extracted and its details are shown. Click <b>Enable</b> to activate it immediately.</li>
</ol>

<h3>Enabling &amp; Disabling</h3>
<p>Toggle the <b>Enable</b> switch next to any plugin to activate or deactivate it instantly. The plugin&rsquo;s footer buttons, hotkeys, and settings tabs appear and disappear in real time.</p>

<h3>Reloading (for Developers)</h3>
<p>Click <b>Reload</b> on any plugin to re-import its Python code from disk without restarting Command Center. This is the fastest way to test plugin changes during development &mdash; just edit the files in <code>%APPDATA%\\CommandCenter\\plugins\\&lt;plugin_id&gt;\\</code> and click Reload.</p>

<h3>Uninstalling</h3>
<p>Click <b>Uninstall</b> to deactivate the plugin and permanently delete its install directory and all associated files.</p>

<h3>Purge All</h3>
<p>The <b>&#9888; Purge All</b> button (title bar of the Plugin Manager) deactivates every plugin, erases all plugin settings, and immediately restarts Command Center for a completely clean state. Use this to recover from a broken plugin.</p>

<h3>Plugin Developer Guide</h3>
<p>Click <b>&#128218; .ccplug Guide</b> at the bottom of the Plugin Manager to open the full developer guide. It covers the complete plugin API, working examples, packaging instructions, and best practices for building your own plugins.</p>

<h3>What Can Plugins Do?</h3>
<ul>
  <li>Add <b>footer buttons</b> that open custom dialogs or run actions</li>
  <li>Register <b>global keyboard shortcuts</b> (application-wide)</li>
  <li>Read, create, update, and delete <b>canvas nodes</b></li>
  <li>Read, create, and modify <b>notebook notes</b></li>
  <li>Inject custom <b>settings tabs</b> into the Settings dialog</li>
  <li>Read and write the <b>system clipboard</b></li>
  <li>Run <b>background timers</b> for periodic tasks</li>
  <li>Launch <b>external programs</b> and manage subprocesses</li>
  <li>Read and write <b>local files</b> in the plugin directory</li>
  <li>React to <b>theme changes</b> and re-style their UI in real time</li>
</ul>

<hr>
<h2>&#10067; FAQ</h2>

<h3>How do I re-paste something I copied earlier?</h3>
<p>Press <code>Ctrl+`</code> (backtick) to open the <b>Clipboard Manager</b>. Select the entry you want and click <b>&#10696; Copy &amp; Close</b> (or just click it once if <b>Auto Copy and Close on Click</b> is enabled in clipboard settings).
The item is restored to your clipboard and the manager closes &mdash; paste as normal.</p>

<h3>My node won&rsquo;t launch &mdash; what&rsquo;s wrong?</h3>
<p>Check the path or URL is correct and that the file still exists. For File nodes, the associated application must be installed.
For script nodes (<code>.py</code>, <code>.bat</code>, <code>.ps1</code>), verify the script works in a standalone terminal first.</p>

<h3>How do I copy a PowerShell script to the clipboard with one click?</h3>
<p>Create a File node pointing to your <code>.ps1</code> file. In the wizard, set <b>Open Behavior</b> to
<b>"Copy file contents to clipboard"</b>. Now every time you click that tile, the entire script is copied to
your clipboard instantly &mdash; paste it directly into PowerShell or any terminal. No editor, no file browser, no manual selection.</p>

<h3>How do I move a node back out of a folder?</h3>
<p>Two ways: right-click the node <i>inside</i> the open folder and choose <b>Remove From Folder</b>, or right-click the folder tile on the canvas and choose <b>Delete Folder (keep nodes)</b> to dump all its contents back to the main canvas at once.</p>

<h3>Can I put a 2&times;2 or 2&times;4 node inside a folder?</h3>
<p>No &mdash; only <b>1&times;1</b> nodes can be nested inside folders.</p>

<h3>What is the difference between &ldquo;Delete Folder (keep nodes)&rdquo; and &ldquo;Delete Folder and Contents&rdquo;?</h3>
<p><b>Delete Folder (keep nodes)</b> removes the folder tile and returns all nested nodes to the main canvas &mdash; nothing is permanently lost.
<b>Delete Folder and Contents</b> removes the folder <i>and</i> permanently deletes every node inside it. There is no undo for either action.</p>

<h3>Where is my data saved?</h3>
<p>Nodes and settings are stored in <code>%APPDATA%\\CommandCenter</code> by default (change this in Settings &#8594; General).
Notebook notes live in <code>%APPDATA%\\CommandCenter\\notebooks\\</code>. Media Library files are in
<code>%APPDATA%\\CommandCenter\\media_library\\</code>. Plugin files are in <code>%APPDATA%\\CommandCenter\\plugins\\</code>.</p>

<h3>How do I change the theme or brightness?</h3>
<p>Go to <b>Settings &#8594; Appearance</b>. Pick a preset, click any color swatch to override individual colors, or drag the <b>UI Brightness</b> slider to scale the overall luminance of every color. All changes apply instantly.</p>

<h3>My Time Tracker reset when I closed it &mdash; is that a bug?</h3>
<p>No, this is intentional. Closing and reopening the tracker always starts a fresh zero-based session so each work period begins clean.</p>

<h3>Why does the title bar text slowly change color?</h3>
<p>The &ldquo;Command Center&rdquo; title has a subtle idle animation that gently pulses between your theme&rsquo;s text color and accent color. It&rsquo;s purely cosmetic and can be disabled in <b>Settings &#8594; Appearance &#8594; Disable title bar animation</b>.</p>

<h3>What file format does the Notebook use?</h3>
<p>Each note is a <code>.ccnote</code> file (standard JSON containing the title, HTML content, and timestamps).
Images are embedded as Base64 strings inside the HTML &mdash; no separate image files are created.</p>

<h3>Notepad++ isn&rsquo;t opening from the footer button.</h3>
<p>Command Center looks for Notepad++ in <code>C:\\Program Files\\Notepad++</code> and <code>C:\\Program Files (x86)\\Notepad++</code>. If you installed it elsewhere, create a regular File node pointing to your <code>notepad++.exe</code> instead.</p>

<h3>How do I add media to the Media Library?</h3>
<p>Open the <b>Media Library</b> from the footer. On the <b>GIFs</b> or <b>Pictures</b> tab, drag and drop files directly onto the tab.
On the <b>Emojis</b> or <b>Emoticons</b> tab, type or paste your text into the input box and click <b>Add</b>.
To use any stored item later, just click it &mdash; it copies to your clipboard and the library closes automatically.</p>

<h3>How do I remove something from the Media Library?</h3>
<p>Right-click any item in the Media Library and choose <b>Remove from library</b>. For GIFs and pictures, the stored file is also deleted from disk.</p>

<h3>How do I back up my data?</h3>
<p>Copy the entire <code>%APPDATA%\\CommandCenter</code> folder. It contains your nodes, settings, notebook notes, media library, and installed plugins.
You can also use <b>Export .node</b> on individual tiles for selective backup.</p>

<h3>How do I install a plugin?</h3>
<p>Click <b>Plugins</b> in the footer, then click <b>+ Install Plugin</b> and select a <code>.ccplug</code> file. After installation, toggle <b>Enable</b> to activate it. See the <b>Plugins</b> section above for the full workflow.</p>

<h3>A plugin is causing Command Center to crash or behave strangely.</h3>
<p>Open the <b>Plugins</b> dialog and disable the plugin using its Enable toggle. If the app cannot start, open <code>%APPDATA%\\CommandCenter\\plugins\\</code> and rename or delete the offending plugin&rsquo;s folder, then restart. As a last resort, use <b>&#9888; Purge All</b> in the Plugin Manager title bar to wipe all plugin state and restart cleanly.</p>

<h3>Where do I find the plugin developer documentation?</h3>
<p>Open the <b>Plugins</b> dialog (footer &#8594; Plugins) and click <b>&#128218; .ccplug Guide</b> at the bottom. The guide covers the full plugin API, file format, lifecycle methods, working examples, packaging, and best practices.</p>
"""


# ===========================================================================
# Plugin System
# ===========================================================================

PLUGIN_DIR = CONFIG_DIR / "plugins"
PLUGIN_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# Embedded .ccplug developer guide
# ---------------------------------------------------------------------------

_CCPLUG_GUIDE = """\
================================================================================
  COMMAND CENTER — PLUGIN DEVELOPER GUIDE
  Version 1.0.0.0+  |  .ccplug Format  |  Python 3.12 + PySide6
================================================================================

This guide covers everything you need to build, test, package, and publish
a plugin for Command Center.  Read it once end-to-end before writing code —
the plugin system is simple, but there are patterns you must follow.


================================================================================
  1. WHAT IS A PLUGIN?
================================================================================

A .ccplug file is a ZIP archive renamed with the .ccplug extension.  It
contains at minimum two files: manifest.json (identity + metadata) and
plugin.py (your Python code).

When a plugin is enabled, Command Center:
  1. Extracts the archive to  %APPDATA%\\CommandCenter\\plugins\\<id>\\
  2. Adds the plugin directory to sys.path
  3. Imports plugin.py
  4. Calls  activate(api)  with the full CommandCenterAPI object

When it is disabled (or uninstalled), Command Center calls  deactivate()  so
your plugin can clean up everything it registered.

Key constraints:
  • You must implement both  activate(api)  and  deactivate()
  • deactivate() must undo every side effect activate() made
  • No blocking code, no sleep(), no synchronous network calls on the main thread
  • All UI code runs on the main thread (Qt requirement)


================================================================================
  2. QUICK-START — A WORKING PLUGIN IN 5 MINUTES
================================================================================

Step 1 — Create a project folder:
  mkdir my_plugin
  cd my_plugin

Step 2 — Create manifest.json:
  {
    "id":          "my_plugin",
    "name":        "My Plugin",
    "version":     "1.0.0.0",
    "description": "A simple footer button.",
    "author":      "Your Name",
    "permissions": ["ui"]
  }

Step 3 — Create plugin.py:
  _api = None
  _btn = None

  def activate(api):
      global _api, _btn
      _api = api
      _btn = api.ui.add_footer_button("My Plugin", _on_click)

  def deactivate():
      if _btn:
          _api.ui.remove_footer_button(_btn)

  def _on_click():
      _api.toast("Hello from My Plugin!", "success")

Step 4 — Package it:
  python -c "
  import zipfile
  with zipfile.ZipFile('my_plugin.ccplug', 'w') as z:
      z.write('manifest.json')
      z.write('plugin.py')
  print('Done!')
  "

Step 5 — Install it:
  Footer → Plugins → '+ Install Plugin' → select my_plugin.ccplug → Enable.

Your button appears in the footer immediately.

For development after the first install, skip the zip step — just edit the
files directly in  %APPDATA%\\CommandCenter\\plugins\\my_plugin\\  and click
Reload in the Plugins dialog.


================================================================================
  3. PLUGIN FILE STRUCTURE
================================================================================

A .ccplug file is a standard ZIP archive with the extension renamed.
Minimum contents:

  my_plugin.ccplug (ZIP)
  ├── manifest.json         REQUIRED
  └── plugin.py             REQUIRED (default entry-point name)

You may also include helper modules and data files:

  my_plugin.ccplug (ZIP)
  ├── manifest.json
  ├── plugin.py
  ├── helpers.py            Additional modules
  ├── ui_components.py      More modules
  ├── data\\
  │   └── defaults.json     Bundled default data
  └── assets\\
      └── icon.png

All files are extracted flat into the plugin's install directory:
  %APPDATA%\\CommandCenter\\plugins\\my_plugin\\

Import helper modules normally (they are on sys.path):
  import helpers
  from ui_components import MyDialog

The install directory stays between plugin reloads and uninstalls (unless
the user clicks Uninstall, which deletes the folder entirely).

IMPORTANT: Never hardcode absolute paths.  Always use  api.plugin_dir:
  data_file = api.plugin_dir / "my_data.json"    # correct
  data_file = "C:\\Users\\me\\my_data.json"       # wrong — never do this


================================================================================
  4. manifest.json — FULL REFERENCE
================================================================================

{
  "id":              "my_plugin",
  "uuid":            "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "name":            "My Plugin",
  "version":         "1.0.0.0",
  "description":     "What this plugin does.",
  "author":          "Your Name",
  "min_app_version": "1.0.0.0",
  "entry":           "plugin.py",
  "permissions":     ["ui", "hotkeys", "notebook", "nodes", "settings",
                      "files", "clipboard", "timers", "network"]
}

Field details:

  id  (REQUIRED)
      Unique snake_case identifier.  Used for install path, settings namespace,
      and log prefix.  PERMANENT — never change after releasing the plugin, as
      doing so will break user installations and settings.

  uuid  (STRONGLY RECOMMENDED)
      A standard UUID v4 string ("xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx").
      Enables in-place plugin updates: when a user installs a .ccplug whose
      uuid matches an already-installed plugin, Command Center automatically
      replaces only the code files (manifest.json + plugin.py) while
      preserving all runtime data files (keys, caches, configs, etc.) and
      then reloads the plugin.  Without a uuid, reinstalling the same plugin
      is treated as a fresh install.
      Generate with:  python -c "import uuid; print(uuid.uuid4())"
      PERMANENT — never change the uuid across releases.

  name  (REQUIRED)
      Human-readable display name shown in the Plugins dialog and toasts.
      Spaces and mixed case are fine.

  version  (REQUIRED)
      Version string "MAJOR.MINOR.PATCH.BUILD".  Displayed in the Plugins
      dialog.  When uuid is present, bump the version on every release so
      users can confirm the update was applied.

  description  (REQUIRED)
      One or two sentences shown below the plugin name in the Plugins dialog.
      Keep it concise — users read this to decide if they want the plugin.

  author  (optional)
      Your name or handle.  Displayed in the Plugins dialog.

  min_app_version  (optional, advisory)
      Minimum Command Center version required.  Displayed but not enforced.

  entry  (optional, default: "plugin.py")
      The Python file inside the archive that contains activate() / deactivate().

  permissions  (optional, informational)
      Labels telling users what the plugin accesses.  Does NOT restrict the API
      — all API methods are always available regardless of what is listed here.
      Use this to be transparent with users.  Valid values:
        "ui"        — footer buttons, dialogs, settings tabs
        "hotkeys"   — global keyboard shortcuts
        "notebook"  — reads or writes notebook notes
        "nodes"     — reads or modifies canvas nodes
        "settings"  — reads or writes plugin-specific settings
        "files"     — reads or writes local files
        "clipboard" — reads or writes the system clipboard
        "timers"    — creates background QTimers
        "network"   — makes HTTP or other network calls


================================================================================
  5. plugin.py — ENTRY-POINT CONTRACT
================================================================================

Your plugin.py MUST define these two top-level functions:

--------------------------------------
  def activate(api):
--------------------------------------
  Called when the plugin is enabled.  Receives the CommandCenterAPI object.
  Store api in a module-level variable and register all your side effects here.

  Lifecycle order:
    1. Plugin manager imports plugin.py
    2. activate(api) is called
    3. Plugin is now live — footer buttons visible, hotkeys registered, etc.

  Avoid:
    • Blocking I/O or sleep() on the main thread
    • Opening dialogs immediately on activate (the window may not be ready)
    • Importing heavy libraries at module level — import inside functions instead

--------------------------------------
  def deactivate():
--------------------------------------
  Called when the plugin is disabled, reloaded, or uninstalled.

  Must undo ALL side effects:
    • Remove every footer button with  api.ui.remove_footer_button(btn)
    • Unregister every hotkey with  api.hotkeys.unregister(sequence)
    • Remove every settings tab with  api.ui.remove_settings_tab(label)
    • Timers created via  api.timers  are auto-stopped — you do not need to
      cancel them manually.
    • Stop any external processes you launched (subprocess.Popen)
    • Release any other resources

  deactivate() may be called even if activate() never ran (e.g., on Reload
  before the first Enable).  Always guard with  if var is not None:  checks.

--------------------------------------
  Recommended module-level pattern:
--------------------------------------
  # Globals — one per managed resource
  _api  = None     # the API object
  _btn  = None     # footer button handle
  _tab  = None     # settings tab label (str)
  _proc = None     # external subprocess

  def activate(api):
      global _api, _btn, _tab
      _api = api
      _btn = api.ui.add_footer_button("My Tool", _open_panel)
      api.hotkeys.register("Ctrl+Shift+P", _open_panel, "Open My Panel")
      _tab = "My Plugin"
      api.ui.add_settings_tab(_tab, _build_settings_widget())
      api.theme.register(_on_theme_change)
      api.log("Activated.")

  def deactivate():
      global _btn, _tab
      if _btn is not None:
          _api.ui.remove_footer_button(_btn)
          _btn = None
      _api.hotkeys.unregister("Ctrl+Shift+P")
      if _tab is not None:
          _api.ui.remove_settings_tab(_tab)
          _tab = None
      _api.log("Deactivated.")

  Note: You do NOT need to un-register theme callbacks — they are cleaned up
  automatically when the plugin is deactivated.


================================================================================
  6. THE CommandCenterAPI — FULL REFERENCE
================================================================================

  Namespaces:
    api.nodes      — canvas node management
    api.notebook   — notebook note management
    api.settings   — per-plugin persistent settings
    api.theme      — theme colors and change events
    api.ui         — footer buttons, dialogs, input prompts, settings tabs
    api.hotkeys    — global keyboard shortcuts
    api.clipboard  — system clipboard (text & images)
    api.timers     — managed auto-cleanup QTimers
    api.files      — local file read/write helpers
    api.plugin_dir — pathlib.Path to your plugin's install directory
    api.config_dir — pathlib.Path to the app config root (read-only)
    api.app_version — current app version string (e.g. "1.0.0.0")

  Top-level methods:
    api.toast(msg, kind)     — on-screen notification
    api.log(msg)             — debug print with plugin ID prefix
    api.launch(command, args) — run external program
    api.open_url(url)        — open URL in default browser


--------------------------------------------------------------------------------
  6a. CORE / UTILITIES
--------------------------------------------------------------------------------

  api.app_version  →  str
      Current version string, e.g. "1.0.0.0".

  api.plugin_dir  →  pathlib.Path
      Your plugin's install directory.  Always use this for data files.
      Example:
        config = api.plugin_dir / "config.json"
        log    = api.plugin_dir / "run.log"

  api.config_dir  →  pathlib.Path
      Application config root.  Treat as read-only — do not write here.

  api.toast(msg, kind="info")
      Show a brief on-screen notification.  Appears in the bottom-right corner
      and fades automatically.
        kind: "info" | "success" | "warn" | "error"
      Examples:
        api.toast("Saved!", "success")
        api.toast("File not found.", "error")

  api.log(msg)
      Print a debug message to stderr prefixed with your plugin ID.
        api.log("Starting scan...")
        # prints: [Plugin:my_plugin] Starting scan...
      Use this instead of print() so log lines are identifiable.

  api.launch(command, args=[])  →  subprocess.Popen
      Launch an external executable asynchronously.  Returns a Popen object you
      can use to check status or terminate later.
      Examples:
        api.launch("notepad.exe")
        api.launch("explorer.exe", [str(api.plugin_dir)])
        proc = api.launch("my_tool.exe", ["--flag", "value"])

  api.open_url(url)
      Open a URL or file path in the default application (browser for https://,
      Explorer for file:// or folder paths, etc.)
      Examples:
        api.open_url("https://example.com")
        api.open_url(str(api.plugin_dir))


--------------------------------------------------------------------------------
  6b. NODES (CANVAS)
--------------------------------------------------------------------------------

  Each node is a plain Python dict.  Fields:

    "id"            int     Unique identifier.  Auto-assigned on creation.
    "name"          str     Tile label text.
    "type"          str     "file" | "url" | "note" | "folder"
    "target"        str     Path, URL, note ID, or "" for folders.
    "size"          str     "1x1" | "2x2" | "2x4"
    "description"   str     Tooltip / subtitle text.
    "accent_color"  str     Hex tile accent color, e.g. "#4fc3f7"
    "icon"          str     Absolute path to a custom icon image, or "".
    "auto_launch"   bool    If True, launches on app startup.
    "archived"      bool    Hidden from canvas if True.
    "folder_id"     int|None  Folder this node belongs to, or None (root).
    "open_behavior" str     "normal" | "admin" | "folder" |
                            "copy_file" | "copy_contents"
    "tags"          list    List of str tags for search.

  --- Querying ---

  api.nodes.list()  →  list[dict]
      Active (non-archived), root-level nodes visible on the canvas.
      Does NOT include archived nodes or nodes inside folders by default.

  api.nodes.all(include_archived=False)  →  list[dict]
      Every node including folder tiles.  Pass include_archived=True to also
      include archived nodes.
      Example: all_nodes = api.nodes.all(include_archived=True)

  api.nodes.get(node_id)  →  dict | None
      Look up a single node by its integer id.  Returns None if not found.
      Example:
        node = api.nodes.get(42)
        if node:
            api.log(f"Found: {node['name']}")

  --- Creating ---

  api.nodes.add(node_dict)  →  dict
      Add a new node to the canvas.  Required fields: "name" and "type".
      Returns the saved dict with its assigned "id".
      Example:
        new_node = api.nodes.add({
            "name":         "GitHub",
            "type":         "url",
            "target":       "https://github.com",
            "size":         "1x1",
            "accent_color": "#4fc3f7",
            "description":  "My GitHub profile",
        })
        api.log(f"Created node id={new_node['id']}")

  --- Updating ---

  api.nodes.update(node_id, updates_dict)  →  dict
      Update one or more fields of an existing node.  Returns the full updated
      node dict.  You only need to pass the fields you want to change.
      Example:
        api.nodes.update(node["id"], {
            "name":         "GitHub (updated)",
            "accent_color": "#ff5722",
        })

  --- Deleting & Archiving ---

  api.nodes.delete(node_id)
      Permanently delete a node.  This cannot be undone.
      Example: api.nodes.delete(node["id"])

  api.nodes.archive(node_id)
      Archive a node.  It disappears from the canvas but can be recovered from
      Settings → Archived Nodes.  Safer than delete.
      Example: api.nodes.archive(node["id"])

  api.nodes.unarchive(node_id)
      Restore an archived node back to the canvas.

  --- Folders ---

  api.nodes.list_folders()  →  list[dict]
      Return all active folder tile dicts.

  api.nodes.move_to_folder(node_id, folder_id)
      Move a node into a folder.  Pass folder_id=None to move it to root.
      Example:
        folder = api.nodes.list_folders()[0]
        api.nodes.move_to_folder(node["id"], folder["id"])

  --- Canvas Refresh ---

  api.nodes.reload()
      Force a full reload and re-render of the canvas.  Needed only if you
      make many changes and want them all visible at once without waiting for
      the automatic refresh.

  api.nodes.on_loaded(fn)
      Register a callback that fires every time the canvas reloads.  fn takes
      no arguments.  Use this to re-build any canvas-dependent state.
      Example:
        api.nodes.on_loaded(lambda: api.log("Canvas reloaded."))

  --- Common Pattern: Finding and Modifying Nodes ---

    def _tag_all_files():
        for node in _api.nodes.all():
            if node["type"] == "file" and "work" not in node.get("tags", []):
                tags = list(node.get("tags", []))
                tags.append("work")
                _api.nodes.update(node["id"], {"tags": tags})
        _api.toast("Tagged all file nodes.", "success")


--------------------------------------------------------------------------------
  6c. NOTEBOOK
--------------------------------------------------------------------------------

  Note summary dict (returned by list_notes / search):
    "id"       str     Unique note ID (e.g. "20260428_085405_286488")
    "title"    str     Note title.
    "modified" str     ISO 8601 timestamp, e.g. "2026-04-28T08:54:05"
    "color"    str     Hex accent color or "" (no color).
    "pinned"   bool    True if the note is pinned to the top of the list.

  Full note dict (returned by load_note):
    Same as summary + "content" (str, the full HTML content of the note).

  api.notebook.list_notes()  →  list[dict]
      All note summaries, sorted by pinned-first then most-recently-modified.

  api.notebook.load_note(note_id)  →  dict
      Full note dict including "content" (HTML string).
      Example:
        note = _api.notebook.load_note(some_id)
        html = note["content"]

  api.notebook.save_note(note_id, title, html)
      Overwrite an existing note's title and HTML content.
      Example:
        _api.notebook.save_note(note["id"], "New Title", "<p>New content</p>")

  api.notebook.new_note(title, html="")  →  str
      Create a new note.  Returns the new note ID string.
      Example:
        note_id = _api.notebook.new_note("Meeting Notes", "<p>Details...</p>")
        _api.toast(f"Created note {note_id}", "success")

  api.notebook.delete_note(note_id)
      Permanently delete a note.  There is no undo.

  api.notebook.update_meta(note_id, **kwargs)
      Update a note's metadata without changing its content.
      Accepted kwargs: color (str, hex), pinned (bool)
      Example:
        _api.notebook.update_meta(note_id, color="#4fc3f7", pinned=True)

  api.notebook.open_note(note_id)
      Open the Notebook window and jump directly to the specified note.
      Useful for deep-linking from a dialog or footer button.
      Example:
        _api.notebook.open_note(note_id)

  api.notebook.search(query)  →  list[dict]
      Search all notes whose title or content contains the query string
      (case-insensitive).  Returns a list of note summary dicts.
      Example:
        results = _api.notebook.search("meeting")
        for r in results:
            api.log(r["title"])

  api.notebook.on_saved(fn)
      Register a callback fired whenever any note is saved.
      fn receives one argument: the saved note_id string.
      Example:
        def _note_saved(note_id):
            api.log(f"Note {note_id} was saved.")
        api.notebook.on_saved(_note_saved)


--------------------------------------------------------------------------------
  6d. SETTINGS
--------------------------------------------------------------------------------

  Plugin settings are stored under QSettings key "plugins/<your_id>/<key>".
  They persist across sessions and survive plugin reloads.

  api.settings.value(key, default=None)
      Read a stored value.  Returns default if the key does not exist.
      Example:
        count = int(api.settings.value("run_count", 0))

  api.settings.set(key, value)
      Write and persist a value.  Supported types: str, int, float, bool.
      For complex data, serialize with json.dumps/loads.
      Example:
        api.settings.set("run_count", count + 1)
        api.settings.set("last_path", str(api.plugin_dir / "data.json"))

  api.settings.remove(key)
      Delete a stored key.
      Example: api.settings.remove("temp_cache")

  api.settings.all_keys()  →  list[str]
      Return all key names stored under your plugin's namespace.
      Useful for backup/export of plugin state.

  api.settings.on_changed(fn)
      Callback fired after the user saves the Settings dialog.
      fn takes no arguments.  Use this to reload settings from widgets.
      Example:
        api.settings.on_changed(_reload_from_settings)

  Storing dicts and lists:
    import json
    api.settings.set("watched_nodes", json.dumps([1, 2, 3]))
    ids = json.loads(api.settings.value("watched_nodes", "[]"))

  Best practice — write a settings helper:
    def _save_config():
        api.settings.set("interval", _interval)
        api.settings.set("enabled",  _enabled)

    def _load_config():
        global _interval, _enabled
        _interval = int(api.settings.value("interval", 30))
        _enabled  = api.settings.value("enabled", "true") == "true"


--------------------------------------------------------------------------------
  6e. THEME / APPEARANCE
--------------------------------------------------------------------------------

  api.theme.name()  →  str
      Name of the active theme, e.g. "Deep Space".

  api.theme.available_themes()  →  list[str]
      Names of all built-in themes.

  api.theme.colors()  →  dict[str, str]
      All current theme colors as hex strings.  Keys:
        "bg_dark"        — darkest background (window/dialog background)
        "bg_mid"         — mid background (panels, cards)
        "glow"           — primary accent/glow color
        "text_primary"   — main text
        "text_secondary" — secondary text (subtitles, labels)
        "text_dim"       — dim text (placeholders, hints)
        "accent_blue"
        "accent_teal"
        "accent_amber"
        "accent_red"

  api.theme.set_theme(name)
      Switch to a named built-in theme.  Applies app-wide immediately.
      Example: api.theme.set_theme("Midnight Blue")

  api.theme.register(fn)
      Register a callback fired whenever the theme changes (user picks a new
      theme or adjusts brightness).  Use this to re-style your widgets.
      fn takes no arguments.
      Example:
        def _on_theme():
            c = _api.theme.colors()
            my_widget.setStyleSheet(
                f"background:{c['bg_dark']}; color:{c['text_primary']};"
            )
        api.theme.register(_on_theme)

  QColor accessors (return PySide6 QColor):
    api.theme.glow_color()
    api.theme.accent_color()     — alias for glow_color()
    api.theme.bg_dark()
    api.theme.bg_mid()
    api.theme.text_primary()
    api.theme.text_secondary()
    api.theme.text_dim()
    api.theme.accent_blue()
    api.theme.accent_teal()
    api.theme.accent_amber()
    api.theme.accent_red()

  --- Styling dialogs to match the current theme ---

  Always read  api.theme.colors()  at paint time rather than caching colors,
  because the user may change the theme while your dialog is open.

  Minimal dialog style:
    c = _api.theme.colors()
    dlg.setStyleSheet(
        f"QDialog   {{ background:{c['bg_dark']}; color:{c['text_primary']}; }}"
        f"QLabel     {{ color:{c['text_primary']}; }}"
        f"QPushButton{{ background:{c['bg_mid']}; color:{c['text_primary']}; "
        f"             border:1px solid {c['glow']}; border-radius:4px; }}"
        f"QPushButton:hover{{ border:1px solid {c['accent_blue']}; }}"
    )

  Full live-theming pattern:
    class MyDialog(QDialog):
        def __init__(self, parent):
            super().__init__(parent)
            self._build_ui()
            _api.theme.register(self._apply_theme)
            self._apply_theme()

        def _apply_theme(self):
            c = _api.theme.colors()
            self.setStyleSheet(
                f"QDialog {{ background:{c['bg_dark']}; "
                f"           color:{c['text_primary']}; }}"
            )


--------------------------------------------------------------------------------
  6f. HOTKEYS
--------------------------------------------------------------------------------

  api.hotkeys.register(sequence, fn, description="")
      Register a global keyboard shortcut.  Fires at application scope
      regardless of which window has focus.  If the sequence conflicts with a
      built-in Command Center shortcut, the built-in is automatically suppressed
      while your plugin is active and restored when it is deactivated.
        sequence    — Qt key sequence string, e.g. "Ctrl+Shift+P"
        fn          — callable with no arguments
        description — optional human-readable description for the Plugins dialog
      Example:
        api.hotkeys.register("Ctrl+Shift+P", _open_panel, "Open My Panel")

  api.hotkeys.unregister(sequence)
      Remove a previously registered shortcut and restore any suppressed
      built-in.  Always call this in deactivate() for every registered hotkey.
      Example: api.hotkeys.unregister("Ctrl+Shift+P")

  api.hotkeys.list_all()  →  dict
      Returns a merged dict of ALL active keyboard shortcuts (built-in CC
      shortcuts + all plugin-registered shortcuts), grouped by source.
      Format: { "built_in": [...], "plugins": {"plugin_id": [...], ...} }
      Use this to check for conflicts before registering a hotkey.
      Example:
        hotkeys = api.hotkeys.list_all()
        all_seqs = [h["sequence"] for h in hotkeys.get("built_in", [])]
        for pid, hs in hotkeys.get("plugins", {}).items():
            all_seqs += [h["sequence"] for h in hs]
        if "Ctrl+Shift+P" not in all_seqs:
            api.hotkeys.register("Ctrl+Shift+P", _fn, "My shortcut")

  Key sequences format:
    Standard: "Ctrl+S", "Alt+F4", "Ctrl+Shift+P", "F5"
    Platform note: Use "Ctrl" not "Control"; use "Shift" not "Shift+".
    Keep sequences unique — conflicts with OTHER plugins are not auto-resolved.

  Always call unregister in deactivate:
    def deactivate():
        _api.hotkeys.unregister("Ctrl+Shift+P")
        # ...


--------------------------------------------------------------------------------
  6g. UI — FOOTER BUTTONS, DIALOGS & INPUT
--------------------------------------------------------------------------------

  api.ui.add_footer_button(label, fn)  →  QPushButton
      Add a button to the footer bar.  The button appears immediately to the
      right of the built-in footer buttons.
      Returns the QPushButton — save it so you can remove it later.
      Example:
        _btn = api.ui.add_footer_button("My Tool", _on_click)

  api.ui.remove_footer_button(btn)
      Remove a footer button.  Pass the QPushButton returned by add_footer_button.
      Always call in deactivate().
      Example:
        if _btn is not None:
            api.ui.remove_footer_button(_btn)
            _btn = None

  api.ui.add_settings_tab(label, widget)
      Inject a custom QWidget tab into the Settings dialog.  The tab appears
      in the Settings dialog as long as the plugin is enabled.
      label  — the tab title string shown on the tab bar
      widget — a QWidget you have already built (with layout, controls, etc.)
      Example:
        tab_widget = _build_settings_panel()
        api.ui.add_settings_tab("My Plugin", tab_widget)

  api.ui.remove_settings_tab(label)
      Remove the settings tab with the given label.  Call in deactivate().
      Example: api.ui.remove_settings_tab("My Plugin")

  api.ui.show_dialog(widget)
      Show any QWidget as a floating window centered on the main window.
      Useful for simple non-modal panels.

  api.ui.ask_confirm(title, message)  →  bool
      Show a Yes/No modal dialog.  Returns True if the user clicked Yes.
      Example:
        if api.ui.ask_confirm("Delete", "Delete all generated notes?"):
            for note in notes:
                _api.notebook.delete_note(note["id"])

  api.ui.ask_input(title, prompt, default="")  →  str | None
      Show a single-line text input dialog.
      Returns the entered string, or None if the user cancelled.
      Example:
        name = api.ui.ask_input("Create Node", "Node name:", "New Node")
        if name:
            api.nodes.add({"name": name, "type": "url", "target": ""})

  api.ui.show_message(title, message, kind="info")
      Show a modal message box.  kind: "info" | "warn" | "error"
      Example:
        api.ui.show_message("Error", "Could not load config.", "error")

  api.ui.refresh_canvas()
      Reload and re-render all canvas tiles.  Use after bulk node changes.

  api.ui.main_window  →  QMainWindow
      The main application window.  Use as parent for your dialogs so they
      inherit the app's window icon and center correctly.
      Example:
        dlg = MyDialog(parent=api.ui.main_window)
        dlg.exec()

  --- Building Custom Dialogs ---

  Use standard PySide6.  Import freely from PySide6.QtWidgets, QtCore, QtGui.

  Template:
    from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
                                   QLabel, QPushButton, QTextEdit)
    from PySide6.QtCore import Qt

    class MyDialog(QDialog):
        def __init__(self, parent):
            super().__init__(parent)
            self.setWindowTitle("My Tool")
            self.setMinimumSize(500, 400)
            self._build_ui()
            self._apply_theme()
            _api.theme.register(self._apply_theme)

        def _build_ui(self):
            layout = QVBoxLayout(self)

            self._label = QLabel("Content goes here.")
            layout.addWidget(self._label)

            btn_row = QHBoxLayout()
            close_btn = QPushButton("Close")
            close_btn.clicked.connect(self.accept)
            btn_row.addStretch()
            btn_row.addWidget(close_btn)
            layout.addLayout(btn_row)

        def _apply_theme(self):
            c = _api.theme.colors()
            self.setStyleSheet(
                f"QDialog     {{ background:{c['bg_dark']}; "
                f"               color:{c['text_primary']}; }}"
                f"QLabel       {{ color:{c['text_primary']}; }}"
                f"QPushButton  {{ background:{c['bg_mid']}; "
                f"               color:{c['text_primary']}; "
                f"               border:1px solid {c['glow']}; "
                f"               border-radius:4px; padding:4px 12px; }}"
                f"QPushButton:hover {{ border-color:{c['accent_blue']}; }}"
            )

  To open it:
    def _open_panel():
        dlg = MyDialog(parent=_api.ui.main_window)
        dlg.exec()   # modal (blocks)
        # or: dlg.show()  for non-modal


--------------------------------------------------------------------------------
  6h. CLIPBOARD
--------------------------------------------------------------------------------

  api.clipboard.get_text()  →  str
      Current clipboard text, or "" if the clipboard holds no text.

  api.clipboard.set_text(text)
      Copy a string to the clipboard.
      Example:
        api.clipboard.set_text("copied content")
        api.toast("Copied!", "success")

  api.clipboard.get_image()  →  QPixmap
      Current clipboard image as a QPixmap.  Returns a null QPixmap if the
      clipboard holds no image (check with  pixmap.isNull()).

  api.clipboard.set_image(pixmap)
      Copy a QPixmap to the clipboard.


--------------------------------------------------------------------------------
  6i. TIMERS
--------------------------------------------------------------------------------

  api.timers.create(interval_ms, fn, single_shot=False)  →  QTimer
      Create and start a timer.  All timers created via this method are
      automatically stopped and cleaned up when the plugin is deactivated —
      you do NOT need to cancel them in deactivate().
        interval_ms  — milliseconds between fires
        fn           — callable with no arguments
        single_shot  — if True, fires once then stops

      Examples:
        # Repeat every 5 seconds:
        _poll_timer = api.timers.create(5000, _poll_server)

        # Fire once after a 500ms delay:
        api.timers.create(500, lambda: api.toast("Ready!", "info"), single_shot=True)

  api.timers.cancel(timer)
      Stop a specific timer before it would naturally stop.
      Example: api.timers.cancel(_poll_timer)

  Threading note:
    QTimers run on the main thread — fn() is called on the main thread.
    For background work, use Python's threading.Thread or concurrent.futures,
    but always update the UI only from the main thread (use a QTimer single-shot
    with interval 0 to schedule a function on the main thread from a background
    thread, or use Qt signals/slots).


--------------------------------------------------------------------------------
  6j. FILE I/O
--------------------------------------------------------------------------------

  api.files.read_text(path)  →  str
      Read a file as UTF-8 text.  Raises OSError if the file does not exist.
      Example: text = api.files.read_text(api.plugin_dir / "notes.txt")

  api.files.write_text(path, content)
      Write UTF-8 text to a file.  Creates parent directories if needed.
      Example: api.files.write_text(api.plugin_dir / "log.txt", log_content)

  api.files.read_json(path)  →  dict | list
      Read and parse a JSON file.  Raises OSError / json.JSONDecodeError.
      Example:
        try:
            data = api.files.read_json(api.plugin_dir / "config.json")
        except (OSError, ValueError):
            data = {}

  api.files.write_json(path, data)
      Serialize data to pretty-printed JSON and write to a file.
      Example:
        api.files.write_json(api.plugin_dir / "config.json", {"key": "value"})

  api.files.list_dir(path, pattern="*")  →  list[pathlib.Path]
      List files matching a glob pattern inside a directory.
      Example:
        py_files = api.files.list_dir(api.plugin_dir, "*.py")

  api.files.open_dialog(title, filters)  →  str | None
      Show a native file-open dialog.  Returns the selected path or None.
      filters — Qt filter string: "Text (*.txt);;All Files (*)"
      Example:
        path = api.files.open_dialog("Select Script", "Scripts (*.py *.bat)")
        if path:
            process_file(path)

  api.files.save_dialog(title, filters)  →  str | None
      Show a native file-save dialog.  Returns the chosen path or None.
      Example:
        path = api.files.save_dialog("Save Report", "Text (*.txt)")
        if path:
            api.files.write_text(path, report_content)


--------------------------------------------------------------------------------
  6k. EXTERNAL PROCESSES
--------------------------------------------------------------------------------

  api.launch(command, args=[])  →  subprocess.Popen
      Launch an external executable asynchronously.  CC does NOT manage
      the subprocess lifetime — you are responsible for terminating it in
      deactivate() if it should not outlive the plugin.
      Example:
        _proc = api.launch("my_watcher.exe", ["--mode", "passive"])

        def deactivate():
            if _proc and _proc.poll() is None:
                _proc.terminate()

  api.open_url(url)
      Open a URL or path with the default application.


================================================================================
  7. FULL WORKING EXAMPLE
================================================================================

This example creates a "Note Counter" plugin with:
  • A footer button
  • A Ctrl+Shift+W hotkey
  • A styled dialog with a clipboard copy button
  • A launch count stored in settings

--- manifest.json ---------------------------------------------------------------

{
  "id":          "note_counter",
  "name":        "Note Counter",
  "version":     "1.0.0.0",
  "description": "Shows a count of your notebook notes in a dialog.",
  "author":      "Your Name",
  "permissions": ["ui", "hotkeys", "notebook", "clipboard", "settings"]
}

--- plugin.py -------------------------------------------------------------------

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
                                QLabel, QPushButton)
from PySide6.QtCore import Qt

_api = None
_btn = None

def activate(api):
    global _api, _btn
    _api = api

    _btn = api.ui.add_footer_button("Note Counter", _show_dialog)
    api.hotkeys.register("Ctrl+Shift+W", _show_dialog, "Open Note Counter")

    count = int(api.settings.value("launch_count", 0)) + 1
    api.settings.set("launch_count", count)
    api.log(f"Activated — launch #{count}")

def deactivate():
    global _btn
    if _btn is not None:
        _api.ui.remove_footer_button(_btn)
        _btn = None
    _api.hotkeys.unregister("Ctrl+Shift+W")
    _api.log("Deactivated.")

def _show_dialog():
    try:
        notes = _api.notebook.list_notes()
        dlg = _NoteCountDialog(
            parent=_api.ui.main_window,
            note_count=len(notes)
        )
        dlg.exec()
    except Exception as e:
        _api.log(f"Error: {e}")
        _api.toast("An error occurred.", "error")

class _NoteCountDialog(QDialog):
    def __init__(self, parent, note_count):
        super().__init__(parent)
        self.setWindowTitle("Note Counter")
        self.setMinimumWidth(300)
        self._count = note_count
        self._build_ui()
        _api.theme.register(self._apply_theme)
        self._apply_theme()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        self._label = QLabel(f"You have {self._count} note(s).")
        self._label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self._label)

        btn_row = QHBoxLayout()
        copy_btn = QPushButton("Copy to Clipboard")
        copy_btn.clicked.connect(lambda: (
            _api.clipboard.set_text(f"{self._count} notes"),
            _api.toast("Copied!", "success"),
        ))
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(copy_btn)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

    def _apply_theme(self):
        c = _api.theme.colors()
        self.setStyleSheet(
            f"QDialog    {{ background:{c['bg_dark']}; "
            f"              color:{c['text_primary']}; }}"
            f"QLabel      {{ color:{c['text_primary']}; }}"
            f"QPushButton {{ background:{c['bg_mid']}; "
            f"              color:{c['text_primary']}; "
            f"              border:1px solid {c['glow']}; "
            f"              border-radius:4px; padding:4px 12px; }}"
            f"QPushButton:hover {{ border-color:{c['accent_blue']}; }}"
        )

--------------------------------------------------------------------------------

Packaging:
  python -c "
  import zipfile
  with zipfile.ZipFile('note_counter.ccplug', 'w') as z:
      z.write('manifest.json')
      z.write('plugin.py')
  print('Packed note_counter.ccplug')
  "


================================================================================
  8. PACKAGING YOUR PLUGIN AS .ccplug
================================================================================

A .ccplug file is a ZIP archive with a renamed extension.  Any ZIP tool works.

METHOD 1 — Python pack script (recommended, cross-platform):

  Save as pack.py in your project folder:

    import zipfile, pathlib, sys

    out_name = sys.argv[1] if len(sys.argv) > 1 else "plugin.ccplug"
    project  = pathlib.Path(".")
    files    = [f for f in project.rglob("*")
                if f.is_file() and not f.suffix == ".ccplug"]

    with zipfile.ZipFile(out_name, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, f.relative_to(project))
    print(f"Packed {len(files)} files → {out_name}")

  Run:  python pack.py my_plugin.ccplug

METHOD 2 — PowerShell (Windows):
  Compress-Archive -Path manifest.json,plugin.py -DestinationPath plugin.zip
  Rename-Item plugin.zip my_plugin.ccplug

INSTALLING
  Footer → Plugins → '+ Install Plugin' → select .ccplug → Enable

INSTALL PATH
  Files are extracted to:  %APPDATA%\\CommandCenter\\plugins\\<your_id>\\
  (or  <script_dir>\\data\\plugins\\<your_id>\\  if the app is running under
  the Windows Store Python package detection fallback)

HOT-RELOAD DEVELOPMENT WORKFLOW
  1. Install your .ccplug once via the Plugins dialog.
  2. Edit files directly in the install directory:
       %APPDATA%\\CommandCenter\\plugins\\<id>\\
  3. Click Reload in the Plugins dialog — no reinstall needed.
  4. When finished, re-zip to produce the final distributable .ccplug.


================================================================================
  9. BEST PRACTICES & COMMON PATTERNS
================================================================================

--------------------------------------
  MUST: implement deactivate() fully
--------------------------------------
  Every resource you acquired in activate() must be released in deactivate():
    • footer button   → api.ui.remove_footer_button(btn)
    • hotkey          → api.hotkeys.unregister(sequence)
    • settings tab    → api.ui.remove_settings_tab(label)
    • subprocess      → proc.terminate() if proc.poll() is None
    • Timers via api.timers — cleaned up automatically, no action needed.

--------------------------------------
  Store the API in a module global
--------------------------------------
  _api = None
  def activate(api):
      global _api
      _api = api
  # Now all module-level functions can reference _api

--------------------------------------
  Use api.plugin_dir for ALL data files
--------------------------------------
  config_path = _api.plugin_dir / "config.json"  # correct
  Never hardcode absolute paths.

--------------------------------------
  Style with api.theme.colors()
--------------------------------------
  Register a theme callback so your dialog automatically re-styles when
  the user changes the theme:
    api.theme.register(_apply_theme)
  Always call _apply_theme() once manually after building the UI.

--------------------------------------
  Guard all callbacks with try/except
--------------------------------------
  Exceptions inside Qt callbacks (button clicks, timer fires) can be
  silently swallowed by the event loop.  Always wrap:
    def _on_click():
        try:
            do_work()
        except Exception as e:
            _api.log(f"Error in _on_click: {e}")
            _api.toast("An error occurred.", "error")

--------------------------------------
  Do NOT block the main thread
--------------------------------------
  Never call sleep(), requests.get() (synchronous), or any blocking I/O
  directly in a callback.  Use:
    • api.timers.create() for periodic background checks
    • threading.Thread for heavy/blocking work; communicate results back via
      a QTimer.singleShot(0, fn) call to marshal back to the main thread

--------------------------------------
  Use api.log() instead of print()
--------------------------------------
  Output is prefixed with your plugin ID:
    [Plugin:my_plugin] message
  Easy to grep in terminal output during development.

--------------------------------------
  Safe deactivate() pattern
--------------------------------------
  deactivate() may be called before activate() (e.g., on a Reload cycle).
  Always check for None:
    def deactivate():
        global _btn
        if _btn is not None:
            _api.ui.remove_footer_button(_btn)
            _btn = None

--------------------------------------
  Settings schema versioning
--------------------------------------
  Store a "settings_version" key so you can migrate old settings on upgrade:
    version = int(api.settings.value("settings_version", 0))
    if version < 1:
        # migrate from v0 format
        api.settings.set("settings_version", 1)

--------------------------------------
  Conditional imports for optional deps
--------------------------------------
  If your plugin uses an optional third-party library:
    try:
        import requests
        _HAS_REQUESTS = True
    except ImportError:
        _HAS_REQUESTS = False
  Then check _HAS_REQUESTS before using it and show a helpful toast if missing.


================================================================================
  10. TROUBLESHOOTING
================================================================================

"Entry file 'plugin.py' missing."
  The file listed in manifest.json "entry" field is not in the archive.
  Verify the filename exactly matches.

"manifest.json must have a non-empty 'id' field."
  The "id" field is missing or empty.  It must be a non-empty snake_case string.

"Import error: ..."
  Python syntax error or missing import in your plugin.py.  Run locally:
    python -c "import py_compile; py_compile.compile('plugin.py')"
  Also check that helper modules are included in the archive and spelled
  correctly in import statements.

"activate() raised: ..."
  Your activate() function threw an exception.  Read the full traceback in
  the terminal and add try/except blocks to isolate the problem.

My footer button did not appear
  Call add_footer_button() inside activate(), not at module import time.
  Verify activate() ran without errors in the terminal.

My hotkey does not fire
  Check for conflicts with existing shortcuts using api.hotkeys.list_all().
  Hotkeys registered via the plugin API suppress conflicting built-ins
  automatically.  Conflicts with OTHER plugins are not resolved automatically
  — choose unique sequences.

My dialog is not themed / looks wrong
  Build the stylesheet inside a function and call it from api.theme.register().
  Use api.theme.colors() to get current colors — do not cache them at startup.

My plugin works but Reload causes an error
  Ensure deactivate() is safe to call even if activate() never ran (None-check
  every module global before using it).

Plugin works locally but fails after reinstall
  Make sure all helper .py files are included in the .ccplug archive and that
  import paths match file names exactly (case-sensitive on some systems).

Settings are not persisting
  Check that you are calling api.settings.set() (not just writing a local var).
  Use api.settings.all_keys() to inspect what is actually stored.

Background timer fires after plugin is disabled
  Only timers created via api.timers.create() are auto-stopped.  If you created
  a QTimer manually, you must call timer.stop() in deactivate().

================================================================================
  END OF GUIDE
  Command Center Plugin System  •  Version 1.0.0.0+
================================================================================
"""


# ---------------------------------------------------------------------------
# Plugin API — the surface exposed to plugin code
# ---------------------------------------------------------------------------
class _PluginSettingsProxy:
    """Wraps QSettings with a per-plugin namespace prefix."""

    def __init__(self, plugin_id: str):
        self._prefix = f"plugins/{plugin_id}/"

    def value(self, key: str, default=None):
        return _settings_store.value(self._prefix + key, default)

    def set(self, key: str, value):
        _settings_store.setValue(self._prefix + key, value)

    def remove(self, key: str):
        """Delete a stored key."""
        _settings_store.remove(self._prefix + key)

    def all_keys(self) -> list:
        """Return all key names stored under this plugin's namespace."""
        _settings_store.beginGroup(self._prefix.rstrip("/"))
        keys = list(_settings_store.childKeys())
        _settings_store.endGroup()
        return keys

    def on_changed(self, fn):
        # Registered globally; plugin manager calls all listeners after settings save
        _plugin_manager._settings_changed_listeners.append(fn)


class _PluginNotebookProxy:
    def __init__(self, plugin_id: str):
        self._pid = plugin_id
        self._on_saved_fns: list = []

    def list_notes(self) -> list:
        return NotebookStore.list_notes()

    def load_note(self, note_id: str) -> dict:
        return NotebookStore.load_note(note_id)

    def save_note(self, note_id: str, title: str, html: str):
        NotebookStore.save_note(note_id, title, html)
        for fn in self._on_saved_fns:
            try:
                fn(note_id)
            except Exception as exc:
                print(f"[Plugin:{self._pid}] notebook.on_saved callback error: {exc}",
                      file=sys.stderr)

    def new_note(self, title: str, html: str = "") -> str:
        note_id = NotebookStore.new_id()
        NotebookStore.save_note(note_id, title, html)
        return note_id

    def delete_note(self, note_id: str):
        NotebookStore.delete_note(note_id)

    def update_meta(self, note_id: str, **kwargs):
        NotebookStore.update_note_meta(note_id, **kwargs)

    def open_note(self, note_id: str):
        """Open the Notebook window and navigate to the given note."""
        mw = _plugin_manager._main_window
        if mw is None:
            return
        try:
            mw._open_notebook()
            nb = mw._notebook_win
            if nb is not None:
                nb._load_note(note_id)
                nb._side._select_item(note_id)
        except Exception as exc:
            print(f"[Plugin:{self._pid}] notebook.open_note error: {exc}", file=sys.stderr)

    def search(self, query: str) -> list:
        """Return note summaries whose title or content contains `query` (case-insensitive)."""
        q = query.lower()
        results = []
        for note in NotebookStore.list_notes():
            if q in note.get("title", "").lower():
                results.append(note)
                continue
            try:
                full = NotebookStore.load_note(note["id"])
                if q in full.get("content", "").lower():
                    results.append(note)
            except Exception:
                pass
        return results

    def on_saved(self, fn):
        self._on_saved_fns.append(fn)


class _PluginNodesProxy:
    def __init__(self, plugin_id: str, store: "NodeStore"):
        self._pid   = plugin_id
        self._store = store
        self._on_loaded_fns: list = []

    def list(self) -> list:
        """Return all active (non-archived) root-level items visible on the canvas."""
        return list(self._store.all_items())

    def all(self, include_archived: bool = False) -> list:
        """Return all nodes. Set include_archived=True to include archived nodes."""
        if include_archived:
            return list(self._store._data)
        return [n for n in self._store._data if not n.get("archived", False)]

    def get(self, node_id) -> Optional[dict]:
        """Return the node dict for the given ID, or None if not found."""
        for n in self._store._data:
            if n.get("id") == node_id:
                return dict(n)
        return None

    def add(self, node_dict: dict) -> dict:
        return self._store.add_node(node_dict)

    def update(self, node_id, updates: dict) -> dict:
        """Update fields on a node. Returns the updated node dict."""
        self._store.update_node(node_id, updates)
        return self.get(node_id) or {}

    def delete(self, node_id):
        self._store.remove_node(node_id)

    def archive(self, node_id):
        """Archive a node (hides it from canvas; recoverable from Settings)."""
        self._store.archive_node(node_id)

    def unarchive(self, node_id):
        """Restore an archived node back to the canvas."""
        self._store.unarchive_node(node_id)

    def list_folders(self) -> list:
        """Return all active folder tiles."""
        return list(self._store.all_folders())

    def move_to_folder(self, node_id, folder_id):
        """Move a node into a folder. Pass folder_id=None to move to root."""
        self._store.move_node_to_folder(node_id, folder_id)

    def reload(self):
        if _plugin_manager._main_window is not None:
            try:
                _plugin_manager._main_window._load_nodes()
            except Exception as exc:
                print(f"[Plugin:{self._pid}] nodes.reload error: {exc}", file=sys.stderr)

    def on_loaded(self, fn):
        self._on_loaded_fns.append(fn)

    # ── Scheduling helpers ────────────────────────────────────────────────────

    def get_schedule(self, node_id) -> dict:
        """Return the schedule dict for the given node ID, or {} if not set.

        Schedule dict fields::

            {
                "enabled":       bool,          # True = scheduler is active
                "type":          str,            # see schedule types below
                "time":          "HH:MM",        # 24-hour local time (most types)
                "date":          "YYYY-MM-DD",   # used by "once" type only
                "days":          [0..6],         # used by "weekly" (0=Mon…6=Sun)
                "day_of_month":  int,            # 1-31, used by "monthly"
                "interval_value": int,           # used by "interval"
                "interval_unit": "minutes"|"hours",
                "last_run":      str|null,       # ISO-8601 datetime of last fire
            }

        Schedule types: "daily", "workdays", "weekends", "weekly",
                        "monthly", "interval", "once".
        """
        node = self.get(node_id)
        if node is None:
            return {}
        sched = node.get("schedule", {})
        return dict(sched) if isinstance(sched, dict) else {}

    def set_schedule(self, node_id, schedule: dict):
        """Set or replace the schedule for a node.

        Pass an empty dict or {"enabled": False} to disable scheduling.
        The scheduler will pick up the change on its next 30-second poll.

        Example — run every day at 9 AM::

            api.nodes.set_schedule(node_id, {
                "enabled": True,
                "type": "daily",
                "time": "09:00",
            })

        Example — run every 30 minutes::

            api.nodes.set_schedule(node_id, {
                "enabled": True,
                "type": "interval",
                "interval_value": 30,
                "interval_unit": "minutes",
            })
        """
        if not isinstance(schedule, dict):
            raise TypeError("schedule must be a dict")
        self._store.update_node(node_id, {"schedule": schedule})


class _PluginThemeProxy:
    def name(self) -> str:
        return _settings_store.value("theme", "Deep Space")

    def available_themes(self) -> list:
        """Return the names of all built-in themes."""
        return [k for k in _BUILTIN_THEMES.keys() if k != "Custom"]

    def colors(self) -> dict:
        t = _theme
        return {
            "bg_dark":        t.BG_DARK.name(),
            "bg_mid":         t.BG_MID.name(),
            "glow":           t.GLOW.name(),
            "text_primary":   t.TEXT_PRIMARY.name(),
            "text_secondary": t.TEXT_SECONDARY.name(),
            "text_dim":       t.TEXT_DIM.name(),
            "accent_blue":    t.ACCENT_BLUE.name(),
            "accent_teal":    t.ACCENT_TEAL.name(),
            "accent_amber":   t.ACCENT_AMBER.name(),
            "accent_red":     t.ACCENT_RED.name(),
        }

    def set_theme(self, name: str):
        if name in _BUILTIN_THEMES:
            _settings_store.setValue("theme", name)
            _theme.load()
            if _plugin_manager._main_window:
                _plugin_manager._main_window._on_theme_changed()

    def register(self, fn):
        _theme.register(fn)

    # ── QColor accessors ──────────────────────────────────────────────────

    def glow_color(self) -> QColor:
        return QColor(_theme.GLOW)

    def accent_color(self) -> QColor:
        """Alias for glow_color() — the primary accent / glow color."""
        return QColor(_theme.GLOW)

    def bg_dark(self) -> QColor:
        return QColor(_theme.BG_DARK)

    def bg_mid(self) -> QColor:
        return QColor(_theme.BG_MID)

    def text_primary(self) -> QColor:
        return QColor(_theme.TEXT_PRIMARY)

    def text_secondary(self) -> QColor:
        return QColor(_theme.TEXT_SECONDARY)

    def text_dim(self) -> QColor:
        return QColor(_theme.TEXT_DIM)

    def accent_blue(self) -> QColor:
        return QColor(_theme.ACCENT_BLUE)

    def accent_teal(self) -> QColor:
        return QColor(_theme.ACCENT_TEAL)

    def accent_amber(self) -> QColor:
        return QColor(_theme.ACCENT_AMBER)

    def accent_red(self) -> QColor:
        return QColor(_theme.ACCENT_RED)


class _PluginFilesProxy:
    @staticmethod
    def read_text(path) -> str:
        return Path(path).read_text(encoding="utf-8")

    @staticmethod
    def write_text(path, content: str):
        p = Path(path)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(content, encoding="utf-8")

    @staticmethod
    def read_json(path):
        return json.loads(Path(path).read_text(encoding="utf-8"))

    @staticmethod
    def write_json(path, data):
        p = Path(path)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    @staticmethod
    def list_dir(path, pattern: str = "*") -> list:
        return list(Path(path).glob(pattern))

    @staticmethod
    def open_dialog(title: str = "Open File", filters: str = "All Files (*)") -> Optional[str]:
        path, _ = QFileDialog.getOpenFileName(None, title, "", filters)
        return path if path else None

    @staticmethod
    def save_dialog(title: str = "Save File", filters: str = "All Files (*)") -> Optional[str]:
        path, _ = QFileDialog.getSaveFileName(None, title, "", filters)
        return path if path else None


class _PluginUIProxy:
    def __init__(self, plugin_id: str):
        self._pid = plugin_id
        self._footer_btns: list = []
        self._settings_tabs: list = []  # list of (label, widget)
        self._shortcuts: list = []       # list of QShortcut

    @property
    def main_window(self):
        return _plugin_manager._main_window

    def add_footer_button(self, label: str, fn) -> "QPushButton":
        mw = _plugin_manager._main_window
        if mw is None:
            return None
        btn = mw._footer._make_btn(label, fn)
        # Add to the dedicated plugin scroll area — keeps built-in buttons
        # static and lets plugin buttons overflow gracefully.
        mw._footer._plugin_area.add_button(btn)
        self._footer_btns.append(btn)
        return btn

    def remove_footer_button(self, btn: "QPushButton"):
        if btn in self._footer_btns:
            mw = _plugin_manager._main_window
            if mw is not None:
                mw._footer._plugin_area.remove_button(btn)
            btn.setParent(None)
            btn.deleteLater()
            self._footer_btns.remove(btn)

    def add_settings_tab(self, label: str, widget: "QWidget"):
        self._settings_tabs.append((label, widget))
        # Live-inject into open settings dialogs
        _plugin_manager._inject_settings_tabs()

    def remove_settings_tab(self, label: str):
        self._settings_tabs = [(l, w) for (l, w) in self._settings_tabs if l != label]

    def show_dialog(self, widget: "QWidget"):
        mw = _plugin_manager._main_window
        if mw is not None:
            mw._center_dialog(widget)
        widget.show()
        widget.raise_()
        widget.activateWindow()

    def ask_confirm(self, title: str, message: str) -> bool:
        """Show a Yes/No question dialog. Returns True if the user clicked Yes."""
        from PySide6.QtWidgets import QMessageBox
        mw = _plugin_manager._main_window
        result = QMessageBox.question(
            mw, title, message,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return result == QMessageBox.StandardButton.Yes

    def ask_input(self, title: str, prompt: str, default: str = "") -> Optional[str]:
        """Show a single-line text input dialog. Returns the text or None if cancelled."""
        from PySide6.QtWidgets import QInputDialog
        mw = _plugin_manager._main_window
        text, ok = QInputDialog.getText(mw, title, prompt, text=default)
        return text if ok else None

    def show_message(self, title: str, message: str, kind: str = "info"):
        """Show a modal message box. kind: 'info', 'warn', or 'error'."""
        from PySide6.QtWidgets import QMessageBox
        mw = _plugin_manager._main_window
        if kind == "error":
            QMessageBox.critical(mw, title, message)
        elif kind == "warn":
            QMessageBox.warning(mw, title, message)
        else:
            QMessageBox.information(mw, title, message)

    def refresh_canvas(self):
        """Reload and re-render all canvas tiles."""
        mw = _plugin_manager._main_window
        if mw is not None:
            try:
                mw._load_nodes()
            except Exception as exc:
                print(f"[Plugin:{self._pid}] ui.refresh_canvas error: {exc}", file=sys.stderr)

    def cleanup(self):
        for btn in list(self._footer_btns):
            self.remove_footer_button(btn)
        self._footer_btns.clear()


class _PluginClipboardProxy:
    """Read and write the system clipboard."""

    @staticmethod
    def get_text() -> str:
        """Return current clipboard text (empty string if none)."""
        return QGuiApplication.clipboard().text()

    @staticmethod
    def set_text(text: str):
        """Copy text to the clipboard."""
        QGuiApplication.clipboard().setText(text)

    @staticmethod
    def get_image() -> "QPixmap":
        """Return current clipboard image as a QPixmap (null pixmap if none)."""
        return QGuiApplication.clipboard().pixmap()

    @staticmethod
    def set_image(pixmap: "QPixmap"):
        """Copy a QPixmap to the clipboard."""
        QGuiApplication.clipboard().setPixmap(pixmap)


class _PluginTimerProxy:
    """Managed QTimers that are automatically stopped when the plugin is deactivated."""

    def __init__(self, plugin_id: str):
        self._pid = plugin_id
        self._timers: list = []

    def create(self, interval_ms: int, fn, single_shot: bool = False) -> "QTimer":
        """Create and start a QTimer.
        interval_ms  — milliseconds between each fire
        fn           — callable with no arguments, called on each tick
        single_shot  — if True, fires only once then stops automatically
        Returns the QTimer (save it if you want to cancel it later).
        """
        t = QTimer()
        t.setInterval(interval_ms)
        t.setSingleShot(single_shot)
        t.timeout.connect(fn)
        t.start()
        self._timers.append(t)
        return t

    def cancel(self, timer: "QTimer"):
        """Stop and destroy a specific timer."""
        if timer in self._timers:
            timer.stop()
            timer.deleteLater()
            self._timers.remove(timer)

    def cleanup(self):
        for t in list(self._timers):
            self.cancel(t)


# Canonical Qt-style name of every built-in keyboard shortcut.
# Used by _PluginHotkeysProxy.list_all() and the AHK Bridge plugin.
_CC_BUILTIN_HOTKEYS: dict[str, str] = {
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
# User-configurable hotkey system  (action_id → default key sequence)
# ---------------------------------------------------------------------------

# Default key sequences for all user-configurable built-in actions.
# Keys are stored in QSettings as  hotkey/<action_id>
_CC_DEFAULT_HOTKEYS: dict[str, str] = {
    "toggle_search":     "Ctrl+F",
    "new_node":          "Ctrl+N",
    "new_folder":        "Ctrl+Shift+N",
    "open_settings":     "Ctrl+,",
    "open_help":         "F1",
    "time_tracker":      "Ctrl+Shift+T",
    "notebook":          "Ctrl+Shift+B",
    "media_library":     "Ctrl+Shift+M",
    "quick_connect":     "Ctrl+Q",
    "clipboard_manager": "Ctrl+`",
}

# Human-readable description for each configurable action
_CC_HOTKEY_LABELS: dict[str, str] = {
    "toggle_search":     "Toggle search bar",
    "new_node":          "New node wizard",
    "new_folder":        "New folder",
    "open_settings":     "Open Settings",
    "open_help":         "Help / FAQ",
    "time_tracker":      "Toggle Time Tracker HUD",
    "notebook":          "Open Notebook",
    "media_library":     "Open Media Library",
    "quick_connect":     "Quick Connect",
    "clipboard_manager": "Open Clipboard Manager",
}


def _get_hotkey(action_id: str) -> str:
    """Return the current effective key sequence for a built-in action.

    Falls back to ``_CC_DEFAULT_HOTKEYS`` if the user has not customised it.
    Returns an empty string if the action_id is unknown.
    """
    return _settings_store.value(
        f"hotkey/{action_id}",
        _CC_DEFAULT_HOTKEYS.get(action_id, ""),
    )


class _PluginHotkeysProxy:
    def __init__(self, plugin_id: str):
        self._pid = plugin_id
        self._registered: dict = {}          # sequence → QShortcut
        self._disabled_builtins: dict = {}   # sequence → list[QShortcut]

    def register(self, sequence: str, fn, description: str = ""):
        from PySide6.QtGui import QShortcut, QKeySequence
        mw = _plugin_manager._main_window
        if mw is None:
            return
        key_seq = QKeySequence(sequence)
        # Disable any conflicting shortcuts already on the main window so the
        # two shortcuts don't become "ambiguous" (which would silence both).
        conflicts = []
        for existing in mw.findChildren(QShortcut):
            if (existing.key() == key_seq
                    and existing not in self._registered.values()
                    and existing.isEnabled()):
                existing.setEnabled(False)
                conflicts.append(existing)
        if conflicts:
            self._disabled_builtins[sequence] = conflicts
        try:
            sc = QShortcut(key_seq, mw)
            sc.setContext(Qt.ApplicationShortcut)
            sc.activated.connect(fn)
            self._registered[sequence] = sc
            # Track in plugin manager for settings display
            _plugin_manager._plugin_hotkeys.setdefault(self._pid, {})[sequence] = description
        except Exception as exc:
            # Re-enable any builtins we disabled if shortcut creation failed
            for sc_builtin in self._disabled_builtins.pop(sequence, []):
                sc_builtin.setEnabled(True)
            print(f"[Plugin:{self._pid}] hotkey register error ({sequence}): {exc}",
                  file=sys.stderr)

    def unregister(self, sequence: str):
        sc = self._registered.pop(sequence, None)
        if sc is not None:
            sc.setEnabled(False)
            sc.setParent(None)
        # Restore any built-in shortcuts we suppressed for this key
        for sc_builtin in self._disabled_builtins.pop(sequence, []):
            sc_builtin.setEnabled(True)
        _plugin_manager._plugin_hotkeys.get(self._pid, {}).pop(sequence, None)

    def cleanup(self):
        for seq in list(self._registered):
            self.unregister(seq)

    def list_all(self) -> dict:
        """Return every known hotkey — CC built-ins plus all active plugin hotkeys.

        Returns a dict keyed by Qt-style sequence string (e.g. 'Ctrl+Shift+T').
        Each value is {"source": str, "description": str, "is_builtin": bool}.
        """
        result: dict = {}
        # Use live user-configured hotkeys so plugins see the actual bindings
        for action_id, desc in _CC_HOTKEY_LABELS.items():
            seq = _get_hotkey(action_id)
            if seq:
                result[seq] = {
                    "source": "Command Center",
                    "description": desc,
                    "is_builtin": True,
                }
        # Also include notebook editor shortcuts (not user-configurable)
        for seq, desc in {
            "Ctrl+B": "Notebook: bold",
            "Ctrl+I": "Notebook: italic",
            "Ctrl+U": "Notebook: underline",
        }.items():
            if seq not in result:
                result[seq] = {"source": "Command Center", "description": desc, "is_builtin": True}
        for pid, keymap in _plugin_manager._plugin_hotkeys.items():
            for seq, desc in keymap.items():
                result[seq] = {
                    "source": f"Plugin: {pid}",
                    "description": desc or "",
                    "is_builtin": False,
                }
        return result


class CommandCenterAPI:
    """Public surface exposed to plugin code via `activate(api)`."""

    def __init__(self, plugin_id: str, plugin_dir: Path, store: "NodeStore"):
        self._plugin_id      = plugin_id
        self.plugin_dir      = plugin_dir
        self.config_dir      = CONFIG_DIR
        self.app_version     = APP_VERSION
        self.raw_settings    = _settings_store

        self.settings   = _PluginSettingsProxy(plugin_id)
        self.notebook   = _PluginNotebookProxy(plugin_id)
        self.nodes      = _PluginNodesProxy(plugin_id, store)
        self.theme      = _PluginThemeProxy()
        self.files      = _PluginFilesProxy()
        self.ui         = _PluginUIProxy(plugin_id)
        self.hotkeys    = _PluginHotkeysProxy(plugin_id)
        self.clipboard  = _PluginClipboardProxy()
        self.timers     = _PluginTimerProxy(plugin_id)

    def toast(self, msg: str, kind: str = "info"):
        mw = _plugin_manager._main_window
        if mw is not None:
            try:
                mw.toast(msg, kind)
            except Exception:
                pass

    def log(self, msg: str):
        print(f"[Plugin:{self._plugin_id}] {msg}", file=sys.stderr)

    def launch(self, command: str, args: list = None) -> "subprocess.Popen":
        return subprocess.Popen([command] + (args or []), shell=False)

    def open_url(self, url: str):
        webbrowser.open(url)

    def _cleanup(self):
        """Called by PluginManager on deactivate to tear down API resources."""
        self.ui.cleanup()
        self.hotkeys.cleanup()
        self.timers.cleanup()


# ---------------------------------------------------------------------------
# Plugin record and PluginManager
# ---------------------------------------------------------------------------

class _PluginRecord:
    """Holds all runtime state for one installed plugin."""
    __slots__ = ("manifest", "install_dir", "enabled", "api",
                 "module", "error", "store")

    def __init__(self, manifest: dict, install_dir: Path, store):
        self.manifest    = manifest
        self.install_dir = install_dir
        self.enabled     = False
        self.api: Optional[CommandCenterAPI] = None
        self.module      = None
        self.error: Optional[str] = None
        self.store       = store


class PluginManager:
    """Singleton managing all plugin lifecycle operations.

    Plugins are stored as subdirectories under PLUGIN_DIR:
      PLUGIN_DIR/<plugin_id>/
        manifest.json
        plugin.py
        (any other files from the .ccplug)
    """

    def __init__(self):
        self._plugins: dict[str, _PluginRecord] = {}   # id → record
        self._main_window: Optional["MainWindow"]  = None
        self._settings_changed_listeners: list     = []
        self._plugin_hotkeys: dict[str, dict]      = {}  # pid → {seq: desc}
        self._open_settings_dialogs: list          = []  # weak-tracked open SettingsDialogs

    # \u2500\u2500 lifecycle \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500

    def set_main_window(self, win: "MainWindow"):
        self._main_window = win

    def scan_installed(self, store):
        """Discover all plugin subdirectories under PLUGIN_DIR.

        Safe to call at startup only.  To avoid duplicate footer buttons
        (and leaked API objects), deactivate all currently-active plugins
        before clearing the registry.
        """
        # Clean up any already-active plugins so their footer buttons,
        # hotkeys, and timers are properly removed before we rebuild state.
        for pid in list(self._plugins.keys()):
            rec = self._plugins[pid]
            if rec.enabled:
                self._deactivate(pid)
        self._plugins.clear()
        for d in PLUGIN_DIR.iterdir():
            if not d.is_dir():
                continue
            manifest_path = d / "manifest.json"
            if not manifest_path.exists():
                continue
            try:
                manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            except Exception as exc:
                print(f"[PluginManager] Bad manifest in {d.name}: {exc}", file=sys.stderr)
                continue
            pid = manifest.get("id") or d.name
            rec = _PluginRecord(manifest, d, store)
            self._plugins[pid] = rec

        # Re-apply enabled states from settings
        enabled_ids = self._load_enabled_set()
        for pid, rec in self._plugins.items():
            if pid in enabled_ids:
                self._activate(pid)

    def install_ccplug(self, path: str, store) -> tuple[bool, str]:
        """Extract a .ccplug ZIP, validate its manifest, install under PLUGIN_DIR.

        If the manifest carries a ``uuid`` field that matches an already-installed
        plugin the existing plugin is *updated in-place*: code files from the ZIP
        are overwritten while any runtime-created files (keys, caches, configs, etc.)
        that are NOT in the ZIP are left untouched.
        Returns (success: bool, message: str).
        """
        path = Path(path)
        if not path.exists():
            return False, f"File not found: {path}"
        try:
            with zipfile.ZipFile(path, "r") as zf:
                names = zf.namelist()
                if "manifest.json" not in names:
                    return False, "Invalid plugin: manifest.json not found inside archive."
                try:
                    manifest = json.loads(zf.read("manifest.json").decode("utf-8"))
                except Exception as exc:
                    return False, f"manifest.json parse error: {exc}"
                pid = manifest.get("id", "").strip()
                if not pid:
                    return False, "manifest.json must have a non-empty 'id' field."
                entry = manifest.get("entry", "plugin.py")
                if entry not in names:
                    return False, f"Entry file '{entry}' not found inside archive."

                # ── UUID-based update detection ───────────────────────────────
                incoming_uuid = manifest.get("uuid", "").strip()
                existing_rec: Optional["_PluginRecord"] = None
                if incoming_uuid:
                    for rec in self._plugins.values():
                        if rec.manifest.get("uuid", "") == incoming_uuid:
                            existing_rec = rec
                            break

                if existing_rec is not None:
                    # Update path — overwrite code files, preserve runtime data
                    was_enabled = existing_rec.enabled
                    old_pid = existing_rec.manifest.get("id", pid)
                    if was_enabled:
                        self._deactivate(old_pid)
                    zf.extractall(existing_rec.install_dir)
                    existing_rec.manifest = manifest
                    existing_rec.error = None
                    if old_pid != pid:
                        del self._plugins[old_pid]
                        self._plugins[pid] = existing_rec
                    if was_enabled:
                        ok, act_msg = self._activate(pid)
                        if not ok:
                            self._save_enabled_set()
                            return False, f"Plugin updated but failed to reload: {act_msg}"
                    self._save_enabled_set()
                    name = manifest.get("name", pid)
                    ver  = manifest.get("version", "?")
                    return True, f"Plugin '{name}' updated to v{ver}."

                # ── Fresh install ─────────────────────────────────────────────
                install_dir = PLUGIN_DIR / pid
                install_dir.mkdir(parents=True, exist_ok=True)
                zf.extractall(install_dir)

        except zipfile.BadZipFile:
            return False, "File is not a valid ZIP/ccplug archive."
        except Exception as exc:
            return False, f"Install error: {exc}"

        # Register the freshly-installed plugin in memory
        rec = _PluginRecord(manifest, PLUGIN_DIR / pid, store)
        self._plugins[pid] = rec
        return True, f"Plugin '{manifest.get('name', pid)}' installed successfully."

    def uninstall(self, pid: str) -> tuple[bool, str]:
        """Deactivate + delete a plugin's directory."""
        rec = self._plugins.get(pid)
        if rec is None:
            return False, "Plugin not found."
        if rec.enabled:
            self.deactivate(pid)
        try:
            shutil.rmtree(rec.install_dir)
        except Exception as exc:
            return False, f"Delete error: {exc}"
        del self._plugins[pid]
        self._save_enabled_set()
        return True, f"Plugin '{rec.manifest.get('name', pid)}' uninstalled."

    def activate(self, pid: str) -> tuple[bool, str]:
        """Enable a plugin (public). Returns (ok, msg)."""
        ok, msg = self._activate(pid)
        self._save_enabled_set()
        return ok, msg

    def deactivate(self, pid: str) -> tuple[bool, str]:
        """Disable a plugin (public). Returns (ok, msg)."""
        ok, msg = self._deactivate(pid)
        self._save_enabled_set()
        return ok, msg

    def _activate(self, pid: str) -> tuple[bool, str]:
        rec = self._plugins.get(pid)
        if rec is None:
            return False, "Plugin not found."
        if rec.enabled:
            return True, "Already active."

        store = rec.store
        plug_dir = rec.install_dir
        entry = rec.manifest.get("entry", "plugin.py")
        entry_path = plug_dir / entry

        if not entry_path.exists():
            rec.error = f"Entry file '{entry}' missing."
            return False, rec.error

        import importlib.util as _ilu
        module_name = f"_ccplugin_{pid}"
        try:
            spec = _ilu.spec_from_file_location(module_name, entry_path,
                                                 submodule_search_locations=[str(plug_dir)])
            mod  = _ilu.module_from_spec(spec)
            sys.modules[module_name] = mod
            spec.loader.exec_module(mod)
        except Exception as exc:
            rec.error = f"Import error: {exc}"
            print(f"[PluginManager] {pid} import error: {exc}", file=sys.stderr)
            return False, rec.error

        if not hasattr(mod, "activate"):
            rec.error = "plugin.py must define an activate(api) function."
            return False, rec.error

        api = CommandCenterAPI(pid, plug_dir, store)
        try:
            mod.activate(api)
        except Exception as exc:
            rec.error = f"activate() raised: {exc}"
            print(f"[PluginManager] {pid} activate error: {exc}", file=sys.stderr)
            api._cleanup()
            return False, rec.error

        rec.api    = api
        rec.module = mod
        rec.enabled = True
        rec.error   = None
        return True, f"Plugin '{rec.manifest.get('name', pid)}' activated."

    def _deactivate(self, pid: str) -> tuple[bool, str]:
        rec = self._plugins.get(pid)
        if rec is None:
            return False, "Plugin not found."
        if not rec.enabled:
            return True, "Already inactive."
        if rec.module is not None and hasattr(rec.module, "deactivate"):
            try:
                rec.module.deactivate()
            except Exception as exc:
                print(f"[PluginManager] {pid} deactivate error: {exc}", file=sys.stderr)
        if rec.api is not None:
            rec.api._cleanup()
        rec.enabled = False
        rec.api     = None
        sys.modules.pop(f"_ccplugin_{pid}", None)
        return True, f"Plugin '{rec.manifest.get('name', pid)}' deactivated."

    def reload_plugin(self, pid: str) -> tuple[bool, str]:
        """Deactivate then reactivate a plugin to pick up code changes."""
        self._deactivate(pid)
        ok, msg = self._activate(pid)
        self._save_enabled_set()
        return ok, msg

    def purge_all(self) -> int:
        """Deactivate all active plugins, clear all plugin QSettings, and unload modules.
        Returns the count of plugins that were deactivated."""
        count = 0
        for pid in list(self._plugins.keys()):
            rec = self._plugins[pid]
            if rec.enabled:
                self._deactivate(pid)
                count += 1
            _settings_store.beginGroup(f"plugins/{pid}")
            _settings_store.remove("")
            _settings_store.endGroup()
        _settings_store.sync()
        self._save_enabled_set()
        return count

    # \u2500\u2500 persistence \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500

    def _load_enabled_set(self) -> set:
        raw = _settings_store.value("plugins_enabled", "")
        if not raw:
            return set()
        return set(raw.split(","))

    def _save_enabled_set(self):
        enabled = [pid for pid, rec in self._plugins.items() if rec.enabled]
        _settings_store.setValue("plugins_enabled", ",".join(enabled))

    # \u2500\u2500 settings-tab injection \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500

    def _inject_settings_tabs(self):
        """Push plugin-provided settings tabs into any open SettingsDialog."""
        for dlg_ref in list(self._open_settings_dialogs):
            dlg = dlg_ref() if callable(dlg_ref) else dlg_ref
            if dlg is None or not dlg.isVisible():
                continue
            if not hasattr(dlg, "_tabs"):
                continue
            for pid, rec in self._plugins.items():
                if not rec.enabled or rec.api is None:
                    continue
                for label, widget in rec.api.ui._settings_tabs:
                    # Avoid double-adding
                    already = False
                    for i in range(dlg._tabs.count()):
                        if dlg._tabs.tabText(i) == label:
                            already = True
                            break
                    if not already:
                        dlg._tabs.addTab(widget, label)

    def notify_settings_changed(self):
        for fn in list(self._settings_changed_listeners):
            try:
                fn()
            except Exception as exc:
                print(f"[PluginManager] settings_changed listener error: {exc}",
                      file=sys.stderr)

    def notify_nodes_loaded(self):
        for pid, rec in self._plugins.items():
            if rec.enabled and rec.api is not None:
                for fn in rec.api.nodes._on_loaded_fns:
                    try:
                        fn()
                    except Exception as exc:
                        print(f"[Plugin:{pid}] on_loaded error: {exc}", file=sys.stderr)

    def all_plugins(self) -> list:
        """Return list of _PluginRecord values in deterministic order."""
        return list(self._plugins.values())


# Singleton
_plugin_manager = PluginManager()


# ---------------------------------------------------------------------------
# .ccplug file-type association helper (Windows only)
# ---------------------------------------------------------------------------

def _register_ccplug_file_association() -> None:
    """Write HKCU registry entries so .ccplug files show the plugin icon.

    Uses HKCU (current-user) only — no admin rights required.
    Silently no-ops on non-Windows platforms or if winreg is unavailable.
    """
    try:
        import winreg
    except ImportError:
        return  # Not on Windows

    ico_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "CommandCenterPlugin.ico")
    if not os.path.isfile(ico_path):
        return  # Icon file not found — skip silently

    prog_id   = "CommandCenter.ccplug"
    ext_key   = r"Software\Classes\.ccplug"
    prog_key  = rf"Software\Classes\{prog_id}"
    icon_key  = rf"Software\Classes\{prog_id}\DefaultIcon"

    try:
        # Map .ccplug → ProgID
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, ext_key, 0,
                                winreg.KEY_WRITE) as k:
            winreg.SetValueEx(k, "", 0, winreg.REG_SZ, prog_id)
            winreg.SetValueEx(k, "Content Type", 0, winreg.REG_SZ,
                              "application/x-ccplug")

        # ProgID description
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, prog_key, 0,
                                winreg.KEY_WRITE) as k:
            winreg.SetValueEx(k, "", 0, winreg.REG_SZ,
                              "CommandCenter Plugin")

        # DefaultIcon
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, icon_key, 0,
                                winreg.KEY_WRITE) as k:
            winreg.SetValueEx(k, "", 0, winreg.REG_SZ, ico_path)

        # Notify the shell so Explorer updates immediately (best-effort)
        try:
            from ctypes import windll
            windll.shell32.SHChangeNotify(0x08000000, 0x0000, None, None)
        except Exception:
            pass

    except OSError:
        pass  # Registry write failed — non-critical, ignore silently


# ---------------------------------------------------------------------------
# Plugins dialog UI
# ---------------------------------------------------------------------------

class PluginsDialog(QDialog):
    """Browse, install, toggle, and uninstall plugins."""

    def __init__(self, store, parent=None):
        super().__init__(parent, Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self._store       = store
        self._drag_pos: Optional[QPoint] = None
        self._selected_pid: Optional[str] = None
        self._build_ui()
        self.setMinimumSize(780, 560)
        self.resize(900, 620)
        # Do NOT call scan_installed here — plugins are already scanned at
        # startup and re-scanning would deactivate/reactivate them, creating
        # duplicate footer buttons every time this dialog is opened.
        self._refresh_list()

    # \u2500\u2500 UI construction \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500

    def _build_ui(self):
        t = _theme
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(0)

        self._card = QWidget()
        self._card.setObjectName("plug_card")
        cl = QVBoxLayout(self._card)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(0)
        self._card.setGraphicsEffect(make_shadow(self._card, 32, QColor(0, 0, 0, 210)))
        root.addWidget(self._card)

        # Title bar
        tbar = QWidget()
        tbar.setFixedHeight(48)
        tbl  = QHBoxLayout(tbar)
        tbl.setContentsMargins(20, 0, 12, 0)
        icon_lbl = QLabel("🔌")
        icon_lbl.setFont(QFont("Segoe UI", 14))
        icon_lbl.setStyleSheet("background:transparent;")
        tbl.addWidget(icon_lbl)
        tbl.addSpacing(8)
        title_lbl = QLabel("Plugins")
        title_lbl.setFont(QFont("Segoe UI", 12, QFont.Weight.DemiBold))
        title_lbl.setStyleSheet(f"color:{t.GLOW.name()}; background:transparent;")
        tbl.addWidget(title_lbl)
        tbl.addStretch()

        # Guide button
        guide_btn = QPushButton("📖 Developer Guide")
        guide_btn.setCursor(Qt.PointingHandCursor)
        guide_btn.setFont(QFont("Segoe UI", 8))
        guide_btn.setFixedHeight(26)
        guide_btn.setStyleSheet(self._btn_style())
        guide_btn.clicked.connect(self._show_guide)
        tbl.addWidget(guide_btn)
        tbl.addSpacing(8)

        # Purge button
        purge_btn = QPushButton("⚠  Purge All")
        purge_btn.setCursor(Qt.PointingHandCursor)
        purge_btn.setFont(QFont("Segoe UI", 8))
        purge_btn.setFixedHeight(26)
        purge_btn.setStyleSheet(self._btn_style_danger())
        purge_btn.clicked.connect(self._on_purge)
        tbl.addWidget(purge_btn)
        tbl.addSpacing(8)

        # Install button
        install_btn = QPushButton("+ Install Plugin")
        install_btn.setCursor(Qt.PointingHandCursor)
        install_btn.setFont(QFont("Segoe UI", 8, QFont.Weight.DemiBold))
        install_btn.setFixedHeight(26)
        install_btn.setStyleSheet(self._btn_style_accent())
        install_btn.clicked.connect(self._on_install)
        tbl.addWidget(install_btn)
        tbl.addSpacing(8)

        close_btn = QPushButton("✕")
        close_btn.setFixedSize(28, 28)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setFont(QFont("Segoe UI", 10))
        close_btn.setStyleSheet(f"""
            QPushButton {{
                background:transparent; border:none;
                color:{t.TEXT_DIM.name()};
                border-radius:14px;
            }}
            QPushButton:hover {{
                background:rgba({t.ACCENT_RED.red()},{t.ACCENT_RED.green()},{t.ACCENT_RED.blue()},140);
                color:#fff;
            }}
        """)
        close_btn.clicked.connect(self.close)
        tbl.addWidget(close_btn)
        cl.addWidget(tbar)

        # Separator
        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},35);")
        cl.addWidget(sep)

        # Main body: list (left) + detail panel (right)
        body = QWidget()
        bl   = QHBoxLayout(body)
        bl.setContentsMargins(0, 0, 0, 0)
        bl.setSpacing(0)

        # -- Left: plugin list
        list_panel = QWidget()
        list_panel.setFixedWidth(280)
        lpl = QVBoxLayout(list_panel)
        lpl.setContentsMargins(12, 12, 8, 12)
        lpl.setSpacing(6)

        search_lbl = QLabel("Installed Plugins")
        search_lbl.setFont(QFont("Segoe UI", 9, QFont.Weight.DemiBold))
        search_lbl.setStyleSheet(f"color:{t.TEXT_SECONDARY.name()}; background:transparent;")
        lpl.addWidget(search_lbl)

        self._list_scroll = QScrollArea()
        self._list_scroll.setWidgetResizable(True)
        self._list_scroll.setFrameShape(QFrame.NoFrame)
        self._list_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._list_scroll.setStyleSheet(f"""
            QScrollArea {{ background:transparent; border:none; }}
            QScrollBar:vertical {{
                background:rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},120);
                width:5px; border-radius:2px;
            }}
            QScrollBar::handle:vertical {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},100);
                border-radius:2px; min-height:16px;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height:0; }}
        """)
        self._list_container = QWidget()
        self._list_container.setStyleSheet("background:transparent;")
        self._list_layout = QVBoxLayout(self._list_container)
        self._list_layout.setContentsMargins(0, 0, 0, 0)
        self._list_layout.setSpacing(4)
        self._list_layout.addStretch()
        self._list_scroll.setWidget(self._list_container)
        lpl.addWidget(self._list_scroll)
        bl.addWidget(list_panel)

        # Vertical separator
        vsep = QFrame(); vsep.setFrameShape(QFrame.VLine)
        vsep.setStyleSheet(f"color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},35);")
        bl.addWidget(vsep)

        # -- Right: detail panel
        self._detail_panel = QWidget()
        self._detail_panel.setStyleSheet("background:transparent;")
        dpl = QVBoxLayout(self._detail_panel)
        dpl.setContentsMargins(20, 16, 20, 16)
        dpl.setSpacing(10)

        self._detail_name = QLabel("Select a plugin")
        self._detail_name.setFont(QFont("Segoe UI", 13, QFont.Weight.DemiBold))
        self._detail_name.setStyleSheet(f"color:{t.TEXT_PRIMARY.name()}; background:transparent;")
        self._detail_name.setWordWrap(True)
        dpl.addWidget(self._detail_name)

        self._detail_meta = QLabel("")
        self._detail_meta.setFont(QFont("Segoe UI", 8))
        self._detail_meta.setStyleSheet(f"color:{t.TEXT_DIM.name()}; background:transparent;")
        dpl.addWidget(self._detail_meta)

        self._detail_desc = QLabel("")
        self._detail_desc.setFont(QFont("Segoe UI", 9))
        self._detail_desc.setStyleSheet(f"color:{t.TEXT_SECONDARY.name()}; background:transparent;")
        self._detail_desc.setWordWrap(True)
        dpl.addWidget(self._detail_desc)

        self._detail_perms = QLabel("")
        self._detail_perms.setFont(QFont("Segoe UI", 8))
        self._detail_perms.setStyleSheet(f"color:{t.ACCENT_AMBER.name()}; background:transparent;")
        self._detail_perms.setWordWrap(True)
        dpl.addWidget(self._detail_perms)

        self._detail_error = QLabel("")
        self._detail_error.setFont(QFont("Segoe UI", 8))
        self._detail_error.setStyleSheet(f"color:{t.ACCENT_RED.name()}; background:transparent;")
        self._detail_error.setWordWrap(True)
        dpl.addWidget(self._detail_error)

        dpl.addStretch()

        btn_row = QWidget()
        brl = QHBoxLayout(btn_row)
        brl.setContentsMargins(0, 0, 0, 0)
        brl.setSpacing(8)

        self._toggle_btn = QPushButton("Enable")
        self._toggle_btn.setFixedHeight(30)
        self._toggle_btn.setCursor(Qt.PointingHandCursor)
        self._toggle_btn.setFont(QFont("Segoe UI", 9, QFont.Weight.DemiBold))
        self._toggle_btn.setStyleSheet(self._btn_style_accent())
        self._toggle_btn.clicked.connect(self._on_toggle)
        self._toggle_btn.setEnabled(False)
        brl.addWidget(self._toggle_btn)

        self._reload_btn = QPushButton("↻ Reload")
        self._reload_btn.setFixedHeight(30)
        self._reload_btn.setCursor(Qt.PointingHandCursor)
        self._reload_btn.setFont(QFont("Segoe UI", 9))
        self._reload_btn.setStyleSheet(self._btn_style())
        self._reload_btn.clicked.connect(self._on_reload)
        self._reload_btn.setEnabled(False)
        brl.addWidget(self._reload_btn)

        self._uninstall_btn = QPushButton("🗑 Uninstall")
        self._uninstall_btn.setFixedHeight(30)
        self._uninstall_btn.setCursor(Qt.PointingHandCursor)
        self._uninstall_btn.setFont(QFont("Segoe UI", 9))
        self._uninstall_btn.setStyleSheet(self._btn_style_danger())
        self._uninstall_btn.clicked.connect(self._on_uninstall)
        self._uninstall_btn.setEnabled(False)
        brl.addWidget(self._uninstall_btn)

        brl.addStretch()
        dpl.addWidget(btn_row)

        bl.addWidget(self._detail_panel, stretch=1)
        cl.addWidget(body, stretch=1)

    # \u2500\u2500 plugin list rendering \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500

    def _refresh_list(self):
        # Clear old rows
        while self._list_layout.count() > 1:
            item = self._list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        plugins = _plugin_manager.all_plugins()
        if not plugins:
            empty = QLabel("No plugins installed.\nUse '+ Install Plugin' to add one.")
            empty.setFont(QFont("Segoe UI", 9))
            empty.setStyleSheet(f"color:{_theme.TEXT_DIM.name()}; background:transparent;")
            empty.setAlignment(Qt.AlignCenter)
            empty.setWordWrap(True)
            self._list_layout.insertWidget(0, empty)
            self._clear_detail()
            return

        for rec in plugins:
            row = self._make_list_row(rec)
            self._list_layout.insertWidget(self._list_layout.count() - 1, row)

    def _make_list_row(self, rec: _PluginRecord) -> QWidget:
        t   = _theme
        pid = rec.manifest.get("id", "?")
        row = QWidget()
        row.setFixedHeight(56)
        row.setCursor(Qt.CursorShape.PointingHandCursor)
        rl = QHBoxLayout(row)
        rl.setContentsMargins(8, 6, 8, 6)
        rl.setSpacing(8)

        # Status dot
        dot = QLabel("●")
        dot.setFont(QFont("Segoe UI", 10))
        dot.setFixedWidth(14)
        dot.setStyleSheet(
            f"color:{t.ACCENT_TEAL.name()};" if rec.enabled else
            f"color:{t.TEXT_DIM.name()};")
        rl.addWidget(dot)

        # Text
        col = QWidget()
        cll = QVBoxLayout(col)
        cll.setContentsMargins(0, 0, 0, 0)
        cll.setSpacing(1)
        name_lbl = QLabel(rec.manifest.get("name", pid))
        name_lbl.setFont(QFont("Segoe UI", 9, QFont.Weight.DemiBold))
        name_lbl.setStyleSheet(f"color:{t.TEXT_PRIMARY.name()}; background:transparent;")
        ver_lbl = QLabel(f"v{rec.manifest.get('version', '?')}  •  "
                         f"{'Active' if rec.enabled else 'Inactive'}")
        ver_lbl.setFont(QFont("Segoe UI", 7))
        ver_lbl.setStyleSheet(
            f"color:{t.ACCENT_TEAL.name() if rec.enabled else t.TEXT_DIM.name()};"
            " background:transparent;")
        cll.addWidget(name_lbl)
        cll.addWidget(ver_lbl)
        rl.addWidget(col, stretch=1)

        # Error badge
        if rec.error:
            err_dot = QLabel("⚠")
            err_dot.setFont(QFont("Segoe UI", 9))
            err_dot.setStyleSheet(f"color:{t.ACCENT_AMBER.name()}; background:transparent;")
            rl.addWidget(err_dot)

        row.setStyleSheet(f"""
            QWidget {{
                background: {'rgba(' +
                    str(t.GLOW.red()) + ',' +
                    str(t.GLOW.green()) + ',' +
                    str(t.GLOW.blue()) + ',22)' if pid == self._selected_pid
                    else 'transparent'};
                border-radius: 6px;
            }}
        """)
        row.mousePressEvent = lambda e, p=pid: self._show_detail(p)
        return row

    def _show_detail(self, pid: str):
        self._selected_pid = pid
        self._refresh_list()  # re-render rows to update selection highlight

        rec = _plugin_manager._plugins.get(pid)
        if rec is None:
            return

        m = rec.manifest
        t = _theme
        self._detail_name.setText(m.get("name", pid))
        self._detail_meta.setText(
            f"ID: {m.get('id','?')}   •   v{m.get('version','?')}   •   "
            f"by {m.get('author', 'Unknown')}"
        )
        self._detail_desc.setText(m.get("description", "No description."))
        perms = m.get("permissions", [])
        if perms:
            self._detail_perms.setText("Permissions: " + ", ".join(perms))
        else:
            self._detail_perms.setText("")

        if rec.error:
            self._detail_error.setText(f"⚠ Error: {rec.error}")
        else:
            self._detail_error.setText("")

        self._toggle_btn.setEnabled(True)
        self._reload_btn.setEnabled(True)
        self._uninstall_btn.setEnabled(True)
        if rec.enabled:
            self._toggle_btn.setText("Disable")
            self._toggle_btn.setStyleSheet(self._btn_style())
        else:
            self._toggle_btn.setText("Enable")
            self._toggle_btn.setStyleSheet(self._btn_style_accent())

    def _clear_detail(self):
        self._selected_pid = None
        self._detail_name.setText("Select a plugin")
        self._detail_meta.setText("")
        self._detail_desc.setText("")
        self._detail_perms.setText("")
        self._detail_error.setText("")
        self._toggle_btn.setEnabled(False)
        self._reload_btn.setEnabled(False)
        self._uninstall_btn.setEnabled(False)

    # \u2500\u2500 actions \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500

    def _on_install(self):
        _register_ccplug_file_association()
        path, _ = QFileDialog.getOpenFileName(
            self, "Install Plugin", "",
            "Command Center Plugin (*.ccplug);;ZIP Archive (*.zip)")
        if not path:
            return
        ok, msg = _plugin_manager.install_ccplug(path, self._store)
        self._refresh_list()
        if not ok:
            QMessageBox.critical(self, "Install Failed", msg)
        else:
            title = "Plugin Updated" if "updated to" in msg else "Plugin Installed"
            QMessageBox.information(self, title, msg)

    def _on_toggle(self):
        if self._selected_pid is None:
            return
        rec = _plugin_manager._plugins.get(self._selected_pid)
        if rec is None:
            return
        if rec.enabled:
            ok, msg = _plugin_manager.deactivate(self._selected_pid)
        else:
            ok, msg = _plugin_manager.activate(self._selected_pid)
        if not ok:
            QMessageBox.warning(self, "Plugin Error", msg)
        self._show_detail(self._selected_pid)

    def _on_reload(self):
        if self._selected_pid is None:
            return
        ok, msg = _plugin_manager.reload_plugin(self._selected_pid)
        if not ok:
            QMessageBox.warning(self, "Reload Error", msg)
        self._show_detail(self._selected_pid)

    def _on_uninstall(self):
        if self._selected_pid is None:
            return
        rec = _plugin_manager._plugins.get(self._selected_pid)
        if rec is None:
            return
        name = rec.manifest.get("name", self._selected_pid)
        reply = QMessageBox.question(
            self, "Confirm Uninstall",
            f"Permanently remove plugin '{name}' and all its files?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply != QMessageBox.Yes:
            return
        ok, msg = _plugin_manager.uninstall(self._selected_pid)
        self._clear_detail()
        self._refresh_list()
        if not ok:
            QMessageBox.critical(self, "Uninstall Error", msg)

    def _on_purge(self):
        reply = QMessageBox.question(
            self, "Confirm Purge All",
            "Deactivate ALL plugins and permanently clear all plugin settings data?\n\n"
            "Command Center will restart automatically to ensure a clean state.\n\n"
            "This cannot be undone.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply != QMessageBox.Yes:
            return
        _plugin_manager.purge_all()
        self.close()
        QTimer.singleShot(200, _restart_application)

    def _show_guide(self):
        dlg = _PluginGuideDialog(self)
        dlg.exec()

    # \u2500\u2500 button styles \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500

    def _btn_style(self) -> str:
        t = _theme
        return f"""
            QPushButton {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},180);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:5px; color:{t.TEXT_SECONDARY.name()}; padding:0 14px;
            }}
            QPushButton:hover {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);
                color:{t.GLOW.name()};
            }}
        """

    def _btn_style_accent(self) -> str:
        t = _theme
        return f"""
            QPushButton {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},160);
                border-radius:5px; color:{t.GLOW.name()}; padding:0 14px;
                font-weight:600;
            }}
            QPushButton:hover {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},100);
            }}
        """

    def _btn_style_danger(self) -> str:
        t = _theme
        return f"""
            QPushButton {{
                background:rgba({t.ACCENT_RED.red()},{t.ACCENT_RED.green()},{t.ACCENT_RED.blue()},40);
                border:1px solid rgba({t.ACCENT_RED.red()},{t.ACCENT_RED.green()},{t.ACCENT_RED.blue()},100);
                border-radius:5px; color:{t.ACCENT_RED.name()}; padding:0 14px;
            }}
            QPushButton:hover {{
                background:rgba({t.ACCENT_RED.red()},{t.ACCENT_RED.green()},{t.ACCENT_RED.blue()},100);
                color:#fff;
            }}
        """

    # \u2500\u2500 painting + drag \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500

    def paintEvent(self, e):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        r = QRectF(8, 8, self.width() - 16, self.height() - 16)
        path = QPainterPath()
        path.addRoundedRect(r, 14, 14)
        grad = QLinearGradient(0, 0, 0, self.height())
        grad.setColorAt(0, t.BG_MID)
        grad.setColorAt(1, t.BG_DARK)
        p.fillPath(path, grad)
        glow_col = QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 60)
        p.setPen(QPen(glow_col, 1.5))
        p.drawPath(path)
        p.end()

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_pos = e.globalPosition().toPoint() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if self._drag_pos and e.buttons() & Qt.LeftButton:
            self.move(e.globalPosition().toPoint() - self._drag_pos)

    def mouseReleaseEvent(self, e):
        self._drag_pos = None


# ---------------------------------------------------------------------------
# AI-optimized plugin-building instructions (not displayed; copy-to-clipboard)
# ---------------------------------------------------------------------------
_CC_PLUGIN_AI_GUIDE = """\
SYSTEM: Command Center Plugin Development — Complete AI Reference
================================================================
Target: Python 3.12, PySide6 6.x, Windows, single-file app CommandCenter.py
Plugin format: .ccplug (ZIP renamed), installed to %APPDATA%\\CommandCenter\\plugins\\<id>\\

=== REQUIRED FILES ===
manifest.json  (REQUIRED)
plugin.py      (REQUIRED, default entry-point; override with "entry" field)

=== manifest.json SCHEMA ===
{
  "id":              "<snake_case_unique_id>",   // permanent, never change
  "uuid":            "<uuid-v4-string>",          // STRONGLY RECOMMENDED — enables in-place updates
  "name":            "<Display Name>",
  "version":         "1.0.0.0",                   // MAJOR.MINOR.PATCH.BUILD
  "description":     "<one-two sentence summary>",
  "author":          "<optional>",
  "min_app_version": "<optional advisory>",
  "entry":           "plugin.py",                // optional default
  "permissions":     ["ui","hotkeys","notebook","nodes","settings",
                      "files","clipboard","timers","network"]  // informational only
}

NOTE — uuid and in-place updates:
  When a .ccplug is installed and its uuid matches an already-installed plugin,
  Command Center performs an in-place update instead of a fresh install:
    1. The active plugin is deactivated.
    2. Code files from the ZIP (manifest.json, plugin.py, etc.) are extracted,
       overwriting their counterparts in the install directory.
    3. Any runtime-created files (data files, keys, caches) that are NOT in the
       ZIP are left completely untouched.
    4. The plugin is reactivated with the new code.
  Without a uuid, reinstalling is treated as a fresh install (same behaviour as
  previous versions).  Generate a uuid once with:
    python -c "import uuid; print(uuid.uuid4())"

=== plugin.py CONTRACT ===
MUST define: activate(api)  AND  deactivate()

def activate(api):
    # called when plugin enabled; store api in module global; register all side effects
    pass

def deactivate():
    # called on disable/reload/uninstall; MUST undo every side effect of activate()
    # safe to call even if activate() never ran — guard with: if var is not None
    pass

=== MODULE-LEVEL STATE PATTERN (always use) ===
_api  = None   # CommandCenterAPI instance
_btn  = None   # QPushButton from add_footer_button (remove in deactivate)
_tab  = None   # str label passed to add_settings_tab (remove in deactivate)
_proc = None   # subprocess.Popen if launched (terminate in deactivate)
# timers via api.timers — auto-cleaned, no manual cancel needed

=== FULL API REFERENCE ===

--- TOP LEVEL ---
api.app_version          -> str         e.g. "1.1.0"
api.plugin_dir           -> pathlib.Path  install dir; use for ALL data files
api.config_dir           -> pathlib.Path  app config root (read-only)
api.toast(msg, kind)                     kind: "info"|"success"|"warn"|"error"
api.log(msg)                             prints [Plugin:<id>] msg to stderr
api.launch(cmd, args=[]) -> Popen        async external process
api.open_url(url)                        open URL/path in default app

--- NODES (canvas tiles) ---
Node dict fields:
  id(int) name(str) type(str:"file"|"url"|"note"|"folder")
  target(str) size(str:"1x1"|"2x2"|"2x4") description(str)
  accent_color(str:#hex) icon(str:path|"") auto_launch(bool)
  archived(bool) folder_id(int|None) open_behavior(str) tags(list[str])
  open_behavior values: "normal"|"admin"|"folder"|"copy_file"|"copy_contents"

api.nodes.list()                      -> list[dict]  active root-level nodes
api.nodes.all(include_archived=False) -> list[dict]  all nodes incl. folders
api.nodes.get(node_id)                -> dict|None
api.nodes.add(node_dict)              -> dict         returned dict has "id"
api.nodes.update(node_id, updates)    -> dict         pass only changed fields
api.nodes.delete(node_id)                             permanent, no undo
api.nodes.archive(node_id)                            recoverable from Settings
api.nodes.unarchive(node_id)
api.nodes.list_folders()              -> list[dict]
api.nodes.move_to_folder(node_id, folder_id|None)
api.nodes.reload()                                    force canvas re-render
api.nodes.on_loaded(fn)                               fn() on every canvas reload

--- NOTEBOOK ---
Note summary dict: id(str) title(str) modified(str:ISO) color(str) pinned(bool)
Full note dict:    + content(str:HTML)

api.notebook.list_notes()                      -> list[dict]  sorted pinned→recent
api.notebook.load_note(note_id)                -> dict   includes "content" HTML
api.notebook.save_note(note_id, title, html)
api.notebook.new_note(title, html="")          -> str    new note_id
api.notebook.delete_note(note_id)
api.notebook.update_meta(note_id, color=None, pinned=None)
api.notebook.open_note(note_id)                        opens Notebook UI at that note
api.notebook.search(query)                     -> list[dict]  case-insensitive
api.notebook.on_saved(fn)                              fn(note_id:str) on any save

--- SETTINGS ---
Namespace: auto-prefixed as "plugins/<id>/<key>" in QSettings

api.settings.value(key, default=None)   -> any    str/int/float/bool
api.settings.set(key, value)                       str/int/float/bool only
api.settings.remove(key)
api.settings.all_keys()                  -> list[str]
api.settings.on_changed(fn)                        fn() after Settings dialog save
# For list/dict: json.dumps/loads

--- THEME ---
Color keys: bg_dark bg_mid glow text_primary text_secondary text_dim
            accent_blue accent_teal accent_amber accent_red

api.theme.name()                         -> str
api.theme.available_themes()             -> list[str]
api.theme.colors()                       -> dict[str,str]  all hex colors
api.theme.set_theme(name)
api.theme.register(fn)                              fn() on any theme change
# QColor accessors: api.theme.glow_color() bg_dark() bg_mid() text_primary()
#   text_secondary() text_dim() accent_blue() accent_teal() accent_amber() accent_red()

--- UI ---
api.ui.add_footer_button(label, fn)      -> QPushButton  (save to remove later)
api.ui.remove_footer_button(btn)                    MUST call in deactivate()
api.ui.add_settings_tab(label, widget)              inject tab into Settings dialog
api.ui.remove_settings_tab(label)                   MUST call in deactivate()
api.ui.show_dialog(widget)                          floating non-modal window
api.ui.ask_confirm(title, msg)           -> bool
api.ui.ask_input(title, prompt, default="") -> str|None
api.ui.show_message(title, msg, kind="info")        kind: "info"|"warn"|"error"
api.ui.refresh_canvas()
api.ui.main_window                       -> QMainWindow  use as dialog parent

--- HOTKEYS ---
api.hotkeys.register(sequence, fn, description="")
  # sequence: "Ctrl+Shift+P" Qt-style; fires app-wide regardless of focus
  # auto-suppresses conflicting CC built-ins; restored on unregister
api.hotkeys.unregister(sequence)        # MUST call in deactivate() for each
api.hotkeys.list_all()                  -> dict
  # {"built_in": [{"sequence":str,"description":str},...],
  #  "plugins":  {"plugin_id": [{"sequence":str,"description":str},...], ...}}

--- CLIPBOARD ---
api.clipboard.get_text()       -> str
api.clipboard.set_text(text)
api.clipboard.get_image()      -> QPixmap  (check .isNull())
api.clipboard.set_image(pixmap)

--- TIMERS ---
api.timers.create(interval_ms, fn, single_shot=False) -> QTimer
  # auto-stopped on deactivate — NO manual cancel needed in deactivate()
api.timers.cancel(timer)   # early stop only

--- FILES ---
api.files.read_text(path)        -> str
api.files.write_text(path, content)     # creates dirs if needed
api.files.read_json(path)        -> dict|list
api.files.write_json(path, data)        # pretty-printed JSON
api.files.list_dir(path, pattern="*")  -> list[pathlib.Path]
api.files.open_dialog(title, filters)  -> str|None  (filters: "Text (*.txt);;All Files (*)")
api.files.save_dialog(title, filters)  -> str|None

=== COMMON PATTERNS ===

--- PATTERN 1: Footer Button + Hotkey + Dialog ---
from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton
from PySide6.QtCore import Qt

_api = None
_btn = None

def activate(api):
    global _api, _btn
    _api = api
    _btn = api.ui.add_footer_button("My Tool", _open)
    api.hotkeys.register("Ctrl+Shift+P", _open, "Open My Tool")

def deactivate():
    global _btn
    if _btn is not None:
        _api.ui.remove_footer_button(_btn)
        _btn = None
    _api.hotkeys.unregister("Ctrl+Shift+P")

def _open():
    try:
        dlg = _MyDialog(_api.ui.main_window)
        dlg.exec()
    except Exception as e:
        _api.log(f"Error: {e}")
        _api.toast("Error opening dialog.", "error")

class _MyDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("My Tool")
        self.setMinimumSize(500, 350)
        self._build_ui()
        _api.theme.register(self._apply_theme)
        self._apply_theme()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        self._lbl = QLabel("Content here")
        self._lbl.setAlignment(Qt.AlignCenter)
        lay.addWidget(self._lbl)
        btn = QPushButton("Close")
        btn.clicked.connect(self.accept)
        lay.addWidget(btn)

    def _apply_theme(self):
        c = _api.theme.colors()
        self.setStyleSheet(
            f"QDialog    {{background:{c['bg_dark']};color:{c['text_primary']}}}"
            f"QPushButton {{background:{c['bg_mid']};color:{c['text_primary']};"
            f"              border:1px solid {c['glow']};border-radius:4px;padding:4px 12px}}"
            f"QPushButton:hover {{border-color:{c['accent_blue']}}}"
        )

--- PATTERN 2: Background Polling Timer ---
_timer = None

def activate(api):
    global _api, _timer
    _api = api
    _timer = api.timers.create(10000, _poll)  # every 10s; auto-stopped on deactivate

def _poll():
    try:
        # do work; update UI if needed
        pass
    except Exception as e:
        _api.log(f"Poll error: {e}")

--- PATTERN 3: Settings Tab ---
from PySide6.QtWidgets import QWidget, QFormLayout, QLineEdit, QPushButton

_api = None
_tab_label = "My Plugin"

def activate(api):
    global _api
    _api = api
    api.ui.add_settings_tab(_tab_label, _build_settings())
    api.settings.on_changed(_reload_settings)

def deactivate():
    _api.ui.remove_settings_tab(_tab_label)

def _build_settings():
    w = QWidget()
    form = QFormLayout(w)
    _field = QLineEdit(_api.settings.value("my_key", "default"))
    save_btn = QPushButton("Save")
    save_btn.clicked.connect(lambda: _api.settings.set("my_key", _field.text()))
    form.addRow("My Key:", _field)
    form.addRow("", save_btn)
    return w

def _reload_settings():
    pass  # re-read settings.value() here if needed

--- PATTERN 4: Node Manipulation ---
def _create_bookmark():
    url = _api.ui.ask_input("Add Bookmark", "Enter URL:")
    if not url:
        return
    name = _api.ui.ask_input("Add Bookmark", "Enter name:", url)
    if not name:
        return
    node = _api.nodes.add({"name": name, "type": "url", "target": url,
                            "accent_color": "#4fc3f7"})
    _api.ui.refresh_canvas()
    _api.toast(f"Added: {node['name']}", "success")

--- PATTERN 5: Notebook Integration ---
def _create_note():
    title = _api.ui.ask_input("New Note", "Title:")
    if not title:
        return
    note_id = _api.notebook.new_note(title, f"<p>Created by My Plugin</p>")
    _api.notebook.open_note(note_id)
    _api.toast("Note created.", "success")

def _watch_notes():
    def _on_saved(note_id):
        _api.log(f"Note saved: {note_id}")
    _api.notebook.on_saved(_on_saved)

--- PATTERN 6: External Process Management ---
import subprocess
_proc = None

def activate(api):
    global _api, _proc
    _api = api
    _proc = api.launch("my_daemon.exe", ["--port", "9001"])

def deactivate():
    global _proc
    if _proc is not None and _proc.poll() is None:
        _proc.terminate()
    _proc = None

=== PACKAGING ===
# Python (cross-platform):
import zipfile
with zipfile.ZipFile("my_plugin.ccplug", "w", zipfile.ZIP_DEFLATED) as zf:
    zf.write("manifest.json")
    zf.write("plugin.py")
    # add any helper .py files or data files

# PowerShell:
# Compress-Archive -Path manifest.json,plugin.py -DestinationPath p.zip
# Rename-Item p.zip my_plugin.ccplug

# Install: Footer → Plugins → "+ Install Plugin" → select .ccplug → Enable

# Hot-reload dev: edit files in %APPDATA%\\CommandCenter\\plugins\\<id>\\
#   then click Reload in Plugins dialog — no reinstall needed

=== CONSTRAINTS & RULES ===
1. activate() and deactivate() are MANDATORY top-level functions
2. deactivate() MUST undo every side effect; must be safe to call before activate()
3. NEVER block main thread: no sleep(), no synchronous HTTP, no heavy I/O in callbacks
4. ALWAYS use api.plugin_dir for data files — never hardcode paths
5. ALWAYS guard callbacks with try/except to prevent silent Qt event loop swallowing
6. ALWAYS unregister every hotkey in deactivate()
7. ALWAYS remove every footer button in deactivate()
8. ALWAYS remove every settings tab in deactivate()
9. Style dialogs using api.theme.colors() — re-apply on api.theme.register(fn)
10. Use api.timers.create() for periodic tasks — auto-cleaned, no manual cancel needed
11. Plugin "id" in manifest.json is permanent — never change after first release
12. PySide6 imports are always available: from PySide6.QtWidgets import ...

=== ERROR HANDLING REQUIREMENTS ===
# Wrap every callback / event handler:
def _on_click():
    try:
        _do_work()
    except Exception as e:
        _api.log(f"_on_click error: {e}")
        _api.toast("An error occurred.", "error")

# Wrap file operations:
def _load_data():
    try:
        return _api.files.read_json(_api.plugin_dir / "data.json")
    except (OSError, ValueError):
        return {}

# Wrap deactivate():
def deactivate():
    global _btn
    try:
        if _btn is not None:
            _api.ui.remove_footer_button(_btn)
            _btn = None
    except Exception as e:
        pass  # deactivate should never raise

=== IMPORTS AVAILABLE WITHOUT INSTALL ===
PySide6.QtWidgets  PySide6.QtCore  PySide6.QtGui  PySide6.QtNetwork
pathlib  json  os  sys  subprocess  threading  re  datetime  time
functools  collections  itertools  typing  dataclasses  copy  math

=== CC BUILT-IN HOTKEYS (do not conflict) ===
Ctrl+F         Open search bar
Ctrl+N         New node wizard
Ctrl+Shift+N   New folder
Ctrl+,         Open Settings
F1             Help dialog
Ctrl+Shift+T   Toggle Time Tracker
Ctrl+Shift+B   Open Notebook
Ctrl+Shift+M   Open Media Library
Ctrl+Z / Ctrl+Y  Undo/Redo (canvas)
Delete         Delete selected node
Ctrl+A         Select all nodes
"""


class _PluginGuideDialog(QDialog):
    """Scrollable developer guide shown from the Plugins dialog."""

    def __init__(self, parent=None):
        super().__init__(parent, Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self._drag_pos: Optional[QPoint] = None
        self._build_ui()
        self.setMinimumSize(640, 580)
        self.resize(740, 680)

    def _build_ui(self):
        t = _theme
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(0)

        card = QWidget()
        card.setObjectName("guide_card")
        cl = QVBoxLayout(card)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(0)
        card.setGraphicsEffect(make_shadow(card, 28, QColor(0, 0, 0, 200)))
        root.addWidget(card)

        # Title bar
        tbar = QWidget()
        tbar.setFixedHeight(44)
        tbl  = QHBoxLayout(tbar)
        tbl.setContentsMargins(18, 0, 10, 0)
        tlbl = QLabel("📖  .ccplug Developer Guide")
        tlbl.setFont(QFont("Segoe UI", 10, QFont.Weight.DemiBold))
        tlbl.setStyleSheet(f"color:{t.GLOW.name()}; background:transparent;")
        tbl.addWidget(tlbl)
        tbl.addStretch()
        copy_guide_btn = QPushButton("⎘  Copy")
        copy_guide_btn.setCursor(Qt.PointingHandCursor)
        copy_guide_btn.setFont(QFont("Segoe UI", 8))
        copy_guide_btn.setFixedHeight(26)
        copy_guide_btn.setStyleSheet(f"""
            QPushButton {{ background:transparent; border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
                           border-radius:4px; color:{t.TEXT_SECONDARY.name()}; padding:0 10px; }}
            QPushButton:hover {{ background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},35);
                                 color:{t.GLOW.name()}; }}
        """)
        copy_guide_btn.clicked.connect(lambda: (
            QGuiApplication.clipboard().setText(_CCPLUG_GUIDE),
            copy_guide_btn.setText("✔  Copied"),
            QTimer.singleShot(1500, lambda: copy_guide_btn.setText("⎘  Copy")),
        ))
        tbl.addWidget(copy_guide_btn)
        tbl.addSpacing(4)
        copy_ai_btn = QPushButton("🤖  Copy AI")
        copy_ai_btn.setCursor(Qt.PointingHandCursor)
        copy_ai_btn.setFont(QFont("Segoe UI", 8))
        copy_ai_btn.setFixedHeight(26)
        copy_ai_btn.setToolTip("Copy AI-optimized plugin-building instructions to clipboard")
        copy_ai_btn.setStyleSheet(f"""
            QPushButton {{ background:transparent; border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
                           border-radius:4px; color:{t.TEXT_SECONDARY.name()}; padding:0 10px; }}
            QPushButton:hover {{ background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},35);
                                 color:{t.GLOW.name()}; }}
        """)
        copy_ai_btn.clicked.connect(lambda: (
            QGuiApplication.clipboard().setText(_CC_PLUGIN_AI_GUIDE),
            copy_ai_btn.setText("✔  Copied"),
            QTimer.singleShot(1500, lambda: copy_ai_btn.setText("🤖  Copy AI")),
        ))
        tbl.addWidget(copy_ai_btn)
        tbl.addSpacing(6)
        close_btn = QPushButton("✕")
        close_btn.setFixedSize(26, 26)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setStyleSheet(f"""
            QPushButton {{ background:transparent; border:none; color:{t.TEXT_DIM.name()};
                           border-radius:13px; font-size:10pt; }}
            QPushButton:hover {{
                background:rgba({t.ACCENT_RED.red()},{t.ACCENT_RED.green()},{t.ACCENT_RED.blue()},140);
                color:#fff; }}
        """)
        close_btn.clicked.connect(self.close)
        tbl.addWidget(close_btn)
        cl.addWidget(tbar)

        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},35);")
        cl.addWidget(sep)

        # Scrollable text content
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet(f"""
            QScrollArea {{ background:transparent; border:none; }}
            QScrollBar:vertical {{
                background:rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},120);
                width:6px; border-radius:3px;
            }}
            QScrollBar::handle:vertical {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},100);
                border-radius:3px; min-height:20px;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height:0; }}
        """)
        content = QWidget()
        content.setStyleSheet("background:transparent;")
        cll = QVBoxLayout(content)
        cll.setContentsMargins(20, 16, 20, 20)

        txt = QTextEdit()
        txt.setReadOnly(True)
        txt.setPlainText(_CCPLUG_GUIDE)
        txt.setFont(QFont("Consolas", 9))
        txt.setStyleSheet(f"""
            QTextEdit {{
                background:transparent; border:none;
                color:{t.TEXT_SECONDARY.name()};
                selection-background-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
            }}
        """)
        cll.addWidget(txt)
        scroll.setWidget(content)
        cl.addWidget(scroll, stretch=1)

    def paintEvent(self, e):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        r = QRectF(8, 8, self.width() - 16, self.height() - 16)
        path = QPainterPath(); path.addRoundedRect(r, 14, 14)
        grad = QLinearGradient(0, 0, 0, self.height())
        grad.setColorAt(0, t.BG_MID)
        grad.setColorAt(1, t.BG_DARK)
        p.fillPath(path, grad)
        p.setPen(QPen(QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 55), 1.5))
        p.drawPath(path)
        p.end()

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_pos = e.globalPosition().toPoint() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if self._drag_pos and e.buttons() & Qt.LeftButton:
            self.move(e.globalPosition().toPoint() - self._drag_pos)

    def mouseReleaseEvent(self, e):
        self._drag_pos = None


# ===========================================================================
# End plugin system
# ===========================================================================


class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent, Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self._drag_pos: Optional[QPoint] = None
        self._build_ui()
        self.setMinimumSize(560, 560)

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(10, 10, 10, 10); outer.setSpacing(0)
        card = QWidget()
        cl = QVBoxLayout(card); cl.setContentsMargins(0, 0, 0, 0); cl.setSpacing(0)
        card.setGraphicsEffect(make_shadow(card, 30, QColor(0, 0, 0, 200)))

        tbar = QWidget(); tbar.setFixedHeight(44)
        tbar.setStyleSheet("background:transparent;")
        tbl = QHBoxLayout(tbar); tbl.setContentsMargins(16, 0, 12, 0)
        ttl = QLabel("❓  Help & FAQ")
        ttl.setFont(QFont("Segoe UI", 11, QFont.Weight.DemiBold))
        ttl.setStyleSheet(f"color:{_theme.TEXT_PRIMARY.name()};")
        tbl.addWidget(ttl); tbl.addStretch()
        cb = TitleBarButton(COLOR_BTN_CLOSE, "x"); cb.clicked.connect(self.accept)
        tbl.addWidget(cb)
        tbar.mousePressEvent = lambda ev: (
            setattr(self, "_drag_pos", ev.globalPosition().toPoint())
            if ev.button() == Qt.LeftButton else None)
        tbar.mouseMoveEvent = self._drag_move
        tbar.mouseReleaseEvent = lambda ev: setattr(self, "_drag_pos", None)
        cl.addWidget(tbar)

        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(
            f"color:rgba({_theme.GLOW.red()},{_theme.GLOW.green()},{_theme.GLOW.blue()},55);")
        cl.addWidget(sep)

        from PySide6.QtWidgets import QTextBrowser
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        t = _theme
        browser.setStyleSheet(f"""
            QTextBrowser {{
                background:transparent; border:none;
                color:{t.TEXT_PRIMARY.name()};
                font-family:'Segoe UI'; font-size:9pt;
                selection-background-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
            }}
            QScrollBar:vertical {{
                background:rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},160);
                width:6px; border-radius:3px;
            }}
            QScrollBar::handle:vertical {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},110);
                border-radius:3px; min-height:20px;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height:0; }}
        """)
        # Inject theme-aware colours into the HTML head
        html = _HELP_CONTENT.replace(
            "<style>",
            f"<style> body {{ color:{t.TEXT_PRIMARY.name()}; }} "
            f"h2 {{ color:{t.GLOW.name()}; }} "
            f"h3 {{ color:{t.TEXT_SECONDARY.name()}; }} "
        )
        browser.setHtml(html)
        browser_container = QWidget(); browser_container.setStyleSheet("background:transparent;")
        bcl = QVBoxLayout(browser_container); bcl.setContentsMargins(16, 8, 16, 16)
        bcl.addWidget(browser)
        cl.addWidget(browser_container)
        outer.addWidget(card)

    def _drag_move(self, e: QMouseEvent):
        if self._drag_pos:
            delta = e.globalPosition().toPoint() - self._drag_pos
            self.move(self.pos() + delta)
            self._drag_pos = e.globalPosition().toPoint()

    def paintEvent(self, e: QPaintEvent):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        rect = QRectF(10, 10, self.width() - 20, self.height() - 20)
        path = QPainterPath(); path.addRoundedRect(rect, BORDER_RADIUS + 2, BORDER_RADIUS + 2)
        grad = QLinearGradient(0, 0, self.width(), self.height())
        grad.setColorAt(0, QColor(t.BG_MID.red()+4, t.BG_MID.green()+4, t.BG_MID.blue()+6, 248))
        grad.setColorAt(1, QColor(t.BG_DARK.red(), t.BG_DARK.green(), t.BG_DARK.blue(), 248))
        p.fillPath(path, grad)
        border = QColor(t.GLOW); border.setAlpha(90)
        p.setPen(QPen(border, 1.3)); p.setBrush(Qt.NoBrush); p.drawPath(path)
        p.end()


# ---------------------------------------------------------------------------
# Hotkey capture widgets  (used by the Settings → Hotkeys tab)
# ---------------------------------------------------------------------------

def _resolve_key_name(key) -> str:
    """Resolve a Qt key value (int or enum) to its canonical display name.

    Handles modifier keys explicitly, then falls back to QKeySequence for
    regular keys (e.g. 81 → 'Q', 16777264 → 'F1').
    """
    # Explicit modifier map so QKeySequence(Ctrl-alone) edge cases are avoided
    try:
        _MOD_MAP = {
            int(Qt.Key_Control): "Ctrl",
            int(Qt.Key_Shift):   "Shift",
            int(Qt.Key_Alt):     "Alt",
            int(Qt.Key_Meta):    "Meta",
            int(Qt.Key_AltGr):   "AltGr",
        }
        m = _MOD_MAP.get(int(key))
        if m:
            return m
    except Exception:
        pass
    try:
        from PySide6.QtGui import QKeySequence as _QKS
        name = _QKS(int(key)).toString()
        if name:
            return name
    except Exception:
        pass
    try:
        raw = getattr(key, "name", "")
        return raw.replace("Key_", "") if raw.startswith("Key_") else raw
    except Exception:
        return ""


class _HotkeyCaptureFiler(QObject):
    """Application-level event filter active only while a hotkey is being recorded.

    Intercepts KeyPress *and* KeyRelease so multi-key chords (e.g. Ctrl+Q+W)
    can be tracked.  Also swallows ShortcutOverride so no QShortcut fires
    during capture.
    """
    def __init__(self):
        super().__init__()
        self._active_row: Optional["_HotkeyRowWidget"] = None
        self._held_non_mods: list[str] = []   # non-modifier key names currently held

    def set_active(self, row: Optional["_HotkeyRowWidget"]):
        self._active_row = row
        self._held_non_mods.clear()

    def eventFilter(self, obj, event):
        t = event.type()
        if t == QEvent.Type.ShortcutOverride:
            event.accept()
            return True
        if self._active_row is None:
            return False
        if t == QEvent.Type.KeyPress:
            key = event.key()
            # Track non-modifier, non-special keys that are currently held
            if key not in (Qt.Key_Control, Qt.Key_Shift, Qt.Key_Alt, Qt.Key_Meta,
                           Qt.Key_AltGr, Qt.Key_CapsLock, Qt.Key_NumLock,
                           Qt.Key_ScrollLock, Qt.Key_unknown,
                           Qt.Key_Escape, Qt.Key_Return, Qt.Key_Enter):
                name = _resolve_key_name(key)
                if name and name not in self._held_non_mods:
                    self._held_non_mods.append(name)
            self._active_row._process_key(event)
            return True
        if t == QEvent.Type.KeyRelease:
            name = _resolve_key_name(event.key())
            if name in self._held_non_mods:
                self._held_non_mods.remove(name)
            return True
        return False


# Singleton filter — installed/removed as needed
_hotkey_capture_filter: Optional[_HotkeyCaptureFiler] = None


class _MainWindowKeyTracker(QObject):
    """App-level event filter that tracks simultaneously-held keys for the
    multi-non-modifier shortcut system (e.g. ``Ctrl+Q+W``).

    Installed once at MainWindow startup and stays installed.  Always returns
    False so it never consumes events — every widget still receives its events.
    """

    def __init__(self, window):
        super().__init__()
        self._win = window

    def eventFilter(self, obj, event):
        t = event.type()
        if t == QEvent.Type.KeyPress and not event.isAutoRepeat():
            name = _resolve_key_name(event.key()).lower()
            if name:
                self._win._held_keys.add(name)
                self._win._check_multi_key_bindings()
        elif t == QEvent.Type.KeyRelease and not event.isAutoRepeat():
            name = _resolve_key_name(event.key()).lower()
            if name:
                self._win._held_keys.discard(name)
        return False   # never consume events


class _HotkeyRowWidget(QWidget):
    """One row in the Hotkeys settings tab.

    Shows an action label, the current key sequence, a ⏺ Record button that
    enters capture mode, and a ↺ Reset button.  Emits ``hotkey_changed``
    whenever the user confirms a new sequence or resets to the default.
    """
    hotkey_changed = Signal(str, str)   # (action_id, new_sequence)

    def __init__(self, action_id: str, pending_hotkeys: dict, parent=None):
        super().__init__(parent)
        self._action_id     = action_id
        self._pending       = pending_hotkeys   # shared dict owned by the tab
        self._capturing     = False
        self.setFocusPolicy(Qt.StrongFocus)
        self._build()

    # ── Construction ──────────────────────────────────────────────────────

    def _build(self):
        t, g = _theme, _theme.GLOW
        h = QHBoxLayout(self)
        h.setContentsMargins(4, 2, 4, 2)
        h.setSpacing(8)

        # Action label
        label = QLabel(_CC_HOTKEY_LABELS.get(self._action_id, self._action_id))
        label.setFont(QFont("Segoe UI", 9))
        label.setStyleSheet(f"color:{t.TEXT_PRIMARY.name()};")
        label.setMinimumWidth(180)
        label.setMaximumWidth(180)
        h.addWidget(label)

        # Key sequence display
        self._seq_lbl = QLabel()
        self._seq_lbl.setMinimumWidth(130)
        self._seq_lbl.setMaximumWidth(160)
        self._seq_lbl.setFixedHeight(26)
        self._seq_lbl.setAlignment(Qt.AlignCenter)
        self._seq_lbl.setFont(QFont("Consolas", 9))
        self._refresh_display()
        h.addWidget(self._seq_lbl)

        h.addStretch()

        # Record button
        self._rec_btn = QPushButton("⏺")
        self._rec_btn.setFixedSize(28, 26)
        self._rec_btn.setCursor(Qt.PointingHandCursor)
        self._rec_btn.setToolTip("Click to record a new hotkey\n(press Escape to cancel)")
        self._rec_btn.clicked.connect(self._toggle_capture)
        h.addWidget(self._rec_btn)

        # Save/confirm button (visible only during capture)
        self._save_btn = QPushButton("✓")
        self._save_btn.setFixedSize(28, 26)
        self._save_btn.setCursor(Qt.PointingHandCursor)
        self._save_btn.setToolTip("Save this hotkey (or press Enter)")
        self._save_btn.clicked.connect(self._save_capture)
        self._save_btn.setVisible(False)
        h.addWidget(self._save_btn)

        # Reset button
        self._rst_btn = QPushButton("↺")
        self._rst_btn.setFixedSize(28, 26)
        self._rst_btn.setCursor(Qt.PointingHandCursor)
        self._rst_btn.setToolTip("Reset to default")
        self._rst_btn.clicked.connect(self._reset_to_default)
        h.addWidget(self._rst_btn)

        self._apply_btn_styles()

    # ── Public helpers ─────────────────────────────────────────────────────

    def current_sequence(self) -> str:
        """Return the sequence currently shown (pending or saved)."""
        return self._pending.get(self._action_id, _get_hotkey(self._action_id))

    def set_conflict(self, has_conflict: bool):
        pass  # conflict feedback is now handled by the tab-level notice banner

    def update_from_pending(self):
        """Refresh display after an external change to the shared pending dict."""
        self._refresh_display()

    # ── Internal helpers ───────────────────────────────────────────────────

    def _refresh_display(self, capturing: bool = False):
        t, g = _theme, _theme.GLOW
        if capturing:
            staged = getattr(self, "_staged_seq", "")
            text = ("  " + staged) if staged else "  Press keys…"
            self._seq_lbl.setText(text)
            self._seq_lbl.setStyleSheet(
                f"background:rgba({g.red()},{g.green()},{g.blue()},30);"
                f"border:1px solid {g.name()};"
                f"border-radius:4px; color:{g.name()}; padding:2px 8px;")
        else:
            seq = self.current_sequence()
            self._seq_lbl.setText(seq or "(none)")
            self._seq_lbl.setStyleSheet(
                f"background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);"
                f"border:1px solid rgba({g.red()},{g.green()},{g.blue()},60);"
                f"border-radius:4px; color:{t.TEXT_PRIMARY.name()}; padding:2px 8px;")

    def _apply_btn_styles(self):
        t, g = _theme, _theme.GLOW
        css = (
            f"QPushButton {{"
            f"  background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},180);"
            f"  border:1px solid rgba({g.red()},{g.green()},{g.blue()},70);"
            f"  border-radius:4px; color:{t.TEXT_SECONDARY.name()}; font-size:12pt;"
            f"}}"
            f"QPushButton:hover {{ color:{g.name()}; border-color:rgba({g.red()},{g.green()},{g.blue()},140); }}"
            f"QPushButton:disabled {{ color:rgba(128,128,128,80); border-color:rgba(128,128,128,30); }}"
        )
        self._rec_btn.setStyleSheet(css)
        self._save_btn.setStyleSheet(css)
        self._rst_btn.setStyleSheet(css)

    # ── Capture control ────────────────────────────────────────────────────

    def _toggle_capture(self):
        if self._capturing:
            self._stop_capture(cancelled=True)
        else:
            self._start_capture()

    def _start_capture(self):
        global _hotkey_capture_filter
        self._staged_seq = ""
        self._capturing = True
        self._refresh_display(capturing=True)
        self._rec_btn.setText("✕")
        self._rec_btn.setToolTip("Cancel recording (Escape)")
        self._save_btn.setVisible(True)
        self._rst_btn.setEnabled(False)
        # Install the app-level event filter that routes KeyPress to us
        if _hotkey_capture_filter is None:
            _hotkey_capture_filter = _HotkeyCaptureFiler()
        _hotkey_capture_filter.set_active(self)
        QApplication.instance().installEventFilter(_hotkey_capture_filter)

    def _stop_capture(self, cancelled: bool = False):
        global _hotkey_capture_filter
        if not self._capturing:
            return
        self._capturing = False
        if _hotkey_capture_filter is not None:
            _hotkey_capture_filter.set_active(None)
            QApplication.instance().removeEventFilter(_hotkey_capture_filter)
        self._rec_btn.setText("⏺")
        self._rec_btn.setToolTip("Click to record a new hotkey\n(press Escape to cancel)")
        self._save_btn.setVisible(False)
        self._rst_btn.setEnabled(True)
        seq_str = getattr(self, "_staged_seq", "")
        if not cancelled and seq_str:
            self._pending[self._action_id] = seq_str
            self.hotkey_changed.emit(self._action_id, seq_str)
        self._refresh_display(capturing=False)

    def _save_capture(self):
        """Confirm and save the staged hotkey sequence (called by the ✓ button)."""
        self._stop_capture(cancelled=False)

    def _reset_to_default(self):
        default = _CC_DEFAULT_HOTKEYS.get(self._action_id, "")
        self._pending[self._action_id] = default
        self._refresh_display()
        self.hotkey_changed.emit(self._action_id, default)

    # ── Key processing (called by the app-level event filter) ─────────────

    def _process_key(self, event):
        """Handle a KeyPress event routed from _HotkeyCaptureFiler."""
        key  = event.key()
        mods = event.modifiers()

        # Escape cancels without saving
        if key == Qt.Key_Escape:
            self._stop_capture(cancelled=True)
            return

        # Enter/Return confirms and saves the staged sequence
        if key in (Qt.Key_Return, Qt.Key_Enter):
            self._stop_capture(cancelled=False)
            return

        # Ignore standalone modifier keys — wait for a real key
        if key in (Qt.Key_Control, Qt.Key_Shift, Qt.Key_Alt, Qt.Key_Meta,
                   Qt.Key_AltGr, Qt.Key_CapsLock, Qt.Key_NumLock,
                   Qt.Key_ScrollLock, Qt.Key_unknown):
            return

        # Build chord: standard modifiers first, then ALL currently-held
        # non-modifier keys.  _held_non_mods was updated by _HotkeyCaptureFiler
        # BEFORE this method was called, so it already includes the current key.
        parts = []
        if mods & Qt.ControlModifier: parts.append("Ctrl")
        if mods & Qt.AltModifier:     parts.append("Alt")
        if mods & Qt.ShiftModifier:   parts.append("Shift")
        if mods & Qt.MetaModifier:    parts.append("Meta")
        for nm in _hotkey_capture_filter._held_non_mods:
            if nm not in parts:
                parts.append(nm)

        self._staged_seq = "+".join(parts) if parts else ""
        self._refresh_display(capturing=True)

    # keyPressEvent / keyReleaseEvent are intentionally NOT overridden here.
    # All key capture is handled by _HotkeyCaptureFiler at the app level.


# ---------------------------------------------------------------------------
# Custom color picker dialog
# Replaces QColorDialog to avoid the stylesheet-induced click-offset bug
# where clicking a swatch would apply an entirely different color.
# ---------------------------------------------------------------------------

class _HueSliderWidget(QWidget):
    """A horizontal bar painted with the full HSV hue spectrum (0–359)."""
    hue_changed = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(20)
        self.setCursor(Qt.PointingHandCursor)
        self._hue = 0

    def set_hue(self, hue: int):
        self._hue = max(0, min(359, hue))
        self.update()

    def hue(self) -> int:
        return self._hue

    def paintEvent(self, _e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, False)
        w, h = self.width(), self.height()
        # Draw rainbow gradient
        for x in range(w):
            hue = int(x / w * 360)
            p.setPen(QColor.fromHsv(hue, 255, 255))
            p.drawLine(x, 0, x, h - 3)
        # Draw indicator
        ix = int(self._hue / 360 * w)
        p.setPen(QPen(Qt.white, 2))
        p.drawLine(ix, 0, ix, h - 1)

    def _pick(self, x: int):
        hue = int(max(0, min(x, self.width() - 1)) / self.width() * 359)
        self._hue = hue
        self.update()
        self.hue_changed.emit(hue)

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._pick(e.position().toPoint().x())

    def mouseMoveEvent(self, e):
        if e.buttons() == Qt.LeftButton:
            self._pick(e.position().toPoint().x())


class _SVPickerWidget(QWidget):
    """2-D saturation × value square for the currently selected hue."""
    sv_changed = Signal(int, int)   # saturation, value (0–255 each)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(200, 200)
        self.setCursor(Qt.CrossCursor)
        self._hue = 0
        self._sat = 255
        self._val = 255
        self._img: Optional[QImage] = None

    def set_hsv(self, hue: int, sat: int, val: int):
        changed_hue = (hue != self._hue)
        self._hue = hue
        self._sat = sat
        self._val = val
        if changed_hue:
            self._img = None   # invalidate cached gradient
        self.update()

    def set_hue(self, hue: int):
        if hue != self._hue:
            self._hue = hue
            self._img = None
            self.update()

    def paintEvent(self, _e):
        p = QPainter(self)
        w, h = self.width(), self.height()

        # Build or reuse the SV gradient image
        if self._img is None or self._img.width() != w or self._img.height() != h:
            self._img = QImage(w, h, QImage.Format_RGB32)
            for xi in range(w):
                s = int(xi / (w - 1) * 255)
                for yi in range(h):
                    v = int((1 - yi / (h - 1)) * 255)
                    self._img.setPixel(xi, yi, QColor.fromHsv(self._hue, s, v).rgb())

        p.drawImage(0, 0, self._img)

        # Draw crosshair
        cx = int(self._sat / 255 * (w - 1))
        cy = int((1 - self._val / 255) * (h - 1))
        p.setPen(QPen(Qt.white, 1, Qt.SolidLine))
        p.drawEllipse(QPoint(cx, cy), 6, 6)
        p.setPen(QPen(Qt.black, 1, Qt.SolidLine))
        p.drawEllipse(QPoint(cx, cy), 7, 7)

    def _pick(self, x: int, y: int):
        w, h = self.width(), self.height()
        s = int(max(0, min(x, w - 1)) / (w - 1) * 255)
        v = int((1 - max(0, min(y, h - 1)) / (h - 1)) * 255)
        self._sat = s
        self._val = v
        self.update()
        self.sv_changed.emit(s, v)

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._pick(e.position().toPoint().x(), e.position().toPoint().y())

    def mouseMoveEvent(self, e):
        if e.buttons() == Qt.LeftButton:
            self._pick(e.position().toPoint().x(), e.position().toPoint().y())


class _CCColorPickerDialog(QDialog):
    """
    Fully custom color picker — avoids QColorDialog entirely.

    The root cause of the wrong-color-on-click bug was that applying a dark
    stylesheet to QColorDialog shifts the visual positions of its internal
    color-swatch cells while leaving their mouse hit-boxes in the original
    positions, causing clicks to land on a different color than the one
    the user sees.  This dialog paints everything itself.

    Layout:
        [ SV picker 200×200 ]  [ Hue bar (full width) ]
                               [ R  slider + spinbox  ]
                               [ G  slider + spinbox  ]
                               [ B  slider + spinbox  ]
                               [ Hex  #RRGGBB input   ]
                               [ Before | After swatches ]
        [ 16-color quick palette                       ]
        [                         Cancel ]  [ Apply   ]
    """

    _PALETTE = [
        "#FF0000", "#FF6600", "#FFCC00", "#FFFF00",
        "#00FF00", "#00CC66", "#00CCCC", "#0066FF",
        "#0000FF", "#6600CC", "#FF00FF", "#FF66CC",
        "#FFFFFF", "#AAAAAA", "#555555", "#000000",
    ]

    def __init__(self, initial: QColor, label: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Choose Color — {label}")
        self.setModal(True)
        self.setFixedWidth(500)
        self._color     = QColor(initial)
        self._old_color = QColor(initial)
        self._updating  = False
        self._build_ui()
        self._apply_theme_style()
        self._sync_from_color()

    # ── UI construction ──────────────────────────────────────────────────────

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setSpacing(10)
        lay.setContentsMargins(16, 16, 16, 14)

        # ── top row: SV picker + right controls ──
        top = QHBoxLayout(); top.setSpacing(14)

        self._sv = _SVPickerWidget(self)
        self._sv.sv_changed.connect(self._on_sv_changed)
        top.addWidget(self._sv)

        right = QVBoxLayout(); right.setSpacing(6)

        # Hue bar
        hue_row = QHBoxLayout(); hue_row.setSpacing(6)
        hue_lbl = QLabel("H"); hue_lbl.setFixedWidth(14)
        self._hue_bar = _HueSliderWidget(self)
        self._hue_bar.hue_changed.connect(self._on_hue_changed)
        self._hue_spin = QLineEdit()
        self._hue_spin.setValidator(QIntValidator(0, 359))
        self._hue_spin.setFixedWidth(50)
        self._hue_spin.setAlignment(Qt.AlignCenter)
        self._hue_spin.textEdited.connect(self._on_hue_spin_changed)
        hue_row.addWidget(hue_lbl); hue_row.addWidget(self._hue_bar, 1)
        hue_row.addWidget(self._hue_spin)
        right.addLayout(hue_row)

        right.addSpacing(6)

        # R / G / B sliders
        self._rgb_sliders: dict[str, tuple] = {}
        for ch_label, attr, color_hex in [
            ("R", "_r", "#DD4444"),
            ("G", "_g", "#44BB44"),
            ("B", "_b", "#4488DD"),
        ]:
            row = QHBoxLayout(); row.setSpacing(6)
            lbl = QLabel(ch_label); lbl.setFixedWidth(14)
            sl = QSlider(Qt.Horizontal)
            sl.setRange(0, 255); sl.setFixedHeight(16)
            sl.setStyleSheet(
                f"QSlider::groove:horizontal{{height:6px;border-radius:3px;"
                f"  background:rgba(80,80,80,160);}}"
                f"QSlider::sub-page:horizontal{{background:{color_hex};border-radius:3px;}}"
                f"QSlider::handle:horizontal{{width:14px;height:14px;margin:-4px 0;"
                f"  border-radius:7px;background:#e8e8e8;}}")
            sp = QLineEdit()
            sp.setValidator(QIntValidator(0, 255))
            sp.setFixedWidth(50)
            sp.setAlignment(Qt.AlignCenter)
            sl.valueChanged.connect(lambda v, a=attr: self._on_rgb_slider(a, v))
            sp.textEdited.connect(lambda text, a=attr: self._on_rgb_spin(a, text))
            self._rgb_sliders[attr] = (sl, sp)
            row.addWidget(lbl); row.addWidget(sl, 1); row.addWidget(sp)
            right.addLayout(row)

        right.addSpacing(6)

        # Hex input
        hex_row = QHBoxLayout(); hex_row.setSpacing(6)
        hex_lbl = QLabel("#"); hex_lbl.setFixedWidth(14)
        self._hex_edit = QLineEdit(); self._hex_edit.setMaxLength(6)
        self._hex_edit.setPlaceholderText("RRGGBB")
        self._hex_edit.setFixedWidth(80)
        self._hex_edit.textEdited.connect(self._on_hex_edited)
        hex_row.addWidget(hex_lbl); hex_row.addWidget(self._hex_edit)
        hex_row.addStretch()
        right.addLayout(hex_row)

        right.addStretch()

        # Before / After swatches
        sw_row = QHBoxLayout(); sw_row.setSpacing(8)
        sw_row.addStretch()
        for title, attr in [("Before", "_old_sw"), ("After", "_new_sw")]:
            col = QVBoxLayout(); col.setSpacing(2)
            lbl = QLabel(title); lbl.setAlignment(Qt.AlignCenter)
            lbl.setStyleSheet("font-size:8pt; color:#aaa;")
            sw = QLabel(); sw.setFixedSize(52, 36)
            sw.setStyleSheet("border:1px solid #555; border-radius:3px;")
            col.addWidget(lbl); col.addWidget(sw)
            sw_row.addLayout(col)
            setattr(self, attr, sw)
        self._old_sw.setToolTip("Original color — click to revert")
        self._old_sw.setCursor(Qt.PointingHandCursor)
        self._old_sw.mousePressEvent = lambda _e: self._revert()
        right.addLayout(sw_row)

        top.addLayout(right)
        lay.addLayout(top)

        # ── palette strip ──
        pal = QHBoxLayout(); pal.setSpacing(3)
        for hex_c in self._PALETTE:
            b = QPushButton(); b.setFixedSize(22, 22)
            b.setToolTip(hex_c)
            b.setCursor(Qt.PointingHandCursor)
            b.setStyleSheet(
                f"QPushButton{{background:{hex_c};"
                f"  border:1px solid rgba(255,255,255,40);border-radius:3px;}}"
                f"QPushButton:hover{{border:2px solid #fff;}}")
            b.clicked.connect(lambda _=False, c=hex_c: self._pick_palette(c))
            pal.addWidget(b)
        pal.addStretch()
        lay.addLayout(pal)

        # ── buttons ──
        btn_row = QHBoxLayout(); btn_row.setSpacing(8)
        btn_row.addStretch()
        cancel = QPushButton("Cancel"); cancel.setFixedSize(90, 32)
        cancel.clicked.connect(self.reject)
        apply  = QPushButton("Apply");  apply.setFixedSize(90, 32)
        apply.setDefault(True)
        apply.clicked.connect(self.accept)
        btn_row.addWidget(cancel); btn_row.addWidget(apply)
        lay.addLayout(btn_row)
        self._apply_btn = apply
        self._cancel_btn = cancel

    def _apply_theme_style(self):
        try:
            t = _theme
            g = t.GLOW
            self.setStyleSheet(
                f"QDialog{{background:{t.BG_DARK.name()};color:{t.TEXT_PRIMARY.name()};}}"
                f"QLabel{{color:{t.TEXT_PRIMARY.name()};}}"
                f"QLineEdit{{background:{t.TILE_BG_BASE.name()};"
                f"  color:{t.TEXT_PRIMARY.name()};border:1px solid rgba(255,255,255,30);"
                f"  border-radius:3px;padding:1px 5px;}}"
            )
            btn_style = (
                f"QPushButton{{background:rgba({t.TILE_BG_BASE.red()},"
                f"{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},180);"
                f"  color:{t.TEXT_SECONDARY.name()};"
                f"  border:1px solid rgba({g.red()},{g.green()},{g.blue()},70);"
                f"  border-radius:5px;}}"
                f"QPushButton:hover{{background:rgba({g.red()},{g.green()},{g.blue()},50);"
                f"  color:{t.TEXT_PRIMARY.name()};}}"
            )
            apply_style = (
                f"QPushButton{{background:rgba({g.red()},{g.green()},{g.blue()},80);"
                f"  color:{g.name()};font-weight:600;"
                f"  border:1px solid rgba({g.red()},{g.green()},{g.blue()},160);"
                f"  border-radius:5px;}}"
                f"QPushButton:hover{{background:rgba({g.red()},{g.green()},{g.blue()},130);}}"
            )
            self._cancel_btn.setStyleSheet(btn_style)
            self._apply_btn.setStyleSheet(apply_style)
        except Exception:
            pass

    # ── sync helpers ──────────────────────────────────────────────────────────

    def _sync_from_color(self):
        self._updating = True
        c = self._color
        h, s, v, _ = c.getHsv()
        if h < 0:
            h = 0
        # SV picker
        self._sv.set_hsv(h, s, v)
        # Hue bar + spin
        self._hue_bar.set_hue(h)
        self._hue_spin.setText(str(h))
        # RGB sliders + spinboxes
        for attr, val in [("_r", c.red()), ("_g", c.green()), ("_b", c.blue())]:
            sl, sp = self._rgb_sliders[attr]
            sl.setValue(val)
            sp.setText(str(val))
        # Hex
        self._hex_edit.setText(c.name()[1:].upper())
        # Swatches
        self._new_sw.setStyleSheet(
            f"background:{c.name()};border:1px solid #555;border-radius:3px;")
        self._old_sw.setStyleSheet(
            f"background:{self._old_color.name()};border:1px solid #555;border-radius:3px;")
        self._updating = False

    # ── slot handlers ─────────────────────────────────────────────────────────

    def _on_hue_changed(self, hue: int):
        if self._updating:
            return
        _, s, v, _ = self._color.getHsv()
        self._color.setHsv(hue, max(0, s), max(0, v))
        self._sync_from_color()

    def _on_hue_spin_changed(self, text: str):
        if self._updating:
            return
        try:
            hue = max(0, min(359, int(text)))
        except (ValueError, OverflowError):
            return
        self._hue_bar.set_hue(hue)
        self._on_hue_changed(hue)

    def _on_sv_changed(self, s: int, v: int):
        if self._updating:
            return
        h, _, _, _ = self._color.getHsv()
        self._color.setHsv(max(0, h), s, v)
        self._sync_from_color()

    def _on_rgb_slider(self, attr: str, val: int):
        if self._updating:
            return
        sl, sp = self._rgb_sliders[attr]
        self._color = QColor(
            self._rgb_sliders["_r"][0].value(),
            self._rgb_sliders["_g"][0].value(),
            self._rgb_sliders["_b"][0].value(),
        )
        self._sync_from_color()

    def _on_rgb_spin(self, attr: str, text: str):
        if self._updating:
            return
        try:
            val = max(0, min(255, int(text)))
        except (ValueError, OverflowError):
            return
        sl, _ = self._rgb_sliders[attr]
        sl.setValue(val)   # triggers _on_rgb_slider

    def _on_hex_edited(self, text: str):
        if self._updating:
            return
        c = QColor(f"#{text}")
        if c.isValid():
            self._color = c
            self._sync_from_color()

    def _pick_palette(self, hex_c: str):
        c = QColor(hex_c)
        if c.isValid():
            self._color = c
            self._sync_from_color()

    def _revert(self):
        self._color = QColor(self._old_color)
        self._sync_from_color()

    def selected_color(self) -> QColor:
        return QColor(self._color)


# ---------------------------------------------------------------------------
# Settings dialog
# ---------------------------------------------------------------------------

class SettingsDialog(QDialog):
    settings_changed = Signal()

    def __init__(self, store: "NodeStore", parent=None):
        super().__init__(parent, Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self._store = store
        self._drag_pos: Optional[QPoint] = None
        self._sec_labels: list = []
        self._build_ui()
        self.setMinimumSize(560, 500)

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(10,10,10,10); outer.setSpacing(0)
        self._card = QWidget()
        cl = QVBoxLayout(self._card)
        cl.setContentsMargins(0,0,0,0); cl.setSpacing(0)
        self._card.setGraphicsEffect(make_shadow(self._card,30,QColor(0,0,0,200)))

        tbar = QWidget(); tbar.setFixedHeight(44)
        tbar.setStyleSheet("background:transparent;")
        tbl = QHBoxLayout(tbar); tbl.setContentsMargins(16,0,12,0)
        ttl = QLabel("Settings")
        ttl.setFont(QFont("Segoe UI",11,QFont.Weight.DemiBold))
        ttl.setStyleSheet(f"color:{_theme.TEXT_PRIMARY.name()};")
        tbl.addWidget(ttl); tbl.addStretch()
        cb2 = TitleBarButton(COLOR_BTN_CLOSE,"x"); cb2.clicked.connect(self.reject)
        tbl.addWidget(cb2)
        tbar.mousePressEvent = lambda ev: (
            setattr(self,"_drag_pos",ev.globalPosition().toPoint())
            if ev.button()==Qt.LeftButton else None)
        tbar.mouseMoveEvent = self._drag_move
        tbar.mouseReleaseEvent = lambda ev: setattr(self,"_drag_pos",None)
        cl.addWidget(tbar)
        self._sep = QFrame(); self._sep.setFrameShape(QFrame.HLine)
        self._sep.setStyleSheet(f"color:rgba({_theme.GLOW.red()},{_theme.GLOW.green()},{_theme.GLOW.blue()},55);"); cl.addWidget(self._sep)

        self._tabs = QTabWidget()
        self._tabs.setStyleSheet(f"""
            QTabWidget::pane {{ border:none; background:transparent; }}
            QTabBar::tab {{
                background:rgba({_theme.TILE_BG_BASE.red()},{_theme.TILE_BG_BASE.green()},{_theme.TILE_BG_BASE.blue()},180);
                color:{_theme.TEXT_SECONDARY.name()};
                padding:7px 18px; border:none; border-radius:4px 4px 0 0;
                font-family:'Segoe UI'; font-size:9pt;
            }}
            QTabBar::tab:selected {{
                background:rgba({_theme.TILE_BG_HOVER.red()},{_theme.TILE_BG_HOVER.green()},{_theme.TILE_BG_HOVER.blue()},200);
                color:{_theme.GLOW.name()};
                border-bottom:2px solid {_theme.GLOW.name()};
            }}
            QTabBar::tab:hover {{ background:rgba({_theme.TILE_BG_HOVER.red()},{_theme.TILE_BG_HOVER.green()},{_theme.TILE_BG_HOVER.blue()},160); }}
        """)
        self._tabs.addTab(self._build_general_tab(), "General")
        self._tabs.addTab(self._build_theme_tab(), "Appearance")
        self._tabs.addTab(self._build_hotkeys_tab(), "Hotkeys")
        self._tabs.addTab(self._build_archived_tab(), "Archived Nodes")
        cl.addWidget(self._tabs)

        brow = QHBoxLayout(); brow.setContentsMargins(16,8,16,14); brow.addStretch()
        self._ok_btn = QPushButton("Apply and Close")
        self._ok_btn.setCursor(Qt.PointingHandCursor)
        self._ok_btn.setFont(QFont("Segoe UI",9,QFont.Weight.DemiBold))
        t = _theme
        g = t.GLOW
        dr, dg, db = int(g.red()*0.65), int(g.green()*0.65), int(g.blue()*0.65)
        self._ok_btn.setStyleSheet(f"""
            QPushButton {{
                background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 rgba({g.red()},{g.green()},{g.blue()},195),
                    stop:1 rgba({dr},{dg},{db},195));
                border:none; border-radius:6px;
                color:{t.TEXT_PRIMARY.name()}; padding:6px 18px;
            }}
            QPushButton:hover {{
                background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 rgba({g.red()},{g.green()},{g.blue()},240),
                    stop:1 rgba({dr},{dg},{db},240));
            }}
        """)
        self._ok_btn.clicked.connect(self._apply); brow.addWidget(self._ok_btn); cl.addLayout(brow)
        outer.addWidget(self._card)

    def _build_general_tab(self):
        t = _theme
        g = t.GLOW
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; }}
            QScrollBar:vertical {{
                background: rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},120);
                width: 6px; border-radius: 3px;
            }}
            QScrollBar::handle:vertical {{
                background: rgba({g.red()},{g.green()},{g.blue()},110);
                border-radius: 3px; min-height: 20px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: rgba({g.red()},{g.green()},{g.blue()},180);
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
        """)
        w = QWidget(); w.setStyleSheet("background:transparent;")
        scroll.setWidget(w)
        layout = QVBoxLayout(w); layout.setContentsMargins(16,16,16,16); layout.setSpacing(14)

        # ── Save File Location ───────────────────────────────────────────
        layout.addWidget(self._sec("Save File Location"))
        # Show the actual resolved path so the user always knows where data lives
        active_path_lbl = QLabel(f"Active data folder:  {CONFIG_DIR}")
        active_path_lbl.setFont(FONT_SMALL)
        active_path_lbl.setWordWrap(True)
        active_path_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        active_path_lbl.setStyleSheet(
            f"color:{t.GLOW.name()}; background:rgba({t.TILE_BG_BASE.red()},"
            f"{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},120); "
            f"border-radius:4px; padding:4px 8px;")
        layout.addWidget(active_path_lbl)
        pr = QHBoxLayout()
        self._save_path_edit = QLineEdit()
        self._save_path_edit.setPlaceholderText("Override (leave blank to use default above)")
        self._save_path_edit.setText(_settings_store.value("save_path",""))
        self._save_path_edit.setStyleSheet(f"""
            QLineEdit {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
                border:1px solid rgba({g.red()},{g.green()},{g.blue()},70);
                border-radius:5px; color:{t.TEXT_PRIMARY.name()};
                padding:5px 9px; font-family:'Segoe UI'; font-size:9pt;
            }}
            QLineEdit:focus {{ border:1px solid rgba({g.red()},{g.green()},{g.blue()},150); }}
        """)
        self._browse_btn = QPushButton("Browse")
        self._browse_btn.setCursor(Qt.PointingHandCursor)
        self._browse_btn.setFont(QFont("Segoe UI",8)); self._browse_btn.setFixedHeight(26)
        self._browse_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},180);
                border:1px solid rgba({g.red()},{g.green()},{g.blue()},70);
                border-radius:5px; color:{t.TEXT_SECONDARY.name()}; padding:0 10px;
            }}
            QPushButton:hover {{ background:rgba({t.TILE_BG_HOVER.red()},{t.TILE_BG_HOVER.green()},{t.TILE_BG_HOVER.blue()},200); color:{g.name()}; }}
        """)
        self._browse_btn.clicked.connect(self._browse_save_path)
        pr.addWidget(self._save_path_edit); pr.addWidget(self._browse_btn)
        layout.addLayout(pr)
        save_note = QLabel("Changes to the save path take effect after restarting.")
        save_note.setFont(FONT_SMALL); save_note.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        save_note.setWordWrap(True)
        layout.addWidget(save_note)

        # ── Startup ──────────────────────────────────────────────────────
        layout.addWidget(self._sec("Startup"))
        self._chk_show_tips = self._make_checkbox(
            "Show a tip at startup",
            _settings_store.value("show_tips", "true") == "true")
        layout.addWidget(self._chk_show_tips)
        self._chk_auto_launch_time_tracker = self._make_checkbox(
            "Auto Launch Time Tracker",
            _settings_store.value("auto_launch_time_tracker", "false") == "true")
        layout.addWidget(self._chk_auto_launch_time_tracker)

        # ── Behavior ─────────────────────────────────────────────────────
        layout.addWidget(self._sec("Behavior"))
        self._chk_confirm_delete = self._make_checkbox(
            "Confirm before deleting a node",
            _settings_store.value("confirm_delete", "false") == "true")
        self._chk_confirm_archive = self._make_checkbox(
            "Confirm before archiving a node",
            _settings_store.value("confirm_archive", "false") == "true")
        self._chk_global_auto_launch = self._make_checkbox(
            "Run auto-launch nodes on startup",
            _settings_store.value("global_auto_launch", "true") == "true")
        layout.addWidget(self._chk_confirm_delete)
        layout.addWidget(self._chk_confirm_archive)
        layout.addWidget(self._chk_global_auto_launch)
        self._chk_single_click = self._make_checkbox(
            "Single-click to launch nodes  (uncheck for double-click)",
            _settings_store.value("launch_on_single_click", "true") == "true")
        layout.addWidget(self._chk_single_click)
        behavior_note = QLabel(
            "Auto-launch nodes open automatically each time Command Center starts.")
        behavior_note.setFont(FONT_SMALL)
        behavior_note.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        behavior_note.setWordWrap(True)
        layout.addWidget(behavior_note)

        # \u2500\u2500 System \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
        layout.addWidget(self._sec("System"))
        self._chk_startup = self._make_checkbox(
            "Launch Command Center on Windows startup",
            _get_windows_startup())
        layout.addWidget(self._chk_startup)
        startup_note = QLabel(
            "Adds Command Center to the Windows Run registry key for your user account.")
        startup_note.setFont(FONT_SMALL)
        startup_note.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        startup_note.setWordWrap(True)
        layout.addWidget(startup_note)

        # ── Scripting ─────────────────────────────────────────────────────────
        layout.addWidget(self._sec("Scripting"))
        self._chk_ps1_bypass = self._make_checkbox(
            "Bypass Script Execution Policy for PowerShell scripts",
            _settings_store.value("ps1_bypass_execution_policy", "true") == "true")
        layout.addWidget(self._chk_ps1_bypass)
        ps1_note = QLabel(
            "When enabled, .ps1 nodes are run with -ExecutionPolicy Bypass so they always execute regardless of system policy.")
        ps1_note.setFont(FONT_SMALL)
        ps1_note.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        ps1_note.setWordWrap(True)
        layout.addWidget(ps1_note)

        # ── Clipboard Manager ─────────────────────────────────────────────
        layout.addWidget(self._sec("Clipboard Manager"))
        self._chk_clipboard_enabled = self._make_checkbox(
            "Enable Clipboard Manager",
            _settings_store.value("clipboard_manager_enabled", "true") == "true")
        layout.addWidget(self._chk_clipboard_enabled)
        clipboard_note = QLabel(
            "When enabled, Command Center records clipboard history and the "
            "Ctrl+` global hotkey opens the Clipboard Manager from any window. "
            "Changes take effect immediately.")
        clipboard_note.setFont(FONT_SMALL)
        clipboard_note.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        clipboard_note.setWordWrap(True)
        layout.addWidget(clipboard_note)

        # ── Hotkeys ──────────────────────────────────────────────────────────
        layout.addWidget(self._sec("Hotkeys"))
        self._chk_disable_hotkeys = self._make_checkbox(
            "Disable all keyboard shortcuts",
            _settings_store.value("disable_hotkeys", "false") == "true")
        layout.addWidget(self._chk_disable_hotkeys)
        hotkeys_note = QLabel("Disables all app-wide shortcuts (Ctrl+N, Ctrl+F, F1, etc.)."
                              "  Note: shortcuts inside the Notebook editor are not affected.")
        hotkeys_note.setFont(FONT_SMALL)
        hotkeys_note.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        hotkeys_note.setWordWrap(True)
        layout.addWidget(hotkeys_note)

        layout.addStretch()
        return scroll

    def _build_theme_tab(self):
        t = _theme
        # ── Outer scroll area so content never gets clipped ──────────────
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; }}
            QScrollBar:vertical {{
                background: rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},120);
                width: 6px; border-radius: 3px;
            }}
            QScrollBar::handle:vertical {{
                background: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},110);
                border-radius: 3px; min-height: 20px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},180);
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
        """)
        w = QWidget(); w.setStyleSheet("background:transparent;")
        scroll.setWidget(w)
        layout = QVBoxLayout(w); layout.setContentsMargins(16,16,16,16); layout.setSpacing(16)

        # ── Preset themes ───────────────────────────────────────────────
        layout.addWidget(self._sec("Theme Preset"))
        self._theme_btns: list[QPushButton] = []
        self._swatch_btns: dict[str, QPushButton] = {}
        themes_grid = QGridLayout(); themes_grid.setSpacing(8)
        preset_names = [n for n in _BUILTIN_THEMES if n != "Custom"]
        for idx, name in enumerate(preset_names):
            btn = QPushButton(name)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setFont(QFont("Segoe UI", 8, QFont.Weight.DemiBold))
            btn.setFixedHeight(30)
            btn.setCheckable(True)
            btn.setChecked(name == t.theme_name)
            btn.setProperty("theme_name", name)
            btn.clicked.connect(self._on_preset_clicked)
            self._theme_btns.append(btn)
            row, col = divmod(idx, 4)
            themes_grid.addWidget(btn, row, col)
        layout.addLayout(themes_grid)
        self._update_theme_btn_styles()

        # ── Custom colors ────────────────────────────────────────────────
        layout.addWidget(self._sec("Custom Colors  (override any preset)"))
        COLOR_LABELS = [
            ("glow",           "Accent / Glow"),
            ("text_primary",   "Primary Text"),
            ("text_secondary", "Secondary Text"),
            ("bg_mid",         "Canvas Background"),
            ("bg_dark",        "Background Dark"),
            ("tile_bg_base",   "Tile Background"),
            ("tile_bg_hover",  "Tile Hover"),
            ("accent_blue",    "Accent Blue"),
            ("accent_teal",    "Accent Teal"),
            ("accent_amber",   "Accent Amber"),
        ]
        grid = QGridLayout(); grid.setSpacing(8)
        for idx, (key, label) in enumerate(COLOR_LABELS):
            lbl = QLabel(label); lbl.setFont(FONT_SMALL)
            lbl.setStyleSheet(f"color:{t.TEXT_SECONDARY.name()};")
            swatch_btn = QPushButton()
            swatch_btn.setFixedSize(42, 24)
            swatch_btn.setCursor(Qt.PointingHandCursor)
            swatch_btn.setToolTip(f"Click to change: {label}")
            existing_tuple = t.color_tuple(key)
            c = QColor(*existing_tuple[:3])
            swatch_btn.setStyleSheet(
                f"background:{c.name()}; border:1px solid rgba(255,255,255,40); border-radius:4px;")
            swatch_btn.setProperty("color_key", key)
            swatch_btn.clicked.connect(self._on_swatch_click)
            self._swatch_btns[key] = swatch_btn
            row, col_base = divmod(idx, 3)
            grid.addWidget(lbl,       row, col_base * 3)
            grid.addWidget(swatch_btn, row, col_base * 3 + 1)
        layout.addLayout(grid)

        self._reset_btn = QPushButton("Reset Custom Colors")
        self._reset_btn.setCursor(Qt.PointingHandCursor)
        self._reset_btn.setFont(FONT_SMALL); self._reset_btn.setFixedHeight(26)
        self._reset_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},180);
                border:1px solid rgba({t.ACCENT_RED.red()},{t.ACCENT_RED.green()},{t.ACCENT_RED.blue()},90);
                border-radius:5px; color:{t.TEXT_SECONDARY.name()}; padding:0 12px;
            }}
            QPushButton:hover {{ color:{t.ACCENT_RED.name()}; border-color:rgba({t.ACCENT_RED.red()},{t.ACCENT_RED.green()},{t.ACCENT_RED.blue()},160); }}
        """)
        self._reset_btn.clicked.connect(self._reset_custom_colors)
        layout.addWidget(self._reset_btn)

        layout.addWidget(self._sec("Preview"))
        preview_note = QLabel("Theme changes apply instantly.  Restart is not required.")
        preview_note.setFont(FONT_SMALL)
        preview_note.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        layout.addWidget(preview_note)

        # \u2500\u2500 Canvas background \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
        layout.addWidget(self._sec("Canvas Background"))
        canvas_bg_row = QHBoxLayout(); canvas_bg_row.setSpacing(6)
        self._canvas_bg_btns: dict = {}
        current_bg = _settings_store.value("canvas_bg_style", "dots")
        for bg_key, bg_label in [
            ("solid",    "Solid"),
            ("dots",     "Dots"),
            ("grid",     "Grid"),
            ("noise",    "Noise"),
            ("gradient", "Gradient"),
            ("hexagons", "Hexagons"),
            ("web",      "Web"),
            ("image",    "Custom Image"),
        ]:
            btn = QPushButton(bg_label)
            btn.setCheckable(True)
            btn.setChecked(bg_key == current_bg)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setFont(QFont("Segoe UI", 8, QFont.Weight.DemiBold))
            btn.setFixedHeight(28)
            btn.setProperty("bg_key", bg_key)
            btn.clicked.connect(self._on_canvas_bg_clicked)
            self._canvas_bg_btns[bg_key] = btn
            canvas_bg_row.addWidget(btn)
        canvas_bg_row.addStretch()
        layout.addLayout(canvas_bg_row)
        self._update_canvas_bg_btn_styles()

        # ── Canvas background image controls (visible only when "image" is selected) ──
        self._bg_image_controls = QWidget()
        self._bg_image_controls.setVisible(current_bg == "image")
        bic_layout = QVBoxLayout(self._bg_image_controls)
        bic_layout.setContentsMargins(0, 4, 0, 0)
        bic_layout.setSpacing(6)

        # File path row
        bir = QHBoxLayout(); bir.setSpacing(6)
        saved_bg_image = _settings_store.value("canvas_bg_image", "")
        bg_display_name = Path(saved_bg_image).name if saved_bg_image and os.path.isfile(saved_bg_image) else "No image selected"
        self._bg_image_path_lbl = QLabel(bg_display_name)
        self._bg_image_path_lbl.setFont(FONT_SMALL)
        self._bg_image_path_lbl.setStyleSheet(f"color:{t.TEXT_DIM.name()}; font-style:italic;")
        self._bg_image_path_lbl.setMaximumWidth(260)
        self._bg_image_path_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        bir.addWidget(self._bg_image_path_lbl)
        self._bg_image_browse_btn = QPushButton("Browse…")
        self._bg_image_browse_btn.setFixedSize(72, 24)
        self._bg_image_browse_btn.setCursor(Qt.PointingHandCursor)
        self._bg_image_browse_btn.setFont(FONT_SMALL)
        _bib_g = t.GLOW
        self._bg_image_browse_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},180);
                border:1px solid rgba({_bib_g.red()},{_bib_g.green()},{_bib_g.blue()},70);
                border-radius:5px; color:{t.TEXT_SECONDARY.name()}; padding:0 8px;
            }}
            QPushButton:hover {{
                background:rgba({t.TILE_BG_HOVER.red()},{t.TILE_BG_HOVER.green()},{t.TILE_BG_HOVER.blue()},200);
                color:{_bib_g.name()};
            }}
        """)
        self._bg_image_browse_btn.clicked.connect(self._on_canvas_bg_image_browse)
        bir.addWidget(self._bg_image_browse_btn)
        bir.addStretch()
        bic_layout.addLayout(bir)

        # Opacity slider row
        opr = QHBoxLayout(); opr.setSpacing(10)
        op_lbl = QLabel("Image Opacity:")
        op_lbl.setFont(FONT_SMALL)
        op_lbl.setStyleSheet(f"color:{t.TEXT_SECONDARY.name()};")
        opr.addWidget(op_lbl)
        self._bg_opacity_slider = QSlider(Qt.Horizontal)
        self._bg_opacity_slider.setRange(5, 100)
        saved_opacity = int(_settings_store.value("canvas_bg_image_opacity", 80))
        self._bg_opacity_slider.setValue(saved_opacity)
        self._bg_opacity_slider.setFixedWidth(180)
        _g = t.GLOW
        self._bg_opacity_slider.setStyleSheet(f"""
            QSlider::groove:horizontal {{
                height: 4px;
                background: rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
                border-radius: 2px;
            }}
            QSlider::sub-page:horizontal {{
                background: rgba({_g.red()},{_g.green()},{_g.blue()},180);
                border-radius: 2px;
            }}
            QSlider::handle:horizontal {{
                width: 14px; height: 14px;
                margin: -5px 0;
                border-radius: 7px;
                background: {_g.name()};
            }}
            QSlider::handle:horizontal:hover {{
                background: rgba({_g.red()},{_g.green()},{_g.blue()},220);
            }}
        """)
        opr.addWidget(self._bg_opacity_slider)
        self._bg_opacity_val_lbl = QLabel(f"{saved_opacity}%")
        self._bg_opacity_val_lbl.setFont(FONT_SMALL)
        self._bg_opacity_val_lbl.setFixedWidth(36)
        self._bg_opacity_val_lbl.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        opr.addWidget(self._bg_opacity_val_lbl)
        opr.addStretch()
        bic_layout.addLayout(opr)
        self._bg_opacity_slider.valueChanged.connect(self._on_canvas_bg_opacity_changed)
        layout.addWidget(self._bg_image_controls)

        layout.addWidget(self._sec("Animation"))
        self._chk_titlebar_anim = self._make_checkbox(
            "Disable title bar animation",
            _settings_store.value("disable_titlebar_anim", "false") == "true")
        self._chk_skip_splash = self._make_checkbox(
            "Skip startup animation",
            _settings_store.value("skip_startup_anim", "false") == "true")
        layout.addWidget(self._chk_titlebar_anim)
        layout.addWidget(self._chk_skip_splash)

        # \u2500\u2500 Custom cursor \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
        # ── UI Brightness ─────────────────────────────────────────────────────
        layout.addWidget(self._sec("UI Brightness"))
        brightness_hint = QLabel("Scales the brightness of the entire UI. 100% is the balanced default; raise it if the theme feels too dark.")
        brightness_hint.setFont(FONT_SMALL)
        brightness_hint.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        brightness_hint.setWordWrap(True)
        layout.addWidget(brightness_hint)

        brightness_row = QHBoxLayout(); brightness_row.setSpacing(10)
        br_lbl = QLabel("Brightness:")
        br_lbl.setFont(FONT_SMALL)
        br_lbl.setStyleSheet(f"color:{t.TEXT_SECONDARY.name()};")
        brightness_row.addWidget(br_lbl)

        self._brightness_slider = QSlider(Qt.Horizontal)
        self._brightness_slider.setRange(30, 300)
        saved_brightness = int(_settings_store.value("ui_brightness", 100))
        self._brightness_slider.setValue(saved_brightness)
        self._brightness_slider.setFixedWidth(200)
        self._brightness_slider.setStyleSheet(f"""
            QSlider::groove:horizontal {{
                height: 4px;
                background: rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
                border-radius: 2px;
            }}
            QSlider::sub-page:horizontal {{
                background: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},180);
                border-radius: 2px;
            }}
            QSlider::handle:horizontal {{
                width: 14px; height: 14px;
                margin: -5px 0;
                border-radius: 7px;
                background: {t.GLOW.name()};
            }}
            QSlider::handle:horizontal:hover {{
                background: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},220);
            }}
        """)
        brightness_row.addWidget(self._brightness_slider)

        self._brightness_val_lbl = QLabel(f"{saved_brightness}%")
        self._brightness_val_lbl.setFont(FONT_SMALL)
        self._brightness_val_lbl.setFixedWidth(40)
        self._brightness_val_lbl.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        brightness_row.addWidget(self._brightness_val_lbl)
        brightness_row.addStretch()
        layout.addLayout(brightness_row)

        self._brightness_slider.valueChanged.connect(self._on_brightness_changed)

        layout.addWidget(self._sec("Custom Cursor"))
        cursor_hint = QLabel("Changes apply instantly. Your preference is saved automatically.")
        cursor_hint.setFont(FONT_SMALL)
        cursor_hint.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        cursor_hint.setWordWrap(True)
        layout.addWidget(cursor_hint)

        cursor_grid = QGridLayout()
        cursor_grid.setSpacing(8)
        cursor_grid.setContentsMargins(0, 4, 0, 0)
        self._cursor_btns: dict[str, QFrame] = {}
        script_dir = os.path.dirname(os.path.abspath(__file__))

        for idx, (cname, cfile) in enumerate(_CURSOR_OPTIONS):
            frame = QFrame()
            frame.setFixedSize(74, 86)
            frame.setCursor(Qt.PointingHandCursor)
            frame.setProperty("cursor_name", cname)

            fl = QVBoxLayout(frame)
            fl.setContentsMargins(4, 7, 4, 5)
            fl.setSpacing(3)

            img_lbl = QLabel()
            img_lbl.setAlignment(Qt.AlignCenter)
            img_lbl.setFixedHeight(46)
            img_lbl.setAttribute(Qt.WA_TransparentForMouseEvents)

            if cfile:
                img_path = os.path.join(script_dir, cfile)
                if os.path.isfile(img_path):
                    raw_px = QPixmap(img_path)
                    if not raw_px.isNull():
                        preview_px = raw_px.scaled(
                            42, 42, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                        img_lbl.setPixmap(preview_px)
                    else:
                        img_lbl.setText("?")
                        img_lbl.setFont(QFont("Segoe UI", 14))
                else:
                    img_lbl.setText("?")
                    img_lbl.setFont(QFont("Segoe UI", 14))

            name_lbl = QLabel(cname)
            name_lbl.setAlignment(Qt.AlignCenter)
            name_lbl.setFont(QFont("Segoe UI", 7))
            name_lbl.setAttribute(Qt.WA_TransparentForMouseEvents)

            fl.addWidget(img_lbl)
            fl.addWidget(name_lbl)

            # Use a closure to capture cname
            frame.mousePressEvent = lambda ev, n=cname: self._on_cursor_select(n)
            self._cursor_btns[cname] = frame
            row, col = divmod(idx, 4)
            cursor_grid.addWidget(frame, row, col)

        # ── "Custom" cursor slot ─────────────────────────────────────────────
        custom_frame = QFrame()
        custom_frame.setFixedSize(74, 86)
        custom_frame.setCursor(Qt.PointingHandCursor)
        custom_frame.setProperty("cursor_name", "Custom")
        cfl = QVBoxLayout(custom_frame)
        cfl.setContentsMargins(4, 7, 4, 5); cfl.setSpacing(3)
        self._custom_cursor_img_lbl = QLabel()
        self._custom_cursor_img_lbl.setAlignment(Qt.AlignCenter)
        self._custom_cursor_img_lbl.setFixedHeight(46)
        self._custom_cursor_img_lbl.setAttribute(Qt.WA_TransparentForMouseEvents)
        # Try to load existing custom cursor preview
        _saved_custom_path = _settings_store.value("cursor_custom_path", "")
        if _saved_custom_path and os.path.isfile(_saved_custom_path):
            _cpx = QPixmap(_saved_custom_path)
            if not _cpx.isNull():
                self._custom_cursor_img_lbl.setPixmap(
                    _cpx.scaled(42, 42, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            else:
                self._custom_cursor_img_lbl.setText("?")
                self._custom_cursor_img_lbl.setFont(QFont("Segoe UI", 14))
        else:
            self._custom_cursor_img_lbl.setText("+")
            self._custom_cursor_img_lbl.setFont(QFont("Segoe UI", 16, QFont.Weight.Light))
        custom_name_lbl = QLabel("Custom")
        custom_name_lbl.setAlignment(Qt.AlignCenter)
        custom_name_lbl.setFont(QFont("Segoe UI", 7))
        custom_name_lbl.setAttribute(Qt.WA_TransparentForMouseEvents)
        cfl.addWidget(self._custom_cursor_img_lbl)
        cfl.addWidget(custom_name_lbl)
        custom_frame.mousePressEvent = lambda ev: self._on_custom_cursor_browse()
        self._cursor_btns["Custom"] = custom_frame
        # Place after the last built-in cursor (idx len(_CURSOR_OPTIONS) = 8)
        _custom_row, _custom_col = divmod(len(_CURSOR_OPTIONS), 4)
        cursor_grid.addWidget(custom_frame, _custom_row, _custom_col)

        layout.addLayout(cursor_grid)
        self._update_cursor_btn_styles()

        # ── Custom cursor info row ───────────────────────────────────────────
        cursor_info_row = QHBoxLayout(); cursor_info_row.setSpacing(6)
        _saved_cursor_name = Path(_saved_custom_path).name if _saved_custom_path and os.path.isfile(_saved_custom_path) else ""
        self._custom_cursor_path_lbl = QLabel(
            f"Custom: {_saved_cursor_name}" if _saved_cursor_name else "Custom: click the Custom tile to upload an image")
        self._custom_cursor_path_lbl.setFont(FONT_SMALL)
        self._custom_cursor_path_lbl.setWordWrap(True)
        self._custom_cursor_path_lbl.setStyleSheet(f"color:{t.TEXT_DIM.name()}; font-style:italic;")
        cursor_info_row.addWidget(self._custom_cursor_path_lbl)
        cursor_info_row.addStretch()
        layout.addLayout(cursor_info_row)

        cursor_hotspot_note = QLabel(
            "For custom cursors, place the click point (hotspot) at the top-left corner of the image.")
        cursor_hotspot_note.setFont(FONT_SMALL)
        cursor_hotspot_note.setWordWrap(True)
        cursor_hotspot_note.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        layout.addWidget(cursor_hotspot_note)

        # ── Cursor size slider ────────────────────────────────────────────
        size_row = QHBoxLayout(); size_row.setSpacing(10)
        size_lbl = QLabel("Cursor Size:")
        size_lbl.setFont(FONT_SMALL)
        size_lbl.setStyleSheet(f"color:{t.TEXT_SECONDARY.name()};")
        size_row.addWidget(size_lbl)

        self._cursor_size_slider = QSlider(Qt.Horizontal)
        self._cursor_size_slider.setRange(16, 128)
        saved_sz = int(_settings_store.value("cursor_size", 40))
        self._cursor_size_slider.setValue(saved_sz)
        self._cursor_size_slider.setFixedWidth(180)
        g = t.GLOW
        self._cursor_size_slider.setStyleSheet(f"""
            QSlider::groove:horizontal {{
                height: 4px;
                background: rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
                border-radius: 2px;
            }}
            QSlider::sub-page:horizontal {{
                background: rgba({g.red()},{g.green()},{g.blue()},180);
                border-radius: 2px;
            }}
            QSlider::handle:horizontal {{
                width: 14px; height: 14px;
                margin: -5px 0;
                border-radius: 7px;
                background: {g.name()};
            }}
            QSlider::handle:horizontal:hover {{
                background: rgba({g.red()},{g.green()},{g.blue()},220);
            }}
        """)
        size_row.addWidget(self._cursor_size_slider)

        self._cursor_size_val_lbl = QLabel(f"{saved_sz} px")
        self._cursor_size_val_lbl.setFont(FONT_SMALL)
        self._cursor_size_val_lbl.setFixedWidth(42)
        self._cursor_size_val_lbl.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        size_row.addWidget(self._cursor_size_val_lbl)
        size_row.addStretch()
        layout.addLayout(size_row)

        self._cursor_size_slider.valueChanged.connect(self._on_cursor_size_changed)

        layout.addStretch()
        return scroll

    def _on_canvas_bg_clicked(self):
        btn = self.sender()
        key = btn.property("bg_key")
        _settings_store.setValue("canvas_bg_style", key)
        _settings_store.sync()
        for k, b in self._canvas_bg_btns.items():
            b.setChecked(k == key)
        self._update_canvas_bg_btn_styles()
        # Show/hide the custom image controls depending on selection
        if hasattr(self, "_bg_image_controls"):
            self._bg_image_controls.setVisible(key == "image")
        self.settings_changed.emit()

    def _update_canvas_bg_btn_styles(self):
        t = _theme; g = t.GLOW
        current_bg = _settings_store.value("canvas_bg_style", "dots")
        dr, dg, db = int(g.red()*0.65), int(g.green()*0.65), int(g.blue()*0.65)
        for key, btn in self._canvas_bg_btns.items():
            if key == current_bg:
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                            stop:0 rgba({g.red()},{g.green()},{g.blue()},195),
                            stop:1 rgba({dr},{dg},{db},195));
                        border:none; border-radius:5px;
                        color:{t.TEXT_PRIMARY.name()}; padding:0 10px;
                    }}
                """)
            else:
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},160);
                        border:1px solid rgba({g.red()},{g.green()},{g.blue()},50);
                        border-radius:5px; color:{t.TEXT_SECONDARY.name()}; padding:0 10px;
                    }}
                    QPushButton:hover {{
                        background:rgba({t.TILE_BG_HOVER.red()},{t.TILE_BG_HOVER.green()},{t.TILE_BG_HOVER.blue()},180);
                        color:{t.TEXT_PRIMARY.name()};
                    }}
                """)

    def _on_preset_clicked(self):
        btn = self.sender()
        name = btn.property("theme_name")
        _theme.apply_theme(name)
        self._update_theme_btn_styles()
        self._refresh_swatches()
        self._restyle_all()
        self.settings_changed.emit()

    def _update_theme_btn_styles(self):
        t = _theme
        for btn in self._theme_btns:
            is_active = btn.property("theme_name") == t.theme_name
            if is_active:
                g = t.GLOW
                dr, dg, db = int(g.red()*0.65), int(g.green()*0.65), int(g.blue()*0.65)
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                            stop:0 rgba({g.red()},{g.green()},{g.blue()},195),
                            stop:1 rgba({dr},{dg},{db},195));
                        border:none; border-radius:5px; color:{t.TEXT_PRIMARY.name()}; padding:0 10px;
                    }}
                """)
            else:
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},160);
                        border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);
                        border-radius:5px; color:{t.TEXT_SECONDARY.name()}; padding:0 10px;
                    }}
                    QPushButton:hover {{
                        background:rgba({t.TILE_BG_HOVER.red()},{t.TILE_BG_HOVER.green()},{t.TILE_BG_HOVER.blue()},180);
                        color:{t.TEXT_PRIMARY.name()};
                    }}
                """)

    def _on_swatch_click(self):
        btn = self.sender()
        if not isinstance(btn, QPushButton):
            return
        key   = btn.property("color_key")
        label = btn.toolTip().replace("Click to change: ", "")
        existing = QColor(*_theme.color_tuple(key)[:3])
        dlg = _CCColorPickerDialog(existing, label, self)
        if dlg.exec() == QDialog.Accepted:
            chosen = dlg.selected_color()
            _theme.set_custom_color(key, chosen)
            btn.setStyleSheet(
                f"background:{chosen.name()}; border:1px solid rgba(255,255,255,40); border-radius:4px;")
            self._update_theme_btn_styles()
            self._restyle_all()
            self.settings_changed.emit()

    def _reset_custom_colors(self):
        _settings_store.remove("theme_custom")
        _settings_store.sync()
        _theme._load()
        _theme._notify()
        self._refresh_swatches()
        self._update_theme_btn_styles()
        self._restyle_all()
        self.settings_changed.emit()

    def _refresh_swatches(self):
        for key, btn in self._swatch_btns.items():
            c = QColor(*_theme.color_tuple(key)[:3])
            btn.setStyleSheet(
                f"background:{c.name()}; border:1px solid rgba(255,255,255,40); border-radius:4px;")

    def _on_cursor_select(self, name: str):
        """Apply a cursor immediately and persist the choice."""
        _settings_store.setValue("cursor_name", name)
        _settings_store.sync()
        _apply_app_cursor()
        self._update_cursor_btn_styles()

    def _on_cursor_size_changed(self, value: int):
        """Update the size label and re-apply the cursor live."""
        self._cursor_size_val_lbl.setText(f"{value} px")
        _settings_store.setValue("cursor_size", value)
        _settings_store.sync()
        _apply_app_cursor()

    def _on_custom_cursor_browse(self):
        """Open file dialog to pick a custom cursor image, copy to config dir, and apply."""
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Cursor Image", "",
            "Image Files (*.png *.jpg *.jpeg *.bmp *.ico *.cur)")
        if not path or not os.path.isfile(path):
            return
        ext = Path(path).suffix.lower()
        dest = CONFIG_DIR / f"custom_cursor{ext}"
        try:
            shutil.copy2(path, dest)
        except OSError as exc:
            print(f"[CommandCenter] Custom cursor copy error: {exc}", file=sys.stderr)
            QMessageBox.warning(self, "Error",
                                f"Could not save custom cursor file:\n{exc}")
            return
        dest_str = str(dest)
        _settings_store.setValue("cursor_custom_path", dest_str)
        _settings_store.sync()
        # Update the preview image label in the Custom tile
        if hasattr(self, "_custom_cursor_img_lbl"):
            px = QPixmap(dest_str)
            if not px.isNull():
                self._custom_cursor_img_lbl.setPixmap(
                    px.scaled(42, 42, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            else:
                self._custom_cursor_img_lbl.setText("?")
        # Update the info label below the cursor grid
        if hasattr(self, "_custom_cursor_path_lbl"):
            self._custom_cursor_path_lbl.setText(f"Custom: {Path(path).name}")
        # Select "Custom" and apply immediately
        _settings_store.setValue("cursor_name", "Custom")
        _settings_store.sync()
        _apply_app_cursor()
        self._update_cursor_btn_styles()

    def _on_canvas_bg_image_browse(self):
        """Open file dialog to pick a background image, copy to config dir, and apply."""
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Background Image", "",
            "Image Files (*.png *.jpg *.jpeg *.bmp *.gif *.webp)")
        if not path or not os.path.isfile(path):
            return
        ext = Path(path).suffix.lower()
        dest = CONFIG_DIR / f"canvas_bg{ext}"
        try:
            shutil.copy2(path, dest)
        except OSError as exc:
            print(f"[CommandCenter] Canvas bg image copy error: {exc}", file=sys.stderr)
            QMessageBox.warning(self, "Error",
                                f"Could not save background image file:\n{exc}")
            return
        _settings_store.setValue("canvas_bg_image", str(dest))
        _settings_store.sync()
        # Update the path label
        if hasattr(self, "_bg_image_path_lbl"):
            self._bg_image_path_lbl.setText(Path(path).name)
        self.settings_changed.emit()

    def _on_canvas_bg_opacity_changed(self, value: int):
        """Save opacity setting and refresh the canvas immediately."""
        self._bg_opacity_val_lbl.setText(f"{value}%")
        _settings_store.setValue("canvas_bg_image_opacity", value)
        _settings_store.sync()
        self.settings_changed.emit()

    def _on_brightness_changed(self, value: int):
        self._brightness_val_lbl.setText(f"{value}%")
        _theme.set_brightness(value)
        self.settings_changed.emit()

    def _update_cursor_btn_styles(self):
        t = _theme
        g = t.GLOW
        active = _settings_store.value("cursor_name", "Standard")
        for name, frame in self._cursor_btns.items():
            if name == active:
                frame.setStyleSheet(f"""
                    QFrame {{
                        background: rgba({g.red()},{g.green()},{g.blue()},45);
                        border: 1px solid {g.name()};
                        border-radius: 6px;
                    }}
                    QLabel {{
                        background: transparent;
                        color: {t.TEXT_PRIMARY.name()};
                    }}
                """)
            else:
                frame.setStyleSheet(f"""
                    QFrame {{
                        background: rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},160);
                        border: 1px solid rgba({g.red()},{g.green()},{g.blue()},35);
                        border-radius: 6px;
                    }}
                    QFrame:hover {{
                        background: rgba({t.TILE_BG_HOVER.red()},{t.TILE_BG_HOVER.green()},{t.TILE_BG_HOVER.blue()},190);
                        border-color: rgba({g.red()},{g.green()},{g.blue()},80);
                    }}
                    QLabel {{
                        background: transparent;
                        color: {t.TEXT_SECONDARY.name()};
                    }}
                """)

    def _build_archived_tab(self):
        t = _theme
        w = QWidget(); w.setStyleSheet("background:transparent;")
        layout = QVBoxLayout(w); layout.setContentsMargins(16,16,16,16); layout.setSpacing(10)
        layout.addWidget(self._sec("Archived Nodes"))

        self._archive_list = QListWidget()
        self._archive_list.setStyleSheet(f"""
            QListWidget {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
                border-radius:6px; color:{t.TEXT_PRIMARY.name()};
                font-family:'Segoe UI'; font-size:9pt; outline:none;
            }}
            QListWidget::item {{ padding:6px 10px; }}
            QListWidget::item:selected {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);
                color:{t.GLOW.name()};
            }}
            QListWidget::item:hover {{
                background:rgba({t.TILE_BG_HOVER.red()},{t.TILE_BG_HOVER.green()},{t.TILE_BG_HOVER.blue()},180);
            }}
        """)
        self._refresh_archive_list(); layout.addWidget(self._archive_list)

        brow = QHBoxLayout()
        self._restore_btn = QPushButton("Restore Selected")
        self._del_btn     = QPushButton("Delete Permanently")
        for btn in (self._restore_btn, self._del_btn):
            btn.setCursor(Qt.PointingHandCursor)
            btn.setFont(QFont("Segoe UI",8)); btn.setFixedHeight(26)
        self._restore_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},180);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:5px;color:{t.TEXT_SECONDARY.name()};padding:0 12px;
            }}
            QPushButton:hover {{ background:rgba({t.TILE_BG_HOVER.red()},{t.TILE_BG_HOVER.green()},{t.TILE_BG_HOVER.blue()},210);color:{t.GLOW.name()}; }}
        """)
        self._del_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({t.ACCENT_RED.red()//2},{t.ACCENT_RED.green()//3},{t.ACCENT_RED.blue()//3},120);
                border:1px solid rgba({t.ACCENT_RED.red()},{t.ACCENT_RED.green()},{t.ACCENT_RED.blue()},80);
                border-radius:5px;color:{t.TEXT_SECONDARY.name()};padding:0 12px;
            }}
            QPushButton:hover {{ background:rgba({t.ACCENT_RED.red()},{t.ACCENT_RED.green()},{t.ACCENT_RED.blue()},150);color:#ffffff; }}
        """)
        self._restore_btn.clicked.connect(self._restore_selected)
        self._del_btn.clicked.connect(self._delete_selected)
        brow.addWidget(self._restore_btn); brow.addWidget(self._del_btn); brow.addStretch()
        layout.addLayout(brow); return w

    # ── Hotkeys tab ────────────────────────────────────────────────────────

    def _build_hotkeys_tab(self):
        t = _theme
        g = t.GLOW
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; }}
            QScrollBar:vertical {{
                background: rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},120);
                width: 6px; border-radius: 3px;
            }}
            QScrollBar::handle:vertical {{
                background: rgba({g.red()},{g.green()},{g.blue()},110);
                border-radius: 3px; min-height: 20px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: rgba({g.red()},{g.green()},{g.blue()},180);
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
        """)
        w = QWidget(); w.setStyleSheet("background:transparent;")
        scroll.setWidget(w)
        layout = QVBoxLayout(w)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        # ── Header row: intro text + Reset All button ─────────────────────
        top_row = QHBoxLayout()
        intro = QLabel(
            "Customize keyboard shortcuts. Use 2–3 keys (e.g. Ctrl+Key) to avoid "
            "accidental activation.  Click ⏺ to record, Escape to cancel.")
        intro.setFont(FONT_SMALL)
        intro.setWordWrap(True)
        intro.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        top_row.addWidget(intro, 1)

        self._hk_reset_all_btn = QPushButton("↺  Reset All to Defaults")
        self._hk_reset_all_btn.setFixedHeight(28)
        self._hk_reset_all_btn.setCursor(Qt.PointingHandCursor)
        self._hk_reset_all_btn.setFont(FONT_SMALL)
        self._hk_reset_all_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},180);
                border:1px solid rgba({t.ACCENT_AMBER.red()},{t.ACCENT_AMBER.green()},{t.ACCENT_AMBER.blue()},90);
                border-radius:5px; color:{t.TEXT_SECONDARY.name()}; padding:0 12px;
            }}
            QPushButton:hover {{
                color:{t.ACCENT_AMBER.name()};
                border-color:rgba({t.ACCENT_AMBER.red()},{t.ACCENT_AMBER.green()},{t.ACCENT_AMBER.blue()},160);
            }}
        """)
        self._hk_reset_all_btn.clicked.connect(self._reset_all_hotkeys)
        top_row.addWidget(self._hk_reset_all_btn)
        layout.addLayout(top_row)

        # Disable-all note
        disable_note = QLabel(
            "To disable all shortcuts at once, use the  \"Disable all keyboard shortcuts\"  "
            "option in the  General  tab.")
        disable_note.setFont(FONT_SMALL)
        disable_note.setWordWrap(True)
        disable_note.setStyleSheet(
            f"color:{t.TEXT_DIM.name()}; "
            f"background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},100);"
            f"border-radius:4px; padding:4px 8px;")
        layout.addWidget(disable_note)

        layout.addWidget(self._sec("Configurable Shortcuts"))

        # ── Column header ─────────────────────────────────────────────────
        hdr = QWidget(); hdr.setStyleSheet("background:transparent;")
        hdr_h = QHBoxLayout(hdr)
        hdr_h.setContentsMargins(4, 0, 4, 0)
        hdr_h.setSpacing(8)
        for col_text, col_width, stretch in [
            ("Action",   180, False),
            ("Shortcut", 130, False),
            ("",          20, True),   # warn icon placeholder
            ("",          20, False),  # conflict icon placeholder
        ]:
            lbl = QLabel(col_text)
            lbl.setFont(QFont("Segoe UI", 8, QFont.Weight.DemiBold))
            lbl.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
            if not stretch:
                lbl.setFixedWidth(col_width) if col_width <= 20 else lbl.setMinimumWidth(col_width)
            hdr_h.addWidget(lbl)
            if stretch:
                hdr_h.addStretch()
        layout.addWidget(hdr)

        sep_top = QFrame(); sep_top.setFrameShape(QFrame.HLine)
        sep_top.setStyleSheet(f"background:rgba({g.red()},{g.green()},{g.blue()},60);")
        sep_top.setFixedHeight(1)
        layout.addWidget(sep_top)

        # ── One row per action ────────────────────────────────────────────
        # Shared pending dict — populated by row widgets, saved in _apply()
        self._pending_hotkeys: dict[str, str] = {}
        self._hotkey_rows: list[_HotkeyRowWidget] = []

        for idx, action_id in enumerate(_CC_DEFAULT_HOTKEYS):
            row = _HotkeyRowWidget(action_id, self._pending_hotkeys)
            row.hotkey_changed.connect(self._on_hotkey_changed)
            self._hotkey_rows.append(row)
            layout.addWidget(row)
            # Thin separator between rows (skip after last)
            if idx < len(_CC_DEFAULT_HOTKEYS) - 1:
                sep = QFrame(); sep.setFrameShape(QFrame.HLine)
                sep.setFixedHeight(1)
                sep.setStyleSheet(f"background:rgba({g.red()},{g.green()},{g.blue()},20);")
                layout.addWidget(sep)

        layout.addStretch()
        return scroll

    def _on_hotkey_changed(self, action_id: str, seq: str):
        """Called whenever a row emits a new sequence.

        - If ``seq`` conflicts with another action, that action's binding is
          cleared and the user is told which action now has no hotkey.
        - If ``seq`` is a single-key binding, a caution notice is shown.
        - Conflict + caution notices can stack (both shown together).
        """
        from PySide6.QtWidgets import QMessageBox
        conflict_messages: list[str] = []

        # ── Conflict resolution ──────────────────────────────────────────
        seq_upper = seq.strip().upper()
        for row in self._hotkey_rows:
            if row._action_id == action_id:
                continue
            other_seq = self._pending_hotkeys.get(
                row._action_id, _get_hotkey(row._action_id))
            if other_seq and other_seq.strip().upper() == seq_upper:
                old_label = _CC_HOTKEY_LABELS.get(row._action_id, row._action_id)
                self._pending_hotkeys[row._action_id] = ""
                row._refresh_display()
                conflict_messages.append(
                    f'"{seq}" was already assigned to "{old_label}".\n'
                    f'That binding has been cleared. Please set a new shortcut for "{old_label}".'
                )

        if conflict_messages:
            box = QMessageBox(self)
            box.setWindowTitle("Hotkey Conflict")
            box.setIcon(QMessageBox.Icon.Warning)
            box.setText("\n\n".join(conflict_messages))
            box.setStandardButtons(QMessageBox.StandardButton.Ok)
            box.exec()

        # ── Single-key caution ───────────────────────────────────────────
        # Function keys (F1-F15) are acceptable as single-key shortcuts.
        import re as _re
        parts = [p.strip() for p in seq.split("+") if p.strip()]
        is_function_key = bool(_re.fullmatch(r'F([1-9]|1[0-5])', seq.strip(), _re.IGNORECASE))
        is_single = bool(seq) and len(parts) == 1 and not is_function_key
        if is_single:
            box = QMessageBox(self)
            box.setWindowTitle("Single-Key Shortcut")
            box.setIcon(QMessageBox.Icon.Information)
            box.setText(
                f'"{seq}" is a single-key shortcut.\n\n'
                f"It may trigger accidentally while typing. "
                f"Consider adding a modifier (e.g. Ctrl+{seq} or Alt+{seq}).")
            box.setStandardButtons(QMessageBox.StandardButton.Ok)
            box.exec()

    def _reset_all_hotkeys(self):
        """Reset every hotkey row to its built-in default."""
        for row in self._hotkey_rows:
            row._reset_to_default()
        # Conflict markers are cleared automatically via hotkey_changed signals

    def _refresh_archive_list(self):
        self._archive_list.clear()
        for n in self._store.archived_nodes():
            item = QListWidgetItem(f"{n.get('name','Unnamed')}  [{n.get('type','?')}]")
            item.setData(Qt.UserRole, n.get("id"))
            self._archive_list.addItem(item)

    def _restore_selected(self):
        item = self._archive_list.currentItem()
        if not item: return
        self._store.unarchive_node(item.data(Qt.UserRole))
        self._refresh_archive_list(); self.settings_changed.emit()

    def _delete_selected(self):
        item = self._archive_list.currentItem()
        if not item: return
        self._store.remove_node(item.data(Qt.UserRole))
        self._refresh_archive_list()

    def _sec(self, txt):
        t = _theme
        l = QLabel(txt); l.setFont(QFont("Segoe UI",8,QFont.Weight.DemiBold))
        l.setStyleSheet(f"color:{t.GLOW.name()}; letter-spacing:1px;")
        self._sec_labels.append(l)
        return l

    def _make_checkbox(self, label: str, checked: bool = False) -> QCheckBox:
        t = _theme
        g = t.GLOW
        cb = QCheckBox(label)
        cb.setChecked(checked)
        cb.setFont(QFont("Segoe UI", 9))
        cb.setStyleSheet(f"""
            QCheckBox {{
                color: {t.TEXT_SECONDARY.name()};
                spacing: 8px;
            }}
            QCheckBox::indicator {{
                width: 14px; height: 14px;
                border-radius: 3px;
                border: 1px solid rgba({g.red()},{g.green()},{g.blue()},80);
                background: rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
            }}
            QCheckBox::indicator:checked {{
                background: rgba({g.red()},{g.green()},{g.blue()},200);
                border-color: {g.name()};
            }}
            QCheckBox::indicator:hover {{
                border-color: {g.name()};
            }}
        """)
        return cb

    def _restyle_all(self):
        """Re-apply theme-dependent stylesheets to all owned widgets after a theme change."""
        t = _theme
        self._sep.setStyleSheet(f"color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},55);")
        self._tabs.setStyleSheet(f"""
            QTabWidget::pane {{ border:none; background:transparent; }}
            QTabBar::tab {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},180);
                color:{t.TEXT_SECONDARY.name()};
                padding:7px 18px; border:none; border-radius:4px 4px 0 0;
                font-family:'Segoe UI'; font-size:9pt;
            }}
            QTabBar::tab:selected {{
                background:rgba({t.TILE_BG_HOVER.red()},{t.TILE_BG_HOVER.green()},{t.TILE_BG_HOVER.blue()},200);
                color:{t.GLOW.name()};
                border-bottom:2px solid {t.GLOW.name()};
            }}
            QTabBar::tab:hover {{ background:rgba({t.TILE_BG_HOVER.red()},{t.TILE_BG_HOVER.green()},{t.TILE_BG_HOVER.blue()},160); }}
        """)
        g = t.GLOW
        dr, dg, db = int(g.red()*0.65), int(g.green()*0.65), int(g.blue()*0.65)
        self._ok_btn.setStyleSheet(f"""
            QPushButton {{
                background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 rgba({g.red()},{g.green()},{g.blue()},195),
                    stop:1 rgba({dr},{dg},{db},195));
                border:none; border-radius:6px;
                color:{t.TEXT_PRIMARY.name()}; padding:6px 18px;
            }}
            QPushButton:hover {{
                background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 rgba({g.red()},{g.green()},{g.blue()},240),
                    stop:1 rgba({dr},{dg},{db},240));
            }}
        """)
        self._browse_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},180);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:5px; color:{t.TEXT_SECONDARY.name()}; padding:0 10px;
            }}
            QPushButton:hover {{ background:rgba({t.TILE_BG_HOVER.red()},{t.TILE_BG_HOVER.green()},{t.TILE_BG_HOVER.blue()},200); color:{t.GLOW.name()}; }}
        """)
        self._save_path_edit.setStyleSheet(f"""
            QLineEdit {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:5px; color:{t.TEXT_PRIMARY.name()};
                padding:5px 9px; font-family:'Segoe UI'; font-size:9pt;
            }}
            QLineEdit:focus {{ border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},150); }}
        """)
        self._reset_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},180);
                border:1px solid rgba({t.ACCENT_RED.red()},{t.ACCENT_RED.green()},{t.ACCENT_RED.blue()},90);
                border-radius:5px; color:{t.TEXT_SECONDARY.name()}; padding:0 12px;
            }}
            QPushButton:hover {{ color:{t.ACCENT_RED.name()}; border-color:rgba({t.ACCENT_RED.red()},{t.ACCENT_RED.green()},{t.ACCENT_RED.blue()},160); }}
        """)
        self._archive_list.setStyleSheet(f"""
            QListWidget {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
                border-radius:6px; color:{t.TEXT_PRIMARY.name()};
                font-family:'Segoe UI'; font-size:9pt; outline:none;
            }}
            QListWidget::item {{ padding:6px 10px; }}
            QListWidget::item:selected {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);
                color:{t.GLOW.name()};
            }}
            QListWidget::item:hover {{
                background:rgba({t.TILE_BG_HOVER.red()},{t.TILE_BG_HOVER.green()},{t.TILE_BG_HOVER.blue()},180);
            }}
        """)
        self._restore_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},180);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:5px;color:{t.TEXT_SECONDARY.name()};padding:0 12px;
            }}
            QPushButton:hover {{ background:rgba({t.TILE_BG_HOVER.red()},{t.TILE_BG_HOVER.green()},{t.TILE_BG_HOVER.blue()},210);color:{t.GLOW.name()}; }}
        """)
        self._del_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({t.ACCENT_RED.red()//2},{t.ACCENT_RED.green()//3},{t.ACCENT_RED.blue()//3},120);
                border:1px solid rgba({t.ACCENT_RED.red()},{t.ACCENT_RED.green()},{t.ACCENT_RED.blue()},80);
                border-radius:5px;color:{t.TEXT_SECONDARY.name()};padding:0 12px;
            }}
            QPushButton:hover {{ background:rgba({t.ACCENT_RED.red()},{t.ACCENT_RED.green()},{t.ACCENT_RED.blue()},150);color:#ffffff; }}
        """)
        for lbl in self._sec_labels:
            lbl.setStyleSheet(f"color:{t.GLOW.name()}; letter-spacing:1px;")
        self._update_cursor_btn_styles()
        if hasattr(self, "_canvas_bg_btns"):
            self._update_canvas_bg_btn_styles()
        # Re-style canvas bg image controls
        if hasattr(self, "_bg_image_path_lbl"):
            self._bg_image_path_lbl.setStyleSheet(
                f"color:{t.TEXT_DIM.name()}; font-style:italic;")
        if hasattr(self, "_bg_image_browse_btn"):
            self._bg_image_browse_btn.setStyleSheet(f"""
                QPushButton {{
                    background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},180);
                    border:1px solid rgba({g.red()},{g.green()},{g.blue()},70);
                    border-radius:5px; color:{t.TEXT_SECONDARY.name()}; padding:0 8px;
                }}
                QPushButton:hover {{
                    background:rgba({t.TILE_BG_HOVER.red()},{t.TILE_BG_HOVER.green()},{t.TILE_BG_HOVER.blue()},200);
                    color:{g.name()};
                }}
            """)
        if hasattr(self, "_bg_opacity_slider"):
            self._bg_opacity_slider.setStyleSheet(f"""
                QSlider::groove:horizontal {{
                    height: 4px;
                    background: rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
                    border-radius: 2px;
                }}
                QSlider::sub-page:horizontal {{
                    background: rgba({g.red()},{g.green()},{g.blue()},180);
                    border-radius: 2px;
                }}
                QSlider::handle:horizontal {{
                    width: 14px; height: 14px; margin: -5px 0;
                    border-radius: 7px; background: {g.name()};
                }}
                QSlider::handle:horizontal:hover {{
                    background: rgba({g.red()},{g.green()},{g.blue()},220);
                }}
            """)
        if hasattr(self, "_bg_opacity_val_lbl"):
            self._bg_opacity_val_lbl.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        # Re-style custom cursor info labels
        if hasattr(self, "_custom_cursor_path_lbl"):
            self._custom_cursor_path_lbl.setStyleSheet(
                f"color:{t.TEXT_DIM.name()}; font-style:italic;")
        self.update()

    def _browse_save_path(self):
        d = QFileDialog.getExistingDirectory(self, "Select Save Directory")
        if d: self._save_path_edit.setText(d)

    def _apply(self):
        _settings_store.setValue("save_path", self._save_path_edit.text().strip())
        _settings_store.setValue("show_tips",
            "true" if self._chk_show_tips.isChecked() else "false")
        _settings_store.setValue("auto_launch_time_tracker",
            "true" if self._chk_auto_launch_time_tracker.isChecked() else "false")
        _settings_store.setValue("confirm_delete",
            "true" if self._chk_confirm_delete.isChecked() else "false")
        _settings_store.setValue("confirm_archive",
            "true" if self._chk_confirm_archive.isChecked() else "false")
        _settings_store.setValue("global_auto_launch",
            "true" if self._chk_global_auto_launch.isChecked() else "false")
        _settings_store.setValue("launch_on_single_click",
            "true" if self._chk_single_click.isChecked() else "false")
        _settings_store.setValue("disable_titlebar_anim",
            "true" if self._chk_titlebar_anim.isChecked() else "false")
        _settings_store.setValue("skip_startup_anim",
            "true" if self._chk_skip_splash.isChecked() else "false")
        _settings_store.setValue("disable_hotkeys",
            "true" if self._chk_disable_hotkeys.isChecked() else "false")
        _settings_store.setValue("ps1_bypass_execution_policy",
            "true" if self._chk_ps1_bypass.isChecked() else "false")
        _settings_store.setValue("clipboard_manager_enabled",
            "true" if self._chk_clipboard_enabled.isChecked() else "false")
        # ── Save customised hotkeys ────────────────────────────────────────
        if hasattr(self, "_pending_hotkeys"):
            for action_id, seq in self._pending_hotkeys.items():
                default = _CC_DEFAULT_HOTKEYS.get(action_id, "")
                if seq == default:
                    # Remove override so we fall back to the default naturally
                    _settings_store.remove(f"hotkey/{action_id}")
                else:
                    _settings_store.setValue(f"hotkey/{action_id}", seq)
        # Windows startup registry
        _set_windows_startup(self._chk_startup.isChecked())
        _settings_store.sync()
        _theme._load()
        _theme._notify()
        self.accept()

    def _drag_move(self, e: QMouseEvent):
        if self._drag_pos:
            delta = e.globalPosition().toPoint() - self._drag_pos
            self.move(self.pos() + delta)
            self._drag_pos = e.globalPosition().toPoint()

    def paintEvent(self, e: QPaintEvent):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        rect = QRectF(10,10,self.width()-20,self.height()-20)
        path = QPainterPath(); path.addRoundedRect(rect,BORDER_RADIUS+2,BORDER_RADIUS+2)
        grad = QLinearGradient(0,0,self.width(),self.height())
        grad.setColorAt(0,QColor(t.BG_MID.red()+4,t.BG_MID.green()+4,t.BG_MID.blue()+6,248))
        grad.setColorAt(1,QColor(t.BG_DARK.red(),t.BG_DARK.green(),t.BG_DARK.blue(),248))
        p.fillPath(path, grad)
        border = QColor(t.GLOW); border.setAlpha(100)
        p.setPen(QPen(border,1.3)); p.setBrush(Qt.NoBrush); p.drawPath(path)
        p.end()


# ---------------------------------------------------------------------------
# Helper: grid cell width in pixels from a node/folder dict
# ---------------------------------------------------------------------------

def _item_grid_span(item: dict) -> tuple[int, int]:
    """Return (col_span, row_span) for an item on the grid."""
    if item.get("type") == NODE_TYPE_FOLDER:
        return (2, 2)   # folders always occupy 2×2 cells
    return NODE_SIZES.get(item.get("size", "1x1"), (1, 1))


def _item_pixel_size(item: dict) -> tuple[int, int]:
    cols, rows = _item_grid_span(item)
    w = cols * TILE_BASE_SIZE + (cols - 1) * TILE_GAP
    h = rows * TILE_BASE_SIZE + (rows - 1) * TILE_GAP
    return w, h


# ---------------------------------------------------------------------------
# Folder dialog  (open a folder — shows nodes inside it)
# ---------------------------------------------------------------------------

class FolderViewDialog(QDialog):
    node_launch_requested        = Signal(dict)
    node_edit_requested          = Signal(dict)
    node_delete_requested        = Signal(dict)
    node_archive_requested       = Signal(dict)
    node_export_requested        = Signal(dict)
    node_duplicate_requested     = Signal(dict)
    node_remove_from_folder_requested = Signal(dict)

    def __init__(self, folder: dict, store: "NodeStore",
                 tooltip_widget: NodeToolTip, parent=None):
        super().__init__(parent, Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self._folder = folder
        self._store  = store
        self._tooltip = tooltip_widget
        self._drag_pos: Optional[QPoint] = None
        self._build_ui()
        # 4-column grid: 4×140 + 3×12 (gaps) + 40 (grid margins) + 20 (outer) = 656 min
        self.setMinimumSize(680, 500)
        self.resize(720, 580)

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(10,10,10,10); outer.setSpacing(0)
        card = QWidget()
        cl = QVBoxLayout(card); cl.setContentsMargins(0,0,0,0); cl.setSpacing(0)
        card.setGraphicsEffect(make_shadow(card, 28, QColor(0,0,0,200)))

        # title bar
        tbar = QWidget(); tbar.setFixedHeight(44)
        tbar.setStyleSheet("background:transparent;")
        tbl = QHBoxLayout(tbar); tbl.setContentsMargins(16,0,12,0)
        folder_icon = QLabel("⬡")
        folder_icon.setFont(QFont("Segoe UI",14)); folder_icon.setStyleSheet(
            f"color:{_theme.ACCENT_AMBER.name()};")
        ttl = QLabel(f"  {self._folder.get('name','Folder')}")
        ttl.setTextFormat(Qt.PlainText)
        ttl.setFont(QFont("Segoe UI",11,QFont.Weight.DemiBold))
        ttl.setStyleSheet(f"color:{_theme.TEXT_PRIMARY.name()};")
        tbl.addWidget(folder_icon); tbl.addWidget(ttl); tbl.addStretch()
        cb = TitleBarButton(COLOR_BTN_CLOSE,"x"); cb.clicked.connect(self.accept)
        tbl.addWidget(cb)
        tbar.mousePressEvent = lambda ev: (
            setattr(self,"_drag_pos",ev.globalPosition().toPoint())
            if ev.button()==Qt.LeftButton else None)
        tbar.mouseMoveEvent  = self._drag_move
        tbar.mouseReleaseEvent = lambda ev: setattr(self,"_drag_pos",None)
        cl.addWidget(tbar)
        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color:rgba(0,160,180,55);"); cl.addWidget(sep)

        # grid of tiles
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("background:transparent;")
        fw = QWidget(); fw.setStyleSheet("background:transparent;")
        grid = QGridLayout(fw)
        grid.setContentsMargins(20,16,20,16); grid.setSpacing(TILE_GAP)
        children_ids = set(self._folder.get("children",[]))
        child_nodes = [n for n in self._store.all_nodes()
                       if n.get("id") in children_ids]
        col_count = 4
        for idx, node in enumerate(child_nodes):
            tile = NodeTile(node, self._tooltip, fw)
            tile.launch_requested.connect(self.node_launch_requested)
            tile.edit_requested.connect(self.node_edit_requested)
            tile.delete_requested.connect(self.node_delete_requested)
            tile.archive_requested.connect(self.node_archive_requested)
            tile.export_requested.connect(self.node_export_requested)
            tile.duplicate_requested.connect(self.node_duplicate_requested)
            tile.remove_from_folder_requested.connect(self.node_remove_from_folder_requested)
            tile.show()
            grid.addWidget(tile, idx // col_count, idx % col_count)
        if not child_nodes:
            lbl = QLabel("This folder is empty.\nDrag nodes here to add them.")
            lbl.setFont(FONT_LABEL)
            lbl.setStyleSheet(f"color:{_theme.TEXT_DIM.name()};")
            lbl.setAlignment(Qt.AlignCenter)
            grid.addWidget(lbl, 0, 0)
        fw.setMinimumHeight(
            math.ceil(max(len(child_nodes),1)/col_count)*(TILE_BASE_SIZE+TILE_GAP)+40)
        scroll.setWidget(fw); cl.addWidget(scroll)
        outer.addWidget(card)

    def _drag_move(self, e: QMouseEvent):
        if self._drag_pos:
            delta = e.globalPosition().toPoint() - self._drag_pos
            self.move(self.pos() + delta)
            self._drag_pos = e.globalPosition().toPoint()

    def paintEvent(self, e: QPaintEvent):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        rect = QRectF(10,10,self.width()-20,self.height()-20)
        path = QPainterPath(); path.addRoundedRect(rect,BORDER_RADIUS+2,BORDER_RADIUS+2)
        grad = QLinearGradient(0,0,self.width(),self.height())
        grad.setColorAt(0,QColor(t.BG_MID.red()+4,t.BG_MID.green()+4,t.BG_MID.blue()+6,248))
        grad.setColorAt(1,QColor(t.BG_DARK.red(),t.BG_DARK.green(),t.BG_DARK.blue(),248))
        p.fillPath(path, grad)
        border = QColor(t.ACCENT_AMBER); border.setAlpha(120)
        p.setPen(QPen(border,1.3)); p.setBrush(Qt.NoBrush); p.drawPath(path)
        p.end()


# ---------------------------------------------------------------------------
# Folder tile
# ---------------------------------------------------------------------------

class FolderTile(QWidget):
    open_requested               = Signal(dict)
    rename_requested             = Signal(dict)
    delete_requested             = Signal(dict)
    delete_with_contents_requested = Signal(dict)
    empty_requested              = Signal(dict)
    # Drag handled by NodeCanvas via eventFilter

    def __init__(self, folder: dict, parent=None):
        super().__init__(parent)
        self._folder = folder
        self._hovered = False
        self._pressed = False
        self._dragging_visual = False   # set by canvas
        self._drop_target     = False   # set by canvas when a node hovers over it
        self._anim = 0.0
        cols, rows = _item_grid_span(folder)
        w = cols * TILE_BASE_SIZE + (cols - 1) * TILE_GAP
        h = rows * TILE_BASE_SIZE + (rows - 1) * TILE_GAP
        self.setFixedSize(w, h)
        self.setCursor(Qt.PointingHandCursor)
        self.setGraphicsEffect(make_shadow(self, 20, QColor(0, 0, 0, 140)))
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)

    def folder_data(self) -> dict: return self._folder

    def update_folder(self, folder: dict):
        self._folder = folder; self.update()

    def enterEvent(self, e): self._hovered = True;  self._timer.start(16)
    def leaveEvent(self, e): self._hovered = False; self._timer.start(16)

    def _tick(self):
        target = 1.0 if self._hovered else 0.0
        if abs(self._anim - target) < 0.08:
            self._anim = target; self._timer.stop()
        else:
            self._anim += 0.08 if target > self._anim else -0.08
        self.update()

    def set_dragging(self, on: bool):
        self._dragging_visual = on
        self.setCursor(Qt.ClosedHandCursor if on else Qt.PointingHandCursor)
        self.update()

    def set_drop_target(self, on: bool):
        """Highlight this folder when a node is hovering over it for a drop."""
        self._drop_target = on
        self.update()

    def mousePressEvent(self, e: QMouseEvent):
        if e.button() == Qt.LeftButton:
            self._pressed = True; self.update()

    def mouseReleaseEvent(self, e: QMouseEvent):
        if e.button() == Qt.LeftButton:
            self._pressed = False; self.update()
            if not self._dragging_visual and self.rect().contains(e.position().toPoint()):
                self.open_requested.emit(self._folder)

    def contextMenuEvent(self, e):
        menu = QMenu(self)
        menu.setStyleSheet(self._menu_css())
        act_open   = menu.addAction("Open Folder")
        act_rename = menu.addAction("Rename")
        act_empty  = menu.addAction("Empty Folder")
        menu.addSeparator()
        act_delete          = menu.addAction("Delete Folder  (keep nodes)")
        act_delete_contents = menu.addAction("Delete Folder and Contents")
        chosen = menu.exec(e.globalPos())
        if chosen == act_open:              self.open_requested.emit(self._folder)
        elif chosen == act_rename:          self.rename_requested.emit(self._folder)
        elif chosen == act_empty:           self.empty_requested.emit(self._folder)
        elif chosen == act_delete:          self.delete_requested.emit(self._folder)
        elif chosen == act_delete_contents: self.delete_with_contents_requested.emit(self._folder)

    def _menu_css(self):
        t = _theme
        return f"""
        QMenu {{
            background:rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},245);
            border:1px solid rgba({t.ACCENT_AMBER.red()},{t.ACCENT_AMBER.green()},{t.ACCENT_AMBER.blue()},110);
            border-radius:8px; padding:4px 0;
            color:{t.TEXT_PRIMARY.name()};
            font-family:'Segoe UI'; font-size:9pt;
        }}
        QMenu::item {{ padding:6px 18px; border-radius:4px; }}
        QMenu::item:selected {{
            background:rgba({t.ACCENT_AMBER.red()},{t.ACCENT_AMBER.green()},{t.ACCENT_AMBER.blue()},55);
            color:{t.ACCENT_AMBER.name()};
        }}
        QMenu::separator {{ height:1px; background:rgba({t.ACCENT_AMBER.red()},{t.ACCENT_AMBER.green()},{t.ACCENT_AMBER.blue()},50); margin:3px 8px; }}
        """

    def paintEvent(self, e: QPaintEvent):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        w, h = self.width(), self.height()
        rect = QRectF(1,1,w-2,h-2)
        path = QPainterPath(); path.addRoundedRect(rect,BORDER_RADIUS,BORDER_RADIUS)
        grad = QLinearGradient(0,0,w,h)
        ba = int(210 + 30*self._anim)
        am = t.ACCENT_AMBER
        if self._hovered:
            grad.setColorAt(0, QColor(am.red()//5, am.green()//8, 0, ba))
            grad.setColorAt(1, QColor(am.red()//8, am.green()//12, 0, ba - 20))
        else:
            grad.setColorAt(0, QColor(am.red()//6, am.green()//10, 0, 210))
            grad.setColorAt(1, QColor(am.red()//9, am.green()//14, 0, 200))
        p.fillPath(path, grad)

        # Shimmer highlight
        if self._hovered or self._anim > 0:
            shimmer = QLinearGradient(0, 0, w, 0)
            sa = int(14 * self._anim)
            shimmer.setColorAt(0,   QColor(255,255,255,0))
            shimmer.setColorAt(0.4, QColor(255,255,255,sa))
            shimmer.setColorAt(1.0, QColor(255,255,255,0))
            p.fillRect(QRectF(1,1,w-2,2), QBrush(shimmer))

        # Border
        if self._drop_target:
            amber = QColor(255, 210, 70); amber.setAlpha(255)
            p.setPen(QPen(amber, 3.0))
        elif self._dragging_visual:
            amber = QColor(t.ACCENT_AMBER); amber.setAlpha(255)
            p.setPen(QPen(amber, 2.5))
        elif self._hovered:
            amber = QColor(t.ACCENT_AMBER); amber.setAlpha(int(90 + 120*self._anim))
            p.setPen(QPen(amber, 1.6))
        else:
            amber = QColor(t.ACCENT_AMBER); amber.setAlpha(75)
            p.setPen(QPen(amber, 1.0))
        p.setBrush(Qt.NoBrush); p.drawPath(path)

        # Drop target outer glow
        if self._drop_target:
            glow = QColor(255, 220, 80, 55)
            p.setPen(QPen(glow, 10)); p.drawPath(path)

        # Folder icon
        cx, cy = w//2, h//2 - 12
        fa = QColor(t.ACCENT_AMBER); fa.setAlpha(int(150+80*self._anim))
        p.setPen(Qt.NoPen)
        tab = QPainterPath()
        tab.moveTo(cx-26, cy-22); tab.lineTo(cx-10, cy-22)
        tab.lineTo(cx-3,  cy-15); tab.lineTo(cx+26, cy-15)
        tab.lineTo(cx+26, cy+15); tab.lineTo(cx-26, cy+15); tab.closeSubpath()
        p.fillPath(tab, fa)

        # Child count
        count = len(self._folder.get("children",[]))
        if count:
            p.setFont(QFont("Segoe UI",9,QFont.Weight.DemiBold))
            p.setPen(QColor(0,0,0,210))
            p.drawText(QRect(cx-26,cy-24,52,32), Qt.AlignCenter, str(count))

        # Name
        name = self._folder.get("name","Folder")
        p.setFont(QFont("Segoe UI",8,QFont.Weight.DemiBold))
        # Frosted backing for name
        label_rect = QRect(4, h-28, w-8, 22)
        nbg = QColor(t.BG_DARK); nbg.setAlpha(140)
        p.setBrush(nbg); p.setPen(Qt.NoPen)
        p.drawRoundedRect(label_rect, 4, 4)
        nc = QColor(t.ACCENT_AMBER); nc.setAlpha(230); p.setPen(nc)
        fm = QFontMetrics(p.font())
        p.drawText(label_rect, Qt.AlignHCenter|Qt.AlignVCenter,
                   fm.elidedText(name, Qt.ElideRight, label_rect.width()-6))
        p.end()


# ---------------------------------------------------------------------------
# Node canvas  —  iOS/Android-style grid with animated reflow on drag
# ---------------------------------------------------------------------------

class NodeCanvas(QWidget):
    """
    Mobile home-screen style grid canvas.

    Architecture:
    - Canvas installs an eventFilter on ALL tile children.
    - All left-button drag logic is handled here, never inside tiles.
    - On drag start: tile is raised, canvas tracks cursor directly.
    - While dragging: other tiles animate to their new positions in real-time
      using QPropertyAnimation so you see them "flowing" out of the way.
    - On drop onto folder: node moves inside; on drop elsewhere: reorder.
    """
    node_edit_requested     = Signal(dict)
    node_delete_requested   = Signal(dict)
    node_archive_requested  = Signal(dict)
    node_launch_requested   = Signal(dict)
    node_export_requested    = Signal(dict)
    node_duplicate_requested = Signal(dict)
    node_remove_from_folder_requested = Signal(dict)
    add_node_requested      = Signal()
    folder_open_requested   = Signal(dict)
    folder_rename_requested = Signal(dict)
    folder_delete_requested              = Signal(dict)
    folder_delete_with_contents_requested = Signal(dict)
    folder_empty_requested               = Signal(dict)
    order_changed           = Signal(list)
    files_dropped           = Signal(list)   # list[str] of local file paths
    icon_set_requested      = Signal(int, str)  # (node_id, image_path)

    # Canvas-level context menu signals
    canvas_new_node_requested    = Signal()
    canvas_new_folder_requested  = Signal()
    canvas_refresh_requested     = Signal()
    canvas_settings_requested    = Signal()
    canvas_help_requested        = Signal()

    # Batch (multi-select) operation signals
    batch_delete_nodes_requested    = Signal(list)   # list[int] node ids
    batch_delete_folders_requested  = Signal(list)   # list[int] folder ids
    batch_delete_all_requested      = Signal(list, list)  # node_ids, folder_ids
    batch_archive_nodes_requested   = Signal(list)   # list[int] node ids
    batch_empty_folders_requested   = Signal(list)   # list[int] folder ids
    batch_place_in_folder_requested = Signal(list)   # list[int] node ids

    _M = 20   # canvas margin px

    def __init__(self, tooltip_widget: NodeToolTip,
                 store: "NodeStore", parent=None):
        super().__init__(parent)
        self._tooltip    = tooltip_widget
        self._store      = store
        self._tiles:     list[NodeTile]   = []
        self._folders:   list[FolderTile] = []
        self._add_tile   = AddNodeTile(self)
        self._add_tile.add_requested.connect(self.add_node_requested)
        self.setStyleSheet("background:transparent;")
        self.setMouseTracking(True)

        self._grid_cols  = GRID_COLS
        self._item_order: list[int] = []

        # --- drag state ---
        self._drag_tile:    Optional[QWidget] = None   # widget being dragged
        self._drag_offset:  QPoint            = QPoint()
        self._drag_idx:     int               = 0      # insertion slot
        self._drop_folder:  Optional[FolderTile] = None

        # press / hold detection
        self._press_tile:   Optional[QWidget] = None
        self._press_gpos:   QPoint            = QPoint()
        self._press_offset: QPoint            = QPoint()
        self._hold_timer    = QTimer(self)
        self._hold_timer.setSingleShot(True)
        self._hold_timer.timeout.connect(self._on_hold_timeout)

        # QPropertyAnimations keyed by item id
        self._anims: dict[int, QPropertyAnimation] = {}
        self._content_height: int = 0   # updated by _relayout*

        # Active filter order — list of matched IDs only; None = no filter
        self._filter_order: Optional[list[int]] = None
        # IDs hidden by the active filter (used by _pack_positions)
        self._filtered_out_ids: set[int] = set()
        # Temporary tiles surfaced from folders during a search
        self._search_tiles: list[NodeTile] = []

        # --- rubber-band / multi-select state ---
        self._selected_ids:  set[int]         = set()
        self._sel_start:     Optional[QPoint] = None
        self._sel_end:       Optional[QPoint] = None
        self._selecting:     bool             = False

        # --- background style cache ---
        # Read once; invalidated by refresh_bg_style() when settings change.
        self._bg_style: str = _settings_store.value("canvas_bg_style", "dots")
        # Per-style QPixmap caches stored as (QPixmap|None, QSize, color_key_tuple).
        # Regenerated only when canvas size or theme glow color changes.
        self._noise_cache: tuple = (None, QSize(), ())
        self._dots_cache:  tuple = (None, QSize(), ())
        self._grid_cache:  tuple = (None, QSize(), ())
        self._hex_cache:   tuple = (None, QSize(), ())
        self._web_cache:   tuple = (None, QSize(), ())
        # Custom image background cache: (QPixmap|None, path_str, QSize)
        self._bg_image_path: str = _settings_store.value("canvas_bg_image", "")
        self._bg_image_opacity: float = int(_settings_store.value("canvas_bg_image_opacity", 80)) / 100.0
        self._bg_image_cache: tuple = (None, "", QSize())

        self.setAcceptDrops(True)

    # ------------------------------------------------------------------
    # Event filter  — intercepts left-button events on all tile children
    # ------------------------------------------------------------------

    def eventFilter(self, obj: QObject, event: QEvent) -> bool:
        if not isinstance(obj, (NodeTile, FolderTile)):
            return False

        et = event.type()

        if et == QEvent.Type.MouseButtonPress:
            if event.button() == Qt.LeftButton:
                # Any tile press cancels rubber-band selection
                if self._selecting:
                    self._cancel_selection()
                self._press_tile   = obj
                self._press_gpos   = event.globalPosition().toPoint()
                self._press_offset = event.position().toPoint()
                self._hold_timer.start(_DRAG_HOLD_MS)
            return False          # tile still sees press for visual feedback

        elif et == QEvent.Type.MouseMove:
            if self._drag_tile is not None:
                # Active drag — move tile and animate others
                self._update_drag(event.globalPosition().toPoint())
                return True       # consume so tile doesn't see it
            elif self._press_tile is not None:
                # Decide: has the finger moved enough to start a drag?
                gp = event.globalPosition().toPoint()
                if (gp - self._press_gpos).manhattanLength() > 10:
                    self._hold_timer.stop()
                    self._start_drag()
                    if self._drag_tile:
                        self._update_drag(gp)
                    return True
            return False

        elif et == QEvent.Type.MouseButtonRelease:
            if event.button() == Qt.LeftButton:
                self._hold_timer.stop()
                if self._drag_tile is not None:
                    self._finish_drag()
                    return True   # consume — prevent spurious launch emit
                self._press_tile = None
            return False

        return False

    # ------------------------------------------------------------------
    # Drag lifecycle
    # ------------------------------------------------------------------

    def _on_hold_timeout(self):
        """Long-press timer fired → initiate drag."""
        if self._press_tile:
            self._start_drag()

    def _start_drag(self):
        tile = self._press_tile
        if tile is None:
            return
        drag_id = self._id_of(tile)
        try:
            self._drag_idx = self._item_order.index(drag_id)
        except ValueError:
            return

        self._drag_tile   = tile
        self._drag_offset = self._press_offset

        # Stop any running animation for this tile so it doesn't snap back
        self._stop_anim(drag_id)

        tile.set_dragging(True)
        tile.raise_()

    def _update_drag(self, global_pos: QPoint):
        """Called every MouseMove while a drag is active."""
        if not self._drag_tile:
            return

        # Move the tile directly under the cursor
        canvas_pos = self.mapFromGlobal(global_pos) - self._drag_offset
        self._drag_tile.move(canvas_pos)
        self._drag_tile.raise_()

        # Center of the dragged tile in canvas coordinates
        center = canvas_pos + QPoint(
            self._drag_tile.width() // 2,
            self._drag_tile.height() // 2)

        # Detect if we're hovering a folder for drop-into (1x1 tiles only)
        new_drop: Optional[FolderTile] = None
        if isinstance(self._drag_tile, NodeTile):
            drag_size = self._drag_tile.node_data().get("size", "1x1")
            if drag_size == "1x1":
                for ft in self._folders:
                    ftg = ft.geometry()
                    # Inflate hit-rect slightly for easier targeting
                    ftg.adjust(-8, -8, 8, 8)
                    if ftg.contains(center):
                        new_drop = ft
                        break

        # Only update folder highlight when it changes
        if new_drop is not self._drop_folder:
            if self._drop_folder:
                self._drop_folder.set_drop_target(False)
            self._drop_folder = new_drop
            if self._drop_folder:
                self._drop_folder.set_drop_target(True)

        # Compute new insertion index and animate others if it changed
        new_idx = self._compute_insert_idx(center)
        if new_idx != self._drag_idx:
            self._drag_idx = new_idx

        # Always re-animate non-drag tiles when not over a folder
        if not self._drop_folder:
            self._animate_to_virtual_layout()

    def _compute_insert_idx(self, canvas_center: QPoint) -> int:
        """
        Find the best linear insertion index for the dragged tile given
        where its center currently is in canvas coords.
        """
        drag_id  = self._id_of(self._drag_tile)
        non_drag = [i for i in self._item_order if i != drag_id]
        if not non_drag:
            return 0

        # Walk through non-dragged items and compute their grid center positions
        col, row = 0, 0
        slot_centers: list[tuple[int, QPoint]] = []
        for list_idx, item_id in enumerate(non_drag):
            w = self._widget_by_id(item_id)
            if w is None:
                continue
            sc, _ = self._span(w)
            if col + sc > self._grid_cols:
                col = 0; row += 1
            slot_px = self._cell_px(col, row)
            c = slot_px + QPoint(w.width() // 2, w.height() // 2)
            slot_centers.append((list_idx, c))
            col += sc
            if col >= self._grid_cols:
                col = 0; row += 1

        if not slot_centers:
            return 0

        # Find closest slot
        best_li = 0
        best_d  = float('inf')
        for li, c in slot_centers:
            d = (canvas_center - c).manhattanLength()
            if d < best_d:
                best_d = d; best_li = li

        # If cursor is to the right of the best slot's center, insert after
        _, best_c = slot_centers[best_li]
        if canvas_center.x() >= best_c.x():
            best_li += 1

        return min(best_li, len(non_drag))

    def _animate_to_virtual_layout(self):
        """
        Animate every non-dragged tile to where it would sit if the dragged
        tile were at _drag_idx.  This gives the iOS "tiles flow around finger"
        behaviour.
        """
        if not self._drag_tile:
            return

        drag_id  = self._id_of(self._drag_tile)
        non_drag = [i for i in self._item_order if i != drag_id]
        idx      = min(self._drag_idx, len(non_drag))
        virtual  = non_drag[:idx] + [drag_id] + non_drag[idx:]

        col, row = 0, 0
        for item_id in virtual:
            w = self._widget_by_id(item_id)
            if w is None:
                continue
            sc, _ = self._span(w)
            if col + sc > self._grid_cols:
                col = 0; row += 1

            if w is not self._drag_tile:
                target = self._cell_px(col, row)
                if w.pos() != target:
                    self._run_anim(item_id, w, target, 150, QEasingCurve.OutCubic)

            col += sc
            if col >= self._grid_cols:
                col = 0; row += 1

    def _finish_drag(self):
        """Drop: either into a folder or reorder in the grid."""
        if not self._drag_tile:
            return

        tile    = self._drag_tile
        drag_id = self._id_of(tile)
        tile.set_dragging(False)

        # Clear folder highlight
        if self._drop_folder:
            self._drop_folder.set_drop_target(False)

        if self._drop_folder and isinstance(tile, NodeTile):
            # ── Drop INTO folder ──────────────────────────────────────
            folder = self._drop_folder.folder_data()
            self._store.move_node_to_folder(drag_id, folder["id"])
            self._drop_folder    = None
            self._drag_tile      = None
            self._press_tile     = None
            self._item_order     = [i for i in self._item_order if i != drag_id]
            for t in self._tiles[:]:
                if t.node_data()["id"] == drag_id:
                    t.deleteLater(); self._tiles.remove(t)
            updated = next(
                (n for n in self._store.all_items() if n["id"] == folder["id"]),
                None)
            if updated:
                for ft in self._folders:
                    if ft.folder_data()["id"] == folder["id"]:
                        ft.update_folder(updated); break
        else:
            # ── Reorder ───────────────────────────────────────────────
            self._drop_folder = None
            non_drag          = [i for i in self._item_order if i != drag_id]
            idx               = min(self._drag_idx, len(non_drag))
            self._item_order  = non_drag[:idx] + [drag_id] + non_drag[idx:]
            self.order_changed.emit(self._item_order)
            self._drag_tile   = None
            self._press_tile  = None

        # Animate everything to final grid positions (with a satisfying snap-back)
        self._relayout_animated()

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------

    def _pack_positions(self) -> tuple[list[tuple[int, int]], tuple[int, int], int]:
        """Pack tiles into the grid using a 2-D occupancy map.

        Handles mixed tile sizes (1×1, 2×2, 1×2, 2×4 …) correctly:
        a large tile marks *all* cells it covers so smaller tiles never
        float on top of it.

        Returns (positions, (add_col, add_row), max_row) where max_row is
        the first empty row index *after* all placed tiles — used for the
        true content height regardless of where the add-tile lands.
        """
        occupied: set[tuple[int, int]] = set()
        cols = self._grid_cols
        positions: list[tuple[int, int]] = []
        for w in self._ordered_widgets():
            wid = self._id_of(w)
            if wid in self._filtered_out_ids:
                positions.append((-1, -1))   # placeholder; filtered tiles aren't placed
                continue
            sc, sr = self._span(w)
            sc = min(sc, cols)       # clamp oversized tiles to grid width
            placed = False
            search_row = 0
            while not placed:
                for c in range(cols):
                    if c + sc > cols:
                        break
                    if all((c + dc, search_row + dr) not in occupied
                           for dc in range(sc) for dr in range(sr)):
                        positions.append((c, search_row))
                        for dc in range(sc):
                            for dr in range(sr):
                                occupied.add((c + dc, search_row + dr))
                        placed = True
                        break
                if not placed:
                    search_row += 1
        # True bottom extent: first row index *after* every placed tile
        if positions:
            max_row = max(
                (r + self._span(w)[1]
                 for w, (_, r) in zip(self._ordered_widgets(), positions)
                 if r >= 0),   # skip hidden placeholders
                default=0)
        else:
            max_row = 0
        # Find first free 1×1 cell for the add-tile (may be a hole mid-grid)
        add_col, add_row = 0, max_row
        for r in range(max_row + 1):
            found = False
            for c in range(cols):
                if (c, r) not in occupied:
                    add_col, add_row = c, r
                    found = True
                    break
            if found:
                break
        else:
            add_col, add_row = 0, max_row
        return positions, (add_col, add_row), max_row

    def _relayout(self):
        """Immediate (non-animated) full layout — used on load/resize."""
        self._grid_cols = self._calc_grid_cols()
        positions, (ac, ar), max_row = self._pack_positions()
        for w, (col, row) in zip(self._ordered_widgets(), positions):
            if row >= 0:   # skip hidden placeholder positions
                w.move(self._cell_px(col, row))
        self._add_tile.move(self._cell_px(ac, ar))
        # Content height must reach the bottom of the last *tile*, not just
        # the add-tile row (which may sit in a hole earlier in the grid).
        bottom_row = max(ar + 1, max_row)
        self._content_height = self._cell_px(0, bottom_row).y() + self._M
        self.setMinimumHeight(self._content_height)
        self.updateGeometry()

    def _relayout_animated(self):
        """Animate all tiles to their final grid positions after a drop."""
        self._grid_cols = self._calc_grid_cols()
        positions, (ac, ar), max_row = self._pack_positions()
        for w, (col, row) in zip(self._ordered_widgets(), positions):
            if row < 0:   # hidden placeholder — stop any in-flight anim but don't move
                self._stop_anim(self._id_of(w))
                continue
            target = self._cell_px(col, row)
            wid = self._id_of(w)
            if w.pos() != target:
                self._run_anim(wid, w, target, 220, QEasingCurve.OutBack)
        self._add_tile.move(self._cell_px(ac, ar))
        bottom_row = max(ar + 1, max_row)
        self._content_height = self._cell_px(0, bottom_row).y() + self._M
        self.setMinimumHeight(self._content_height)
        self.updateGeometry()

    def sizeHint(self):
        return QSize(self.width(), self._content_height)

    def refresh_bg_style(self):
        """Re-read background style from settings and clear pixmap caches.
        Connect to settings_changed so the cached style and pre-rendered
        background pixmaps are invalidated whenever the user changes the style."""
        self._bg_style = _settings_store.value("canvas_bg_style", "dots")
        self._bg_image_path = _settings_store.value("canvas_bg_image", "")
        self._bg_image_opacity = int(_settings_store.value("canvas_bg_image_opacity", 80)) / 100.0
        self._noise_cache = (None, QSize(), ())
        self._dots_cache  = (None, QSize(), ())
        self._grid_cache  = (None, QSize(), ())
        self._hex_cache   = (None, QSize(), ())
        self._web_cache   = (None, QSize(), ())
        self._bg_image_cache = (None, "", QSize())
        self.update()

    def paintEvent(self, e: QPaintEvent):
        style = self._bg_style
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        if style == "gradient":
            grad = QLinearGradient(0, 0, 0, self.height())
            grad.setColorAt(0.0, QColor(t.BG_MID.red(),  t.BG_MID.green(),  t.BG_MID.blue(),  120))
            grad.setColorAt(1.0, QColor(t.BG_DARK.red(), t.BG_DARK.green(), t.BG_DARK.blue(), 220))
            p.fillRect(self.rect(), QBrush(grad))
        elif style == "dots":
            # Cache the dot pattern as a QPixmap; regenerate only on size or theme change.
            ck = (t.GLOW.red(), t.GLOW.green(), t.GLOW.blue())
            sz = self.size()
            pm, cached_sz, cached_ck = self._dots_cache
            if pm is None or cached_sz != sz or cached_ck != ck:
                pm = QPixmap(sz); pm.fill(Qt.transparent)
                pp = QPainter(pm); pp.setRenderHint(QPainter.Antialiasing)
                pp.setPen(Qt.NoPen); pp.setBrush(QColor(ck[0], ck[1], ck[2], 70))
                spacing = 24
                for x in range(0, sz.width() + spacing, spacing):
                    for y in range(0, sz.height() + spacing, spacing):
                        pp.drawEllipse(x - 1, y - 1, 2, 2)
                pp.end()
                self._dots_cache = (pm, sz, ck)
            p.drawPixmap(0, 0, pm)
        elif style == "grid":
            # Cache the grid pattern as a QPixmap; regenerate only on size or theme change.
            ck = (t.GLOW.red(), t.GLOW.green(), t.GLOW.blue())
            sz = self.size()
            pm, cached_sz, cached_ck = self._grid_cache
            if pm is None or cached_sz != sz or cached_ck != ck:
                pm = QPixmap(sz); pm.fill(Qt.transparent)
                pp = QPainter(pm)
                pp.setPen(QPen(QColor(ck[0], ck[1], ck[2], 50), 1))
                spacing = 32
                for x in range(0, sz.width() + spacing, spacing):
                    pp.drawLine(x, 0, x, sz.height())
                for y in range(0, sz.height() + spacing, spacing):
                    pp.drawLine(0, y, sz.width(), y)
                pp.end()
                self._grid_cache = (pm, sz, ck)
            p.drawPixmap(0, 0, pm)
        elif style == "noise":
            # Cache the noise pattern as a QPixmap; regenerate only on size or theme change.
            # Using a fixed seed (42) means the pattern is deterministic — pre-rendering is safe.
            ck = (t.GLOW.red(), t.GLOW.green(), t.GLOW.blue())
            sz = self.size()
            pm, cached_sz, cached_ck = self._noise_cache
            if pm is None or cached_sz != sz or cached_ck != ck:
                pm = QPixmap(sz); pm.fill(Qt.transparent)
                pp = QPainter(pm); pp.setRenderHint(QPainter.Antialiasing)
                pp.setPen(Qt.NoPen); pp.setBrush(QColor(ck[0], ck[1], ck[2], 60))
                rng = random.Random(42)
                count = min((sz.width() * sz.height()) // 80, 2800)
                for _ in range(count):
                    x = rng.randint(0, max(sz.width() - 2, 0))
                    y = rng.randint(0, max(sz.height() - 2, 0))
                    pp.drawRect(x, y, 2, 2)
                pp.end()
                self._noise_cache = (pm, sz, ck)
            p.drawPixmap(0, 0, pm)
        elif style == "hexagons":
            # Flat-top hexagon grid, cached as QPixmap.
            ck = (t.GLOW.red(), t.GLOW.green(), t.GLOW.blue())
            sz = self.size()
            pm, cached_sz, cached_ck = self._hex_cache
            if pm is None or cached_sz != sz or cached_ck != ck:
                pm = QPixmap(sz); pm.fill(Qt.transparent)
                pp = QPainter(pm); pp.setRenderHint(QPainter.Antialiasing)
                pp.setPen(QPen(QColor(ck[0], ck[1], ck[2], 55), 1))
                pp.setBrush(Qt.NoBrush)
                R = 22          # circumradius
                dx = R * math.sqrt(3)          # horizontal step between hex centres
                dy = R * 1.5                    # vertical step between rows
                col_idx = 0
                x = 0.0
                while x < sz.width() + dx:
                    y_off = (dx / 2) if (col_idx % 2 == 1) else 0.0
                    y = y_off - dy
                    while y < sz.height() + dy:
                        pts = []
                        for k in range(6):
                            angle = math.pi / 180 * (60 * k - 30)
                            pts.append((x + R * math.cos(angle),
                                        y + R * math.sin(angle)))
                        poly = [QPointF(px, py) for px, py in pts]
                        pp.drawPolygon(poly)
                        y += dy * 2
                    x += dx
                    col_idx += 1
                pp.end()
                self._hex_cache = (pm, sz, ck)
            p.drawPixmap(0, 0, pm)
        elif style == "web":
            # Spider-web pattern: concentric polygons connected by radial strands, cached.
            ck = (t.GLOW.red(), t.GLOW.green(), t.GLOW.blue())
            sz = self.size()
            pm, cached_sz, cached_ck = self._web_cache
            if pm is None or cached_sz != sz or cached_ck != ck:
                pm = QPixmap(sz); pm.fill(Qt.transparent)
                pp = QPainter(pm); pp.setRenderHint(QPainter.Antialiasing)
                strand_color  = QColor(ck[0], ck[1], ck[2], 50)
                ring_color    = QColor(ck[0], ck[1], ck[2], 40)
                # Tile the web across the canvas with overlapping anchor points
                tile_w = 180
                tile_h = 180
                strands   = 8
                rings     = 5
                ring_step = min(tile_w, tile_h) // (rings * 2)
                cx_list = range(-tile_w, sz.width()  + tile_w, tile_w)
                cy_list = range(-tile_h, sz.height() + tile_h, tile_h)
                for cx in cx_list:
                    for cy in cy_list:
                        # Draw radial strands
                        pp.setPen(QPen(strand_color, 1))
                        for s in range(strands):
                            angle = 2 * math.pi * s / strands
                            ex = cx + int((rings * ring_step) * math.cos(angle))
                            ey = cy + int((rings * ring_step) * math.sin(angle))
                            pp.drawLine(cx, cy, ex, ey)
                        # Draw concentric polygon rings
                        pp.setPen(QPen(ring_color, 1))
                        for r in range(1, rings + 1):
                            pts = []
                            for s in range(strands):
                                angle = 2 * math.pi * s / strands
                                rx = cx + int(r * ring_step * math.cos(angle))
                                ry = cy + int(r * ring_step * math.sin(angle))
                                pts.append(QPointF(rx, ry))
                            pts.append(pts[0])  # close the ring
                            for k in range(len(pts) - 1):
                                pp.drawLine(pts[k].toPoint(), pts[k + 1].toPoint())
                pp.end()
                self._web_cache = (pm, sz, ck)
            p.drawPixmap(0, 0, pm)
        elif style == "image":
            # Draw user-uploaded background image scaled to cover the canvas.
            # The pixmap is cached and only regenerated when path or canvas size changes.
            path = self._bg_image_path
            if path and os.path.isfile(path):
                sz = self.size()
                pm, cached_path, cached_sz = self._bg_image_cache
                if pm is None or cached_path != path or cached_sz != sz:
                    try:
                        src = QPixmap(path)
                        if not src.isNull():
                            # Scale to cover, preserving aspect ratio, then crop to center
                            scaled = src.scaled(
                                sz, Qt.KeepAspectRatioByExpanding, Qt.SmoothTransformation)
                            if scaled.size() != sz:
                                cx = (scaled.width()  - sz.width())  // 2
                                cy = (scaled.height() - sz.height()) // 2
                                pm = scaled.copy(cx, cy, sz.width(), sz.height())
                            else:
                                pm = scaled
                        else:
                            pm = None
                    except Exception as exc:
                        print(f"[CommandCenter] Canvas bg image render error: {exc}",
                              file=sys.stderr)
                        pm = None
                    self._bg_image_cache = (pm, path, sz)
                if pm is not None and not pm.isNull():
                    p.setOpacity(max(0.0, min(1.0, self._bg_image_opacity)))
                    p.drawPixmap(0, 0, pm)
                    p.setOpacity(1.0)
        # else "solid" → transparent, draw nothing

        # Draw selection highlights over selected tiles
        if self._selected_ids:
            sel_fill  = QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 55)
            sel_border = QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 220)
            p.setPen(QPen(sel_border, 2.5))
            p.setBrush(sel_fill)
            for w in self._ordered_widgets():
                wid = self._id_of(w)
                if wid in self._selected_ids:
                    gr = w.geometry()
                    p.drawRoundedRect(QRectF(gr.x() - 2, gr.y() - 2,
                                            gr.width() + 4, gr.height() + 4),
                                     BORDER_RADIUS + 1, BORDER_RADIUS + 1)

        # Draw rubber-band selection rectangle
        if self._selecting and self._sel_start and self._sel_end:
            rb = QRect(self._sel_start, self._sel_end).normalized()
            rb_fill   = QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 28)
            rb_border = QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 200)
            p.setBrush(rb_fill)
            p.setPen(QPen(rb_border, 1.5, Qt.DashLine))
            p.drawRoundedRect(QRectF(rb), 4, 4)

        p.end()

    def resizeEvent(self, e: QResizeEvent):
        super().resizeEvent(e)
        if not self._drag_tile:
            self._relayout()

    # ------------------------------------------------------------------
    # Empty-space mouse events — rubber-band multi-select
    # ------------------------------------------------------------------

    def mousePressEvent(self, e: QMouseEvent):
        """Left-drag on empty space → start rubber-band selection."""
        if e.button() == Qt.LeftButton:
            self._clear_selection()
            self._sel_start  = e.position().toPoint()
            self._sel_end    = e.position().toPoint()
            self._selecting  = True
            self.update()
        e.accept()

    def mouseMoveEvent(self, e: QMouseEvent):
        """Update rubber-band rectangle and compute which tiles are inside."""
        if self._selecting and self._sel_start is not None:
            self._sel_end = e.position().toPoint()
            self._update_selection()
            self.update()
        e.accept()

    def mouseReleaseEvent(self, e: QMouseEvent):
        """Finish rubber-band: if anything selected, show batch menu."""
        if e.button() == Qt.LeftButton and self._selecting:
            self._selecting = False
            self._sel_end   = e.position().toPoint()
            self._update_selection()
            self.update()
            if self._selected_ids:
                self._show_selection_menu(e.globalPosition().toPoint())
            else:
                self._sel_start = None
                self._sel_end   = None
        e.accept()

    def contextMenuEvent(self, e):
        """Right-click on empty canvas space → canvas-level actions menu."""
        menu = QMenu(self)
        menu.setStyleSheet(self._menu_css())
        act_new_node   = menu.addAction("✚  New Node")
        act_new_folder = menu.addAction("📁  New Folder")
        menu.addSeparator()
        act_refresh    = menu.addAction("↺  Refresh Display")
        menu.addSeparator()
        act_settings   = menu.addAction("⚙  Settings")
        act_help       = menu.addAction("?  Help")
        chosen = menu.exec(e.globalPos())
        if   chosen == act_new_node:   self.canvas_new_node_requested.emit()
        elif chosen == act_new_folder: self.canvas_new_folder_requested.emit()
        elif chosen == act_refresh:    self.canvas_refresh_requested.emit()
        elif chosen == act_settings:   self.canvas_settings_requested.emit()
        elif chosen == act_help:       self.canvas_help_requested.emit()

    # ------------------------------------------------------------------
    # Selection helpers
    # ------------------------------------------------------------------

    def _clear_selection(self):
        """Deselect all tiles and reset rubber-band state."""
        self._selected_ids.clear()
        self._sel_start   = None
        self._sel_end     = None
        self._selecting   = False
        self.update()

    def _cancel_selection(self):
        """Cancel an in-progress rubber-band without showing a menu."""
        self._selecting   = False
        self._selected_ids.clear()
        self._sel_start   = None
        self._sel_end     = None
        self.update()

    def _update_selection(self):
        """Recompute which tiles intersect the current rubber-band rect."""
        if self._sel_start is None or self._sel_end is None:
            return
        rb = QRect(self._sel_start, self._sel_end).normalized()
        self._selected_ids.clear()
        for w in self._ordered_widgets():
            if w.geometry().intersects(rb):
                self._selected_ids.add(self._id_of(w))

    def _show_selection_menu(self, global_pos: QPoint):
        """Show context menu for the current multi-selection."""
        node_ids   = [i for i in self._selected_ids
                      if any(t.node_data()["id"] == i for t in self._tiles + self._search_tiles)]
        folder_ids = [i for i in self._selected_ids
                      if any(f.folder_data()["id"] == i for f in self._folders)]

        has_nodes   = bool(node_ids)
        has_folders = bool(folder_ids)
        n_total     = len(self._selected_ids)

        menu = QMenu(self)
        menu.setStyleSheet(self._menu_css())
        title_lbl = menu.addAction(
            f"  {n_total} item{'s' if n_total != 1 else ''} selected")
        title_lbl.setEnabled(False)
        menu.addSeparator()

        act_del_nodes   = act_del_folders = act_del_all = None
        act_arch_nodes  = act_empty_folders = act_place  = None

        if has_nodes:
            act_del_nodes  = menu.addAction(
                f"🗑  Delete selected node{'s' if len(node_ids) != 1 else ''}"
                f" ({len(node_ids)})")
        if has_folders:
            act_del_folders = menu.addAction(
                f"🗑  Delete selected folder{'s' if len(folder_ids) != 1 else ''}"
                f" ({len(folder_ids)})")
        if has_nodes and has_folders:
            act_del_all    = menu.addAction(
                f"🗑  Delete all selected  ({n_total})")
        menu.addSeparator()
        if has_nodes:
            act_arch_nodes = menu.addAction(
                f"📦  Archive selected node{'s' if len(node_ids) != 1 else ''}"
                f" ({len(node_ids)})")
        if has_folders:
            act_empty_folders = menu.addAction(
                f"📤  Empty selected folder{'s' if len(folder_ids) != 1 else ''}"
                f" ({len(folder_ids)})")
        if has_nodes:
            menu.addSeparator()
            act_place = menu.addAction(
                f"📂  Place selected node{'s' if len(node_ids) != 1 else ''}"
                f" in new folder  ({len(node_ids)})")

        menu.addSeparator()
        act_clear = menu.addAction("✖  Clear selection")

        chosen = menu.exec(global_pos)

        if chosen == act_clear or chosen is None:
            self._clear_selection()
            return

        # Emit the relevant signal; MainWindow handles the actual operation
        try:
            if chosen == act_del_nodes and node_ids:
                self.batch_delete_nodes_requested.emit(node_ids)
            elif chosen == act_del_folders and folder_ids:
                self.batch_delete_folders_requested.emit(folder_ids)
            elif chosen == act_del_all:
                self.batch_delete_all_requested.emit(node_ids, folder_ids)
            elif chosen == act_arch_nodes and node_ids:
                self.batch_archive_nodes_requested.emit(node_ids)
            elif chosen == act_empty_folders and folder_ids:
                self.batch_empty_folders_requested.emit(folder_ids)
            elif chosen == act_place and node_ids:
                self.batch_place_in_folder_requested.emit(node_ids)
        finally:
            self._clear_selection()

    def _menu_css(self) -> str:
        t = _theme
        return (
            f"QMenu {{"
            f"  background: rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},245);"
            f"  border: 1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},110);"
            f"  border-radius: 8px; padding: 4px 0;"
            f"  color: {t.TEXT_PRIMARY.name()};"
            f"  font-family:'Segoe UI'; font-size:9pt;"
            f"}}"
            f"QMenu::item {{ padding:6px 18px; border-radius:4px; }}"
            f"QMenu::item:selected {{"
            f"  background: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);"
            f"  color: {t.GLOW.name()};"
            f"}}"
            f"QMenu::item:disabled {{ color: rgba({t.TEXT_DIM.red()},{t.TEXT_DIM.green()},{t.TEXT_DIM.blue()},180); }}"
            f"QMenu::separator {{"
            f"  height:1px; background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},45);"
            f"  margin:3px 8px;"
            f"}}"
        )

    # ------------------------------------------------------------------
    # Animation helper
    # ------------------------------------------------------------------

    def _run_anim(self, item_id: int, widget: QWidget, target: QPoint,
                  duration: int, curve: QEasingCurve):
        """Start (or restart) a smooth position animation for a tile."""
        if item_id in self._anims:
            old = self._anims.pop(item_id)
            old.stop()
        if widget.pos() == target:
            return
        anim = QPropertyAnimation(widget, b"pos", self)
        anim.setDuration(duration)
        anim.setEasingCurve(curve)
        anim.setStartValue(widget.pos())
        anim.setEndValue(target)
        def _done(wid=item_id): self._anims.pop(wid, None)
        anim.finished.connect(_done)
        anim.start()
        self._anims[item_id] = anim

    def _stop_anim(self, item_id: int):
        if item_id in self._anims:
            self._anims.pop(item_id).stop()

    # ------------------------------------------------------------------
    # Small grid helpers
    # ------------------------------------------------------------------

    def _calc_grid_cols(self) -> int:
        avail = max(self.width(), 600) - 2 * self._M
        return max(1, avail // (TILE_BASE_SIZE + TILE_GAP))

    def _cell_px(self, col: int, row: int) -> QPoint:
        return QPoint(self._M + col * (TILE_BASE_SIZE + TILE_GAP),
                      self._M + row * (TILE_BASE_SIZE + TILE_GAP))

    def _span(self, w: QWidget) -> tuple[int, int]:
        if isinstance(w, NodeTile):
            return NODE_SIZES.get(w.node_data().get("size", "1x1"), (1, 1))
        return _item_grid_span(w.folder_data())

    def _id_of(self, w: QWidget) -> int:
        if isinstance(w, NodeTile):
            return w.node_data()["id"]
        return w.folder_data()["id"]

    def _widget_by_id(self, item_id: int) -> Optional[QWidget]:
        for t in self._tiles:
            if t.node_data()["id"] == item_id: return t
        for f in self._folders:
            if f.folder_data()["id"] == item_id: return f
        for t in self._search_tiles:
            if t.node_data()["id"] == item_id: return t
        return None

    def _ordered_widgets(self) -> list[QWidget]:
        by_id: dict[int, QWidget] = {}
        for t in self._tiles:        by_id[t.node_data()["id"]] = t
        for f in self._folders:      by_id[f.folder_data()["id"]] = f
        for t in self._search_tiles: by_id[t.node_data()["id"]] = t
        order = self._filter_order if self._filter_order is not None else self._item_order
        return [by_id[i] for i in order if i in by_id]

    # ------------------------------------------------------------------
    # Widget creation (installs eventFilter)
    # ------------------------------------------------------------------

    def _add_tile_widget(self, node: dict):
        tile = NodeTile(node, self._tooltip, self)
        tile.launch_requested.connect(self.node_launch_requested)
        tile.edit_requested.connect(self.node_edit_requested)
        tile.delete_requested.connect(self.node_delete_requested)
        tile.archive_requested.connect(self.node_archive_requested)
        tile.export_requested.connect(self.node_export_requested)
        tile.duplicate_requested.connect(self.node_duplicate_requested)
        tile.installEventFilter(self)
        tile.show(); self._tiles.append(tile)

    def _add_folder_widget(self, folder: dict):
        ft = FolderTile(folder, self)
        ft.open_requested.connect(self.folder_open_requested)
        ft.rename_requested.connect(self.folder_rename_requested)
        ft.delete_requested.connect(self.folder_delete_requested)
        ft.delete_with_contents_requested.connect(self.folder_delete_with_contents_requested)
        ft.empty_requested.connect(self.folder_empty_requested)
        ft.installEventFilter(self)
        ft.show(); self._folders.append(ft)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def load_items(self, items: list[dict]):
        for t in self._tiles:   t.deleteLater()
        for f in self._folders: f.deleteLater()
        self._tiles.clear(); self._folders.clear()
        for item in items:
            if item.get("type") == NODE_TYPE_FOLDER:
                self._add_folder_widget(item)
            else:
                self._add_tile_widget(item)
        self._item_order = [item["id"] for item in items]
        self._relayout()

    def add_tile_for_node(self, node: dict):
        self._add_tile_widget(node)
        self._item_order.append(node["id"]); self._relayout()

    def apply_filter(self, query: str):
        """Float matching tiles to the top; restore original order when cleared.
        Also surfaces nodes that live inside folders if they match the query.
        """
        # Remove any temporary folder-child tiles from a previous search
        for st in self._search_tiles:
            st.deleteLater()
        self._search_tiles.clear()

        q = query.strip().lower()
        if not q:
            self._filter_order     = None
            self._filtered_out_ids = set()
            for tile in self._tiles:
                tile.setVisible(True)
            for ft in self._folders:
                ft.setVisible(True)
            self._relayout_animated()
            return

        matched_ids:   list[int] = []
        unmatched_ids: list[int] = []

        # Root-level items (regular tiles + folder tiles)
        for item_id in self._item_order:
            widget = self._widget_by_id(item_id)
            if widget is None:
                continue
            if isinstance(widget, NodeTile):
                node = widget.node_data()
                name = node.get("name", "").lower()
                tags = node.get("tags", [])
                match = q in name or any(q in str(tg).lower() for tg in tags)
            else:  # FolderTile
                match = q in widget.folder_data().get("name", "").lower()
            widget.setVisible(match)
            (matched_ids if match else unmatched_ids).append(item_id)

        # Nodes that live inside folders — surface them as temporary tiles
        folder_child_nodes = [
            n for n in self._store.all_nodes()
            if n.get("folder_id") and not n.get("archived", False)
        ]
        for node in folder_child_nodes:
            name = node.get("name", "").lower()
            tags = node.get("tags", [])
            if not (q in name or any(q in str(tg).lower() for tg in tags)):
                continue
            tile = NodeTile(node, self._tooltip, self)
            tile.launch_requested.connect(self.node_launch_requested)
            tile.edit_requested.connect(self.node_edit_requested)
            tile.delete_requested.connect(self.node_delete_requested)
            tile.archive_requested.connect(self.node_archive_requested)
            tile.export_requested.connect(self.node_export_requested)
            tile.duplicate_requested.connect(self.node_duplicate_requested)
            tile.remove_from_folder_requested.connect(
                self.node_remove_from_folder_requested)
            tile.installEventFilter(self)
            tile.show()
            self._search_tiles.append(tile)
            matched_ids.append(node["id"])

        # Matched items float to top; unmatched are hidden
        self._filtered_out_ids = set(unmatched_ids)
        self._filter_order = matched_ids + unmatched_ids
        self._relayout_animated()


    def add_tile_for_folder(self, folder: dict):
        self._add_folder_widget(folder)
        self._item_order.append(folder["id"]); self._relayout()

    def update_tile_for_node(self, node: dict):
        for tile in self._tiles:
            if tile.node_data().get("id") == node.get("id"):
                tile.update_node(node); self._relayout(); return

    def update_folder_tile(self, folder: dict):
        for ft in self._folders:
            if ft.folder_data().get("id") == folder.get("id"):
                ft.update_folder(folder); return

    def remove_tile_for_node(self, node_id: int):
        for tile in self._tiles[:]:
            if tile.node_data().get("id") == node_id:
                tile.deleteLater(); self._tiles.remove(tile)
        self._item_order = [i for i in self._item_order if i != node_id]
        self._relayout()

    def remove_folder_tile(self, folder_id: int):
        for ft in self._folders[:]:
            if ft.folder_data().get("id") == folder_id:
                ft.deleteLater(); self._folders.remove(ft)
        self._item_order = [i for i in self._item_order if i != folder_id]
        self._relayout()

    # ── External file drag-and-drop ────────────────────────────────────────

    def dragEnterEvent(self, e: QDragEnterEvent):
        if e.mimeData().hasUrls():
            # Only accept if at least one URL is a local file
            if any(u.isLocalFile() for u in e.mimeData().urls()):
                e.acceptProposedAction(); return
        e.ignore()

    def dragMoveEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
        else:
            e.ignore()

    def dropEvent(self, e: QDropEvent):
        paths = [u.toLocalFile() for u in e.mimeData().urls() if u.isLocalFile()]
        if not paths:
            e.acceptProposedAction(); return

        # If exactly one image is dropped directly onto an existing node tile,
        # set it as that tile's icon instead of creating a new node.
        _IMG = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".ico"}
        if len(paths) == 1 and Path(paths[0]).suffix.lower() in _IMG:
            drop_pos = e.position().toPoint()
            for tile in self._tiles:
                if tile.geometry().contains(drop_pos):
                    nid = tile.node_data().get("id")
                    if nid is not None:
                        self.icon_set_requested.emit(nid, paths[0])
                        e.acceptProposedAction(); return

        self.files_dropped.emit(paths)
        e.acceptProposedAction()


# ---------------------------------------------------------------------------
# Reminder System
# ---------------------------------------------------------------------------

import uuid as _uuid_mod
import winsound as _winsound

REMINDERS_FILE = CONFIG_DIR / "reminders.json"

# Reminder alert levels
REMINDER_LEVEL_1 = 1   # Notification / popup / sound
REMINDER_LEVEL_2 = 2   # Full-screen takeover on all monitors
REMINDER_LEVEL_3 = 3   # Must type "acknowledge" to dismiss

# Level-1 notification type keys
RNOTIF_WIN_TOAST = "windows_notification"
RNOTIF_POPUP     = "popup"
RNOTIF_SOUND     = "sound"

# Built-in sound labels → Windows SND_ALIAS scheme name
_BUILTIN_SOUNDS = {
    "Default Beep":  "SystemDefault",
    "Critical Stop": "SystemHand",
}


def _play_reminder_sound(sound_key: str, custom_path: str = "", volume: int = 80):
    """Play a reminder sound asynchronously.  Catches all errors silently."""
    try:
        if sound_key == "custom" and custom_path:
            p = Path(custom_path)
            if p.exists():
                if p.suffix.lower() == ".wav":
                    _winsound.PlaySound(str(p),
                                       _winsound.SND_FILENAME | _winsound.SND_ASYNC)
                else:
                    # .mp3 and other formats: WPF MediaPlayer via PowerShell
                    try:
                        uri = str(p).replace("'", "''")
                        subprocess.Popen(
                            ["powershell", "-NoProfile", "-WindowStyle", "Hidden",
                             "-Command",
                             f"Add-Type -AssemblyName presentationCore;"
                             f"$p=[System.Windows.Media.MediaPlayer]::new();"
                             f"$p.Open([Uri]'{uri}');"
                             f"$p.Play();Start-Sleep 20;$p.Close()"],
                            creationflags=0x08000000)
                    except Exception:
                        pass
        elif sound_key in _BUILTIN_SOUNDS:
            alias = _BUILTIN_SOUNDS[sound_key]
            try:
                _winsound.PlaySound(alias, _winsound.SND_ALIAS | _winsound.SND_ASYNC)
            except Exception:
                _winsound.MessageBeep(_winsound.MB_OK)
    except Exception as exc:
        print(f"[Reminder] sound error: {exc}", file=sys.stderr)


def _show_windows_toast(title: str, message: str):
    """Show a Windows 10/11 toast notification.

    Registers CommandCenter as a toast-capable AUMID in HKCU (one-time,
    idempotent) so CreateToastNotifier has a valid registered app to work with.
    """
    try:
        import winreg, base64 as _b64
        # Register the AUMID so Windows lets us send toasts
        _aumid_key = r"SOFTWARE\Classes\AppUserModelId\CommandCenter"
        try:
            with winreg.CreateKeyEx(
                    winreg.HKEY_CURRENT_USER, _aumid_key,
                    access=winreg.KEY_SET_VALUE) as _k:
                winreg.SetValueEx(_k, "DisplayName", 0, winreg.REG_SZ,
                                  "CommandCenter")
        except Exception:
            pass

        def _xml_esc(s: str) -> str:
            return (s.replace("&", "&amp;").replace("<", "&lt;")
                     .replace(">", "&gt;").replace('"', "&quot;")
                     .replace("'", "&apos;"))

        t_safe = _xml_esc(str(title)[:200])
        m_safe = _xml_esc(str(message)[:400])
        ps = "\n".join([
            "[Windows.UI.Notifications.ToastNotificationManager,"
            "Windows.UI.Notifications,ContentType=WindowsRuntime] | Out-Null",
            "[Windows.Data.Xml.Dom.XmlDocument,"
            "Windows.Data.Xml.Dom,ContentType=WindowsRuntime] | Out-Null",
            "$x = [Windows.Data.Xml.Dom.XmlDocument]::new()",
            f'$x.LoadXml(\'<toast><visual><binding template="ToastGeneric">'
            f"<text>{t_safe}</text><text>{m_safe}</text>"
            f"</binding></visual></toast>\')" ,
            "[Windows.UI.Notifications.ToastNotificationManager]"
            "::CreateToastNotifier('CommandCenter')"
            ".Show([Windows.UI.Notifications.ToastNotification]::new($x))",
        ])
        enc = _b64.b64encode(ps.encode("utf-16-le")).decode("ascii")
        subprocess.Popen(
            ["powershell", "-NoProfile", "-WindowStyle", "Hidden",
             "-EncodedCommand", enc],
            creationflags=0x08000000)
    except Exception as exc:
        print(f"[Reminder] Windows toast error: {exc}", file=sys.stderr)


# ── Reminder Store ────────────────────────────────────────────────────────────

class ReminderStore:
    """JSON-backed persistence for reminders at CONFIG_DIR/reminders.json."""

    def __init__(self):
        self._data: list[dict] = []
        self._load()

    def _load(self):
        try:
            if REMINDERS_FILE.exists():
                with open(REMINDERS_FILE, "r", encoding="utf-8") as fh:
                    loaded = json.load(fh)
                    self._data = loaded if isinstance(loaded, list) else []
            else:
                self._data = []
        except Exception as exc:
            print(f"[ReminderStore] load error: {exc}", file=sys.stderr)
            self._data = []

    def save(self):
        try:
            REMINDERS_FILE.parent.mkdir(parents=True, exist_ok=True)
            with open(REMINDERS_FILE, "w", encoding="utf-8") as fh:
                json.dump(self._data, fh, indent=2, ensure_ascii=False)
        except Exception as exc:
            print(f"[ReminderStore] save error: {exc}", file=sys.stderr)

    def all(self) -> list:
        return list(self._data)

    def get(self, reminder_id: str) -> Optional[dict]:
        for r in self._data:
            if r.get("id") == reminder_id:
                return dict(r)
        return None

    def add(self, reminder: dict) -> dict:
        if "id" not in reminder or not reminder["id"]:
            reminder["id"] = _uuid_mod.uuid4().hex[:10]
        if "created" not in reminder:
            reminder["created"] = datetime.now().isoformat(timespec="seconds")
        self._data.append(reminder)
        self.save()
        return reminder

    def update(self, reminder_id: str, updates: dict):
        for r in self._data:
            if r.get("id") == reminder_id:
                r.update(updates)
                self.save()
                return

    def delete(self, reminder_id: str):
        self._data = [r for r in self._data if r.get("id") != reminder_id]
        self.save()

    def enabled(self) -> list:
        """Return only enabled (active) reminders."""
        return [r for r in self._data if r.get("enabled", True)]


_reminder_store: Optional[ReminderStore] = None


def _get_reminder_store() -> ReminderStore:
    global _reminder_store
    if _reminder_store is None:
        _reminder_store = ReminderStore()
    return _reminder_store


# ── Level-1 Popup Dialog ──────────────────────────────────────────────────────

class _ReminderPopupDialog(QDialog):
    """Level-1 popup reminder dialog — themed, dismissable, with optional snooze."""

    snoozed = Signal(int)   # snooze minutes

    def __init__(self, reminder: dict, parent=None):
        super().__init__(parent, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setModal(False)
        self._reminder = reminder
        self._build_ui()
        self.adjustSize()

    def _build_ui(self):
        t = _theme
        outer = QVBoxLayout(self)
        outer.setContentsMargins(10, 10, 10, 10)

        card = QWidget()
        card.setObjectName("rem_card")
        card.setStyleSheet(f"""
            QWidget#rem_card {{
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},250),
                    stop:1 rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},252));
                border: 1.5px solid rgba({t.ACCENT_AMBER.red()},{t.ACCENT_AMBER.green()},{t.ACCENT_AMBER.blue()},160);
                border-radius: 12px;
            }}
        """)
        card.setGraphicsEffect(make_shadow(card, 28, QColor(0, 0, 0, 180)))
        cl = QVBoxLayout(card)
        cl.setContentsMargins(22, 18, 22, 18)
        cl.setSpacing(10)

        # Header row
        hdr = QHBoxLayout(); hdr.setSpacing(10)
        icon_lbl = QLabel("🔔")
        icon_lbl.setFont(QFont("Segoe UI Emoji", 18))
        hdr.addWidget(icon_lbl)
        title_lbl = QLabel(self._reminder.get("title", "Reminder"))
        title_lbl.setFont(QFont("Segoe UI", 12, QFont.Weight.DemiBold))
        title_lbl.setStyleSheet(f"color:{t.ACCENT_AMBER.name()};")
        title_lbl.setWordWrap(True)
        hdr.addWidget(title_lbl, 1)
        cl.addLayout(hdr)

        # Message
        msg = self._reminder.get("message", "").strip()
        if msg:
            msg_lbl = QLabel(msg)
            msg_lbl.setFont(QFont("Segoe UI", 10))
            msg_lbl.setStyleSheet(f"color:{t.TEXT_PRIMARY.name()};")
            msg_lbl.setWordWrap(True)
            msg_lbl.setMaximumWidth(380)
            cl.addWidget(msg_lbl)

        cl.addSpacing(6)

        # Buttons
        btn_row = QHBoxLayout(); btn_row.setSpacing(8)
        snooze_min = int(self._reminder.get("snooze_minutes", 5))
        if snooze_min > 0:
            snooze_btn = QPushButton(f"⏱ Snooze {snooze_min}m")
            snooze_btn.setFont(QFont("Segoe UI", 9))
            snooze_btn.setFixedHeight(30)
            snooze_btn.setCursor(Qt.PointingHandCursor)
            snooze_btn.setStyleSheet(f"""
                QPushButton {{
                    background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},30);
                    border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},80);
                    border-radius:6px; color:{t.TEXT_SECONDARY.name()}; padding:0 14px;
                }}
                QPushButton:hover {{ background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},55);
                    color:{t.GLOW.name()}; }}
            """)
            snooze_btn.clicked.connect(lambda: (self.snoozed.emit(snooze_min), self.accept()))
            btn_row.addWidget(snooze_btn)
        btn_row.addStretch()
        ok_btn = QPushButton("Dismiss")
        ok_btn.setFont(QFont("Segoe UI", 9, QFont.Weight.DemiBold))
        ok_btn.setFixedHeight(30)
        ok_btn.setCursor(Qt.PointingHandCursor)
        ok_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({t.ACCENT_AMBER.red()},{t.ACCENT_AMBER.green()},{t.ACCENT_AMBER.blue()},160);
                border:none; border-radius:6px;
                color:{t.BG_DARK.name()}; padding:0 18px; font-weight:600;
            }}
            QPushButton:hover {{
                background:rgba({t.ACCENT_AMBER.red()},{t.ACCENT_AMBER.green()},{t.ACCENT_AMBER.blue()},220);
            }}
        """)
        ok_btn.clicked.connect(self.accept)
        btn_row.addWidget(ok_btn)
        cl.addLayout(btn_row)

        outer.addWidget(card)

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key_Return, Qt.Key_Enter, Qt.Key_Escape):
            self.accept()
        else:
            super().keyPressEvent(e)

    def paintEvent(self, e):
        p = QPainter(self)
        p.fillRect(self.rect(), Qt.transparent)
        p.end()


# ── Level-2 Full-Screen Overlay ───────────────────────────────────────────────

# Holds live overlay objects so Python doesn't GC them before dismissal
_active_fullscreen_overlays: list = []


class _ReminderFullscreenWidget(QWidget):
    """A single full-screen overlay widget for one monitor."""

    dismissed = Signal()

    def __init__(self, reminder: dict, screen, parent=None):
        super().__init__(parent, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self._reminder = reminder
        # Cover this specific screen exactly (do NOT call setWindowState/showFullScreen
        # — that overrides the geometry to the primary screen)
        self.setGeometry(screen.geometry())
        self.setAttribute(Qt.WA_TranslucentBackground, False)
        # Set countdown BEFORE _build_ui so the label can read it
        timeout_sec = max(1, int(reminder.get("fullscreen_timeout", 10)))
        self._countdown = timeout_sec
        self._build_ui()
        # Auto-dismiss timer
        self._timer = QTimer(self)
        self._timer.setInterval(1000)
        self._timer.timeout.connect(self._tick)
        self._timer.start()

    def _build_ui(self):
        t = _theme
        bg_hex  = self._reminder.get("fullscreen_bg_color") or "#12001a"
        txt_hex = self._reminder.get("fullscreen_text_color") or "#ff4488"
        try:
            bg_c  = QColor(bg_hex)
            txt_c = QColor(txt_hex)
        except Exception:
            bg_c  = QColor("#12001a")
            txt_c = QColor("#ff4488")

        self.setStyleSheet(f"background: {bg_c.name()};")

        lay = QVBoxLayout(self)
        lay.setAlignment(Qt.AlignCenter)
        lay.setSpacing(24)

        icon_lbl = QLabel("⚠")
        icon_lbl.setFont(QFont("Segoe UI Emoji", 72))
        icon_lbl.setAlignment(Qt.AlignCenter)
        icon_lbl.setStyleSheet(f"color: {txt_c.name()};")
        lay.addWidget(icon_lbl)

        title_lbl = QLabel(self._reminder.get("title", "REMINDER"))
        title_lbl.setFont(QFont("Segoe UI", 42, QFont.Weight.Bold))
        title_lbl.setAlignment(Qt.AlignCenter)
        title_lbl.setWordWrap(True)
        title_lbl.setStyleSheet(f"color: {txt_c.name()};")
        lay.addWidget(title_lbl)

        msg = self._reminder.get("message", "").strip()
        if msg:
            msg_lbl = QLabel(msg)
            msg_lbl.setFont(QFont("Segoe UI", 22))
            msg_lbl.setAlignment(Qt.AlignCenter)
            msg_lbl.setWordWrap(True)
            msg_lbl.setStyleSheet(f"color: rgba({txt_c.red()},{txt_c.green()},{txt_c.blue()},200);")
            lay.addWidget(msg_lbl)

        lay.addSpacing(16)
        self._cd_lbl = QLabel(f"Dismissing in {self._countdown}s  —  click or press any key")
        self._cd_lbl.setFont(QFont("Segoe UI", 14))
        self._cd_lbl.setAlignment(Qt.AlignCenter)
        self._cd_lbl.setStyleSheet(f"color: rgba({txt_c.red()},{txt_c.green()},{txt_c.blue()},140);")
        lay.addWidget(self._cd_lbl)

    def _tick(self):
        self._countdown -= 1
        try:
            self._cd_lbl.setText(f"Dismissing in {self._countdown}s  —  click or press any key")
        except Exception:
            pass
        if self._countdown <= 0:
            self._dismiss()

    def _dismiss(self):
        try:
            self._timer.stop()
        except Exception:
            pass
        self.dismissed.emit()
        self.close()

    def mousePressEvent(self, e):
        self._dismiss()

    def keyPressEvent(self, e):
        self._dismiss()


class _ReminderFullscreenOverlay:
    """Creates and manages full-screen overlay widgets on all connected monitors."""

    def __init__(self, reminder: dict):
        self._widgets: list[_ReminderFullscreenWidget] = []
        screens = QApplication.screens()
        if not screens:
            screens = [QApplication.primaryScreen()]
        for screen in screens:
            try:
                w = _ReminderFullscreenWidget(reminder, screen)
                w.dismissed.connect(self._on_dismissed)
                self._widgets.append(w)
            except Exception as exc:
                print(f"[Reminder] fullscreen widget error: {exc}", file=sys.stderr)
        # Keep a reference so Python doesn't GC this object before dismissal
        _active_fullscreen_overlays.append(self)
        for w in self._widgets:
            try:
                w.show()
                w.raise_()
                w.activateWindow()
            except Exception:
                pass

    def _on_dismissed(self):
        """Dismiss all overlays when any one is dismissed."""
        for w in self._widgets:
            try:
                if w.isVisible():
                    w._timer.stop()
                    w.close()
            except Exception:
                pass
        # Release our own reference so this object can be GC'd
        try:
            _active_fullscreen_overlays.remove(self)
        except ValueError:
            pass


# ── Level-3 Acknowledge Dialog ────────────────────────────────────────────────

class _ReminderAcknowledgeDialog(QDialog):
    """Level-3: Flashing red warning — must type 'acknowledge' to close."""

    _FLASH_INTERVAL_MS = 450

    def __init__(self, reminder: dict, parent=None):
        super().__init__(parent,
                         Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setModal(True)
        self._reminder = reminder
        self._flash_state = False
        self._build_ui()
        self.adjustSize()
        self._flash_timer = QTimer(self)
        self._flash_timer.setInterval(self._FLASH_INTERVAL_MS)
        self._flash_timer.timeout.connect(self._flash_tick)
        self._flash_timer.start()

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(12, 12, 12, 12)

        self._card = QWidget()
        self._card.setObjectName("ack_card")
        cl = QVBoxLayout(self._card)
        cl.setContentsMargins(28, 24, 28, 24)
        cl.setSpacing(14)
        self._card.setGraphicsEffect(make_shadow(self._card, 36, QColor(220, 0, 0, 160)))

        # Warning banner
        warn_lbl = QLabel("⚠  CRITICAL REMINDER  ⚠")
        warn_lbl.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        warn_lbl.setAlignment(Qt.AlignCenter)
        warn_lbl.setStyleSheet("color: #ff2222; letter-spacing: 2px;")
        cl.addWidget(warn_lbl)

        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color: rgba(255,40,40,120);")
        cl.addWidget(sep)

        title_lbl = QLabel(self._reminder.get("title", "REMINDER"))
        title_lbl.setFont(QFont("Segoe UI", 15, QFont.Weight.Bold))
        title_lbl.setWordWrap(True)
        title_lbl.setAlignment(Qt.AlignCenter)
        title_lbl.setStyleSheet("color: #ffcccc;")
        cl.addWidget(title_lbl)

        msg = self._reminder.get("message", "").strip()
        if msg:
            msg_lbl = QLabel(msg)
            msg_lbl.setFont(QFont("Segoe UI", 11))
            msg_lbl.setWordWrap(True)
            msg_lbl.setAlignment(Qt.AlignCenter)
            msg_lbl.setStyleSheet("color: rgba(255,200,200,220);")
            msg_lbl.setMaximumWidth(460)
            cl.addWidget(msg_lbl)

        cl.addSpacing(8)

        instruct_lbl = QLabel('Type  <b style="color:#ff4444">acknowledge</b>  below to dismiss:')
        instruct_lbl.setFont(QFont("Segoe UI", 10))
        instruct_lbl.setAlignment(Qt.AlignCenter)
        instruct_lbl.setStyleSheet("color: rgba(255,180,180,200);")
        cl.addWidget(instruct_lbl)

        self._ack_edit = QLineEdit()
        self._ack_edit.setFont(QFont("Segoe UI", 11))
        self._ack_edit.setFixedHeight(36)
        self._ack_edit.setPlaceholderText("Type 'acknowledge' here…")
        self._ack_edit.setStyleSheet("""
            QLineEdit {
                background: rgba(60,0,0,220);
                border: 1.5px solid rgba(255,60,60,160);
                border-radius: 6px;
                color: #ffaaaa;
                padding: 0 12px;
            }
            QLineEdit:focus {
                border-color: rgba(255,80,80,240);
            }
        """)
        self._ack_edit.textChanged.connect(self._on_text_changed)
        self._ack_edit.returnPressed.connect(self._try_dismiss)
        cl.addWidget(self._ack_edit)

        self._ok_btn = QPushButton("Acknowledge")
        self._ok_btn.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        self._ok_btn.setFixedHeight(36)
        self._ok_btn.setEnabled(False)
        self._ok_btn.setCursor(Qt.PointingHandCursor)
        self._ok_btn.setStyleSheet("""
            QPushButton:disabled {
                background: rgba(80,0,0,120);
                border: 1px solid rgba(120,30,30,100);
                border-radius: 6px;
                color: rgba(180,100,100,150);
            }
            QPushButton:enabled {
                background: rgba(200,30,30,200);
                border: 1px solid rgba(255,80,80,200);
                border-radius: 6px;
                color: #ffffff;
            }
            QPushButton:enabled:hover {
                background: rgba(230,50,50,230);
            }
        """)
        self._ok_btn.clicked.connect(self._try_dismiss)
        cl.addWidget(self._ok_btn)

        outer.addWidget(self._card)
        self._update_card_style(False)

    def _update_card_style(self, flash_on: bool):
        if flash_on:
            bg = "rgba(50,0,0,252)"
            border = "rgba(255,40,40,230)"
        else:
            bg = "rgba(30,0,0,248)"
            border = "rgba(180,20,20,180)"
        self._card.setStyleSheet(f"""
            QWidget#ack_card {{
                background: {bg};
                border: 2px solid {border};
                border-radius: 14px;
            }}
        """)

    def _flash_tick(self):
        self._flash_state = not self._flash_state
        self._update_card_style(self._flash_state)

    def _on_text_changed(self, text: str):
        self._ok_btn.setEnabled(text.strip().lower() == "acknowledge")

    def _try_dismiss(self):
        if self._ack_edit.text().strip().lower() == "acknowledge":
            try:
                self._flash_timer.stop()
            except Exception:
                pass
            self.accept()

    def keyPressEvent(self, e):
        # Only Escape-like shortcuts ignored — user must type acknowledge
        if e.key() not in (Qt.Key_Escape,):
            super().keyPressEvent(e)

    def closeEvent(self, e):
        # Block window-close button — only dismissable via acknowledgment
        if self._ack_edit.text().strip().lower() != "acknowledge":
            e.ignore()
        else:
            try:
                self._flash_timer.stop()
            except Exception:
                pass
            super().closeEvent(e)

    def paintEvent(self, e):
        p = QPainter(self)
        p.fillRect(self.rect(), Qt.transparent)
        p.end()


# ── Reminder Scheduler ────────────────────────────────────────────────────────

class ReminderScheduler(QObject):
    """Polls enabled reminders every 30 seconds and fires those that are due."""

    reminder_fired = Signal(dict)   # emitted with the full reminder dict

    _POLL_MS = 30_000

    def __init__(self, store: ReminderStore, parent=None):
        super().__init__(parent)
        self._store   = store
        self._win     = None
        self._timer   = QTimer(self)
        self._timer.setInterval(self._POLL_MS)
        self._timer.timeout.connect(self._poll)
        self._snooze_map: dict[str, datetime] = {}   # id → snooze-until dt

    def set_main_window(self, win):
        self._win = win

    def start(self):
        self._timer.start()
        QTimer.singleShot(6_000, self._poll)

    def stop(self):
        self._timer.stop()

    def snooze(self, reminder_id: str, minutes: int):
        self._snooze_map[reminder_id] = (
            datetime.now() + timedelta(minutes=minutes))

    def _poll(self):
        try:
            now = datetime.now()
            for reminder in self._store.enabled():
                rid = reminder.get("id", "")
                # Respect snooze
                if rid in self._snooze_map:
                    if now < self._snooze_map[rid]:
                        continue
                    else:
                        del self._snooze_map[rid]
                sched = reminder.get("schedule")
                if not sched or not isinstance(sched, dict):
                    continue
                if not sched.get("enabled", False):
                    continue
                if _sched_should_fire(sched):
                    sched["last_run"] = now.isoformat(timespec="seconds")
                    self._store.update(rid, {"schedule": sched})
                    try:
                        self._fire(reminder)
                    except Exception as exc:
                        print(f"[ReminderScheduler] fire error: {exc}",
                              file=sys.stderr)
        except Exception as exc:
            print(f"[ReminderScheduler] poll error: {exc}", file=sys.stderr)

    def _fire(self, reminder: dict):
        level = int(reminder.get("level", 1))
        # Level-1 notification types
        if level == 1:
            notif_types = reminder.get("notification_types",
                                       [RNOTIF_POPUP])
            if RNOTIF_WIN_TOAST in notif_types:
                _show_windows_toast(
                    reminder.get("title", "Reminder"),
                    reminder.get("message", ""))
            if RNOTIF_SOUND in notif_types:
                _play_reminder_sound(
                    reminder.get("sound_key", "Default Beep"),
                    reminder.get("sound_custom_path", ""))
            if RNOTIF_POPUP in notif_types:
                if self._win is not None:
                    dlg = _ReminderPopupDialog(reminder)
                    dlg.snoozed.connect(
                        lambda mins, r=reminder: self.snooze(r.get("id",""), mins))
                    dlg.setParent(self._win, Qt.Dialog)
                    dlg.setAttribute(Qt.WA_DeleteOnClose)
                    # Position top-right of main window
                    geo = self._win.frameGeometry()
                    dlg.adjustSize()
                    dlg.move(geo.right() - dlg.width() - 20,
                             geo.top() + 60)
                    dlg.show()
                    dlg.raise_()

        elif level == 2:
            _ReminderFullscreenOverlay(reminder)
            if RNOTIF_SOUND in reminder.get("notification_types", []):
                _play_reminder_sound(
                    reminder.get("sound_key", "Exclamation"),
                    reminder.get("sound_custom_path", ""))

        elif level == 3:
            if RNOTIF_SOUND in reminder.get("notification_types", []):
                _play_reminder_sound(
                    reminder.get("sound_key", "Critical"),
                    reminder.get("sound_custom_path", ""))
            if self._win is not None:
                dlg = _ReminderAcknowledgeDialog(reminder, self._win)
                self._win._center_dialog(dlg)
                dlg.exec()
            else:
                dlg = _ReminderAcknowledgeDialog(reminder)
                dlg.exec()

        self.reminder_fired.emit(reminder)


_reminder_scheduler: Optional[ReminderScheduler] = None


# ── Reminder Wizard (create / edit) ──────────────────────────────────────────

class ReminderWizard(QDialog):
    """Create or edit a reminder.  Emits reminder_saved(dict) on save."""

    reminder_saved = Signal(dict)

    def __init__(self, existing: Optional[dict] = None, parent=None):
        super().__init__(parent, Qt.FramelessWindowHint | Qt.Dialog)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self._existing = existing
        self._drag_pos: Optional[QPoint] = None
        self._build_ui()
        if existing:
            self._populate(existing)
        self.setMinimumSize(540, 700)
        self.adjustSize()

    # ── builders ──────────────────────────────────────────────────────────────

    def _le(self, placeholder: str = "") -> QLineEdit:
        t = _theme
        w = QLineEdit()
        w.setPlaceholderText(placeholder)
        w.setFont(FONT_LABEL)
        w.setFixedHeight(30)
        w.setStyleSheet(f"""
            QLineEdit {{
                background:rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},200);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:5px; color:{t.TEXT_PRIMARY.name()}; padding:0 8px;
            }}
            QLineEdit:focus {{
                border-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},180);
            }}
        """)
        return w

    def _te(self) -> QTextEdit:
        t = _theme
        w = QTextEdit()
        w.setFont(FONT_LABEL)
        w.setFixedHeight(68)
        w.setStyleSheet(f"""
            QTextEdit {{
                background:rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},200);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:5px; color:{t.TEXT_PRIMARY.name()}; padding:4px 8px;
            }}
            QTextEdit:focus {{
                border-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},180);
            }}
        """)
        return w

    def _combo_style(self, cb: QComboBox):
        t = _theme
        cb.setStyleSheet(f"""
            QComboBox {{
                background:rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},200);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:5px; color:{t.TEXT_PRIMARY.name()};
                padding:0 8px; height:28px;
            }}
            QComboBox:hover {{ border-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},150); }}
            QComboBox::drop-down {{ border:none; width:20px; }}
            QComboBox QAbstractItemView {{
                background:{t.BG_DARK.name()};
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},90);
                color:{t.TEXT_PRIMARY.name()};
                selection-background-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
            }}
        """)

    def _section_label(self, text: str) -> QLabel:
        t = _theme
        lbl = QLabel(text)
        lbl.setFont(QFont("Segoe UI", 9, QFont.Weight.DemiBold))
        lbl.setStyleSheet(f"color:{t.GLOW.name()}; letter-spacing:0.6px;")
        return lbl

    def _row_label(self, text: str) -> QLabel:
        t = _theme
        lbl = QLabel(text)
        lbl.setFont(FONT_LABEL)
        lbl.setStyleSheet(f"color:{t.TEXT_SECONDARY.name()};")
        lbl.setFixedWidth(130)
        return lbl

    def _chk(self, text: str) -> QCheckBox:
        t = _theme
        cb = QCheckBox(text)
        cb.setFont(FONT_LABEL)
        cb.setStyleSheet(f"""
            QCheckBox {{ color:{t.TEXT_SECONDARY.name()}; spacing:6px; }}
            QCheckBox::indicator {{ width:14px; height:14px;
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},100);
                border-radius:3px; background:rgba(0,0,0,80); }}
            QCheckBox::indicator:checked {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},160);
                border-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},200); }}
        """)
        return cb

    def _spin(self, width: int = 60, height: int = 28) -> QSpinBox:
        """Return a styled QSpinBox matching the wizard theme."""
        t = _theme
        sp = QSpinBox()
        sp.setFont(FONT_LABEL)
        sp.setFixedSize(width, height)
        sp.setAlignment(Qt.AlignCenter)
        sp.setStyleSheet(f"""
            QSpinBox {{
                background:rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},200);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:5px; color:{t.TEXT_PRIMARY.name()}; padding:0 2px;
            }}
            QSpinBox:focus {{
                border-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},180);
            }}
            QSpinBox::up-button, QSpinBox::down-button {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},25);
                border:none; width:16px;
            }}
            QSpinBox::up-button:hover, QSpinBox::down-button:hover {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
            }}
        """)
        return sp

    def _build_ui(self):
        t = _theme
        outer = QVBoxLayout(self)
        outer.setContentsMargins(10, 10, 10, 10)
        outer.setSpacing(0)

        card = QWidget()
        card.setObjectName("rwiz_card")
        card.setStyleSheet(f"""
            QWidget#rwiz_card {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},252),
                    stop:1 rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},255));
                border:1.5px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},90);
                border-radius:14px;
            }}
        """)
        card.setGraphicsEffect(make_shadow(card, 32, QColor(0, 0, 0, 200)))
        cl = QVBoxLayout(card)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(0)

        # Title bar
        tbar = QWidget(); tbar.setFixedHeight(44)
        tbar.setStyleSheet("background:transparent;")
        tbl = QHBoxLayout(tbar); tbl.setContentsMargins(18, 0, 12, 0)
        ttl = QLabel("New Reminder" if not self._existing else "Edit Reminder")
        ttl.setFont(QFont("Segoe UI", 11, QFont.Weight.DemiBold))
        ttl.setStyleSheet(f"color:{t.TEXT_PRIMARY.name()};")
        tbl.addWidget(ttl); tbl.addStretch()
        close_btn = QPushButton("✕")
        close_btn.setFixedSize(28, 28)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setFont(QFont("Segoe UI", 10))
        close_btn.setStyleSheet(f"""
            QPushButton {{ background:transparent; border:none;
                color:{t.TEXT_DIM.name()}; border-radius:5px; }}
            QPushButton:hover {{ background:rgba(200,60,60,100);
                color:#ff8888; }}
        """)
        close_btn.clicked.connect(self.reject)
        tbl.addWidget(close_btn)
        tbar.mousePressEvent   = lambda e: setattr(self, "_drag_pos", e.globalPosition().toPoint())
        tbar.mouseMoveEvent    = lambda e: (
            self.move(self.pos() + (e.globalPosition().toPoint() - self._drag_pos)),
            setattr(self, "_drag_pos", e.globalPosition().toPoint())
        ) if self._drag_pos else None
        tbar.mouseReleaseEvent = lambda e: setattr(self, "_drag_pos", None)
        cl.addWidget(tbar)

        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},35);")
        cl.addWidget(sep)

        # Scroll area for the form
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("background:transparent;")
        form_w = QWidget(); form_w.setStyleSheet("background:transparent;")
        form = QVBoxLayout(form_w)
        form.setContentsMargins(22, 16, 22, 10)
        form.setSpacing(10)

        # ── Basic info ────────────────────────────────────────────────────
        form.addWidget(self._section_label("BASIC INFO"))

        row1 = QHBoxLayout()
        row1.addWidget(self._row_label("Title *"))
        self._title_edit = self._le("Submit Ticket Time")
        row1.addWidget(self._title_edit, 1)
        form.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(self._row_label("Message"))
        self._msg_edit = self._te()
        row2.addWidget(self._msg_edit, 1)
        form.addLayout(row2)

        # Tags
        row_tags = QHBoxLayout()
        row_tags.addWidget(self._row_label("Tags"))
        self._tags_edit = self._le("Comma-separated, e.g. work, health")
        row_tags.addWidget(self._tags_edit, 1)
        form.addLayout(row_tags)

        form.addWidget(_build_h_sep())

        # ── Alert level ───────────────────────────────────────────────────
        form.addWidget(self._section_label("ALERT LEVEL"))
        level_row = QHBoxLayout(); level_row.setSpacing(12)
        self._level_group = QButtonGroup(self)
        for val, label, desc in [
            (1, "Level 1 — Notify",    "Notification · Popup · Sound"),
            (2, "Level 2 — Full Screen", "All-monitor takeover"),
            (3, "Level 3 — Critical",  "Requires 'acknowledge' to dismiss"),
        ]:
            rb = QRadioButton(label)
            rb.setProperty("level", val)
            rb.setFont(QFont("Segoe UI", 9))
            rb.setStyleSheet(f"""
                QRadioButton {{ color:{t.TEXT_SECONDARY.name()}; spacing:6px; }}
                QRadioButton::indicator {{ width:14px; height:14px;
                    border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},100);
                    border-radius:7px; background:rgba(0,0,0,80); }}
                QRadioButton::indicator:checked {{
                    background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},180);
                    border-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},220); }}
            """)
            rb.setToolTip(desc)
            if val == 1:
                rb.setChecked(True)
            self._level_group.addButton(rb)
            level_row.addWidget(rb)
        level_row.addStretch()
        form.addLayout(level_row)

        form.addWidget(_build_h_sep())

        # ── Level-1 options ───────────────────────────────────────────────
        self._l1_box = QWidget()
        l1_lay = QVBoxLayout(self._l1_box)
        l1_lay.setContentsMargins(0, 0, 0, 0); l1_lay.setSpacing(8)
        l1_lay.addWidget(self._section_label("LEVEL 1 — NOTIFICATION TYPES"))

        notif_row = QHBoxLayout(); notif_row.setSpacing(16)
        self._chk_toast  = self._chk("Windows Notification")
        self._chk_popup  = self._chk("App Popup")
        self._chk_sound  = self._chk("Play Sound")
        self._chk_popup.setChecked(True)
        for c in (self._chk_toast, self._chk_popup, self._chk_sound):
            notif_row.addWidget(c)
        notif_row.addStretch()
        l1_lay.addLayout(notif_row)

        # Sound picker (shown when chk_sound is checked)
        self._sound_row = QWidget()
        sr_lay = QHBoxLayout(self._sound_row)
        sr_lay.setContentsMargins(0, 0, 0, 0); sr_lay.setSpacing(8)
        sr_lay.addWidget(self._row_label("Sound"))
        self._sound_combo = QComboBox()
        self._sound_combo.setFont(FONT_LABEL)
        self._sound_combo.setFixedHeight(28)
        for name in _BUILTIN_SOUNDS:
            self._sound_combo.addItem(name, name)
        self._sound_combo.addItem("Custom (.wav / .mp3)…", "custom")
        self._combo_style(self._sound_combo)
        self._sound_combo.currentIndexChanged.connect(self._on_sound_changed)
        sr_lay.addWidget(self._sound_combo)
        self._sound_browse_btn = QPushButton("Browse…")
        self._sound_browse_btn.setFixedHeight(28)
        self._sound_browse_btn.setCursor(Qt.PointingHandCursor)
        self._sound_browse_btn.setFont(QFont("Segoe UI", 8))
        self._sound_browse_btn.setStyleSheet(f"""
            QPushButton {{ background:transparent;
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:5px; color:{t.TEXT_SECONDARY.name()}; padding:0 10px; }}
            QPushButton:hover {{ color:{t.GLOW.name()};
                border-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},150); }}
        """)
        self._sound_browse_btn.clicked.connect(self._browse_sound)
        self._sound_browse_btn.setVisible(False)
        sr_lay.addWidget(self._sound_browse_btn)
        self._sound_path_edit = self._le("Path to .wav or .mp3 file")
        self._sound_path_edit.setVisible(False)
        sr_lay.addWidget(self._sound_path_edit, 1)
        l1_lay.addWidget(self._sound_row)
        self._chk_sound.toggled.connect(
            lambda checked: self._sound_row.setVisible(checked))
        self._sound_row.setVisible(False)

        # Volume
        vol_row = QHBoxLayout(); vol_row.setSpacing(8)
        vol_row.addWidget(self._row_label("Volume"))
        self._vol_slider = QSlider(Qt.Horizontal)
        self._vol_slider.setRange(0, 100); self._vol_slider.setValue(80)
        self._vol_slider.setFixedHeight(20)
        self._vol_slider.setStyleSheet(f"""
            QSlider::groove:horizontal {{ height:4px;
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},40);
                border-radius:2px; }}
            QSlider::handle:horizontal {{ width:14px; height:14px; margin:-5px 0;
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},180);
                border-radius:7px; }}
            QSlider::sub-page:horizontal {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},130);
                border-radius:2px; }}
        """)
        vol_row.addWidget(self._vol_slider, 1)
        self._vol_lbl = QLabel("80%")
        self._vol_lbl.setFont(FONT_SMALL)
        self._vol_lbl.setStyleSheet(f"color:{t.TEXT_SECONDARY.name()};")
        self._vol_lbl.setFixedWidth(36)
        self._vol_slider.valueChanged.connect(
            lambda v: self._vol_lbl.setText(f"{v}%"))
        vol_row.addWidget(self._vol_lbl)
        l1_lay.addLayout(vol_row)

        # Snooze
        snz_row = QHBoxLayout(); snz_row.setSpacing(8)
        snz_row.addWidget(self._row_label("Snooze duration"))
        self._snooze_spin_w = self._le("5")
        self._snooze_spin_w.setMaximumWidth(70)
        snz_row.addWidget(self._snooze_spin_w)
        snz_lbl = QLabel("minutes  (0 = no snooze)")
        snz_lbl.setFont(FONT_SMALL)
        snz_lbl.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        snz_row.addWidget(snz_lbl)
        snz_row.addStretch()
        l1_lay.addLayout(snz_row)

        form.addWidget(self._l1_box)

        # ── Level-2 options ───────────────────────────────────────────────
        self._l2_box = QWidget()
        l2_lay = QVBoxLayout(self._l2_box)
        l2_lay.setContentsMargins(0, 0, 0, 0); l2_lay.setSpacing(8)
        l2_lay.addWidget(self._section_label("LEVEL 2 — FULL SCREEN OPTIONS"))

        to_row = QHBoxLayout(); to_row.setSpacing(8)
        to_row.addWidget(self._row_label("Auto-dismiss after"))
        self._fs_timeout_edit = self._le("10")
        self._fs_timeout_edit.setMaximumWidth(70)
        to_row.addWidget(self._fs_timeout_edit)
        to_lbl2 = QLabel("seconds  (min 1)")
        to_lbl2.setFont(FONT_SMALL)
        to_lbl2.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        to_row.addWidget(to_lbl2); to_row.addStretch()
        l2_lay.addLayout(to_row)

        col_row = QHBoxLayout(); col_row.setSpacing(12)
        col_row.addWidget(self._row_label("Background colour"))
        self._fs_bg_btn = QPushButton()
        self._fs_bg_btn.setFixedSize(52, 26)
        self._fs_bg_color = "#12001a"
        self._fs_bg_btn.setCursor(Qt.PointingHandCursor)
        self._update_color_btn(self._fs_bg_btn, self._fs_bg_color)
        self._fs_bg_btn.clicked.connect(lambda: self._pick_color("bg"))
        col_row.addWidget(self._fs_bg_btn)
        col_row.addWidget(self._row_label("Text colour"))
        self._fs_txt_btn = QPushButton()
        self._fs_txt_btn.setFixedSize(52, 26)
        self._fs_txt_color = "#ff4488"
        self._fs_txt_btn.setCursor(Qt.PointingHandCursor)
        self._update_color_btn(self._fs_txt_btn, self._fs_txt_color)
        self._fs_txt_btn.clicked.connect(lambda: self._pick_color("txt"))
        col_row.addWidget(self._fs_txt_btn)
        col_row.addStretch()
        l2_lay.addLayout(col_row)

        l2_snd_row = QHBoxLayout(); l2_snd_row.setSpacing(8)
        self._l2_chk_sound = self._chk("Play sound on Level-2 trigger")
        l2_snd_row.addWidget(self._l2_chk_sound); l2_snd_row.addStretch()
        l2_lay.addLayout(l2_snd_row)

        form.addWidget(self._l2_box)
        self._l2_box.setVisible(False)

        # ── Level-3 options ───────────────────────────────────────────────
        self._l3_box = QWidget()
        l3_lay = QVBoxLayout(self._l3_box)
        l3_lay.setContentsMargins(0, 0, 0, 0); l3_lay.setSpacing(8)
        l3_lay.addWidget(self._section_label("LEVEL 3 — CRITICAL OPTIONS"))

        l3_snd_row = QHBoxLayout()
        self._l3_chk_sound = self._chk("Play sound on Level-3 trigger")
        l3_snd_row.addWidget(self._l3_chk_sound); l3_snd_row.addStretch()
        l3_lay.addLayout(l3_snd_row)

        l3_note = QLabel(
            "⚠  Level 3 reminders require the user to type \"acknowledge\" before they "
            "can be dismissed. Use sparingly.")
        l3_note.setFont(QFont("Segoe UI", 8))
        l3_note.setWordWrap(True)
        l3_note.setStyleSheet("color:rgba(255,120,120,180);")
        l3_lay.addWidget(l3_note)

        form.addWidget(self._l3_box)
        self._l3_box.setVisible(False)

        # Connect level radio buttons
        for rb in self._level_group.buttons():
            rb.toggled.connect(self._on_level_changed)

        form.addWidget(_build_h_sep())

        # ── Schedule (required) ────────────────────────────────────────────
        form.addWidget(self._section_label("SCHEDULE"))

        self._sched_box = QWidget()
        sb_lay = QVBoxLayout(self._sched_box)
        sb_lay.setContentsMargins(0, 4, 0, 0); sb_lay.setSpacing(10)

        # Schedule type
        type_row = QHBoxLayout(); type_row.setSpacing(8)
        type_row.addWidget(self._row_label("Schedule type"))
        self._sched_type_combo = QComboBox()
        self._sched_type_combo.setFont(FONT_LABEL)
        self._sched_type_combo.setFixedHeight(28)
        for key, label in SCHED_TYPES:
            self._sched_type_combo.addItem(label, key)
        self._combo_style(self._sched_type_combo)
        self._sched_type_combo.currentIndexChanged.connect(self._refresh_sched_ui)
        type_row.addWidget(self._sched_type_combo, 1)
        sb_lay.addLayout(type_row)

        # ── Time picker: hour / min / AM-PM spin wheels ──────────────────
        self._sched_time_row_w = QWidget()
        tp_lay = QHBoxLayout(self._sched_time_row_w)
        tp_lay.setContentsMargins(0, 0, 0, 0); tp_lay.setSpacing(6)
        tp_lay.addWidget(self._row_label("Time"))
        self._sched_hour_spin = self._spin(52, 36)
        self._sched_hour_spin.setRange(1, 12)
        self._sched_hour_spin.setValue(9)
        self._sched_hour_spin.setWrapping(True)
        self._sched_hour_spin.setToolTip("Hour (1–12)")
        tp_lay.addWidget(self._sched_hour_spin)
        _colon_lbl = QLabel(":")
        _colon_lbl.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        _colon_lbl.setStyleSheet(f"color:{t.TEXT_SECONDARY.name()};")
        _colon_lbl.setFixedWidth(10)
        _colon_lbl.setAlignment(Qt.AlignCenter)
        tp_lay.addWidget(_colon_lbl)
        self._sched_min_spin = self._spin(52, 36)
        self._sched_min_spin.setRange(0, 59)
        self._sched_min_spin.setValue(0)
        self._sched_min_spin.setWrapping(True)
        self._sched_min_spin.setToolTip("Minute (0–59)")
        tp_lay.addWidget(self._sched_min_spin)
        self._sched_ampm_btn = QPushButton("AM")
        self._sched_ampm_btn.setFixedSize(46, 36)
        self._sched_ampm_btn.setCursor(Qt.PointingHandCursor)
        self._sched_ampm_btn.setCheckable(True)
        self._sched_ampm_btn.setFont(QFont("Segoe UI", 9, QFont.Weight.DemiBold))
        self._on_ampm_toggle(False)  # initial AM style
        self._sched_ampm_btn.clicked.connect(self._on_ampm_toggle)
        tp_lay.addWidget(self._sched_ampm_btn)
        tp_lay.addStretch()
        sb_lay.addWidget(self._sched_time_row_w)

        # ── Calendar picker (Once) ────────────────────────────────────────
        self._sched_date_row = QWidget()
        cal_vlay = QVBoxLayout(self._sched_date_row)
        cal_vlay.setContentsMargins(0, 4, 0, 0); cal_vlay.setSpacing(4)
        _cl_hdr = QHBoxLayout()
        _cl_hdr.addWidget(self._row_label("Date"))
        _cl_hdr.addStretch()
        cal_vlay.addLayout(_cl_hdr)
        self._sched_calendar = QCalendarWidget()
        self._sched_calendar.setMinimumDate(QDate.currentDate())
        self._sched_calendar.setGridVisible(True)
        self._sched_calendar.setMaximumHeight(220)
        _cal_t = _theme
        self._sched_calendar.setStyleSheet(f"""
            QCalendarWidget {{
                background:{_cal_t.BG_DARK.name()};
                border:1px solid rgba({_cal_t.GLOW.red()},{_cal_t.GLOW.green()},{_cal_t.GLOW.blue()},60);
                border-radius:8px;
            }}
            QCalendarWidget QWidget#qt_calendar_navigationbar {{
                background:rgba({_cal_t.BG_MID.red()},{_cal_t.BG_MID.green()},{_cal_t.BG_MID.blue()},220);
                border-bottom:1px solid rgba({_cal_t.GLOW.red()},{_cal_t.GLOW.green()},{_cal_t.GLOW.blue()},40);
            }}
            QCalendarWidget QToolButton {{
                color:{_cal_t.TEXT_PRIMARY.name()};
                background:transparent; border:none; border-radius:4px;
                padding:4px 8px; font-size:10pt;
            }}
            QCalendarWidget QToolButton:hover {{
                background:rgba({_cal_t.GLOW.red()},{_cal_t.GLOW.green()},{_cal_t.GLOW.blue()},40);
                color:{_cal_t.GLOW.name()};
            }}
            QCalendarWidget QMenu {{
                background:{_cal_t.BG_DARK.name()}; color:{_cal_t.TEXT_PRIMARY.name()};
                border:1px solid rgba({_cal_t.GLOW.red()},{_cal_t.GLOW.green()},{_cal_t.GLOW.blue()},80);
            }}
            QCalendarWidget QSpinBox {{
                background:{_cal_t.BG_MID.name()}; color:{_cal_t.TEXT_PRIMARY.name()};
                border:1px solid rgba({_cal_t.GLOW.red()},{_cal_t.GLOW.green()},{_cal_t.GLOW.blue()},60);
                border-radius:4px;
            }}
            QCalendarWidget QAbstractItemView:enabled {{
                color:{_cal_t.TEXT_PRIMARY.name()};
                background:{_cal_t.BG_DARK.name()};
                selection-background-color:rgba({_cal_t.GLOW.red()},{_cal_t.GLOW.green()},{_cal_t.GLOW.blue()},140);
                selection-color:{_cal_t.BG_DARK.name()};
                font-size:9pt;
            }}
            QCalendarWidget QAbstractItemView:disabled {{
                color:rgba(120,120,120,120);
            }}
        """)
        cal_vlay.addWidget(self._sched_calendar)
        sb_lay.addWidget(self._sched_date_row)

        # ── Weekly day checkboxes ─────────────────────────────────────────
        self._sched_days_row = QWidget()
        days_lay = QHBoxLayout(self._sched_days_row)
        days_lay.setContentsMargins(0, 0, 0, 0); days_lay.setSpacing(4)
        days_lay.addWidget(self._row_label("Days"))
        self._day_checks: list[QCheckBox] = []
        for idx, day_name in enumerate(["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]):
            dc = QCheckBox(day_name)
            dc.setProperty("day_idx", idx)
            dc.setFont(QFont("Segoe UI", 8))
            t2 = _theme
            dc.setStyleSheet(f"""
                QCheckBox {{ color:{t2.TEXT_SECONDARY.name()}; spacing:4px; }}
                QCheckBox::indicator {{ width:12px; height:12px;
                    border:1px solid rgba({t2.GLOW.red()},{t2.GLOW.green()},{t2.GLOW.blue()},100);
                    border-radius:2px; background:rgba(0,0,0,80); }}
                QCheckBox::indicator:checked {{
                    background:rgba({t2.GLOW.red()},{t2.GLOW.green()},{t2.GLOW.blue()},160); }}
            """)
            self._day_checks.append(dc)
            days_lay.addWidget(dc)
        days_lay.addStretch()
        sb_lay.addWidget(self._sched_days_row)

        # ── Monthly day-of-month spinbox ──────────────────────────────────
        self._sched_dom_row = QWidget()
        dom_row = QHBoxLayout(self._sched_dom_row)
        dom_row.setContentsMargins(0, 0, 0, 0); dom_row.setSpacing(8)
        dom_row.addWidget(self._row_label("Day of month"))
        self._sched_dom_spin = self._spin(70, 28)
        self._sched_dom_spin.setRange(1, 31)
        self._sched_dom_spin.setValue(1)
        self._sched_dom_spin.setToolTip("Day of month (1–31)")
        dom_row.addWidget(self._sched_dom_spin)
        dom_row.addStretch()
        sb_lay.addWidget(self._sched_dom_row)

        # ── Interval spinbox ──────────────────────────────────────────────
        self._sched_interval_row = QWidget()
        int_lay = QHBoxLayout(self._sched_interval_row)
        int_lay.setContentsMargins(0, 0, 0, 0); int_lay.setSpacing(8)
        int_lay.addWidget(self._row_label("Every"))
        self._sched_int_spin = self._spin(80, 28)
        self._sched_int_spin.setRange(1, 9999)
        self._sched_int_spin.setValue(30)
        self._sched_int_spin.setToolTip("Interval value")
        int_lay.addWidget(self._sched_int_spin)
        self._sched_int_unit = QComboBox()
        self._sched_int_unit.setFont(FONT_LABEL)
        self._sched_int_unit.setFixedHeight(28)
        self._sched_int_unit.addItem("minutes", "minutes")
        self._sched_int_unit.addItem("hours",   "hours")
        self._combo_style(self._sched_int_unit)
        int_lay.addWidget(self._sched_int_unit); int_lay.addStretch()
        sb_lay.addWidget(self._sched_interval_row)

        form.addWidget(self._sched_box)
        self._refresh_sched_ui()

        form.addWidget(_build_h_sep())

        # ── Additional options ────────────────────────────────────────────
        form.addWidget(self._section_label("ADDITIONAL OPTIONS"))

        rep_row = QHBoxLayout(); rep_row.setSpacing(8)
        rep_row.addWidget(self._row_label("Repeat limit"))
        self._repeat_edit = self._le("0")
        self._repeat_edit.setMaximumWidth(70)
        rep_row.addWidget(self._repeat_edit)
        rep_lbl = QLabel("times  (0 = unlimited)")
        rep_lbl.setFont(FONT_SMALL)
        rep_lbl.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        rep_row.addWidget(rep_lbl); rep_row.addStretch()
        form.addLayout(rep_row)

        col_accent_row = QHBoxLayout(); col_accent_row.setSpacing(8)
        col_accent_row.addWidget(self._row_label("Tag colour"))
        self._tag_color_btn = QPushButton("Default")
        self._tag_color_btn.setFixedSize(120, 28)
        self._tag_color_btn.setCursor(Qt.PointingHandCursor)
        self._tag_color_val = ""
        self._tag_color_btn.setFont(QFont("Segoe UI", 8))
        self._tag_color_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},30);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},80);
                border-radius:5px; color:{t.TEXT_SECONDARY.name()}; }}
            QPushButton:hover {{ color:{t.GLOW.name()}; }}
        """)
        self._tag_color_btn.clicked.connect(self._pick_tag_color)
        col_accent_row.addWidget(self._tag_color_btn); col_accent_row.addStretch()
        form.addLayout(col_accent_row)

        priority_row = QHBoxLayout(); priority_row.setSpacing(8)
        priority_row.addWidget(self._row_label("Priority"))
        self._priority_combo = QComboBox()
        self._priority_combo.setFont(FONT_LABEL)
        self._priority_combo.setFixedHeight(28)
        for label, key in [("Normal","normal"),("High","high"),("Urgent","urgent")]:
            self._priority_combo.addItem(label, key)
        self._combo_style(self._priority_combo)
        priority_row.addWidget(self._priority_combo); priority_row.addStretch()
        form.addLayout(priority_row)

        enabled_row = QHBoxLayout()
        self._enabled_chk = self._chk("Reminder enabled")
        self._enabled_chk.setChecked(True)
        enabled_row.addWidget(self._enabled_chk); enabled_row.addStretch()
        form.addLayout(enabled_row)

        form.addStretch()
        scroll.setWidget(form_w)
        cl.addWidget(scroll)

        # Save / Cancel buttons
        btn_bar = QWidget(); btn_bar.setFixedHeight(52)
        btn_bar.setStyleSheet("background:transparent;")
        bbl = QHBoxLayout(btn_bar); bbl.setContentsMargins(18, 8, 18, 8); bbl.setSpacing(10)
        bbl.addStretch()
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setFixedHeight(32); cancel_btn.setCursor(Qt.PointingHandCursor)
        cancel_btn.setFont(QFont("Segoe UI", 9))
        cancel_btn.setStyleSheet(f"""
            QPushButton {{ background:transparent;
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:6px; color:{t.TEXT_SECONDARY.name()}; padding:0 18px; }}
            QPushButton:hover {{ color:{t.GLOW.name()}; }}
        """)
        cancel_btn.clicked.connect(self.reject)
        bbl.addWidget(cancel_btn)
        save_btn = QPushButton("Save Reminder")
        save_btn.setFixedHeight(32); save_btn.setCursor(Qt.PointingHandCursor)
        save_btn.setFont(QFont("Segoe UI", 9, QFont.Weight.DemiBold))
        save_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},140);
                border:none; border-radius:6px; color:{t.BG_DARK.name()};
                padding:0 22px; font-weight:600; }}
            QPushButton:hover {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},200); }}
        """)
        save_btn.clicked.connect(self._on_save)
        bbl.addWidget(save_btn)
        cl.addWidget(btn_bar)

        outer.addWidget(card)

    # ── helpers ───────────────────────────────────────────────────────────────

    def _update_color_btn(self, btn: QPushButton, hex_color: str):
        try:
            c = QColor(hex_color)
        except Exception:
            c = QColor("#222222")
        btn.setStyleSheet(f"""
            QPushButton {{
                background:{c.name()};
                border:1px solid rgba(255,255,255,80);
                border-radius:4px;
            }}
        """)

    def _pick_color(self, which: str):
        cur = self._fs_bg_color if which == "bg" else self._fs_txt_color
        try:
            color = QColorDialog.getColor(QColor(cur), self)
        except Exception:
            return
        if not color.isValid():
            return
        if which == "bg":
            self._fs_bg_color = color.name()
            self._update_color_btn(self._fs_bg_btn, self._fs_bg_color)
        else:
            self._fs_txt_color = color.name()
            self._update_color_btn(self._fs_txt_btn, self._fs_txt_color)

    def _pick_tag_color(self):
        try:
            cur = QColor(self._tag_color_val) if self._tag_color_val else QColor("#00c8ff")
            color = QColorDialog.getColor(cur, self)
            if color.isValid():
                self._tag_color_val = color.name()
                self._tag_color_btn.setText(color.name())
                self._tag_color_btn.setStyleSheet(f"""
                    QPushButton {{
                        background:{color.name()};
                        border:1px solid rgba(255,255,255,80); border-radius:5px;
                        color:{_theme.BG_DARK.name()}; font-weight:600; }}
                """)
        except Exception as exc:
            print(f"[ReminderWizard] color pick error: {exc}", file=sys.stderr)

    def _on_level_changed(self):
        lv = self._current_level()
        self._l1_box.setVisible(lv == 1)
        self._l2_box.setVisible(lv == 2)
        self._l3_box.setVisible(lv == 3)

    def _current_level(self) -> int:
        for rb in self._level_group.buttons():
            if rb.isChecked():
                return int(rb.property("level"))
        return 1

    def _on_sound_changed(self):
        is_custom = self._sound_combo.currentData() == "custom"
        self._sound_browse_btn.setVisible(is_custom)
        self._sound_path_edit.setVisible(is_custom)

    def _browse_sound(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Sound File", "",
            "Audio Files (*.wav *.mp3 *.ogg *.aac *.m4a);;All Files (*)")
        if path:
            self._sound_path_edit.setText(path)

    def _refresh_sched_ui(self):
        stype = self._sched_type_combo.currentData() or SCHED_DAILY
        has_time     = stype in (SCHED_DAILY, SCHED_WORKDAYS, SCHED_WEEKENDS,
                                  SCHED_WEEKLY, SCHED_MONTHLY, SCHED_ONCE)
        has_days     = stype == SCHED_WEEKLY
        has_dom      = stype == SCHED_MONTHLY
        has_date     = stype == SCHED_ONCE
        has_interval = stype == SCHED_INTERVAL
        self._sched_time_row_w.setVisible(has_time)
        self._sched_days_row.setVisible(has_days)
        self._sched_dom_row.setVisible(has_dom)
        self._sched_date_row.setVisible(has_date)
        self._sched_interval_row.setVisible(has_interval)

    def _on_ampm_toggle(self, checked: bool):
        """Update AM/PM button label and colour."""
        t = _theme
        self._sched_ampm_btn.setText("PM" if checked else "AM")
        if checked:  # PM
            self._sched_ampm_btn.setStyleSheet(f"""
                QPushButton {{
                    background:rgba({t.ACCENT_AMBER.red()},{t.ACCENT_AMBER.green()},{t.ACCENT_AMBER.blue()},180);
                    border:none; border-radius:5px;
                    color:{t.BG_DARK.name()}; font-weight:700;
                }}
                QPushButton:hover {{
                    background:rgba({t.ACCENT_AMBER.red()},{t.ACCENT_AMBER.green()},{t.ACCENT_AMBER.blue()},220);
                }}
            """)
        else:  # AM
            self._sched_ampm_btn.setStyleSheet(f"""
                QPushButton {{
                    background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},180);
                    border:none; border-radius:5px;
                    color:{t.BG_DARK.name()}; font-weight:700;
                }}
                QPushButton:hover {{
                    background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},220);
                }}
            """)

    def _get_schedule_dict(self) -> dict:
        stype = self._sched_type_combo.currentData() or SCHED_DAILY
        sched: dict = {"enabled": True, "type": stype}
        if stype != SCHED_INTERVAL:
            # Convert 12-hour spinbox + AM/PM to 24-hour HH:MM
            h = self._sched_hour_spin.value()
            m = self._sched_min_spin.value()
            is_pm = self._sched_ampm_btn.isChecked()
            if is_pm and h != 12:
                h += 12
            elif not is_pm and h == 12:
                h = 0
            sched["time"] = f"{h:02d}:{m:02d}"
        if stype == SCHED_WEEKLY:
            sched["days"] = [dc.property("day_idx")
                             for dc in self._day_checks if dc.isChecked()]
        elif stype == SCHED_MONTHLY:
            sched["day_of_month"] = self._sched_dom_spin.value()
        elif stype == SCHED_ONCE:
            qd = self._sched_calendar.selectedDate()
            sched["date"] = qd.toString("yyyy-MM-dd")
        elif stype == SCHED_INTERVAL:
            sched["interval_value"] = self._sched_int_spin.value()
            sched["interval_unit"] = (
                self._sched_int_unit.currentData() or "minutes")
        sched.setdefault("last_run", None)
        return sched

    def _populate_schedule(self, reminder: dict):
        sched = reminder.get("schedule") or {}
        if not isinstance(sched, dict):
            sched = {}
        stype = sched.get("type", SCHED_DAILY)
        for i in range(self._sched_type_combo.count()):
            if self._sched_type_combo.itemData(i) == stype:
                self._sched_type_combo.setCurrentIndex(i)
                break
        # Parse 24-hour "HH:MM" into hour spinbox + AM/PM button
        try:
            time_str = sched.get("time", "09:00")
            hh, mm = map(int, time_str.split(":"))
            is_pm = hh >= 12
            h12 = hh % 12 or 12
            self._sched_hour_spin.setValue(h12)
            self._sched_min_spin.setValue(mm)
            self._sched_ampm_btn.setChecked(is_pm)
            self._on_ampm_toggle(is_pm)
        except Exception:
            pass
        for dc in self._day_checks:
            dc.setChecked(dc.property("day_idx") in sched.get("days", []))
        try:
            self._sched_dom_spin.setValue(
                max(1, min(31, int(sched.get("day_of_month", 1)))))
        except Exception:
            pass
        try:
            date_str = sched.get("date", "")
            if date_str:
                qd = QDate.fromString(date_str, "yyyy-MM-dd")
                if qd.isValid():
                    self._sched_calendar.setSelectedDate(qd)
        except Exception:
            pass
        try:
            self._sched_int_spin.setValue(
                max(1, int(sched.get("interval_value", 30))))
        except Exception:
            pass
        unit = sched.get("interval_unit", "minutes")
        for i in range(self._sched_int_unit.count()):
            if self._sched_int_unit.itemData(i) == unit:
                self._sched_int_unit.setCurrentIndex(i); break
        self._refresh_sched_ui()

    # ── populate ──────────────────────────────────────────────────────────────

    def _populate(self, r: dict):
        self._title_edit.setText(r.get("title", ""))
        self._msg_edit.setPlainText(r.get("message", ""))
        self._tags_edit.setText(", ".join(r.get("tags", [])))
        self._enabled_chk.setChecked(r.get("enabled", True))

        lv = int(r.get("level", 1))
        for rb in self._level_group.buttons():
            if int(rb.property("level")) == lv:
                rb.setChecked(True); break

        notif_types = r.get("notification_types", [RNOTIF_POPUP])
        self._chk_toast.setChecked(RNOTIF_WIN_TOAST in notif_types)
        self._chk_popup.setChecked(RNOTIF_POPUP     in notif_types)
        self._chk_sound.setChecked(RNOTIF_SOUND     in notif_types)
        if RNOTIF_SOUND in notif_types:
            self._sound_row.setVisible(True)
        sound_key = r.get("sound_key", "Default Beep")
        for i in range(self._sound_combo.count()):
            if self._sound_combo.itemData(i) == sound_key:
                self._sound_combo.setCurrentIndex(i); break
        self._sound_path_edit.setText(r.get("sound_custom_path", ""))
        self._vol_slider.setValue(int(r.get("sound_volume", 80)))
        try:
            self._snooze_spin_w.setText(str(int(r.get("snooze_minutes", 5))))
        except Exception:
            self._snooze_spin_w.setText("5")

        self._fs_timeout_edit.setText(str(r.get("fullscreen_timeout", 10)))
        self._fs_bg_color  = r.get("fullscreen_bg_color",  "#12001a")
        self._fs_txt_color = r.get("fullscreen_text_color","#ff4488")
        self._update_color_btn(self._fs_bg_btn,  self._fs_bg_color)
        self._update_color_btn(self._fs_txt_btn, self._fs_txt_color)
        self._l2_chk_sound.setChecked(r.get("l2_play_sound", False))
        self._l3_chk_sound.setChecked(r.get("l3_play_sound", False))

        try:
            self._repeat_edit.setText(str(int(r.get("repeat_limit", 0))))
        except Exception:
            self._repeat_edit.setText("0")
        for i in range(self._priority_combo.count()):
            if self._priority_combo.itemData(i) == r.get("priority", "normal"):
                self._priority_combo.setCurrentIndex(i); break
        tag_c = r.get("tag_color", "")
        if tag_c:
            self._tag_color_val = tag_c
            self._tag_color_btn.setText(tag_c)
            try:
                self._tag_color_btn.setStyleSheet(f"""
                    QPushButton {{
                        background:{tag_c};
                        border:1px solid rgba(255,255,255,80); border-radius:5px;
                        color:{_theme.BG_DARK.name()}; font-weight:600; }}
                """)
            except Exception:
                pass

        self._populate_schedule(r)
        self._on_level_changed()

    # ── save ──────────────────────────────────────────────────────────────────

    def _on_save(self):
        title = self._title_edit.text().strip()
        if not title:
            self._title_edit.setFocus()
            QMessageBox.warning(self, "Validation", "Please enter a reminder title.")
            return
        lv = self._current_level()
        notif_types = []
        if lv == 1:
            if self._chk_toast.isChecked():  notif_types.append(RNOTIF_WIN_TOAST)
            if self._chk_popup.isChecked():  notif_types.append(RNOTIF_POPUP)
            if self._chk_sound.isChecked():  notif_types.append(RNOTIF_SOUND)
            if not notif_types:
                QMessageBox.warning(self, "Validation",
                    "Please select at least one notification type for Level 1.")
                return
        elif lv == 2:
            if self._l2_chk_sound.isChecked(): notif_types.append(RNOTIF_SOUND)
            try:
                to = int(self._fs_timeout_edit.text().strip())
                if to < 1:
                    raise ValueError
            except ValueError:
                QMessageBox.warning(self, "Validation",
                    "Auto-dismiss timeout must be a whole number ≥ 1.")
                return
        elif lv == 3:
            if self._l3_chk_sound.isChecked(): notif_types.append(RNOTIF_SOUND)

        try:
            snooze_min = max(0, int(self._snooze_spin_w.text().strip() or "0"))
        except ValueError:
            snooze_min = 5
        try:
            repeat_limit = max(0, int(self._repeat_edit.text().strip() or "0"))
        except ValueError:
            repeat_limit = 0

        raw_tags = self._tags_edit.text().strip()
        tags = [t.strip() for t in raw_tags.split(",") if t.strip()] if raw_tags else []

        reminder: dict = {
            "title":               title,
            "message":             self._msg_edit.toPlainText().strip(),
            "tags":                tags,
            "enabled":             self._enabled_chk.isChecked(),
            "level":               lv,
            "notification_types":  notif_types,
            "sound_key":           self._sound_combo.currentData() or "Default Beep",
            "sound_custom_path":   self._sound_path_edit.text().strip(),
            "sound_volume":        self._vol_slider.value(),
            "snooze_minutes":      snooze_min,
            "fullscreen_timeout":  max(1, int(self._fs_timeout_edit.text().strip() or "10")),
            "fullscreen_bg_color": self._fs_bg_color,
            "fullscreen_text_color": self._fs_txt_color,
            "l2_play_sound":       self._l2_chk_sound.isChecked(),
            "l3_play_sound":       self._l3_chk_sound.isChecked(),
            "repeat_limit":        repeat_limit,
            "priority":            self._priority_combo.currentData() or "normal",
            "tag_color":           self._tag_color_val,
            "schedule":            self._get_schedule_dict(),
            "fire_count":          self._existing.get("fire_count", 0) if self._existing else 0,
        }
        if self._existing:
            reminder["id"]      = self._existing.get("id", "")
            reminder["created"] = self._existing.get("created", "")
        self.reminder_saved.emit(reminder)
        self.accept()

    def paintEvent(self, e):
        p = QPainter(self)
        p.fillRect(self.rect(), Qt.transparent)
        p.end()


def _build_h_sep() -> QFrame:
    sep = QFrame(); sep.setFrameShape(QFrame.HLine)
    sep.setStyleSheet(f"color:rgba({_theme.GLOW.red()},{_theme.GLOW.green()},{_theme.GLOW.blue()},30);")
    return sep


# ── Reminder Row Widget (inside ReminderWindow list) ─────────────────────────

class _ReminderRow(QWidget):
    """A single row in the reminder list — shows title, level badge, schedule, status."""

    edit_requested   = Signal(str)    # reminder id
    delete_requested = Signal(str)
    test_requested   = Signal(str)
    toggle_requested = Signal(str, bool)

    _LEVEL_COLORS = {1: "#ffa020", 2: "#20aaff", 3: "#ff3333"}
    _LEVEL_LABELS = {1: "L1", 2: "L2", 3: "L3"}
    _PRIORITY_ICONS = {"high": "🔺", "urgent": "🚨", "normal": ""}

    def __init__(self, reminder: dict, parent=None):
        super().__init__(parent)
        self._rid      = reminder.get("id", "")
        self._reminder = reminder
        self.setFixedHeight(56)
        self._build_ui()
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_menu)

    def _build_ui(self):
        t = _theme
        lay = QHBoxLayout(self)
        lay.setContentsMargins(12, 6, 12, 6)
        lay.setSpacing(10)

        # Level badge
        lv = int(self._reminder.get("level", 1))
        lv_color = self._LEVEL_COLORS.get(lv, "#aaaaaa")
        lv_lbl = QLabel(self._LEVEL_LABELS.get(lv, "L1"))
        lv_lbl.setFont(QFont("Segoe UI", 8, QFont.Weight.Bold))
        lv_lbl.setFixedSize(26, 26)
        lv_lbl.setAlignment(Qt.AlignCenter)
        lv_lbl.setStyleSheet(f"""
            color: {lv_color};
            border: 1.5px solid {lv_color};
            border-radius: 5px;
            background: transparent;
        """)
        lay.addWidget(lv_lbl)

        # Priority icon
        pri = self._reminder.get("priority", "normal")
        pri_icon = self._PRIORITY_ICONS.get(pri, "")
        if pri_icon:
            pi_lbl = QLabel(pri_icon)
            pi_lbl.setFont(QFont("Segoe UI Emoji", 10))
            pi_lbl.setFixedWidth(20)
            lay.addWidget(pi_lbl)

        # Title + schedule info
        info = QVBoxLayout(); info.setSpacing(1); info.setContentsMargins(0, 0, 0, 0)
        title_str = self._reminder.get("title", "(untitled)")
        tag_c = self._reminder.get("tag_color", "")
        title_color = tag_c if tag_c else t.TEXT_PRIMARY.name()
        title_lbl = QLabel(title_str)
        title_lbl.setFont(QFont("Segoe UI", 10, QFont.Weight.DemiBold))
        title_lbl.setStyleSheet(f"color:{title_color};")
        title_lbl.setMaximumWidth(280)
        fm = QFontMetrics(title_lbl.font())
        title_lbl.setText(fm.elidedText(title_str, Qt.ElideRight, 280))
        info.addWidget(title_lbl)

        sched = self._reminder.get("schedule") or {}
        if sched.get("enabled"):
            stype = sched.get("type", "")
            stime = sched.get("time", "")
            if stype == SCHED_INTERVAL:
                iv = sched.get("interval_value", "?")
                iu = sched.get("interval_unit", "min")
                sub = f"Every {iv} {iu}"
            elif stype == SCHED_ONCE:
                sub = f"Once on {sched.get('date','')} {stime}"
            else:
                sub = f"{stype.title()} at {stime}"
            sub_lbl = QLabel(f"⏰ {sub}")
            sub_lbl.setFont(QFont("Segoe UI", 8))
            sub_lbl.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
            info.addWidget(sub_lbl)
        else:
            last = self._reminder.get("last_fired")
            if last:
                try:
                    ldt = datetime.fromisoformat(last)
                    sub_lbl = QLabel(f"Last fired: {ldt.strftime('%Y-%m-%d %H:%M')}")
                    sub_lbl.setFont(QFont("Segoe UI", 8))
                    sub_lbl.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
                    info.addWidget(sub_lbl)
                except Exception:
                    pass

        lay.addLayout(info, 1)

        # Tags
        tags = self._reminder.get("tags", [])
        if tags:
            tag_lbl = QLabel(", ".join(tags[:3]))
            tag_lbl.setFont(QFont("Segoe UI", 8))
            tag_lbl.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
            tag_lbl.setMaximumWidth(120)
            lay.addWidget(tag_lbl)

        # Enable toggle
        self._toggle = QCheckBox()
        self._toggle.setChecked(self._reminder.get("enabled", True))
        self._toggle.setToolTip("Enable / Disable reminder")
        self._toggle.setStyleSheet(f"""
            QCheckBox {{ spacing:0; }}
            QCheckBox::indicator {{ width:22px; height:22px; border-radius:11px;
                border:2px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},100); }}
            QCheckBox::indicator:unchecked {{
                background:rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},180); }}
            QCheckBox::indicator:checked {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},200);
                border-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},220); }}
        """)
        self._toggle.toggled.connect(
            lambda checked: self.toggle_requested.emit(self._rid, checked))
        lay.addWidget(self._toggle)

        # Action buttons
        for icon, tip, sig in [
            ("▶", "Test now", self.test_requested),
            ("✏", "Edit",     self.edit_requested),
            ("🗑", "Delete",   self.delete_requested),
        ]:
            btn = QPushButton(icon)
            btn.setFont(QFont("Segoe UI Emoji", 9))
            btn.setFixedSize(26, 26)
            btn.setToolTip(tip)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(f"""
                QPushButton {{ background:transparent;
                    border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
                    border-radius:5px; color:{t.TEXT_SECONDARY.name()}; }}
                QPushButton:hover {{ background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},35);
                    color:{t.GLOW.name()}; }}
            """)
            rid = self._rid
            btn.clicked.connect(lambda _, s=sig, r=rid: s.emit(r))
            lay.addWidget(btn)

    def _show_menu(self, pos):
        t = _theme
        menu = QMenu(self)
        menu.setStyleSheet(f"""
            QMenu {{ background:{t.BG_DARK.name()};
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},100);
                border-radius:8px; padding:4px 0; color:{t.TEXT_PRIMARY.name()};
                font-family:'Segoe UI'; font-size:9pt; }}
            QMenu::item {{ padding:6px 18px; border-radius:4px; }}
            QMenu::item:selected {{ background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);
                color:{t.GLOW.name()}; }}
        """)
        menu.addAction("▶  Test now",   lambda: self.test_requested.emit(self._rid))
        menu.addAction("✏  Edit",       lambda: self.edit_requested.emit(self._rid))
        menu.addSeparator()
        menu.addAction("🗑  Delete",    lambda: self.delete_requested.emit(self._rid))
        menu.exec(self.mapToGlobal(pos))

    def paintEvent(self, e):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        col = QColor(t.BG_MID.red(), t.BG_MID.green(), t.BG_MID.blue(), 140)
        p.setBrush(col); p.setPen(Qt.NoPen)
        p.drawRoundedRect(QRectF(self.rect().adjusted(2, 2, -2, -2)), 8, 8)
        # bottom divider
        div = QColor(t.GLOW); div.setAlpha(28)
        p.setPen(QPen(div, 1))
        p.drawLine(12, self.height()-1, self.width()-12, self.height()-1)
        p.end()


# ── Reminder Window (main management UI) ─────────────────────────────────────

class ReminderWindow(QWidget):
    """Floating reminder management panel."""

    def __init__(self, store: ReminderStore,
                 scheduler: Optional[ReminderScheduler] = None,
                 parent=None):
        super().__init__(parent, Qt.Window | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self._store     = store
        self._scheduler = scheduler
        self._drag_pos: Optional[QPoint] = None
        self.setMinimumSize(560, 520)
        self.resize(620, 580)
        self._build_ui()
        self._refresh_list()

    def _build_ui(self):
        t = _theme
        outer = QVBoxLayout(self)
        outer.setContentsMargins(10, 10, 10, 10)
        outer.setSpacing(0)

        card = QWidget()
        card.setObjectName("rwin_card")
        card.setStyleSheet(f"""
            QWidget#rwin_card {{
                background:qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},252),
                    stop:1 rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},255));
                border:1.5px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},90);
                border-radius:14px;
            }}
        """)
        card.setGraphicsEffect(make_shadow(card, 32, QColor(0, 0, 0, 200)))
        cl = QVBoxLayout(card)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(0)

        # Title bar
        tbar = QWidget(); tbar.setFixedHeight(46)
        tbar.setStyleSheet("background:transparent;")
        tbl = QHBoxLayout(tbar); tbl.setContentsMargins(18, 0, 12, 0)
        bell_lbl = QLabel("🔔")
        bell_lbl.setFont(QFont("Segoe UI Emoji", 14))
        tbl.addWidget(bell_lbl)
        tbl.addSpacing(6)
        ttl = QLabel("Reminders")
        ttl.setFont(QFont("Segoe UI", 12, QFont.Weight.DemiBold))
        ttl.setStyleSheet(f"color:{t.TEXT_PRIMARY.name()};")
        tbl.addWidget(ttl); tbl.addStretch()

        new_btn = QPushButton("+ New Reminder")
        new_btn.setFont(QFont("Segoe UI", 9, QFont.Weight.DemiBold))
        new_btn.setFixedHeight(30); new_btn.setCursor(Qt.PointingHandCursor)
        new_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},130);
                border:none; border-radius:6px; color:{t.BG_DARK.name()};
                padding:0 14px; font-weight:600; }}
            QPushButton:hover {{
                background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},190); }}
        """)
        new_btn.clicked.connect(self._on_new)
        tbl.addWidget(new_btn)
        tbl.addSpacing(8)

        close_btn = QPushButton("✕")
        close_btn.setFixedSize(28, 28); close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setFont(QFont("Segoe UI", 10))
        close_btn.setStyleSheet(f"""
            QPushButton {{ background:transparent; border:none;
                color:{t.TEXT_DIM.name()}; border-radius:5px; }}
            QPushButton:hover {{ background:rgba(200,60,60,100); color:#ff8888; }}
        """)
        close_btn.clicked.connect(self.close)
        tbl.addWidget(close_btn)

        tbar.mousePressEvent   = lambda e: setattr(self, "_drag_pos",
                                                    e.globalPosition().toPoint())
        tbar.mouseMoveEvent    = lambda e: (
            self.move(self.pos() + e.globalPosition().toPoint() - self._drag_pos),
            setattr(self, "_drag_pos", e.globalPosition().toPoint())
        ) if self._drag_pos else None
        tbar.mouseReleaseEvent = lambda e: setattr(self, "_drag_pos", None)
        cl.addWidget(tbar)

        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},35);")
        cl.addWidget(sep)

        # Filter row
        filt_w = QWidget()
        filt_w.setStyleSheet("background:transparent;")
        filt_lay = QHBoxLayout(filt_w); filt_lay.setContentsMargins(16, 8, 16, 4)
        filt_lay.setSpacing(8)
        self._search_edit = QLineEdit()
        self._search_edit.setPlaceholderText("Search reminders…")
        self._search_edit.setFont(FONT_LABEL)
        self._search_edit.setFixedHeight(28)
        self._search_edit.setStyleSheet(f"""
            QLineEdit {{
                background:rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},180);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
                border-radius:5px; color:{t.TEXT_PRIMARY.name()}; padding:0 8px;
            }}
            QLineEdit:focus {{
                border-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},150);
            }}
        """)
        self._search_edit.textChanged.connect(self._refresh_list)
        filt_lay.addWidget(self._search_edit, 1)

        self._filter_combo = QComboBox()
        self._filter_combo.setFont(FONT_LABEL)
        self._filter_combo.setFixedHeight(28)
        for label, key in [("All","all"),("Enabled","enabled"),
                            ("Disabled","disabled"),
                            ("Level 1","l1"),("Level 2","l2"),("Level 3","l3")]:
            self._filter_combo.addItem(label, key)
        self._filter_combo.setStyleSheet(f"""
            QComboBox {{
                background:rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},180);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
                border-radius:5px; color:{t.TEXT_PRIMARY.name()};
                padding:0 8px; height:26px;
            }}
            QComboBox:hover {{ border-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},130); }}
            QComboBox::drop-down {{ border:none; width:18px; }}
            QComboBox QAbstractItemView {{
                background:{t.BG_DARK.name()};
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},90);
                color:{t.TEXT_PRIMARY.name()};
                selection-background-color:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
            }}
        """)
        self._filter_combo.currentIndexChanged.connect(self._refresh_list)
        filt_lay.addWidget(self._filter_combo)
        cl.addWidget(filt_w)

        # Stats bar
        self._stats_lbl = QLabel("")
        self._stats_lbl.setFont(QFont("Segoe UI", 8))
        self._stats_lbl.setStyleSheet(f"color:{t.TEXT_DIM.name()}; padding:0 18px;")
        cl.addWidget(self._stats_lbl)

        # Scroll area for reminder list
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("background:transparent;")
        self._list_widget = QWidget()
        self._list_widget.setStyleSheet("background:transparent;")
        self._list_layout = QVBoxLayout(self._list_widget)
        self._list_layout.setContentsMargins(12, 8, 12, 8)
        self._list_layout.setSpacing(6)
        self._list_layout.addStretch()
        scroll.setWidget(self._list_widget)
        cl.addWidget(scroll, 1)

        # Footer
        footer_w = QWidget(); footer_w.setFixedHeight(40)
        footer_w.setStyleSheet("background:transparent;")
        ftl = QHBoxLayout(footer_w); ftl.setContentsMargins(16, 4, 16, 4); ftl.setSpacing(8)

        help_lbl = QLabel(
            "Level 1: Notify/Popup/Sound  ·  "
            "Level 2: Full-screen all monitors  ·  "
            "Level 3: Requires 'acknowledge'")
        help_lbl.setFont(QFont("Segoe UI", 7))
        help_lbl.setStyleSheet(f"color:{t.TEXT_DIM.name()};")
        ftl.addWidget(help_lbl, 1)
        cl.addWidget(footer_w)

        outer.addWidget(card)

    def _refresh_list(self):
        t = _theme
        # Clear list (except stretch at end)
        while self._list_layout.count() > 1:
            item = self._list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        query    = self._search_edit.text().strip().lower()
        fkey     = self._filter_combo.currentData() or "all"
        all_rem  = self._store.all()

        filtered: list[dict] = []
        for r in all_rem:
            # Text search
            if query:
                searchable = " ".join([
                    r.get("title", ""), r.get("message", ""),
                    " ".join(r.get("tags", []))
                ]).lower()
                if query not in searchable:
                    continue
            # Filter
            if fkey == "enabled"  and not r.get("enabled",  True): continue
            if fkey == "disabled" and r.get("enabled", True):       continue
            if fkey == "l1" and int(r.get("level", 1)) != 1:        continue
            if fkey == "l2" and int(r.get("level", 1)) != 2:        continue
            if fkey == "l3" and int(r.get("level", 1)) != 3:        continue
            filtered.append(r)

        # Sort: urgent first, then high, then normal; then by title
        _pri_order = {"urgent": 0, "high": 1, "normal": 2}
        filtered.sort(key=lambda r: (
            _pri_order.get(r.get("priority", "normal"), 2),
            r.get("title", "").lower()
        ))

        if not filtered:
            empty_lbl = QLabel("No reminders found.\nClick '+ New Reminder' to add one.")
            empty_lbl.setFont(QFont("Segoe UI", 10))
            empty_lbl.setAlignment(Qt.AlignCenter)
            empty_lbl.setStyleSheet(f"color:{t.TEXT_DIM.name()}; padding:40px 0;")
            self._list_layout.insertWidget(0, empty_lbl)
        else:
            for i, r in enumerate(filtered):
                row = _ReminderRow(r)
                row.edit_requested.connect(self._on_edit)
                row.delete_requested.connect(self._on_delete)
                row.test_requested.connect(self._on_test)
                row.toggle_requested.connect(self._on_toggle)
                self._list_layout.insertWidget(i, row)

        total   = len(all_rem)
        enabled = sum(1 for r in all_rem if r.get("enabled", True))
        self._stats_lbl.setText(
            f"{total} reminder{'s' if total != 1 else ''}  ·  "
            f"{enabled} enabled  ·  showing {len(filtered)}")

    # ── actions ───────────────────────────────────────────────────────────────

    def _on_new(self):
        try:
            wiz = ReminderWizard(parent=self)
            wiz.reminder_saved.connect(self._on_reminder_saved)
            wiz.exec()
        except Exception as exc:
            print(f"[ReminderWindow] new wizard error: {exc}", file=sys.stderr)
            QMessageBox.critical(self, "Error", f"Could not open wizard:\n{exc}")

    def _on_edit(self, reminder_id: str):
        try:
            r = self._store.get(reminder_id)
            if r is None:
                return
            wiz = ReminderWizard(existing=r, parent=self)
            wiz.reminder_saved.connect(self._on_reminder_saved)
            wiz.exec()
        except Exception as exc:
            print(f"[ReminderWindow] edit wizard error: {exc}", file=sys.stderr)
            QMessageBox.critical(self, "Error", f"Could not open editor:\n{exc}")

    def _on_reminder_saved(self, reminder: dict):
        try:
            if reminder.get("id") and self._store.get(reminder["id"]):
                self._store.update(reminder["id"], reminder)
            else:
                self._store.add(reminder)
            self._refresh_list()
        except Exception as exc:
            print(f"[ReminderWindow] save error: {exc}", file=sys.stderr)
            QMessageBox.critical(self, "Error", f"Could not save reminder:\n{exc}")

    def _on_delete(self, reminder_id: str):
        try:
            r = self._store.get(reminder_id)
            if r is None:
                return
            title = r.get("title", "this reminder")
            reply = QMessageBox.question(
                self, "Delete Reminder",
                f"Permanently delete \"{title}\"?\nThis cannot be undone.",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No)
            if reply == QMessageBox.Yes:
                self._store.delete(reminder_id)
                self._refresh_list()
        except Exception as exc:
            print(f"[ReminderWindow] delete error: {exc}", file=sys.stderr)

    def _on_test(self, reminder_id: str):
        try:
            r = self._store.get(reminder_id)
            if r is None:
                return
            if self._scheduler is not None:
                self._scheduler._fire(r)
            else:
                # Fallback: fire directly
                lv = int(r.get("level", 1))
                if lv == 1:
                    dlg = _ReminderPopupDialog(r, self)
                    dlg.exec()
                elif lv == 2:
                    _ReminderFullscreenOverlay(r)
                elif lv == 3:
                    dlg = _ReminderAcknowledgeDialog(r, self)
                    dlg.exec()
        except Exception as exc:
            print(f"[ReminderWindow] test error: {exc}", file=sys.stderr)
            QMessageBox.critical(self, "Error", f"Test fire failed:\n{exc}")

    def _on_toggle(self, reminder_id: str, enabled: bool):
        try:
            self._store.update(reminder_id, {"enabled": enabled})
        except Exception as exc:
            print(f"[ReminderWindow] toggle error: {exc}", file=sys.stderr)

    def paintEvent(self, e):
        p = QPainter(self)
        p.fillRect(self.rect(), Qt.transparent)
        p.end()

    # ── resize / drag ────────────────────────────────────────────────────────

    _RESIZE_MARGIN = 6

    def mousePressEvent(self, e: QMouseEvent):
        if e.button() == Qt.LeftButton:
            self._drag_pos = e.globalPosition().toPoint()

    def mouseMoveEvent(self, e: QMouseEvent):
        if self._drag_pos and (e.buttons() & Qt.LeftButton):
            self.move(self.pos() + e.globalPosition().toPoint() - self._drag_pos)
            self._drag_pos = e.globalPosition().toPoint()

    def mouseReleaseEvent(self, e: QMouseEvent):
        self._drag_pos = None


# ---------------------------------------------------------------------------
# Media Library
# ---------------------------------------------------------------------------

class _MediaStore:
    """Thin JSON-backed store for emojis and emoticons."""

    # Module-level in-memory cache so repeated calls within a session don't
    # hit the filesystem on every emoji/emoticon operation.
    _cache: "dict | None" = None

    @classmethod
    def _load(cls) -> dict:
        if cls._cache is not None:
            return cls._cache
        if MEDIA_LIBRARY_JSON.exists():
            try:
                cls._cache = json.loads(MEDIA_LIBRARY_JSON.read_text(encoding="utf-8"))
                return cls._cache
            except Exception:
                pass
        cls._cache = {"emojis": [], "emoticons": list(_DEFAULT_EMOTICONS)}
        return cls._cache

    @classmethod
    def _save(cls, data: dict):
        cls._cache = data
        try:
            MEDIA_LIBRARY_JSON.write_text(
                json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as exc:
            print(f"[MediaLibrary] Save error: {exc}", file=sys.stderr)

    @classmethod
    def all_emojis(cls) -> list[str]:
        return cls._load().get("emojis", [])

    @classmethod
    def all_emoticons(cls) -> list[str]:
        return cls._load().get("emoticons", list(_DEFAULT_EMOTICONS))

    @classmethod
    def add_emoji(cls, text: str) -> bool:
        """Return True if added, False if already present."""
        d = cls._load(); items = d.get("emojis", [])
        if text in items:
            return False
        items.append(text)
        d["emojis"] = items; cls._save(d)
        return True

    @classmethod
    def add_emoticon(cls, text: str) -> bool:
        """Return True if added, False if already present."""
        d = cls._load(); items = d.get("emoticons", [])
        if text in items:
            return False
        items.append(text)
        d["emoticons"] = items; cls._save(d)
        return True

    @classmethod
    def remove_emoji(cls, text: str):
        d = cls._load(); d["emojis"] = [e for e in d.get("emojis", []) if e != text]
        cls._save(d)

    @classmethod
    def remove_emoticon(cls, text: str):
        d = cls._load(); d["emoticons"] = [e for e in d.get("emoticons", []) if e != text]
        cls._save(d)

    @classmethod
    def all_gifs(cls) -> list[Path]:
        return sorted(MEDIA_LIBRARY_GIFS.glob("*.gif"))

    @classmethod
    def all_pictures(cls) -> list[Path]:
        exts = {"*.png", "*.jpg", "*.jpeg", "*.bmp", "*.webp"}
        paths = []
        for ext in exts:
            paths.extend(MEDIA_LIBRARY_PICS.glob(ext))
        return sorted(paths)

    @classmethod
    def add_media_file(cls, src: str, kind: str) -> bool:
        """Copy a file into gifs/ or pictures/ folder.
        Returns True if the file was added, False if an identical file already exists."""
        p = Path(src)
        if not p.is_file():
            return False
        dest_dir = MEDIA_LIBRARY_GIFS if kind == "gif" else MEDIA_LIBRARY_PICS
        # Compute MD5 of incoming file
        src_hash = hashlib.md5(p.read_bytes()).hexdigest()
        # Check every existing file in the target dir for a content match
        existing = list(dest_dir.glob("*"))
        for ex in existing:
            if ex.is_file():
                try:
                    if hashlib.md5(ex.read_bytes()).hexdigest() == src_hash:
                        return False  # identical content already stored
                except OSError:
                    pass
        dest = dest_dir / p.name
        # Avoid filename collision with a different file
        if dest.exists():
            stem = p.stem; suffix = p.suffix; i = 1
            while dest.exists():
                dest = dest_dir / f"{stem}_{i}{suffix}"; i += 1
        try:
            shutil.copy2(str(p), str(dest))
            return True
        except OSError as exc:
            print(f"[MediaLibrary] File copy error: {exc}", file=sys.stderr)
            return False

    @classmethod
    def remove_media_file(cls, path: Path):
        try:
            if path.exists():
                path.unlink()
        except OSError as exc:
            print(f"[MediaLibrary] Remove error: {exc}", file=sys.stderr)


class _GifLabel(QLabel):
    """A label that only animates its QMovie while the mouse is hovered."""

    def __init__(self, path: Path, thumb_size: int = 110, parent=None):
        super().__init__(parent)
        self._path = path
        self._movie: Optional["QMovie"] = None
        self._thumb_size = thumb_size
        self.setFixedSize(thumb_size, thumb_size)
        self.setAlignment(Qt.AlignCenter)
        self.setCursor(Qt.PointingHandCursor)
        self.setScaledContents(False)
        # Show first frame as static preview
        self._show_static()

    def _show_static(self):
        from PySide6.QtGui import QMovie as _QMovie
        # Show first frame without animation
        movie = _QMovie(str(self._path))
        movie.jumpToFrame(0)
        pm = movie.currentPixmap()
        if not pm.isNull():
            pm = pm.scaled(self._thumb_size, self._thumb_size,
                           Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.setPixmap(pm)
        movie.stop()

    def enterEvent(self, e):
        super().enterEvent(e)
        self._start_movie()

    def leaveEvent(self, e):
        super().leaveEvent(e)
        self._stop_movie()

    def _on_frame_changed(self):
        """Named slot — avoids holding a strong ref in a lambda after widget deletion."""
        if self._movie is not None:
            try:
                self.setPixmap(self._movie.currentPixmap())
            except RuntimeError:
                # C++ object already deleted (e.g. rapid refresh); swallow gracefully
                pass

    def _start_movie(self):
        from PySide6.QtGui import QMovie as _QMovie
        if self._movie is None:
            self._movie = _QMovie(str(self._path))
            self._movie.setScaledSize(QSize(self._thumb_size, self._thumb_size))
            self._movie.frameChanged.connect(self._on_frame_changed)
        self._movie.start()

    def _stop_movie(self):
        if self._movie:
            self._movie.stop()
            try:
                self._movie.frameChanged.disconnect(self._on_frame_changed)
            except RuntimeError:
                pass
            self._movie.deleteLater()
            self._movie = None
        self._show_static()


class _MediaGridWidget(QWidget):
    """Scrollable grid of media tiles (GIFs or Pictures)."""
    item_clicked      = Signal(Path)
    item_delete       = Signal(Path)
    duplicate_detected = Signal(str)   # emitted with filename when a dup is dropped

    _TILE = 120   # tile size px
    _COLS = 4     # fixed column count — avoids unstable width-based calc in resizeEvent

    def __init__(self, kind: str, parent=None):
        """kind: 'gif' | 'picture'"""
        super().__init__(parent)
        self._kind = kind
        self._layout = QGridLayout(self)
        self._layout.setContentsMargins(8, 8, 8, 8)
        self._layout.setSpacing(8)
        self.setAcceptDrops(True)
        self.refresh()

    def refresh(self):
        # Stop any running GIF animations *before* deleting their parent frames,
        # so the QMovie signal can't fire against an already-deleted C++ object.
        while self._layout.count():
            item = self._layout.takeAt(0)
            w = item.widget()
            if w:
                for gif_lbl in w.findChildren(_GifLabel):
                    gif_lbl._stop_movie()
                w.deleteLater()

        paths = (_MediaStore.all_gifs() if self._kind == "gif"
                 else _MediaStore.all_pictures())
        t = _theme
        cols = self._COLS
        for idx, path in enumerate(paths):
            row, col = divmod(idx, cols)
            frame = QWidget()
            frame.setFixedSize(self._TILE, self._TILE)
            frame.setCursor(Qt.PointingHandCursor)
            frame.setContextMenuPolicy(Qt.CustomContextMenu)
            frame.customContextMenuRequested.connect(
                lambda pos, p=path, fr=frame: self._on_ctx(p, fr))
            fl = QVBoxLayout(frame)
            fl.setContentsMargins(4, 4, 4, 4)
            fl.setSpacing(2)
            if self._kind == "gif":
                lbl = _GifLabel(path, self._TILE - 8)
                lbl.mousePressEvent = lambda e, p=path: (
                    self.item_clicked.emit(p) if e.button() == Qt.LeftButton else None)
                fl.addWidget(lbl, alignment=Qt.AlignCenter)
            else:
                lbl = QLabel()
                lbl.setFixedSize(self._TILE - 8, self._TILE - 8)
                lbl.setAlignment(Qt.AlignCenter)
                pm = QPixmap(str(path))
                if not pm.isNull():
                    pm = pm.scaled(self._TILE - 8, self._TILE - 8,
                                   Qt.KeepAspectRatio, Qt.SmoothTransformation)
                lbl.setPixmap(pm)
                lbl.setCursor(Qt.PointingHandCursor)
                lbl.mousePressEvent = lambda e, p=path: (
                    self.item_clicked.emit(p) if e.button() == Qt.LeftButton else None)
                fl.addWidget(lbl, alignment=Qt.AlignCenter)
            frame.setStyleSheet(
                f"QWidget {{ background: rgba({t.TILE_BG_BASE.red()},"
                f"{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);"
                f"border: 1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},55);"
                f"border-radius: 6px; }}"
                f"QWidget:hover {{ border-color: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},160); }}")
            self._layout.addWidget(frame, row, col)

        # Push content to the top-left: stretch fills remaining rows/cols
        if paths:
            last_row = (len(paths) - 1) // cols
            self._layout.setRowStretch(last_row + 1, 1)
        self._layout.setColumnStretch(cols, 1)

    def _on_ctx(self, path: Path, frame: QWidget):
        t = _theme
        menu = QMenu(self)
        menu.setStyleSheet(
            f"QMenu {{ background:rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},245);"
            f"border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},110);"
            f"border-radius:8px; padding:4px 0; color:{t.TEXT_PRIMARY.name()};"
            f"font-family:'Segoe UI'; font-size:9pt; }}"
            f"QMenu::item {{ padding:6px 18px; border-radius:4px; }}"
            f"QMenu::item:selected {{ background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);"
            f"color:{t.GLOW.name()}; }}")
        act_del = menu.addAction("🗑  Remove from library")
        chosen = menu.exec(QCursor.pos())
        if chosen == act_del:
            self.item_delete.emit(path)

    # ── Drag and drop ─────────────────────────────────────────────────

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            valid_exts = (
                {".gif"} if self._kind == "gif"
                else {".png", ".jpg", ".jpeg", ".bmp", ".webp"}
            )
            if any(Path(u.toLocalFile()).suffix.lower() in valid_exts
                   for u in e.mimeData().urls() if u.isLocalFile()):
                e.acceptProposedAction(); return
        e.ignore()

    def dragMoveEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
        else:
            e.ignore()

    def dropEvent(self, e):
        valid_exts = (
            {".gif"} if self._kind == "gif"
            else {".png", ".jpg", ".jpeg", ".bmp", ".webp"}
        )
        added = 0
        dup_names: list[str] = []
        for url in e.mimeData().urls():
            if not url.isLocalFile(): continue
            p = Path(url.toLocalFile())
            if p.suffix.lower() in valid_exts:
                ok = _MediaStore.add_media_file(str(p), self._kind)
                if ok:
                    added += 1
                else:
                    dup_names.append(p.name)
        if added:
            self.refresh()
        if dup_names:
            self.duplicate_detected.emit(dup_names[0] if len(dup_names) == 1
                                         else f"{len(dup_names)} files")
        e.acceptProposedAction()


class _TextChipList(QWidget):
    """Scrollable list of text chips (emojis or emoticons)."""
    item_clicked       = Signal(str)
    item_delete        = Signal(str)
    duplicate_detected = Signal(str)   # emitted with the duplicate text

    def __init__(self, kind: str, parent=None):
        """kind: 'emoji' | 'emoticon'"""
        super().__init__(parent)
        self._kind = kind
        self._main_layout = QVBoxLayout(self)
        self._main_layout.setContentsMargins(8, 8, 8, 8)
        self._main_layout.setSpacing(6)

        # Input row
        t = _theme
        input_row = QHBoxLayout(); input_row.setSpacing(6)
        self._input = QLineEdit()
        self._input.setFont(QFont("Segoe UI", 11))
        self._input.setPlaceholderText(
            "Paste emoji here…" if kind == "emoji" else "Type emoticon here…")
        self._input.setStyleSheet(f"""
            QLineEdit {{
                background: rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
                border: 1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius: 5px; color: {t.TEXT_PRIMARY.name()};
                padding: 5px 9px; font-family:'Segoe UI Emoji'; font-size:11pt;
            }}
            QLineEdit:focus {{
                border-color: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},160);
            }}
        """)
        add_btn = QPushButton("Add")
        add_btn.setFixedHeight(30)
        add_btn.setCursor(Qt.PointingHandCursor)
        add_btn.setFont(QFont("Segoe UI", 9, QFont.Weight.DemiBold))
        add_btn.setStyleSheet(f"""
            QPushButton {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 rgba({t.ACCENT_BLUE.red()},{t.ACCENT_BLUE.green()},{t.ACCENT_BLUE.blue()},180),
                    stop:1 rgba({t.ACCENT_TEAL.red()},{t.ACCENT_TEAL.green()},{t.ACCENT_TEAL.blue()},180));
                border: none; border-radius: 5px;
                color: {t.TEXT_PRIMARY.name()}; padding: 0 14px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 rgba({t.ACCENT_BLUE.red()},{t.ACCENT_BLUE.green()},{t.ACCENT_BLUE.blue()},230),
                    stop:1 rgba({t.ACCENT_TEAL.red()},{t.ACCENT_TEAL.green()},{t.ACCENT_TEAL.blue()},230));
            }}
        """)
        add_btn.clicked.connect(self._on_add)
        self._input.returnPressed.connect(self._on_add)
        input_row.addWidget(self._input)
        input_row.addWidget(add_btn)
        self._main_layout.addLayout(input_row)

        # Scroll area for chips
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setStyleSheet("background:transparent;")
        self._chips_widget = QWidget()
        self._chips_widget.setStyleSheet("background:transparent;")
        self._flow = QVBoxLayout(self._chips_widget)
        self._flow.setContentsMargins(0, 0, 0, 0)
        self._flow.setSpacing(4)
        self._scroll.setWidget(self._chips_widget)
        self._main_layout.addWidget(self._scroll)

        self.refresh()

    def _on_add(self):
        text = self._input.text().strip()
        if not text: return
        if self._kind == "emoji":
            added = _MediaStore.add_emoji(text)
        else:
            added = _MediaStore.add_emoticon(text)
        self._input.clear()
        if added:
            self.refresh()
        else:
            self.duplicate_detected.emit(text)

    def refresh(self):
        # Clear
        while self._flow.count():
            item = self._flow.takeAt(0)
            if item.widget(): item.widget().deleteLater()

        items = (_MediaStore.all_emojis() if self._kind == "emoji"
                 else _MediaStore.all_emoticons())
        t = _theme
        for text in items:
            row = QWidget()
            row.setStyleSheet(
                f"QWidget {{ background: rgba({t.TILE_BG_BASE.red()},"
                f"{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},180);"
                f"border: 1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);"
                f"border-radius: 6px; }}"
                f"QWidget:hover {{ border-color: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},160); }}")
            row.setCursor(Qt.PointingHandCursor)
            row.setContextMenuPolicy(Qt.CustomContextMenu)
            row.customContextMenuRequested.connect(
                lambda pos, tx=text: self._on_ctx(tx))
            rl = QHBoxLayout(row); rl.setContentsMargins(10, 6, 10, 6); rl.setSpacing(8)
            lbl = QLabel(text)
            lbl.setFont(QFont("Segoe UI", 12))
            lbl.setStyleSheet(
                f"color:{t.TEXT_PRIMARY.name()}; background:transparent; border:none;")
            lbl.setTextInteractionFlags(Qt.NoTextInteraction)
            rl.addWidget(lbl); rl.addStretch()
            row.mousePressEvent = lambda e, tx=text: (
                self.item_clicked.emit(tx) if e.button() == Qt.LeftButton else None)
            self._flow.addWidget(row)
        self._flow.addStretch()

    def _on_ctx(self, text: str):
        t = _theme
        menu = QMenu(self)
        menu.setStyleSheet(
            f"QMenu {{ background:rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},245);"
            f"border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},110);"
            f"border-radius:8px; padding:4px 0; color:{t.TEXT_PRIMARY.name()};"
            f"font-family:'Segoe UI'; font-size:9pt; }}"
            f"QMenu::item {{ padding:6px 18px; border-radius:4px; }}"
            f"QMenu::item:selected {{ background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},50);"
            f"color:{t.GLOW.name()}; }}")
        act_del = menu.addAction("🗑  Remove from library")
        chosen = menu.exec(QCursor.pos())
        if chosen == act_del:
            self.item_delete.emit(text)


class MediaLibraryWindow(QWidget):
    """Floating media library panel with GIFs, Emojis, Emoticons, Pictures tabs."""
    duplicate_detected = Signal(str)   # bubbled up from child grids/chip-lists

    def __init__(self, parent=None):
        super().__init__(parent, Qt.Window | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setMinimumSize(520, 560)
        self.resize(640, 600)
        self._drag_pos:          Optional[QPoint] = None
        self._resize_dir:        Optional[str]    = None
        self._resize_start_geom: Optional[QRect]  = None
        self._resize_start_pos:  Optional[QPoint] = None
        self._build_ui()
        _theme.register(self._on_theme)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(0)

        self._card = QWidget()
        self._card.setObjectName("ml_card")
        cl = QVBoxLayout(self._card)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(0)
        self._card.setGraphicsEffect(make_shadow(self._card, 32, QColor(0, 0, 0, 210)))

        # ── Title bar ──────────────────────────────────────────────────
        tbar = QWidget(); tbar.setFixedHeight(44)
        tbar.setStyleSheet("background:transparent;")
        tbl = QHBoxLayout(tbar); tbl.setContentsMargins(16, 0, 12, 0)
        ico = QLabel("🖼")
        ico.setFont(QFont("Segoe UI", 12))
        ico.setStyleSheet("background:transparent;")
        ttl = QLabel("  Media Library")
        ttl.setFont(QFont("Segoe UI", 11, QFont.Weight.DemiBold))
        ttl.setStyleSheet(f"color:{_theme.TEXT_PRIMARY.name()}; background:transparent;")
        tbl.addWidget(ico); tbl.addWidget(ttl); tbl.addStretch()
        tenor_btn = QPushButton("Tenor")
        tenor_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        tenor_btn.setFont(QFont("Segoe UI", 9, QFont.Weight.DemiBold))
        tenor_btn.setStyleSheet(
            "QPushButton { background: transparent; color: #4A9EFF;"
            " border: 1px solid rgba(74,158,255,130); border-radius: 4px;"
            " padding: 2px 10px; }"
            "QPushButton:hover { background: rgba(74,158,255,30); }"
        )
        tenor_btn.clicked.connect(lambda: __import__("webbrowser").open(
            "https://tenor.com/view/rickroll-roll-rick-never-gonna-give-you-up-never-gonna-gif-22954713"
        ))
        tbl.addWidget(tenor_btn)
        tbl.addSpacing(6)
        cb = TitleBarButton(COLOR_BTN_CLOSE, "x")
        cb.clicked.connect(self.close)
        tbl.addWidget(cb)
        tbar.mousePressEvent  = self._tbar_press
        tbar.mouseMoveEvent   = self._tbar_move
        tbar.mouseReleaseEvent = lambda e: setattr(self, "_drag_pos", None)
        cl.addWidget(tbar)

        topsep = QFrame(); topsep.setFrameShape(QFrame.HLine)
        topsep.setStyleSheet(
            f"color:rgba({_theme.GLOW.red()},{_theme.GLOW.green()},{_theme.GLOW.blue()},50);")
        cl.addWidget(topsep)

        # ── Tab widget ─────────────────────────────────────────────────
        self._tabs = QTabWidget()
        self._tabs.setDocumentMode(True)
        self._apply_tab_style()
        cl.addWidget(self._tabs)

        # GIFs tab
        self._gif_scroll = QScrollArea()
        self._gif_scroll.setWidgetResizable(True)
        self._gif_scroll.setFrameShape(QFrame.NoFrame)
        self._gif_scroll.setStyleSheet("background:transparent;")
        self._gif_grid = _MediaGridWidget("gif")
        self._gif_grid.item_clicked.connect(self._on_media_file_clicked)
        self._gif_grid.item_delete.connect(self._on_delete_media_file)
        self._gif_grid.duplicate_detected.connect(self._on_duplicate_detected)
        self._gif_scroll.setWidget(self._gif_grid)
        self._gif_drop_label = self._make_drop_hint("Drop .gif files here")
        gif_container = QWidget()
        gif_cl = QVBoxLayout(gif_container)
        gif_cl.setContentsMargins(0, 0, 0, 0); gif_cl.setSpacing(0)
        gif_cl.addWidget(self._gif_drop_label)
        gif_cl.addWidget(self._gif_scroll)
        self._tabs.addTab(gif_container, "  GIFs  ")

        # Favorite Emojis tab
        self._emoji_list = _TextChipList("emoji")
        self._emoji_list.item_clicked.connect(self._on_text_clicked)
        self._emoji_list.item_delete.connect(self._on_delete_emoji)
        self._emoji_list.duplicate_detected.connect(self._on_duplicate_detected)
        self._tabs.addTab(self._emoji_list, "  Emojis  ")

        # Favorite Emoticons tab
        self._emoticon_list = _TextChipList("emoticon")
        self._emoticon_list.item_clicked.connect(self._on_text_clicked)
        self._emoticon_list.item_delete.connect(self._on_delete_emoticon)
        self._emoticon_list.duplicate_detected.connect(self._on_duplicate_detected)
        self._tabs.addTab(self._emoticon_list, "  Emoticons  ")

        # Pictures tab
        self._pic_scroll = QScrollArea()
        self._pic_scroll.setWidgetResizable(True)
        self._pic_scroll.setFrameShape(QFrame.NoFrame)
        self._pic_scroll.setStyleSheet("background:transparent;")
        self._pic_grid = _MediaGridWidget("picture")
        self._pic_grid.item_clicked.connect(self._on_media_file_clicked)
        self._pic_grid.item_delete.connect(self._on_delete_media_file)
        self._pic_grid.duplicate_detected.connect(self._on_duplicate_detected)
        self._pic_scroll.setWidget(self._pic_grid)
        self._pic_drop_label = self._make_drop_hint(
            "Drop .png / .jpg / .jpeg / .bmp / .webp files here")
        pic_container = QWidget()
        pic_cl = QVBoxLayout(pic_container)
        pic_cl.setContentsMargins(0, 0, 0, 0); pic_cl.setSpacing(0)
        pic_cl.addWidget(self._pic_drop_label)
        pic_cl.addWidget(self._pic_scroll)
        self._tabs.addTab(pic_container, "  Pictures  ")

        root.addWidget(self._card)

    def _make_drop_hint(self, text: str) -> QLabel:
        t = _theme
        lbl = QLabel(f"⬇  {text}")
        lbl.setAlignment(Qt.AlignCenter)
        lbl.setFont(QFont("Segoe UI", 9))
        lbl.setStyleSheet(
            f"color: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},140);"
            f"background: rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},60);"
            f"border-bottom: 1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},30);"
            f"padding: 4px 0;")
        lbl.setFixedHeight(26)
        return lbl

    def _apply_tab_style(self):
        t = _theme
        self._tabs.setStyleSheet(f"""
            QTabWidget::pane {{
                background: transparent;
                border: none;
            }}
            QTabBar::tab {{
                background: rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},160);
                color: {t.TEXT_SECONDARY.name()};
                font-family: 'Segoe UI'; font-size: 9pt;
                padding: 6px 14px; border-radius: 0;
                border-bottom: 2px solid transparent;
                margin-right: 2px;
            }}
            QTabBar::tab:selected {{
                background: rgba({t.TILE_BG_HOVER.red()},{t.TILE_BG_HOVER.green()},{t.TILE_BG_HOVER.blue()},200);
                color: {t.GLOW.name()};
                border-bottom: 2px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},220);
            }}
            QTabBar::tab:hover {{
                background: rgba({t.TILE_BG_HOVER.red()},{t.TILE_BG_HOVER.green()},{t.TILE_BG_HOVER.blue()},120);
                color: {t.TEXT_PRIMARY.name()};
            }}
        """)

    # ── Slot handlers ──────────────────────────────────────────────────────

    def _on_duplicate_detected(self, name: str):
        """Show an in-window explanation, then close after a short pause."""
        self._show_copied_toast(f"⚠  Already in your library!", duration_ms=1400)
        # Emit to MainWindow so a toast also appears once the library closes
        self.duplicate_detected.emit(name)
        QTimer.singleShot(1500, self.close)

    def _on_media_file_clicked(self, path: Path):
        """Copy a GIF or picture file to the system clipboard and close."""
        try:
            mime = QMimeData()
            mime.setUrls([QUrl.fromLocalFile(str(path.resolve()))])
            QApplication.clipboard().setMimeData(mime)
            self._show_copied_toast(f"📋 Copied file: {path.name}")
        except Exception as exc:
            print(f"[MediaLibrary] Clipboard error: {exc}", file=sys.stderr)
        QTimer.singleShot(800, self.close)

    def _on_text_clicked(self, text: str):
        """Copy a text item to the clipboard and close."""
        try:
            QApplication.clipboard().setText(text)
            self._show_copied_toast(f"📋 Copied to clipboard!")
        except Exception as exc:
            print(f"[MediaLibrary] Clipboard error: {exc}", file=sys.stderr)
        QTimer.singleShot(800, self.close)

    def _on_delete_media_file(self, path: Path):
        _MediaStore.remove_media_file(path)
        self._gif_grid.refresh()
        self._pic_grid.refresh()

    def _on_delete_emoji(self, text: str):
        _MediaStore.remove_emoji(text)
        self._emoji_list.refresh()

    def _on_delete_emoticon(self, text: str):
        _MediaStore.remove_emoticon(text)
        self._emoticon_list.refresh()

    def _show_copied_toast(self, msg: str, duration_ms: int = 750):
        """Brief overlay label inside the window."""
        t = _theme
        toast = QLabel(msg, self._card)
        toast.setFont(QFont("Segoe UI", 10, QFont.Weight.DemiBold))
        toast.setAlignment(Qt.AlignCenter)
        toast.setStyleSheet(
            f"background: rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},230);"
            f"color: {t.GLOW.name()};"
            f"border: 1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},140);"
            f"border-radius: 8px; padding: 10px 20px;")
        toast.adjustSize()
        cx = (self._card.width()  - toast.width())  // 2
        cy = (self._card.height() - toast.height()) // 2
        toast.move(cx, cy)
        toast.show(); toast.raise_()
        QTimer.singleShot(duration_ms, toast.deleteLater)

    # ── Title bar drag ─────────────────────────────────────────────────────

    def _tbar_press(self, e: QMouseEvent):
        if e.button() == Qt.LeftButton:
            self._drag_pos = e.globalPosition().toPoint()

    def _tbar_move(self, e: QMouseEvent):
        if self._drag_pos:
            delta = e.globalPosition().toPoint() - self._drag_pos
            self.move(self.pos() + delta)
            self._drag_pos = e.globalPosition().toPoint()

    # ── Resize (thin margin) ───────────────────────────────────────────────

    _RM = 6  # resize margin

    def _get_resize_dir(self, pos: QPoint) -> Optional[str]:
        m = self._RM; w, h = self.width(), self.height()
        x, y = pos.x(), pos.y()
        l = x < m; r = x > w - m; tp = y < m; b = y > h - m
        if tp and l: return "tl"
        if tp and r: return "tr"
        if b  and l: return "bl"
        if b  and r: return "br"
        if l: return "l"
        if r: return "r"
        if tp: return "t"
        if b: return "b"
        return None

    _CUR_MAP = {
        "tl": Qt.SizeFDiagCursor, "br": Qt.SizeFDiagCursor,
        "tr": Qt.SizeBDiagCursor, "bl": Qt.SizeBDiagCursor,
        "l":  Qt.SizeHorCursor,   "r":  Qt.SizeHorCursor,
        "t":  Qt.SizeVerCursor,   "b":  Qt.SizeVerCursor,
    }

    def mousePressEvent(self, e: QMouseEvent):
        if e.button() != Qt.LeftButton: return
        d = self._get_resize_dir(e.position().toPoint())
        if d:
            self._resize_dir = d
            self._resize_start_geom = self.geometry()
            self._resize_start_pos  = e.globalPosition().toPoint()

    def mouseMoveEvent(self, e: QMouseEvent):
        if self._resize_dir and self._resize_start_pos:
            self._do_resize(e.globalPosition().toPoint()); return
        d = self._get_resize_dir(e.position().toPoint())
        self.setCursor(self._CUR_MAP.get(d, Qt.ArrowCursor) if d else Qt.ArrowCursor)

    def mouseReleaseEvent(self, e: QMouseEvent):
        self._resize_dir = self._resize_start_geom = self._resize_start_pos = None

    def _do_resize(self, gp: QPoint):
        if not (self._resize_start_geom and self._resize_start_pos): return
        dx = gp.x() - self._resize_start_pos.x()
        dy = gp.y() - self._resize_start_pos.y()
        g = QRect(self._resize_start_geom); d = self._resize_dir or ""
        if "r" in d: g.setRight(g.right() + dx)
        if "b" in d: g.setBottom(g.bottom() + dy)
        if "l" in d: g.setLeft(g.left() + dx)
        if "t" in d: g.setTop(g.top() + dy)
        if g.width() >= self.minimumWidth() and g.height() >= self.minimumHeight():
            self.setGeometry(g)

    # ── Paint ──────────────────────────────────────────────────────────────

    def paintEvent(self, e: QPaintEvent):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        rect = QRectF(8, 8, self.width() - 16, self.height() - 16)
        path = QPainterPath(); path.addRoundedRect(rect, BORDER_RADIUS + 2, BORDER_RADIUS + 2)
        grad = QLinearGradient(0, 0, self.width(), self.height())
        grad.setColorAt(0, QColor(t.BG_MID.red() + 4, t.BG_MID.green() + 4, t.BG_MID.blue() + 6, 248))
        grad.setColorAt(1, QColor(t.BG_DARK.red(), t.BG_DARK.green(), t.BG_DARK.blue(), 248))
        p.fillPath(path, grad)
        border = QColor(t.GLOW); border.setAlpha(100)
        p.setPen(QPen(border, 1.3)); p.setBrush(Qt.NoBrush); p.drawPath(path)
        p.end()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        if hasattr(self, "_card"):
            self._card.setGeometry(8, 8, self.width() - 16, self.height() - 16)

    def closeEvent(self, e):
        _theme.unregister(self._on_theme)
        super().closeEvent(e)

    def _on_theme(self):
        self._apply_tab_style()
        self.update()


# ---------------------------------------------------------------------------
# Toast notification
# ---------------------------------------------------------------------------

class _ToastNotification(QWidget):
    """Small themed popup that fades in, holds, then fades out."""

    _ICONS = {"info": "ℹ", "success": "✓", "error": "✗", "warning": "⚠"}

    def __init__(self, message: str, kind: str = "info", parent=None):
        super().__init__(parent,
                         Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self._kind    = kind
        self._opacity = 0.0
        self._phase   = "in"
        self._build_ui(message)
        self.adjustSize()
        self._anim = QTimer(self)
        self._anim.setInterval(16)
        self._anim.timeout.connect(self._tick)
        self._anim.start()
        # Pass self as context so the timer auto-cancels if the widget is
        # destroyed before 3.4 s elapses, preventing a use-after-free.
        QTimer.singleShot(3400, self, self._fade_out)
        self.setWindowOpacity(0.0)

    def _build_ui(self, msg: str):
        t = _theme
        color = {
            "info":    t.ACCENT_BLUE,
            "success": t.ACCENT_TEAL,
            "error":   t.ACCENT_RED,
            "warning": t.ACCENT_AMBER,
        }.get(self._kind, t.ACCENT_BLUE)
        outer = QHBoxLayout(self); outer.setContentsMargins(0, 0, 0, 0)
        card = QWidget(); card.setObjectName("toast_card")
        card.setStyleSheet(f"""
            QWidget#toast_card {{
                background:rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},238);
                border:1px solid rgba({color.red()},{color.green()},{color.blue()},150);
                border-left:3px solid rgba({color.red()},{color.green()},{color.blue()},255);
                border-radius:9px;
            }}
        """)
        cl = QHBoxLayout(card); cl.setContentsMargins(14, 10, 16, 10); cl.setSpacing(10)
        icon_lbl = QLabel(self._ICONS.get(self._kind, "ℹ"))
        icon_lbl.setFont(QFont("Segoe UI", 12))
        icon_lbl.setStyleSheet(
            f"color:rgb({color.red()},{color.green()},{color.blue()}); background:transparent;")
        cl.addWidget(icon_lbl)
        msg_lbl = QLabel(msg); msg_lbl.setWordWrap(True)
        msg_lbl.setMaximumWidth(260)
        msg_lbl.setFont(QFont("Segoe UI", 9))
        msg_lbl.setStyleSheet(
            f"color:{t.TEXT_PRIMARY.name()}; background:transparent;")
        cl.addWidget(msg_lbl)
        outer.addWidget(card)
        self.setFixedWidth(320)

    def _tick(self):
        if self._phase == "in":
            self._opacity = min(1.0, self._opacity + 0.10)
            self.setWindowOpacity(self._opacity)
            if self._opacity >= 1.0:
                self._phase = "hold"; self._anim.stop()
        elif self._phase == "out":
            self._opacity = max(0.0, self._opacity - 0.06)
            self.setWindowOpacity(self._opacity)
            if self._opacity <= 0.0:
                self._anim.stop(); self.close()

    def _fade_out(self):
        if self._phase != "out":
            self._phase = "out"; self._anim.start()


# ---------------------------------------------------------------------------
# Modal scrim overlay
# ---------------------------------------------------------------------------

class _DialogScrim(QWidget):
    """Semi-transparent dark overlay drawn over the main window while a dialog is open."""
    def __init__(self, parent: QWidget):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.hide()

    def paintEvent(self, e: QPaintEvent):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor(0, 0, 0, 130))
        p.end()


# ---------------------------------------------------------------------------
# Search bar + filter chip
# ---------------------------------------------------------------------------

class _SearchBar(QWidget):
    """Collapsible search bar that slides in below the title bar."""
    search_changed = Signal(str)   # emits current query (empty string = cleared)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self):
        t = _theme
        self.setFixedHeight(38)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(12, 4, 12, 4)
        lay.setSpacing(6)

        icon_lbl = QLabel("\U0001f50d")
        icon_lbl.setFont(QFont("Segoe UI", 9))
        icon_lbl.setStyleSheet(f"color:{t.TEXT_DIM.name()}; background:transparent;")
        lay.addWidget(icon_lbl)
        self._icon_lbl = icon_lbl

        self._edit = QLineEdit()
        self._edit.setPlaceholderText("Search by name or tag\u2026")
        self._edit.setFont(QFont("Segoe UI", 9))
        self._edit.textChanged.connect(self.search_changed)
        self._edit.textChanged.connect(self._update_clear_btn)
        lay.addWidget(self._edit)

        self._clear_btn = QPushButton("\u00d7")
        self._clear_btn.setFixedSize(22, 22)
        self._clear_btn.setCursor(Qt.PointingHandCursor)
        self._clear_btn.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        self._clear_btn.setToolTip("Clear search")
        self._clear_btn.clicked.connect(self._edit.clear)
        self._clear_btn.setVisible(False)
        lay.addWidget(self._clear_btn)

        self.refresh_theme()

    def refresh_theme(self):
        t = _theme
        self.setStyleSheet(f"""
            _SearchBar, QWidget#search_bar_bg {{
                background: rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},240);
                border-bottom: 1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
            }}
        """)
        if hasattr(self, "_icon_lbl"):
            self._icon_lbl.setStyleSheet(f"color:{t.TEXT_DIM.name()}; background:transparent;")
        if hasattr(self, "_edit"):
            self._edit.setStyleSheet(f"""
                QLineEdit {{
                    background: rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},180);
                    border: 1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                    border-radius: 5px;
                    color: {t.TEXT_PRIMARY.name()};
                    padding: 3px 8px;
                    selection-background-color: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},80);
                }}
                QLineEdit:focus {{
                    border-color: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},160);
                }}
            """)
        if hasattr(self, "_clear_btn"):
            self._clear_btn.setStyleSheet(f"""
                QPushButton {{
                    background: transparent;
                    border: none;
                    color: {t.TEXT_DIM.name()};
                    border-radius: 11px;
                }}
                QPushButton:hover {{
                    background: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},60);
                    color: {t.TEXT_PRIMARY.name()};
                }}
            """)

    def _update_clear_btn(self, text: str):
        self._clear_btn.setVisible(bool(text))

    def focus_edit(self):
        self._edit.setFocus()
        self._edit.selectAll()

    def clear(self):
        self._edit.clear()

    def current_query(self) -> str:
        return self._edit.text()


class _FilterChip(QWidget):
    """Floating pill showing the active search filter with a clear button."""
    clear_requested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()
        self.setVisible(False)

    def _build_ui(self):
        t = _theme
        lay = QHBoxLayout(self)
        lay.setContentsMargins(10, 4, 6, 4)
        lay.setSpacing(5)

        self._lbl = QLabel()
        self._lbl.setFont(QFont("Segoe UI", 8, QFont.Weight.DemiBold))
        self._lbl.setStyleSheet(f"color:{t.GLOW.name()}; background:transparent;")
        lay.addWidget(self._lbl)

        close_btn = QPushButton("\u00d7")
        close_btn.setFixedSize(16, 16)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        close_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent; border: none;
                color: {t.TEXT_SECONDARY.name()};
                border-radius: 8px;
            }}
            QPushButton:hover {{
                color: {t.GLOW.name()};
                background: rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},40);
            }}
        """)
        close_btn.clicked.connect(self.clear_requested)
        lay.addWidget(close_btn)
        self.adjustSize()

    def set_query(self, query: str):
        self._lbl.setText(f"\U0001f50d  {query}")
        self.adjustSize()

    def paintEvent(self, e):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setBrush(QColor(t.BG_MID.red(), t.BG_MID.green(), t.BG_MID.blue(), 230))
        p.setPen(QPen(QColor(t.GLOW.red(), t.GLOW.green(), t.GLOW.blue(), 140), 1))
        p.drawRoundedRect(self.rect().adjusted(0, 0, -1, -1), 10, 10)
        p.end()
        super().paintEvent(e)


# ---------------------------------------------------------------------------
# Node Scheduler
# ---------------------------------------------------------------------------

# Schedule types
SCHED_ONCE      = "once"
SCHED_DAILY     = "daily"
SCHED_WORKDAYS  = "workdays"
SCHED_WEEKENDS  = "weekends"
SCHED_WEEKLY    = "weekly"
SCHED_MONTHLY   = "monthly"
SCHED_INTERVAL  = "interval"

SCHED_TYPES = [
    (SCHED_DAILY,    "Daily"),
    (SCHED_WORKDAYS, "Workdays (Mon–Fri)"),
    (SCHED_WEEKENDS, "Weekends (Sat–Sun)"),
    (SCHED_WEEKLY,   "Weekly (specific days)"),
    (SCHED_MONTHLY,  "Monthly (day of month)"),
    (SCHED_INTERVAL, "Interval (every N minutes/hours)"),
    (SCHED_ONCE,     "Once (specific date & time)"),
]

# Tolerance window (seconds) within which a scheduled fire is considered "on time".
# The scheduler polls every 30 s, so ±35 s gives a comfortable margin.
_SCHED_TOLERANCE_SEC = 35


def _sched_should_fire(sched: dict) -> bool:
    """Return True if the node's schedule should fire right now.

    Does NOT update last_run — the caller is responsible for that.
    Returns False on any parse / logic error so a bad schedule never crashes.
    """
    try:
        if not sched.get("enabled", False):
            return False
        stype = sched.get("type", SCHED_DAILY)
        now   = datetime.now()

        # ── interval ────────────────────────────────────────────────────
        if stype == SCHED_INTERVAL:
            value  = int(sched.get("interval_value", 30))
            unit   = sched.get("interval_unit", "minutes")
            if unit == "hours":
                delta_sec = value * 3600
            else:
                delta_sec = value * 60
            last = sched.get("last_run")
            if not last:
                return True   # never run → fire now
            try:
                last_dt = datetime.fromisoformat(last)
            except (ValueError, TypeError):
                return True
            return (now - last_dt).total_seconds() >= delta_sec

        # ── time-of-day schedules (daily / weekly / etc.) ────────────────
        time_str = sched.get("time", "09:00")
        try:
            h, m = (int(x) for x in time_str.split(":"))
        except (ValueError, TypeError):
            return False
        target_dt = now.replace(hour=h, minute=m, second=0, microsecond=0)
        diff = abs((now - target_dt).total_seconds())
        if diff > _SCHED_TOLERANCE_SEC:
            return False   # not close enough to the scheduled time

        # Prevent double-fire within the same minute
        last = sched.get("last_run")
        if last:
            try:
                last_dt = datetime.fromisoformat(last)
                if abs((now - last_dt).total_seconds()) < 60:
                    return False
            except (ValueError, TypeError):
                pass

        weekday = now.weekday()   # 0=Mon, 6=Sun

        if stype == SCHED_DAILY:
            return True

        if stype == SCHED_WORKDAYS:
            return weekday < 5

        if stype == SCHED_WEEKENDS:
            return weekday >= 5

        if stype == SCHED_WEEKLY:
            days = sched.get("days", [])
            return weekday in days

        if stype == SCHED_MONTHLY:
            dom = int(sched.get("day_of_month", 1))
            # clamp to valid days (e.g. Feb 28/29)
            import calendar
            max_day = calendar.monthrange(now.year, now.month)[1]
            return now.day == min(dom, max_day)

        if stype == SCHED_ONCE:
            date_str = sched.get("date", "")
            try:
                y, mo, d = (int(x) for x in date_str.split("-"))
            except (ValueError, TypeError):
                return False
            return now.year == y and now.month == mo and now.day == d

    except Exception as exc:
        print(f"[NodeScheduler] _sched_should_fire error: {exc}", file=sys.stderr)
    return False


class NodeScheduler(QObject):
    """Polls node schedules every 30 seconds and fires matching nodes.

    Attach to a MainWindow via  scheduler.set_main_window(win)  after the
    window is fully constructed.  Call  start()  to begin polling.
    """
    node_fired = Signal(dict)   # emitted with the node dict when it should run

    _POLL_INTERVAL_MS = 30_000  # 30 seconds

    def __init__(self, store: NodeStore, parent=None):
        super().__init__(parent)
        self._store  = store
        self._win    = None
        self._timer  = QTimer(self)
        self._timer.setInterval(self._POLL_INTERVAL_MS)
        self._timer.timeout.connect(self._poll)
        self._enabled = True

    def set_main_window(self, win):
        self._win = win

    def start(self):
        self._timer.start()
        # Also run an initial check shortly after startup so "once" nodes
        # set to the current minute are not missed.
        QTimer.singleShot(5_000, self._poll)

    def stop(self):
        self._timer.stop()

    def set_enabled(self, enabled: bool):
        self._enabled = enabled
        if enabled:
            self._timer.start()
        else:
            self._timer.stop()

    def _poll(self):
        if not self._enabled:
            return
        try:
            for node in self._store._data:
                if node.get("archived", False):
                    continue
                sched = node.get("schedule")
                if not sched or not isinstance(sched, dict):
                    continue
                if _sched_should_fire(sched):
                    # Mark last_run before firing to prevent re-fires
                    sched["last_run"] = datetime.now().isoformat(timespec="seconds")
                    self._store.save()
                    try:
                        if self._win is not None:
                            self._win._launch_node(node)
                            name = node.get("name", "?")
                            self._win.toast(
                                f"⏰ Scheduled: {name}", "info")
                    except Exception as exc:
                        print(f"[NodeScheduler] launch error for node "
                              f"{node.get('id')}: {exc}", file=sys.stderr)
        except Exception as exc:
            print(f"[NodeScheduler] poll error: {exc}", file=sys.stderr)

    def force_check(self):
        """Manually trigger a schedule check (e.g. after saving a node)."""
        self._poll()


# Module-level singleton created alongside _plugin_manager
_node_scheduler: Optional["NodeScheduler"] = None


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def __init__(self, store: NodeStore):
        super().__init__()
        self._store = store
        self._drag_pos: Optional[QPoint] = None
        self._resize_dir: Optional[str] = None
        self._resize_start_geom: Optional[QRect] = None
        self._resize_start_pos: Optional[QPoint] = None

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Window)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setMinimumSize(760, 520)
        self.resize(1380, 860)

        self._tooltip = NodeToolTip()
        self._build_ui()
        self._load_nodes()
        self._auto_launch()
        # Start node scheduler
        global _node_scheduler
        _node_scheduler = NodeScheduler(self._store, self)
        _node_scheduler.set_main_window(self)
        _node_scheduler.start()
        # Start reminder scheduler
        global _reminder_scheduler
        try:
            _reminder_scheduler = ReminderScheduler(_get_reminder_store(), self)
            _reminder_scheduler.set_main_window(self)
            _reminder_scheduler.start()
        except Exception as _exc:
            print(f"[MainWindow] reminder scheduler init error: {_exc}", file=sys.stderr)
        # Initialize plugin manager after window is fully constructed
        _plugin_manager.set_main_window(self)
        _plugin_manager.scan_installed(self._store)

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(RESIZE_MARGIN, RESIZE_MARGIN,
                                RESIZE_MARGIN, RESIZE_MARGIN)
        root.setSpacing(0)

        self._inner = QWidget()
        self._inner.setObjectName("inner_frame")
        il = QVBoxLayout(self._inner)
        il.setContentsMargins(0, 0, 0, 0); il.setSpacing(0)

        self._titlebar = CustomTitleBar(APP_NAME, self._inner)
        self._titlebar.set_main_window(self)
        self._titlebar.help_requested.connect(self._open_help)
        self._titlebar.search_requested.connect(self._toggle_search)
        self._titlebar.quick_connect_requested.connect(self._open_quick_connect)
        il.addWidget(self._titlebar)

        self._search_bar = _SearchBar(self._inner)
        self._search_bar.setVisible(False)
        self._search_bar.search_changed.connect(self._apply_search_filter)
        il.addWidget(self._search_bar)

        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setStyleSheet(self._scroll_css())

        self._canvas = NodeCanvas(self._tooltip, self._store, self._scroll)
        self._canvas.add_node_requested.connect(self._on_add_node)
        self._canvas.node_edit_requested.connect(self._on_edit_node)
        self._canvas.node_delete_requested.connect(self._on_delete_node)
        self._canvas.node_archive_requested.connect(self._on_archive_node)
        self._canvas.node_launch_requested.connect(self._launch_node)
        self._canvas.node_export_requested.connect(self._on_export_node)
        self._canvas.node_duplicate_requested.connect(self._on_duplicate_node)
        self._canvas.node_remove_from_folder_requested.connect(self._on_remove_from_folder)
        self._canvas.folder_open_requested.connect(self._on_open_folder)
        self._canvas.folder_rename_requested.connect(self._on_rename_folder)
        self._canvas.folder_delete_requested.connect(self._on_delete_folder)
        self._canvas.folder_delete_with_contents_requested.connect(self._on_delete_folder_and_contents)
        self._canvas.folder_empty_requested.connect(self._on_empty_folder)
        self._canvas.order_changed.connect(self._store.reorder_items)
        self._canvas.files_dropped.connect(self._on_files_dropped)
        self._canvas.icon_set_requested.connect(self._on_icon_set_requested)
        # Canvas context-menu signals
        self._canvas.canvas_new_node_requested.connect(self._on_add_node)
        self._canvas.canvas_new_folder_requested.connect(self._on_new_folder)
        self._canvas.canvas_refresh_requested.connect(self._load_nodes)
        self._canvas.canvas_settings_requested.connect(self._open_settings)
        self._canvas.canvas_help_requested.connect(self._open_help)
        # Batch multi-select signals
        self._canvas.batch_delete_nodes_requested.connect(self._on_batch_delete_nodes)
        self._canvas.batch_delete_folders_requested.connect(self._on_batch_delete_folders)
        self._canvas.batch_delete_all_requested.connect(self._on_batch_delete_all)
        self._canvas.batch_archive_nodes_requested.connect(self._on_batch_archive_nodes)
        self._canvas.batch_empty_folders_requested.connect(self._on_batch_empty_folders)
        self._canvas.batch_place_in_folder_requested.connect(self._on_batch_place_in_folder)
        self._scroll.setWidget(self._canvas)
        il.addWidget(self._scroll)

        self._footer = FooterToolBar(self._inner)
        self._footer.settings_requested.connect(self._open_settings)
        self._footer.clipboard_requested.connect(self._open_clipboard_manager)
        self._footer.plugins_requested.connect(self._open_plugins)
        self._footer.new_folder_requested.connect(self._on_new_folder)
        self._footer.time_tracker_requested.connect(self._open_time_tracker)
        self._time_tracker_hud: Optional[TimeClockHUD] = None
        self._footer.notebook_requested.connect(self._open_notebook)
        self._notebook_win: Optional[NotebookWindow] = None
        self._footer.reminders_requested.connect(self._open_reminders)
        self._reminder_win: Optional[ReminderWindow] = None
        self._footer.media_library_requested.connect(self._open_media_library)
        self._media_library_win: Optional[MediaLibraryWindow] = None
        self._footer.lock_screen_requested.connect(self._lock_screen)
        il.addWidget(self._footer)

        # ── Clipboard Manager (persistent top-level window) ────────────────
        global _clip_store
        self._clipboard_win: Optional[ClipboardManagerWindow] = None
        self._global_hotkey = _GlobalHotkeyListener()
        self._global_hotkey.hotkey_triggered.connect(self._on_global_hotkey)
        self._apply_clipboard_enabled(startup=True)
        self._global_hotkey.register_paste_plain_hotkey()

        root.addWidget(self._inner)

        self._dialog_scrim = _DialogScrim(self._inner)

        self._filter_chip = _FilterChip(self._inner)
        self._filter_chip.clear_requested.connect(self._clear_search_filter)

        from PySide6.QtGui import QShortcut, QKeySequence
        self._shortcuts: list = []
        self._multi_key_bindings: list = []   # [(frozenset[str], callable)]
        self._held_keys: set = set()
        self._multi_keys_enabled: bool = True
        self._key_tracker = _MainWindowKeyTracker(self)
        QApplication.instance().installEventFilter(self._key_tracker)
        self._build_shortcuts()
        self._apply_hotkeys_enabled()

        self._glow = GlowBorderWidget(central)
        self._glow.raise_()
        # Register for theme change notifications so the window repaints
        _theme.register(self._on_theme_changed)
        self._sync_titlebar_anim()

    def _sync_titlebar_anim(self):
        """Apply the 'disable title bar animation' setting to the title label."""
        if hasattr(self, "_titlebar") and hasattr(self._titlebar, "_title_label"):
            disabled = _settings_store.value("disable_titlebar_anim", "false") == "true"
            self._titlebar._title_label.set_animation_enabled(not disabled)

    def _hotkey_action_map(self) -> list:
        """Return list of (action_id, slot) pairs for all configurable shortcuts."""
        return [
            ("toggle_search",     self._toggle_search),
            ("new_node",          self._on_add_node),
            ("new_folder",        self._on_new_folder),
            ("open_settings",     self._open_settings),
            ("open_help",         self._open_help),
            ("time_tracker",      self._open_time_tracker),
            ("notebook",          self._open_notebook),
            ("media_library",     self._open_media_library),
            ("quick_connect",     self._open_quick_connect),
            ("clipboard_manager", self._open_clipboard_manager),
        ]

    def _build_shortcuts(self):
        """Create QShortcut or multi-key-binding entries from current hotkey settings.

        Combos with more than one non-modifier key (e.g. ``Ctrl+Q+W``) cannot
        be expressed as a single QKeySequence chord, so they are registered in
        ``_multi_key_bindings`` and handled by ``_MainWindowKeyTracker``.
        """
        from PySide6.QtGui import QShortcut, QKeySequence
        _MODS = {"ctrl", "shift", "alt", "meta"}
        for action_id, slot in self._hotkey_action_map():
            key = _get_hotkey(action_id)
            if not key:
                continue
            parts = [p.strip() for p in key.split("+")]
            non_mods = [p for p in parts if p.lower() not in _MODS]
            if len(non_mods) > 1:
                # Multi-non-modifier chord — handled by _MainWindowKeyTracker
                key_set = frozenset(p.lower() for p in parts)
                self._multi_key_bindings.append((key_set, slot))
            else:
                try:
                    sc = QShortcut(QKeySequence(key), self)
                    sc.activated.connect(slot)
                    self._shortcuts.append(sc)
                except Exception as exc:
                    print(f"[CommandCenter] Shortcut create error ({action_id}, {key!r}): {exc}",
                          file=sys.stderr)

    def _rebuild_shortcuts(self):
        """Destroy all existing shortcuts and recreate them from current settings.

        Called after the Settings dialog is accepted so hotkey changes take
        effect immediately without requiring a restart.
        """
        for sc in self._shortcuts:
            try:
                sc.setEnabled(False)
                sc.setParent(None)
                sc.deleteLater()
            except Exception:
                pass
        self._shortcuts.clear()
        self._multi_key_bindings.clear()
        self._held_keys.clear()
        self._build_shortcuts()
        self._apply_hotkeys_enabled()

    def _apply_hotkeys_enabled(self):
        """Enable or disable all app-wide keyboard shortcuts per user setting."""
        enabled = _settings_store.value("disable_hotkeys", "false") != "true"
        for sc in getattr(self, "_shortcuts", []):
            sc.setEnabled(enabled)
        self._multi_keys_enabled = enabled

    def _check_multi_key_bindings(self):
        """Fire a multi-key binding when the held-key set exactly matches one."""
        if not getattr(self, "_multi_keys_enabled", True):
            return
        if not self.isActiveWindow():
            return
        # Don't fire shortcuts when a text input widget has focus
        from PySide6.QtWidgets import QLineEdit, QTextEdit, QPlainTextEdit
        fw = QApplication.focusWidget()
        if isinstance(fw, (QLineEdit, QTextEdit, QPlainTextEdit)):
            return
        for key_set, slot in self._multi_key_bindings:
            if key_set == self._held_keys:
                self._held_keys.clear()
                slot()
                break

    def _on_theme_changed(self):
        self.update()
        if hasattr(self, "_scroll"):
            self._scroll.setStyleSheet(self._scroll_css())
        if hasattr(self, "_canvas"):
            self._canvas.update()
            for w in self._canvas._tiles + self._canvas._folders:
                w.update()
        if hasattr(self, "_titlebar"):
            self._titlebar.refresh_theme()
        if hasattr(self, "_search_bar"):
            self._search_bar.refresh_theme()
        if hasattr(self, "_footer"):
            self._footer.refresh_theme()

    def _scroll_css(self):
        t = _theme
        return f"""
        QScrollArea {{ background:transparent; border:none; }}
        QScrollBar:vertical {{
            background:rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},160); width:7px; border-radius:3px; margin:0;
        }}
        QScrollBar::handle:vertical {{
            background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},110); border-radius:3px; min-height:20px;
        }}
        QScrollBar::handle:vertical:hover {{ background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},160); }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height:0; }}
        QScrollBar:horizontal {{
            background:rgba({t.BG_DARK.red()},{t.BG_DARK.green()},{t.BG_DARK.blue()},160); height:7px; border-radius:3px;
        }}
        QScrollBar::handle:horizontal {{
            background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},110); border-radius:3px; min-width:20px;
        }}
        QScrollBar::handle:horizontal:hover {{ background:rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},160); }}
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width:0; }}
        """

    def _load_nodes(self):
        self._canvas.load_items(self._store.all_items())
        _plugin_manager.notify_nodes_loaded()

    def _auto_launch(self):
        if _settings_store.value("global_auto_launch", "true") == "false":
            return
        for node in self._store.all_nodes():
            if node.get("auto_launch"):
                self._launch_node(node)
        if _settings_store.value("auto_launch_time_tracker", "false") == "true":
            QTimer.singleShot(500, self._open_time_tracker)

    # ------------------------------------------------------------------
    # Search / filter
    # ------------------------------------------------------------------

    def _toggle_search(self):
        visible = not self._search_bar.isVisible()
        self._search_bar.setVisible(visible)
        if visible:
            self._search_bar.focus_edit()
        else:
            self._search_bar.clear()   # also fires _apply_search_filter("")

    def _close_search_if_open(self):
        if self._search_bar.isVisible():
            self._search_bar.clear()
            self._search_bar.setVisible(False)

    def _apply_search_filter(self, query: str):
        q = query.strip().lower()
        self._canvas.apply_filter(q)
        if q:
            self._filter_chip.set_query(query.strip())
            self._filter_chip.setVisible(True)
            self._filter_chip.raise_()
            self._reposition_filter_chip()
        else:
            self._filter_chip.setVisible(False)

    def _clear_search_filter(self):
        """Clear triggered by the chip's × button."""
        self._search_bar.clear()           # fires search_changed → _apply_search_filter("")
        if not self._search_bar.isVisible():
            # If bar is already hidden, just ensure tiles are all unhidden
            self._canvas.apply_filter("")

    def _reposition_filter_chip(self):
        chip = self._filter_chip
        chip.adjustSize()
        margin = 10
        tb_h = self._titlebar.height()
        sb_h = self._search_bar.height() if self._search_bar.isVisible() else 0
        x = self._inner.width() - chip.width() - margin
        y = tb_h + sb_h + margin
        chip.move(x, y)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        if hasattr(self, "_filter_chip") and self._filter_chip.isVisible():
            self._reposition_filter_chip()

    # ------------------------------------------------------------------
    # Scrim helpers
    # ------------------------------------------------------------------

    def _show_scrim(self):

        self._dialog_scrim.setGeometry(self._inner.rect())
        self._dialog_scrim.raise_()
        self._dialog_scrim.show()

    def _hide_scrim(self):
        self._dialog_scrim.hide()

    def _on_add_node(self):
        wiz = NodeWizard(parent=self)
        wiz.node_saved.connect(self._save_new_node)
        self._center_dialog(wiz)
        wiz.finished.connect(self._hide_scrim)
        self._show_scrim()
        wiz.show(); wiz.raise_(); wiz.activateWindow()

    def _on_edit_node(self, node: dict):
        wiz = NodeWizard(existing_node=node, parent=self)
        wiz.node_saved.connect(self._save_updated_node)
        self._center_dialog(wiz)
        wiz.finished.connect(self._hide_scrim)
        self._show_scrim()
        wiz.show(); wiz.raise_(); wiz.activateWindow()

    def _save_new_node(self, node: dict):
        saved = self._store.add_node(node)
        self._canvas.add_tile_for_node(saved)
        self._set_status(f"Node '{saved['name']}' created.")

    def _save_updated_node(self, node: dict):
        self._store.update_node(node["id"], node)
        self._canvas.update_tile_for_node(node)
        self._set_status(f"Node '{node['name']}' updated.")

    def _on_duplicate_node(self, node: dict):
        import copy
        new_node = copy.deepcopy(node)
        new_node.pop("id", None)        # NodeStore assigns a fresh id
        new_node.pop("grid_order", None)  # placed at end of grid
        new_node["name"] = f"Copy of {node.get('name', '')}"
        new_node["archived"] = False
        saved = self._store.add_node(new_node)
        self._canvas.add_tile_for_node(saved)
        self._set_status(f"Duplicated \u2018{node['name']}\u2019 \u2192 \u2018{saved['name']}\u2019.")

    def _on_delete_node(self, node: dict):
        if _settings_store.value("confirm_delete", "false") == "true":
            if not self._confirm_action(
                    "Delete",
                    f"Permanently delete '{node.get('name','')}'?  This cannot be undone."):
                return
        self._store.remove_node(node["id"])
        self._canvas.remove_tile_for_node(node["id"])
        self._set_status(f"Node '{node['name']}' deleted.")

    def _on_archive_node(self, node: dict):
        if _settings_store.value("confirm_archive", "false") == "true":
            if not self._confirm_action(
                    "Archive",
                    f"Archive '{node.get('name','')}'?  You can restore it later from Settings."):
                return
        self._store.archive_node(node["id"])
        self._canvas.remove_tile_for_node(node["id"])
        self._set_status(f"Node '{node['name']}' archived.")

    def _on_export_node(self, node: dict):
        default_name = node.get("name","node").replace(" ","_") + NODE_FILE_EXT
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Node", default_name,
            f"Node Files (*{NODE_FILE_EXT});;All Files (*)")
        if path:
            try:
                NodeStore.export_node(node, path)
                self._set_status(f"Node exported: {Path(path).name}")
            except OSError as exc:
                self._set_status(f"Export error: {exc}")

    def _on_new_folder(self):
        name, ok = self._inline_rename_dialog("New Folder", "Folder name:")
        if not ok or not name.strip(): return
        folder = self._store.add_folder(name.strip())
        self._canvas.add_tile_for_folder(folder)
        self._set_status(f"Folder '{folder['name']}' created.")

    def _on_open_folder(self, folder: dict):
        dlg = FolderViewDialog(folder, self._store, self._tooltip, self)
        dlg.node_launch_requested.connect(self._launch_node)
        dlg.node_edit_requested.connect(self._on_edit_node)
        dlg.node_delete_requested.connect(self._on_delete_node)
        dlg.node_archive_requested.connect(self._on_archive_node)
        dlg.node_export_requested.connect(self._on_export_node)
        dlg.node_duplicate_requested.connect(self._on_duplicate_node)
        dlg.node_remove_from_folder_requested.connect(
            lambda node, d=dlg: self._on_remove_from_folder(node, d))
        self._center_dialog(dlg)
        dlg.show(); dlg.raise_(); dlg.activateWindow()

    def _on_remove_from_folder(self, node: dict, dlg=None):
        """Move a node out of its folder and back to the main canvas.
        Called both from the folder popup (dlg is set) and from search results
        surfaced directly on the canvas (dlg is None).
        """
        node_id   = node["id"]
        folder_id = node.get("folder_id")
        self._store.move_node_to_folder(node_id, None)  # clears folder_id + children list
        # Refresh the tile data from store so folder_id is gone
        updated = next((n for n in self._store._data if n.get("id") == node_id), node)
        self._canvas.add_tile_for_node(updated)
        # Close the folder popup if one is open
        if dlg is not None:
            dlg.reject()
        # Refresh the folder tile's child count badge
        if folder_id is not None:
            updated_folder = next(
                (n for n in self._store.all_items() if n["id"] == folder_id), None)
            if updated_folder:
                self._canvas.update_folder_tile(updated_folder)
        # Re-run the current search so the search tile is replaced by the real tile
        if self._search_bar.isVisible():
            self._apply_search_filter(self._search_bar.current_query())
        self._set_status(f"'{node.get('name','')}' moved back to canvas.")

    def _on_rename_folder(self, folder: dict):
        name, ok = self._inline_rename_dialog("Rename Folder", "New name:", folder.get("name",""))
        if not ok or not name.strip(): return
        self._store.update_node(folder["id"], {"name": name.strip()})
        updated = next((n for n in self._store.all_items() if n["id"] == folder["id"]), folder)
        self._canvas.update_folder_tile(updated)
        self._set_status(f"Folder renamed to '{name.strip()}'.")

    def _on_delete_folder(self, folder: dict):
        """Delete the folder tile but dump its children back to the main canvas."""
        children_ids = list(folder.get("children", []))
        for node_id in children_ids:
            self._store.move_node_to_folder(node_id, None)
            node = next((n for n in self._store.all_nodes() if n["id"] == node_id), None)
            if node:
                self._canvas.add_tile_for_node(node)
        self._store.remove_node(folder["id"])
        self._canvas.remove_folder_tile(folder["id"])
        self._set_status(
            f"Folder '{folder.get('name','')}' deleted"
            + (f" — {len(children_ids)} node(s) returned to canvas." if children_ids else "."))

    def _on_delete_folder_and_contents(self, folder: dict):
        """Delete the folder and permanently remove all its child nodes."""
        children_ids = list(folder.get("children", []))
        for node_id in children_ids:
            self._store.remove_node(node_id)
        self._store.remove_node(folder["id"])
        self._canvas.remove_folder_tile(folder["id"])
        self._set_status(
            f"Folder '{folder.get('name','')}' and {len(children_ids)} node(s) permanently deleted.")

    def _on_empty_folder(self, folder: dict):
        """Move all children of a folder back to the root canvas."""
        children_ids = list(folder.get("children", []))
        if not children_ids:
            self._set_status(f"Folder '{folder.get('name','')}' is already empty.")
            return
        for node_id in children_ids:
            self._store.move_node_to_folder(node_id, None)
            node = next((n for n in self._store.all_nodes() if n["id"] == node_id), None)
            if node:
                self._canvas.add_tile_for_node(node)
        # Refresh the folder tile to show 0 children
        updated = next((n for n in self._store.all_items() if n["id"] == folder["id"]), None)
        if updated:
            self._canvas.update_folder_tile(updated)
        self._set_status(
            f"Emptied folder '{folder.get('name','')}' — {len(children_ids)} node(s) returned to canvas.")

    # ------------------------------------------------------------------
    # Batch multi-select operations
    # ------------------------------------------------------------------

    def _on_batch_delete_nodes(self, node_ids: list):
        if not node_ids: return
        if _settings_store.value("confirm_delete", "false") == "true":
            if not self._confirm_action(
                    "Delete Nodes",
                    f"Permanently delete {len(node_ids)} node(s)?  This cannot be undone."):
                return
        for nid in node_ids:
            node = next((n for n in self._store.all_nodes() if n["id"] == nid), None)
            if node:
                self._store.remove_node(nid)
                self._canvas.remove_tile_for_node(nid)
        self._set_status(f"Deleted {len(node_ids)} node(s).")

    def _on_batch_delete_folders(self, folder_ids: list):
        if not folder_ids: return
        if _settings_store.value("confirm_delete", "false") == "true":
            if not self._confirm_action(
                    "Delete Folders",
                    f"Delete {len(folder_ids)} folder(s) and return their contents to the canvas?"):
                return
        for fid in folder_ids:
            folder = next((n for n in self._store.all_items() if n["id"] == fid), None)
            if folder:
                self._on_delete_folder(folder)
        self._set_status(f"Deleted {len(folder_ids)} folder(s).")

    def _on_batch_delete_all(self, node_ids: list, folder_ids: list):
        total = len(node_ids) + len(folder_ids)
        if not total: return
        if _settings_store.value("confirm_delete", "false") == "true":
            if not self._confirm_action(
                    "Delete All Selected",
                    f"Permanently delete {total} item(s)?  This cannot be undone."):
                return
        for nid in node_ids:
            self._store.remove_node(nid)
            self._canvas.remove_tile_for_node(nid)
        for fid in folder_ids:
            folder = next((n for n in self._store.all_items() if n["id"] == fid), None)
            if folder:
                children_ids = list(folder.get("children", []))
                for child_id in children_ids:
                    self._store.remove_node(child_id)
                self._store.remove_node(fid)
                self._canvas.remove_folder_tile(fid)
        self._set_status(f"Deleted {total} selected item(s).")

    def _on_batch_archive_nodes(self, node_ids: list):
        if not node_ids: return
        if _settings_store.value("confirm_archive", "false") == "true":
            if not self._confirm_action(
                    "Archive Nodes",
                    f"Archive {len(node_ids)} node(s)?  You can restore them later from Settings."):
                return
        for nid in node_ids:
            self._store.archive_node(nid)
            self._canvas.remove_tile_for_node(nid)
        self._set_status(f"Archived {len(node_ids)} node(s).")

    def _on_batch_empty_folders(self, folder_ids: list):
        if not folder_ids: return
        total_returned = 0
        for fid in folder_ids:
            folder = next((n for n in self._store.all_items() if n["id"] == fid), None)
            if folder:
                children_ids = list(folder.get("children", []))
                for child_id in children_ids:
                    self._store.move_node_to_folder(child_id, None)
                    node = next((n for n in self._store.all_nodes() if n["id"] == child_id), None)
                    if node:
                        self._canvas.add_tile_for_node(node)
                        total_returned += 1
                updated = next((n for n in self._store.all_items() if n["id"] == fid), None)
                if updated:
                    self._canvas.update_folder_tile(updated)
        self._set_status(f"Emptied {len(folder_ids)} folder(s) — {total_returned} node(s) returned to canvas.")

    def _on_batch_place_in_folder(self, node_ids: list):
        if not node_ids: return
        name, ok = self._inline_rename_dialog("New Folder", "Folder name for selected nodes:")
        if not ok or not name.strip(): return
        folder = self._store.add_folder(name.strip())
        self._canvas.add_tile_for_folder(folder)
        for nid in node_ids:
            self._store.move_node_to_folder(nid, folder["id"])
            self._canvas.remove_tile_for_node(nid)
        # Refresh folder tile to show correct child count
        updated = next((n for n in self._store.all_items() if n["id"] == folder["id"]), None)
        if updated:
            self._canvas.update_folder_tile(updated)
        self._set_status(f"Placed {len(node_ids)} node(s) into new folder '{name.strip()}'.")

    def _inline_rename_dialog(self, title: str, label: str,
                               current: str = "") -> tuple[str, bool]:
        t = _theme
        dlg = QDialog(self, Qt.FramelessWindowHint)
        dlg.setFixedSize(340, 150)
        # Solid background container so the dialog is never transparent
        container = QWidget(dlg)
        container.setGeometry(0, 0, 340, 150)
        container.setObjectName("dlg_bg")
        container.setStyleSheet(
            f"QWidget#dlg_bg {{ "
            f"background: rgba({t.BG_MID.red()},{t.BG_MID.green()},{t.BG_MID.blue()},245); "
            f"border: 1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},90); "
            f"border-radius: 10px; }}"
        )
        layout = QVBoxLayout(container); layout.setContentsMargins(20,16,20,16); layout.setSpacing(10)
        lbl = QLabel(label); lbl.setFont(FONT_LABEL)
        lbl.setStyleSheet(f"color:{t.TEXT_SECONDARY.name()};")
        layout.addWidget(lbl)
        le = QLineEdit(current); le.setFont(FONT_LABEL)
        le.setMaxLength(80)
        le.setStyleSheet(f"""
            QLineEdit {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},210);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:5px; color:{t.TEXT_PRIMARY.name()};
                padding:5px 9px; font-family:'Segoe UI'; font-size:9pt;
            }}
            QLineEdit:focus {{ border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},150); }}
        """)
        le.selectAll(); layout.addWidget(le)
        brow = QHBoxLayout(); brow.addStretch()
        ok_btn = QPushButton("OK"); ok_btn.setFixedHeight(28)
        ok_btn.setFont(QFont("Segoe UI",9,QFont.Weight.DemiBold))
        ok_btn.setCursor(Qt.PointingHandCursor)
        ok_btn.setStyleSheet(f"""
            QPushButton {{
                background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 rgba({t.ACCENT_BLUE.red()},{t.ACCENT_BLUE.green()},{t.ACCENT_BLUE.blue()},180),
                    stop:1 rgba({t.ACCENT_TEAL.red()},{t.ACCENT_TEAL.green()},{t.ACCENT_TEAL.blue()},180));
                border:none; border-radius:5px;
                color:{t.TEXT_PRIMARY.name()}; padding:0 14px;
            }}
            QPushButton:hover {{ background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                stop:0 rgba({t.ACCENT_BLUE.red()},{t.ACCENT_BLUE.green()},{t.ACCENT_BLUE.blue()},220),
                stop:1 rgba({t.ACCENT_TEAL.red()},{t.ACCENT_TEAL.green()},{t.ACCENT_TEAL.blue()},220)); }}
        """)
        ok_btn.clicked.connect(dlg.accept)
        le.returnPressed.connect(dlg.accept)
        brow.addWidget(ok_btn); layout.addLayout(brow)
        self._center_dialog(dlg)
        result = dlg.exec()
        return le.text(), result == QDialog.Accepted

    def _launch_node(self, node: dict):
        nt   = node.get("type", NODE_TYPE_FILE)
        name = node.get("name", "")
        try:
            if nt == NODE_TYPE_FILE:
                target = node.get("target", "")
                ob     = node.get("open_behavior", "normal")
                if ob == "open_folder":
                    # Open the folder that contains the file (or the folder itself)
                    folder = (os.path.dirname(os.path.abspath(target))
                              if os.path.isfile(target) else target)
                    subprocess.Popen(["explorer", folder])
                    self._set_status(f"Opened folder for: {name}")
                elif ob == "run_admin":
                    ret = ctypes.windll.shell32.ShellExecuteW(
                        None, "runas", target,
                        node.get("params", "") or None, None, 1)
                    if ret <= 32:
                        raise OSError(
                            f"ShellExecuteW returned {ret} — user may have cancelled UAC.")
                    self._set_status(f"Launched as admin: {name}")
                elif ob == "copy_file":
                    if not os.path.isfile(target):
                        raise OSError(f"File not found: {target}")
                    mime = QMimeData()
                    mime.setUrls([QUrl.fromLocalFile(os.path.abspath(target))])
                    QApplication.clipboard().setMimeData(mime)
                    self._set_status(f"File copied to clipboard: {name}")
                    self.toast(f"📋 File copied to clipboard: {name}", "success")
                elif ob == "copy_contents":
                    if not os.path.isfile(target):
                        raise OSError(f"File not found: {target}")
                    with open(target, "r", encoding="utf-8", errors="replace") as f:
                        text = f.read()
                    QApplication.clipboard().setText(text)
                    self._set_status(f"File contents copied to clipboard: {name}")
                    self.toast(f"📋 Contents copied to clipboard: {name}", "success")
                else:  # "normal" or legacy nodes with no open_behavior key
                    launch_path(target, node.get("params", ""))
                    self._set_status(f"Launched: {name}")
            elif nt == NODE_TYPE_URL:
                url = node.get("target","").strip()
                if not url.startswith(("http://","https://","ftp://")):
                    url = "https://" + url
                pu = QUrl(url)
                if pu.isValid() and pu.scheme() in ("http","https","ftp"):
                    webbrowser.open(url)
                    self._set_status(f"Opened: {name}")
            elif nt == NODE_TYPE_NOTE:
                note_id = node.get("target", "")
                self._open_notebook()
                if note_id and self._notebook_win:
                    self._notebook_win._load_note(note_id)
                    self._notebook_win._side._select_item(note_id)
        except OSError as exc:
            msg = f"Error launching \u2018{name}\u2019: {exc}"
            self._set_status(msg)
            self.toast(msg, "error")

    def _show_note(self, node: dict):
        t = _theme
        dlg = QDialog(self, Qt.FramelessWindowHint)
        dlg.setAttribute(Qt.WA_TranslucentBackground); dlg.setMinimumSize(420,300)
        layout = QVBoxLayout(dlg); layout.setContentsMargins(20,20,20,20)
        lbl = QLabel(node.get("name","Note"))
        lbl.setTextFormat(Qt.PlainText)
        lbl.setFont(QFont("Segoe UI",11,QFont.Weight.DemiBold))
        lbl.setStyleSheet(f"color:{t.TEXT_PRIMARY.name()};"); layout.addWidget(lbl)
        te = QTextEdit(); te.setPlainText(node.get("note","")); te.setReadOnly(True)
        te.setStyleSheet(f"""
            QTextEdit {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},220);
                border:1px solid rgba({t.GLOW.red()},{t.GLOW.green()},{t.GLOW.blue()},70);
                border-radius:6px; color:{t.TEXT_PRIMARY.name()};
                font-family:'Segoe UI'; font-size:9pt; padding:8px;
            }}
        """)
        layout.addWidget(te)
        btn = QPushButton("Close"); btn.setFont(FONT_LABEL)
        btn.setCursor(Qt.PointingHandCursor); btn.clicked.connect(dlg.accept)
        btn.setStyleSheet(f"""
            QPushButton {{
                background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 rgba({t.ACCENT_BLUE.red()},{t.ACCENT_BLUE.green()},{t.ACCENT_BLUE.blue()},165),
                    stop:1 rgba({t.ACCENT_TEAL.red()},{t.ACCENT_TEAL.green()},{t.ACCENT_TEAL.blue()},165));
                border:none; border-radius:6px;
                color:{t.TEXT_PRIMARY.name()}; padding:6px 16px;
            }}
            QPushButton:hover {{ background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                stop:0 rgba({t.ACCENT_BLUE.red()},{t.ACCENT_BLUE.green()},{t.ACCENT_BLUE.blue()},210),
                stop:1 rgba({t.ACCENT_TEAL.red()},{t.ACCENT_TEAL.green()},{t.ACCENT_TEAL.blue()},210)); }}
        """)
        layout.addWidget(btn, alignment=Qt.AlignRight)
        self._center_dialog(dlg); dlg.exec()

    def _maybe_show_tip(self):
        """Show the Tip of the Day dialog if the user hasn't disabled it."""
        if _settings_store.value("show_tips", "true") == "true":
            dlg = TipOfDayDialog(self)
            # No _center_dialog — dialog covers the parent as a full-screen scrim
            dlg.exec()

    def _confirm_action(self, action: str, message: str) -> bool:
        """Show a small themed yes/no dialog. Returns True if the user confirms."""
        t = _theme
        dlg = QDialog(self, Qt.FramelessWindowHint)
        dlg.setAttribute(Qt.WA_TranslucentBackground)
        dlg.setFixedSize(360, 155)
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(20, 16, 20, 16); layout.setSpacing(10)
        ttl = QLabel(action)
        ttl.setFont(QFont("Segoe UI", 10, QFont.Weight.DemiBold))
        ttl.setStyleSheet(f"color:{t.TEXT_PRIMARY.name()};")
        layout.addWidget(ttl)
        msg_lbl = QLabel(message)
        msg_lbl.setFont(FONT_SMALL)
        msg_lbl.setStyleSheet(f"color:{t.TEXT_SECONDARY.name()};")
        msg_lbl.setWordWrap(True)
        layout.addWidget(msg_lbl)
        brow = QHBoxLayout(); brow.addStretch()
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setFont(FONT_LABEL); cancel_btn.setFixedHeight(28)
        cancel_btn.setCursor(Qt.PointingHandCursor)
        g = t.GLOW
        cancel_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},180);
                border:1px solid rgba({g.red()},{g.green()},{g.blue()},70);
                border-radius:5px; color:{t.TEXT_SECONDARY.name()}; padding:0 12px;
            }}
            QPushButton:hover {{ color:{g.name()}; }}
        """)
        cancel_btn.clicked.connect(dlg.reject)
        confirm_btn = QPushButton(action)
        confirm_btn.setFont(QFont("Segoe UI", 9, QFont.Weight.DemiBold))
        confirm_btn.setFixedHeight(28); confirm_btn.setCursor(Qt.PointingHandCursor)
        ar = t.ACCENT_RED
        confirm_btn.setStyleSheet(f"""
            QPushButton {{
                background:rgba({ar.red()//2},{ar.green()//3},{ar.blue()//3},160);
                border:1px solid rgba({ar.red()},{ar.green()},{ar.blue()},90);
                border-radius:5px; color:{t.TEXT_PRIMARY.name()}; padding:0 12px;
            }}
            QPushButton:hover {{
                background:rgba({ar.red()},{ar.green()},{ar.blue()},200);
            }}
        """)
        confirm_btn.clicked.connect(dlg.accept)
        brow.addWidget(cancel_btn); brow.addSpacing(6); brow.addWidget(confirm_btn)
        layout.addLayout(brow)
        self._center_dialog(dlg)
        return dlg.exec() == QDialog.Accepted

    def _open_help(self):
        dlg = HelpDialog(self)
        self._center_dialog(dlg)
        dlg.finished.connect(self._hide_scrim)
        self._show_scrim()
        dlg.show(); dlg.raise_(); dlg.activateWindow()

    def _open_quick_connect(self):
        dlg = QuickConnectDialog(self)
        self._center_dialog(dlg)
        dlg.finished.connect(self._hide_scrim)
        self._show_scrim()
        dlg.show(); dlg.raise_(); dlg.activateWindow()
        dlg._client_edit.setFocus()

    def _apply_clipboard_enabled(self, startup: bool = False):
        """Start or stop the clipboard manager based on the current setting."""
        global _clip_store
        enabled = _settings_store.value("clipboard_manager_enabled", "true") == "true"
        if enabled:
            if _clip_store is None:
                _clip_store = ClipboardStore(self)
            self._global_hotkey.register_clipboard_hotkey()
            # Show footer clipboard button
            if hasattr(self, "_footer"):
                for btn in self._footer.findChildren(QPushButton):
                    if btn.text() == "📋 Clipboard":
                        btn.setVisible(True)
                        break
        else:
            # Tear down
            self._global_hotkey.unregister_clipboard_hotkey()
            if _clip_store is not None:
                _clip_store.deleteLater()
                _clip_store = None
            if self._clipboard_win is not None:
                self._clipboard_win.hide()
            # Hide footer clipboard button
            if hasattr(self, "_footer"):
                for btn in self._footer.findChildren(QPushButton):
                    if btn.text() == "📋 Clipboard":
                        btn.setVisible(False)
                        break

    def _on_global_hotkey(self, hotkey_id: int):
        if hotkey_id == _GlobalHotkeyListener.HOTKEY_CLIPBOARD:
            self._open_clipboard_manager()
        elif hotkey_id == _GlobalHotkeyListener.HOTKEY_PASTE_PLAIN:
            self._do_paste_plain()

    def _do_paste_plain(self):
        """Global Ctrl+Shift+V: strip clipboard to plain text, then forward Ctrl+V."""
        try:
            cb  = QApplication.clipboard()
            mime = cb.mimeData()
            if mime is None:
                return
            plain: Optional[str] = None
            if mime.hasText():
                plain = mime.text() or None
            elif mime.hasHtml():
                from PySide6.QtGui import QTextDocument as _QTD
                _doc = _QTD()
                _doc.setHtml(mime.html())
                plain = _doc.toPlainText() or None
            if not plain:
                return
            cb.setText(plain)
        except Exception:
            return
        # Small delay lets the hotkey key-events settle and the clipboard
        # propagate before we inject Ctrl+V into the foreground window.
        QTimer.singleShot(80, _GlobalHotkeyListener._inject_ctrl_v)

    def _open_clipboard_manager(self):
        if _settings_store.value("clipboard_manager_enabled", "true") != "true":
            return
        if self._clipboard_win is None:
            self._clipboard_win = ClipboardManagerWindow()
        win = self._clipboard_win
        # Always move to current cursor position and bring to front.
        # This means one click/hotkey always works regardless of whether
        # the window is hidden, minimized, or buried behind other windows.
        cursor_pos = QCursor.pos()
        screen = QGuiApplication.screenAt(cursor_pos)
        if screen is None:
            screen = QGuiApplication.primaryScreen()
        sg = screen.availableGeometry()
        win.adjustSize()
        x = cursor_pos.x() - win.width() // 2
        y = cursor_pos.y() - win.height() // 2
        x = max(sg.left() + 10, min(x, sg.right()  - win.width()  - 10))
        y = max(sg.top()  + 10, min(y, sg.bottom() - win.height() - 10))
        win.move(x, y)
        win.show()
        win.raise_()
        win.activateWindow()

    def _open_settings(self):
        dlg = SettingsDialog(self._store, self)
        dlg.settings_changed.connect(self._on_theme_changed)
        dlg.settings_changed.connect(self._load_nodes)
        dlg.settings_changed.connect(self._canvas.refresh_bg_style)
        dlg.settings_changed.connect(_plugin_manager.notify_settings_changed)
        dlg.settings_changed.connect(self._apply_clipboard_enabled)
        self._center_dialog(dlg)
        # Register so plugins can inject their settings tabs
        _plugin_manager._open_settings_dialogs.append(dlg)
        _plugin_manager._inject_settings_tabs()
        def _on_settings_closed():
            try:
                _plugin_manager._open_settings_dialogs.remove(dlg)
            except ValueError:
                pass
            self._hide_scrim()
            self._load_nodes()
            self._sync_titlebar_anim()
            self._rebuild_shortcuts()
        dlg.finished.connect(_on_settings_closed)
        self._show_scrim()
        dlg.show(); dlg.raise_(); dlg.activateWindow()

    def _open_plugins(self):
        dlg = PluginsDialog(self._store, self)
        self._center_dialog(dlg)
        dlg.finished.connect(self._hide_scrim)
        self._show_scrim()
        dlg.show(); dlg.raise_(); dlg.activateWindow()

    def _open_task_manager(self):
        try:
            subprocess.Popen(["taskmgr.exe"], shell=False)
        except OSError as e:
            self.toast(f"Could not open Task Manager: {e}", "error")

    def _open_notepad_plus(self):
        paths = [
            r"C:\Program Files\Notepad++\notepad++.exe",
            r"C:\Program Files (x86)\Notepad++\notepad++.exe",
        ]
        for p in paths:
            if os.path.isfile(p):
                try:
                    subprocess.Popen([p], shell=False)
                    return
                except OSError:
                    pass
        self.toast("Notepad++ not found. Please install it.", "error")

    def _lock_screen(self):
        ctypes.windll.user32.LockWorkStation()

    def _open_notebook(self):
        if self._notebook_win is None or not self._notebook_win.isVisible():
            self._notebook_win = NotebookWindow()
            self._notebook_win.node_creation_requested.connect(self._on_notebook_create_node)
            self._center_dialog(self._notebook_win)
            self._notebook_win.show()
            self._notebook_win.raise_()
        else:
            self._notebook_win.raise_()
            self._notebook_win.activateWindow()

    def _open_media_library(self):
        if self._media_library_win is None or not self._media_library_win.isVisible():
            self._media_library_win = MediaLibraryWindow()
            self._media_library_win.duplicate_detected.connect(
                lambda name: self.toast(
                    f"⚠ Already in library: {name}", "info"))
            self._center_dialog(self._media_library_win)
            self._media_library_win.show()
            self._media_library_win.raise_()
        else:
            self._media_library_win.raise_()
            self._media_library_win.activateWindow()

    def _open_reminders(self):
        try:
            store = _get_reminder_store()
            if self._reminder_win is None or not self._reminder_win.isVisible():
                self._reminder_win = ReminderWindow(store, _reminder_scheduler, self)
                self._center_dialog(self._reminder_win)
                self._reminder_win.show()
                self._reminder_win.raise_()
            else:
                self._reminder_win.raise_()
                self._reminder_win.activateWindow()
        except Exception as exc:
            print(f"[MainWindow] open_reminders error: {exc}", file=sys.stderr)
            self.toast("Could not open Reminders", "error")

    def _on_notebook_create_node(self, note_id: str, note_title: str):
        """Create a 1×1 Note node on the canvas linked to the given notebook note."""
        node = {
            "name": note_title or "Note",
            "type": NODE_TYPE_NOTE,
            "size": "1x1",
            "target": note_id,
            "note": "",
            "params": "",
            "description": f"Notebook note: {note_title}",
            "icon": "",
            "auto_launch": False,
            "archived": False,
        }
        saved = self._store.add_node(node)
        self._canvas.add_tile_for_node(saved)
        self.toast(f"📌 Note node created: {note_title}", "success")
        self._set_status(f"Note node created: {note_title}")

    # ── File drag-and-drop ─────────────────────────────────────────────────────

    _IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".ico", ".svg"}
    # Only genuine prose/document files become Notebook notes.
    # Scripts, code, config, and data files all open with their default app.
    _TEXT_EXTS  = {".txt", ".md", ".rst"}

    def _on_files_dropped(self, paths: list):
        for path_str in paths:
            p = Path(path_str)
            if not p.exists():
                continue
            suffix = p.suffix.lower()
            if suffix in self._IMAGE_EXTS:
                node = {
                    "name": p.stem,
                    "type": NODE_TYPE_FILE,
                    "size": "1x1",
                    "target": str(p),
                    "icon": str(p),
                    "note": "",
                    "params": "",
                    "description": f"Image: {p.name}",
                    "auto_launch": False,
                    "archived": False,
                }
                saved = self._store.add_node(node)
                self._canvas.add_tile_for_node(saved)
                self.toast(f"🖼️ Image node created: {p.name}", "success")
            elif suffix in self._TEXT_EXTS:
                try:
                    content = p.read_text(encoding="utf-8", errors="replace")
                    note_id = NotebookStore.new_id()
                    html = ("<p>" +
                            content.replace("&", "&amp;").replace("<", "&lt;")
                                   .replace("\n\n", "</p><p>")
                                   .replace("\n", "<br/>") +
                            "</p>")
                    NotebookStore.save_note(note_id, p.stem, html)
                    node = {
                        "name": p.stem,
                        "type": NODE_TYPE_NOTE,
                        "size": "1x1",
                        "target": note_id,
                        "icon": "",
                        "note": "",
                        "params": "",
                        "description": f"Notebook note from {p.name}",
                        "auto_launch": False,
                        "archived": False,
                    }
                    saved = self._store.add_node(node)
                    self._canvas.add_tile_for_node(saved)
                    if self._notebook_win and self._notebook_win.isVisible():
                        self._notebook_win._load_all_notes()
                    self.toast(f"📝 Note created from: {p.name}", "success")
                except Exception as exc:
                    self.toast(f"Error reading {p.name}: {exc}", "error")
            else:
                node = {
                    "name": p.stem,
                    "type": NODE_TYPE_FILE,
                    "size": "1x1",
                    "target": str(p),
                    "icon": "",
                    "note": "",
                    "params": "",
                    "description": f"File: {p.name}",
                    "auto_launch": False,
                    "archived": False,
                }
                saved = self._store.add_node(node)
                self._canvas.add_tile_for_node(saved)
                self.toast(f"📂 Node created: {p.name}", "success")

    def _on_icon_set_requested(self, node_id: int, image_path: str):
        """Drop of a single image onto an existing tile → update that tile's icon."""
        self._store.update_node(node_id, {"icon": image_path})
        load_pixmap_cached.cache_clear()   # flush stale cached pixmap
        for tile in self._canvas._tiles:
            if tile.node_data().get("id") == node_id:
                tile.node_data()["icon"] = image_path
                tile.update()
                break
        self.toast(f"🖼️ Icon updated: {Path(image_path).name}", "success")

    # ── Toast notifications ────────────────────────────────────────────────────

    def toast(self, msg: str, kind: str = "info"):
        """Show a brief themed notification near the bottom-right of the window."""
        t = _ToastNotification(msg, kind)
        geo = self.frameGeometry()
        t.adjustSize()
        # Stack toasts upward (offset by existing toasts is tricky; anchor to window edge)
        x = geo.right() - t.width() - 18
        y = geo.bottom() - t.height() - 18
        t.move(x, y)
        t.show()
        t.raise_()

    def _open_time_tracker(self):
        if self._time_tracker_hud is not None and self._time_tracker_hud.isVisible():
            # toggle: close and clear so the next open starts fresh
            self._time_tracker_hud._on_close()
            self._time_tracker_hud = None
        else:
            # always create a fresh instance so all timers reset to 0
            if self._time_tracker_hud is not None:
                try:
                    self._time_tracker_hud.deleteLater()
                except Exception:
                    pass
            self._time_tracker_hud = TimeClockHUD()
            self._time_tracker_hud.show()
            self._time_tracker_hud.raise_()

    def _set_status(self, msg: str):
        self._titlebar._title_label.set_transient(f"{APP_NAME}  -  {msg}")
        QTimer.singleShot(4000, lambda: self._titlebar.set_title(APP_NAME))

    def _center_dialog(self, dlg: QWidget):
        """Position dlg so its center aligns with the center of the main window."""
        dlg.adjustSize()
        sz = dlg.size()
        if sz.isEmpty() or sz == QSize(0, 0):
            sz = dlg.sizeHint()
        if sz.isEmpty():
            sz = dlg.minimumSize()
        mw_center = self.geometry().center()
        dlg.move(mw_center - QPoint(sz.width() // 2, sz.height() // 2))

    # ---- resize / drag ----

    def _get_resize_dir(self, pos: QPoint) -> Optional[str]:
        m = RESIZE_MARGIN; w, h = self.width(), self.height()
        x, y = pos.x(), pos.y()
        left = x < m; right = x > w - m; top = y < m; bottom = y > h - m
        if top and left:   return "tl"
        if top and right:  return "tr"
        if bottom and left:  return "bl"
        if bottom and right: return "br"
        if left:   return "l"
        if right:  return "r"
        if top:    return "t"
        if bottom: return "b"
        return None

    _CURSOR_MAP = {
        "tl": Qt.SizeFDiagCursor, "br": Qt.SizeFDiagCursor,
        "tr": Qt.SizeBDiagCursor, "bl": Qt.SizeBDiagCursor,
        "l":  Qt.SizeHorCursor,   "r":  Qt.SizeHorCursor,
        "t":  Qt.SizeVerCursor,   "b":  Qt.SizeVerCursor,
    }

    def mousePressEvent(self, e: QMouseEvent):
        if e.button() != Qt.LeftButton: return
        d = self._get_resize_dir(e.position().toPoint())
        if d:
            self._resize_dir = d
            self._resize_start_geom = self.geometry()
            self._resize_start_pos  = e.globalPosition().toPoint()
        else:
            self._drag_pos = e.globalPosition().toPoint()

    def mouseMoveEvent(self, e: QMouseEvent):
        if self._resize_dir and self._resize_start_pos:
            self._do_resize(e.globalPosition().toPoint()); return
        d = self._get_resize_dir(e.position().toPoint())
        self.setCursor(self._CURSOR_MAP.get(d, Qt.ArrowCursor)
                       if d else Qt.ArrowCursor)

    def mouseReleaseEvent(self, e: QMouseEvent):
        self._resize_dir = self._resize_start_geom = self._resize_start_pos = None
        self._drag_pos = None

    def _do_resize(self, gp: QPoint):
        if not (self._resize_start_geom and self._resize_start_pos): return
        dx = gp.x() - self._resize_start_pos.x()
        dy = gp.y() - self._resize_start_pos.y()
        g = QRect(self._resize_start_geom); d = self._resize_dir or ""
        if "r" in d: g.setRight(g.right() + dx)
        if "b" in d: g.setBottom(g.bottom() + dy)
        if "l" in d: g.setLeft(g.left() + dx)
        if "t" in d: g.setTop(g.top() + dy)
        if g.width() >= self.minimumWidth() and g.height() >= self.minimumHeight():
            self.setGeometry(g)

    # ---- paint ----

    def paintEvent(self, e: QPaintEvent):
        t = _theme
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        r = QRectF(RESIZE_MARGIN, RESIZE_MARGIN,
                   self.width()  - 2 * RESIZE_MARGIN,
                   self.height() - 2 * RESIZE_MARGIN)
        path = QPainterPath()
        path.addRoundedRect(r, BORDER_RADIUS, BORDER_RADIUS)
        grad = QLinearGradient(0, 0, self.width(), self.height())
        grad.setColorAt(0,   t.BG_GRAD_TOP)
        grad.setColorAt(0.4, t.BG_MID)
        grad.setColorAt(0.7, QColor(t.BG_MID.red()+2, t.BG_MID.green()+2, t.BG_MID.blue()+4))
        grad.setColorAt(1,   t.BG_GRAD_BOT)
        p.fillPath(path, grad)
        p.end()

    def resizeEvent(self, e: QResizeEvent):
        super().resizeEvent(e)
        if hasattr(self, "_glow"):
            self._glow.setGeometry(0, 0, self.width(), self.height())
        if hasattr(self, "_inner"):
            self._inner.setGeometry(
                RESIZE_MARGIN, RESIZE_MARGIN,
                self.width()  - 2 * RESIZE_MARGIN,
                self.height() - 2 * RESIZE_MARGIN)

    def closeEvent(self, e):
        """Unregister theme listener and stop the title animation on close."""
        global _node_scheduler, _reminder_scheduler
        if _node_scheduler is not None:
            try:
                _node_scheduler.stop()
            except Exception:
                pass
        if _reminder_scheduler is not None:
            try:
                _reminder_scheduler.stop()
            except Exception:
                pass
        _theme.unregister(self._on_theme_changed)
        if hasattr(self, "_titlebar") and hasattr(self._titlebar, "_title_label"):
            self._titlebar._title_label._timer.stop()
        super().closeEvent(e)


# ---------------------------------------------------------------------------
# Time Tracker  —  converted from TimeClock.ps1
# ---------------------------------------------------------------------------

_TIMECLOCK_LOG      = Path.home() / "Documents" / "TimeClockLog.csv"
_TIMECLOCK_MAX_DAYS = 30


def _tc_fmt(total_sec: float) -> str:
    """Seconds → HH:MM:SS string."""
    s = max(0, int(total_sec))
    return f"{s // 3600:02d}:{(s % 3600) // 60:02d}:{s % 60:02d}"


# ── CSV / log helpers ────────────────────────────────────────────────────────

def _tc_parse_csv_line(line: str) -> list:
    """Split a CSV line respecting double-quoted fields."""
    fields = []; in_q = False; cur = ""
    for ch in line:
        if ch == '"':
            in_q = not in_q
        elif ch == ',' and not in_q:
            fields.append(cur); cur = ""
        else:
            cur += ch
    fields.append(cur)
    return fields


def _tc_ensure_log():
    _TIMECLOCK_LOG.parent.mkdir(parents=True, exist_ok=True)
    if not _TIMECLOCK_LOG.exists():
        _TIMECLOCK_LOG.write_text(
            "Timestamp,Event,Duration_Seconds,Notes,UserNote\n", encoding="utf-8")


def _tc_append(event: str, duration: float, notes: str = "", user_note: str = ""):
    try:
        ts   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        dur  = round(duration, 1)
        sn   = notes.replace('"', '""')
        su   = user_note.replace('"', '""')
        with open(_TIMECLOCK_LOG, "a", encoding="utf-8", newline="") as f:
            f.write(f'{ts},{event},{dur},"{sn}","{su}"\n')
    except Exception:
        pass


def _tc_prune():
    try:
        if not _TIMECLOCK_LOG.exists():
            return
        cutoff = datetime.now() - timedelta(days=_TIMECLOCK_MAX_DAYS)
        lines  = _TIMECLOCK_LOG.read_text(encoding="utf-8").splitlines()
        if len(lines) <= 1:
            return
        kept = [lines[0]]
        for ln in lines[1:]:
            try:
                ts_str = ln.split(",")[0]
                if datetime.strptime(ts_str, "%Y-%m-%d %H:%M:%S") >= cutoff:
                    kept.append(ln)
            except Exception:
                kept.append(ln)
        _TIMECLOCK_LOG.write_text("\n".join(kept) + "\n", encoding="utf-8")
    except Exception:
        pass


def _tc_save_user_note(timestamp: str, new_note: str):
    try:
        lines = _TIMECLOCK_LOG.read_text(encoding="utf-8").splitlines()
        out   = [lines[0]]
        for ln in lines[1:]:
            f = _tc_parse_csv_line(ln)
            while len(f) < 5:
                f.append("")
            if f[0] == timestamp:
                sn = f[3].replace('"', '""')
                su = new_note.replace('"', '""')
                out.append(f'{f[0]},{f[1]},{f[2]},"{sn}","{su}"')
            else:
                out.append(ln)
        _TIMECLOCK_LOG.write_text("\n".join(out) + "\n", encoding="utf-8")
    except Exception:
        pass


def _tc_read_rows() -> list:
    """Return all log rows newest-first as list of dicts."""
    try:
        if not _TIMECLOCK_LOG.exists():
            return []
        lines = _TIMECLOCK_LOG.read_text(encoding="utf-8").splitlines()
        rows  = []
        for ln in reversed(lines[1:]):
            if not ln.strip():
                continue
            f = _tc_parse_csv_line(ln)
            while len(f) < 5:
                f.append("")
            rows.append({"ts": f[0], "event": f[1].strip(),
                          "dur": f[2], "notes": f[3],
                          "user_note": f[4], "raw": ln})
        return rows
    except Exception:
        return []


# ── WTS session-change listener (lock / unlock detection) ────────────────────

class _WtsListener(QWidget):
    """
    Hidden 1×1 window that hooks Windows WM_WTSSESSION_CHANGE so we know
    when the screen is locked or unlocked without polling.
    """
    locked   = Signal()
    unlocked = Signal()

    _WM_WTSSESSION_CHANGE = 0x02B1
    _WTS_SESSION_LOCK     = 7
    _WTS_SESSION_UNLOCK   = 8

    class _MSG(ctypes.Structure):
        _fields_ = [
            ("hwnd",    ctypes.c_void_p),
            ("message", ctypes.c_uint),
            ("wParam",  ctypes.c_size_t),
            ("lParam",  ctypes.c_ssize_t),
            ("time",    ctypes.c_uint),
            ("ptx",     ctypes.c_long),
            ("pty",     ctypes.c_long),
        ]

    def __init__(self):
        super().__init__(None,
                         Qt.FramelessWindowHint | Qt.Tool | Qt.WindowDoesNotAcceptFocus)
        self.resize(1, 1)
        self.move(-9999, -9999)
        # Must create a real window handle before registering
        self.winId()           # force handle creation
        try:
            ctypes.windll.wtsapi32.WTSRegisterSessionNotification(
                int(self.winId()), 0)   # 0 = NOTIFY_FOR_THIS_SESSION
        except Exception:
            pass

    def nativeEvent(self, event_type, message):
        try:
            msg = self._MSG.from_address(int(message))
            if msg.message == self._WM_WTSSESSION_CHANGE:
                if msg.wParam == self._WTS_SESSION_LOCK:
                    self.locked.emit()
                elif msg.wParam == self._WTS_SESSION_UNLOCK:
                    self.unlocked.emit()
        except Exception:
            pass
        return super().nativeEvent(event_type, message)

    def cleanup(self):
        try:
            ctypes.windll.wtsapi32.WTSUnRegisterSessionNotification(int(self.winId()))
        except Exception:
            pass


# ── Global hotkey listener (system-wide, works regardless of focus) ──────────

class _GlobalHotkeyListener(QWidget):
    """Hidden 1×1 window that registers and receives Win32 global hotkeys.

    Uses RegisterHotKey / WM_HOTKEY so the shortcut fires even when
    Command Center does not have focus.
    """
    hotkey_triggered = Signal(int)   # emits the hotkey id

    _WM_HOTKEY   = 0x0312
    _MOD_CONTROL = 0x0002
    _VK_OEM_3    = 0xC0   # backtick / grave accent (US keyboard layout)

    HOTKEY_CLIPBOARD  = 1
    HOTKEY_PASTE_PLAIN = 2

    _MOD_SHIFT = 0x0004
    _VK_V      = 0x56

    class _MSG(ctypes.Structure):
        _fields_ = [("hwnd",    ctypes.c_void_p),
                    ("message", ctypes.c_uint),
                    ("wParam",  ctypes.c_size_t),
                    ("lParam",  ctypes.c_size_t),
                    ("time",    ctypes.c_uint),
                    ("pt",      ctypes.c_long * 2)]

    def __init__(self, parent=None):
        super().__init__(parent, Qt.Window | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.resize(1, 1)
        self.move(-9999, -9999)
        self.winId()   # force native handle creation

    def register_clipboard_hotkey(self) -> bool:
        """Register Ctrl+` as a global hotkey. Returns True on success."""
        try:
            ok = ctypes.windll.user32.RegisterHotKey(
                int(self.winId()),
                self.HOTKEY_CLIPBOARD,
                self._MOD_CONTROL,
                self._VK_OEM_3)
            return bool(ok)
        except Exception:
            return False

    def unregister_clipboard_hotkey(self):
        try:
            ctypes.windll.user32.UnregisterHotKey(
                int(self.winId()), self.HOTKEY_CLIPBOARD)
        except Exception:
            pass

    def register_paste_plain_hotkey(self) -> bool:
        """Register Ctrl+Shift+V as a global paste-as-plain-text hotkey."""
        try:
            ok = ctypes.windll.user32.RegisterHotKey(
                int(self.winId()),
                self.HOTKEY_PASTE_PLAIN,
                self._MOD_CONTROL | self._MOD_SHIFT,
                self._VK_V)
            return bool(ok)
        except Exception:
            return False

    def unregister_paste_plain_hotkey(self):
        try:
            ctypes.windll.user32.UnregisterHotKey(
                int(self.winId()), self.HOTKEY_PASTE_PLAIN)
        except Exception:
            pass

    def nativeEvent(self, event_type, message):
        try:
            msg = self._MSG.from_address(int(message))
            if msg.message == self._WM_HOTKEY:
                self.hotkey_triggered.emit(int(msg.wParam))
        except Exception:
            pass
        return super().nativeEvent(event_type, message)

    def cleanup(self):
        self.unregister_clipboard_hotkey()
        self.unregister_paste_plain_hotkey()

    @staticmethod
    def _inject_ctrl_v():
        """Inject Ctrl+V via Win32 keybd_event to paste into the focused window.

        Called ~80 ms after setting the clipboard to plain text so that the
        hotkey key-events (Ctrl+Shift+V) have already been released by the OS.
        We also send a synthetic Shift key-up first to clear any residual Shift
        state that Windows may still hold from the triggering hotkey combo.
        """
        try:
            ke = ctypes.windll.user32.keybd_event
            ke(0x10, 0, 0x0002, 0)   # VK_SHIFT  key-up  (release leftover Shift)
            ke(0x11, 0, 0x0000, 0)   # VK_CONTROL key-down
            ke(0x56, 0, 0x0000, 0)   # VK_V       key-down  → Ctrl+V
            ke(0x56, 0, 0x0002, 0)   # VK_V       key-up
            ke(0x11, 0, 0x0002, 0)   # VK_CONTROL key-up
        except Exception:
            pass


# ── Adjust popup  (± scroll-wheel to set minutes then apply) ─────────────────

class TimeClockAdjustDialog(QDialog):
    """Popup for manually adjusting active/locked timer totals."""

    apply_active = Signal(int)   # minutes
    apply_locked = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent,
                         Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self._minutes = 0
        self._drag_pos: Optional[QPoint] = None
        self._build_ui()

    def _build_ui(self):
        self.setFixedSize(240, 168)
        root = QVBoxLayout(self); root.setContentsMargins(12, 12, 12, 12); root.setSpacing(8)

        # title row
        tr = QHBoxLayout(); tr.setSpacing(0)
        title = QLabel("ADJUST TIMER")
        title.setFont(QFont("Consolas", 8, QFont.Weight.Bold))
        title.setStyleSheet("color: rgb(90,90,90);")
        tr.addWidget(title); tr.addStretch()
        cb = QPushButton("×"); cb.setFixedSize(18, 18)
        cb.setFont(QFont("Consolas", 9, QFont.Weight.Bold))
        cb.setStyleSheet("QPushButton{background:transparent;color:rgb(90,90,90);border:none;}"
                         "QPushButton:hover{background:rgb(60,30,30);}")
        cb.setCursor(Qt.PointingHandCursor); cb.clicked.connect(self.reject)
        tr.addWidget(cb); root.addLayout(tr)

        # value display
        self._val_lbl = QLabel("+0 min")
        self._val_lbl.setFont(QFont("Consolas", 20, QFont.Weight.Bold))
        self._val_lbl.setAlignment(Qt.AlignCenter)
        self._val_lbl.setStyleSheet("color: rgb(72,213,107);")
        root.addWidget(self._val_lbl)

        hint = QLabel("scroll wheel to change amount")
        hint.setFont(QFont("Consolas", 8))
        hint.setAlignment(Qt.AlignCenter)
        hint.setStyleSheet("color: rgb(65,65,65);")
        root.addWidget(hint)

        # apply buttons
        self._btn_active = self._mkbtn("● Apply to Active", "rgb(72,213,107)")
        self._btn_active.clicked.connect(lambda: (self.apply_active.emit(self._minutes), self.accept()))
        self._btn_locked = self._mkbtn("● Apply to Locked", "rgb(220,65,65)")
        self._btn_locked.clicked.connect(lambda: (self.apply_locked.emit(self._minutes), self.accept()))
        root.addWidget(self._btn_active)
        root.addWidget(self._btn_locked)

    def _mkbtn(self, text: str, fg: str) -> QPushButton:
        b = QPushButton(text)
        b.setFont(QFont("Consolas", 8))
        b.setFixedHeight(26)
        b.setCursor(Qt.PointingHandCursor)
        b.setStyleSheet(f"QPushButton{{background:rgb(25,25,25);color:{fg};"
                         "border:1px solid rgb(45,45,45);border-radius:4px;}}"
                         f"QPushButton:hover{{background:rgb(35,35,35);}}")
        return b

    def wheelEvent(self, e):
        delta = e.angleDelta().y()
        self._minutes += int(round(delta / 120))
        col = "rgb(72,213,107)" if self._minutes >= 0 else "rgb(220,65,65)"
        self._val_lbl.setStyleSheet(f"color: {col};")
        sign = "+" if self._minutes >= 0 else ""
        self._val_lbl.setText(f"{sign}{self._minutes} min")

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_pos = e.globalPosition().toPoint() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if self._drag_pos and e.buttons() == Qt.LeftButton:
            self.move(e.globalPosition().toPoint() - self._drag_pos)

    def mouseReleaseEvent(self, e):
        self._drag_pos = None

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        r = self.rect().adjusted(0, 0, -1, -1)
        grad = QLinearGradient(0, 0, 0, r.height())
        tb = _theme.TITLEBAR_BG; bm = _theme.BG_MID
        grad.setColorAt(0.0, QColor(tb.red(), tb.green(), tb.blue(), 248))
        grad.setColorAt(1.0, QColor(bm.red(), bm.green(), bm.blue(), 248))
        g = _theme.GLOW
        p.setBrush(QBrush(grad))
        p.setPen(QPen(QColor(g.red(), g.green(), g.blue(), 55), 1))
        p.drawRoundedRect(r, 8, 8)


# ── Log viewer ────────────────────────────────────────────────────────────────

_TC_ROW_COLORS = {
    "LOGIN":          ((20,  55, 28),  (130, 240, 160)),
    "UNLOCK":         ((18,  48, 24),  (100, 210, 130)),
    "LOCK":           ((60,  20, 20),  (240, 110, 110)),
    "SESSION_END":    ((45,  45, 65),  (160, 160, 210)),
    "SESSION_END_LOCKED": ((45, 45, 65), (160, 160, 210)),
    "ADJUST_ACTIVE":  ((20,  48, 48),  (100, 210, 210)),
    "ADJUST_LOCKED":  ((48,  28, 48),  (200, 140, 210)),
    "LAP_RESET":      ((50,  30,  5),  (255, 180,  80)),
    "RESET":          ((55,  40, 15),  (220, 175,  90)),
    "RESET_WHILE_LOCKED": ((55, 40, 15), (220, 175, 90)),
}
_TC_ROW_DEFAULT = ((22, 22, 22), (190, 190, 190))


def _tc_row_colors(event: str) -> tuple:
    for k, v in _TC_ROW_COLORS.items():
        if event.upper().startswith(k.upper()):
            return v
    return _TC_ROW_DEFAULT


class TimeClockLogViewer(QDialog):
    """Full log viewer — mirrors the PS1 DataGridView log form."""

    def __init__(self, parent=None):
        super().__init__(parent,
                         Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setMinimumSize(760, 380)
        self.resize(1100, 620)
        self._drag_pos: Optional[QPoint] = None
        self._all_rows: list = []
        self._build_ui()
        self._load_data()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10); root.setSpacing(6)

        # title bar
        tr = QHBoxLayout(); tr.setSpacing(4)
        _td = _theme.TEXT_DIM
        ttl = QLabel("≡  TIME CLOCK LOG")
        ttl.setFont(QFont("Consolas", 9, QFont.Weight.Bold))
        ttl.setStyleSheet(
            f"color: rgba({_td.red()},{_td.green()},{_td.blue()},200);"
            " letter-spacing:1px;")
        tr.addWidget(ttl); tr.addStretch()
        cb = QPushButton("×"); cb.setFixedSize(20, 18)
        cb.setFont(QFont("Consolas", 10, QFont.Weight.Bold))
        cb.setStyleSheet("QPushButton{background:transparent;color:rgb(90,90,90);border:none;}"
                         "QPushButton:hover{background:rgb(60,30,30);color:white;}")
        cb.setCursor(Qt.PointingHandCursor); cb.clicked.connect(self.reject)
        tr.addWidget(cb); root.addLayout(tr)

        # table
        from PySide6.QtWidgets import QTableWidget, QTableWidgetItem, QAbstractItemView, QHeaderView
        self._table = QTableWidget()
        self._table.setColumnCount(5)
        self._table.setHorizontalHeaderLabels(
            ["Timestamp", "Event", "Duration (s)", "Notes", "User Note  (click to edit)"])
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Interactive)
        self._table.setColumnWidth(0, 148); self._table.setColumnWidth(1, 148)
        self._table.setColumnWidth(2, 90);  self._table.setColumnWidth(3, 320)
        self._table.verticalHeader().setVisible(False)
        self._table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SingleSelection)
        self._table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked)
        self._table.setStyleSheet("""
            QTableWidget {
                background: rgb(18,18,18); color: rgb(200,200,200);
                gridline-color: rgb(38,38,38); border: none;
                font-family: Consolas; font-size: 8pt;
            }
            QHeaderView::section {
                background: rgb(28,28,28); color: rgb(90,90,90);
                font-family: Consolas; font-size: 8pt; font-weight: bold;
                border: none; border-bottom: 1px solid rgb(38,38,38); padding: 3px;
            }
            QTableWidget::item:selected {
                background: rgb(40,60,40); color: rgb(220,220,220);
            }
            QScrollBar:vertical { background: rgb(18,18,18); width: 10px; }
            QScrollBar::handle:vertical { background: rgb(50,50,50); border-radius: 5px; }
            QScrollBar:horizontal { background: rgb(18,18,18); height: 10px; }
            QScrollBar::handle:horizontal { background: rgb(50,50,50); border-radius: 5px; }
        """)
        self._table.itemChanged.connect(self._on_cell_changed)
        self._table.setRowHeight(0, 22)
        root.addWidget(self._table)

        # toolbar
        tb = QHBoxLayout(); tb.setSpacing(6)
        self._btn_refresh   = self._mkbtn("Refresh")
        self._btn_copy_row  = self._mkbtn("Copy Row")
        self._btn_export    = self._mkbtn("Export CSV")
        self._btn_refresh.clicked.connect(self._load_data)
        self._btn_copy_row.clicked.connect(self._copy_row)
        self._btn_export.clicked.connect(self._export_csv)
        tb.addWidget(self._btn_refresh)
        tb.addWidget(self._btn_copy_row)
        tb.addWidget(self._btn_export)
        tb.addSpacing(12)

        lbl_f = QLabel("Filter:")
        lbl_f.setFont(QFont("Consolas", 8)); lbl_f.setStyleSheet("color:rgb(90,90,90);")
        tb.addWidget(lbl_f)
        self._filter_cb = QComboBox()
        self._filter_cb.addItems(["All Events","LOGIN","LOCK","UNLOCK",
                                   "SESSION_END","ADJUST_ACTIVE","ADJUST_LOCKED",
                                   "LAP_RESET","RESET","RESET_WHILE_LOCKED"])
        self._filter_cb.setFont(QFont("Consolas", 8))
        self._filter_cb.setStyleSheet("""QComboBox{background:rgb(30,30,30);color:rgb(190,190,190);
            border:1px solid rgb(50,50,50);border-radius:3px;padding:1px 6px;}
            QComboBox QAbstractItemView{background:rgb(25,25,25);color:rgb(190,190,190);
            selection-background-color:rgb(45,60,45);}""")
        self._filter_cb.currentIndexChanged.connect(self._apply_filter)
        tb.addWidget(self._filter_cb)
        tb.addSpacing(8)

        lbl_s = QLabel("Search:")
        lbl_s.setFont(QFont("Consolas", 8)); lbl_s.setStyleSheet("color:rgb(90,90,90);")
        tb.addWidget(lbl_s)
        self._search_le = QLineEdit()
        self._search_le.setFixedWidth(180)
        self._search_le.setFont(QFont("Consolas", 8))
        self._search_le.setStyleSheet("QLineEdit{background:rgb(30,30,30);color:rgb(190,190,190);"
                                       "border:1px solid rgb(50,50,50);border-radius:3px;padding:1px 6px;}")
        self._search_le.textChanged.connect(self._apply_filter)
        tb.addWidget(self._search_le)
        tb.addSpacing(12)

        self._count_lbl = QLabel("")
        self._count_lbl.setFont(QFont("Consolas", 8))
        self._count_lbl.setStyleSheet("color:rgb(70,70,70);")
        tb.addWidget(self._count_lbl)
        tb.addStretch()
        root.addLayout(tb)

    def _mkbtn(self, text: str) -> QPushButton:
        b = QPushButton(text); b.setFont(QFont("Consolas", 8)); b.setFixedHeight(24)
        b.setCursor(Qt.PointingHandCursor)
        b.setStyleSheet("QPushButton{background:rgb(28,28,28);color:rgb(160,160,160);"
                         "border:1px solid rgb(55,55,55);border-radius:4px;padding:0 10px;}"
                         "QPushButton:hover{background:rgb(40,40,40);}")
        return b

    def _load_data(self):
        self._all_rows = _tc_read_rows()
        self._apply_filter()

    def _apply_filter(self):
        filt   = self._filter_cb.currentText()
        search = self._search_le.text().lower()
        self._table.blockSignals(True)
        self._table.setRowCount(0)
        shown      = 0
        last_date  = None   # track day changes for separator rows
        for r in self._all_rows:
            if filt not in ("", "All Events") and r["event"] != filt:
                continue
            if search and search not in r["raw"].lower():
                continue

            # ── Day separator ─────────────────────────────────────────────
            row_date = r["ts"][:10] if len(r["ts"]) >= 10 else ""
            if row_date and row_date != last_date:
                sep_idx = self._table.rowCount()
                self._table.insertRow(sep_idx)
                self._table.setRowHeight(sep_idx, 18)
                # Format date nicely: 2026-04-23 → Wednesday  Apr 23 2026
                try:
                    dt_day = datetime.strptime(row_date, "%Y-%m-%d")
                    day_str = dt_day.strftime("── %A,  %B %d %Y ─────────────────────────────────────────────────────────────────────")
                except Exception:
                    day_str = f"── {row_date} ──────────────────────────────────────────────────────────────────────────────"
                sep_item = QTableWidgetItem(day_str)
                sep_item.setFlags(Qt.ItemIsEnabled)
                sep_item.setBackground(QBrush(QColor(32, 32, 42)))
                sep_item.setForeground(QBrush(QColor(80, 110, 140)))
                sep_item.setFont(QFont("Consolas", 7, QFont.Weight.Bold))
                self._table.setItem(sep_idx, 0, sep_item)
                self._table.setSpan(sep_idx, 0, 1, 5)
                last_date = row_date

            row_idx = self._table.rowCount()
            self._table.insertRow(row_idx)
            self._table.setRowHeight(row_idx, 22)
            bg, fg = _tc_row_colors(r["event"])
            for col, val in enumerate([r["ts"], r["event"], r["dur"], r["notes"], r["user_note"]]):
                item = QTableWidgetItem(val)
                if col < 4:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                item.setBackground(QBrush(QColor(*bg)))
                item.setForeground(QBrush(QColor(*fg)))
                if col == 4:
                    item.setBackground(QBrush(QColor(28, 28, 28)))
                    item.setForeground(QBrush(QColor(210, 210, 160)))
                self._table.setItem(row_idx, col, item)
            shown += 1
        self._table.blockSignals(False)
        self._count_lbl.setText(f"{shown} row{'s' if shown != 1 else ''} shown")

    def _on_cell_changed(self, item):
        if item.column() != 4:
            return
        row   = item.row()
        ts_it = self._table.item(row, 0)
        if ts_it:
            _tc_save_user_note(ts_it.text(), item.text())
            # update in-memory rows
            for r in self._all_rows:
                if r["ts"] == ts_it.text():
                    r["user_note"] = item.text()
                    break

    def _copy_row(self):
        row = self._table.currentRow()
        if row < 0:
            return
        parts = [self._table.item(row, c).text() if self._table.item(row, c) else ""
                 for c in range(self._table.columnCount())]
        QApplication.clipboard().setText("\t".join(parts))

    def _export_csv(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Export CSV", str(Path.home() / "Desktop" /
            f"TimeClockExport_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"),
            "CSV files (*.csv);;All files (*.*)")
        if path and _TIMECLOCK_LOG.exists():
            shutil.copy2(_TIMECLOCK_LOG, path)

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_pos = e.globalPosition().toPoint() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if self._drag_pos and e.buttons() == Qt.LeftButton:
            self.move(e.globalPosition().toPoint() - self._drag_pos)

    def mouseReleaseEvent(self, e):
        self._drag_pos = None

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        r = self.rect().adjusted(0, 0, -1, -1)
        grad = QLinearGradient(0, 0, 0, r.height())
        tb = _theme.TITLEBAR_BG; bm = _theme.BG_MID
        grad.setColorAt(0.0, QColor(tb.red(), tb.green(), tb.blue(), 250))
        grad.setColorAt(1.0, QColor(bm.red(), bm.green(), bm.blue(), 250))
        g = _theme.GLOW
        p.setBrush(QBrush(grad))
        p.setPen(QPen(QColor(g.red(), g.green(), g.blue(), 55), 1))
        p.drawRoundedRect(r, 8, 8)
        p.setPen(QPen(QColor(g.red(), g.green(), g.blue(), 30), 1))
        p.drawLine(8, 26, self.width() - 8, 26)
        p.drawLine(8, self.height() - 38, self.width() - 8, self.height() - 38)


# ── Main HUD window ───────────────────────────────────────────────────────────

class TimeClockHUD(QWidget):
    """
    Floating compact HUD — direct Python port of TimeClock.ps1.

    Layout (expanded, 270×130):
        ┌──────────────────────────────────────────┐
        │ ⏱ TIME CLOCK                [–]  [×]    │
        ├──────────────────────────────────────────┤
        │  ● 00:00:00 (green)          0.00 (white)│
        │  ● 00:00:00 (red)            0.00 (amber)│
        ├──────────────────────────────────────────┤
        │  [±]  [≡]                       ACTIVE   │
        └──────────────────────────────────────────┘
    """

    # ── timer state ──────────────────────────────────────────────────────────
    def __init__(self, parent=None):
        super().__init__(parent,
                         Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_DeleteOnClose, False)   # reusable

        # state
        self._is_locked        = False
        self._login_start      = datetime.now()
        self._lock_start       = datetime.now()
        self._total_login_sec  = 0.0
        self._total_locked_sec = 0.0
        self._lap_base_sec     = 0.0
        self._lap_seg_start    = datetime.now()
        self._collapsed        = False
        self._drag_pos: Optional[QPoint] = None
        self._log_viewer: Optional[TimeClockLogViewer] = None

        # WTS lock/unlock listener
        self._wts = _WtsListener()
        self._wts.locked.connect(self._on_lock)
        self._wts.unlocked.connect(self._on_unlock)

        # init log
        _tc_ensure_log()
        _tc_prune()
        _tc_append("LOGIN", 0.0, "TimeClock started")

        self._build_ui()
        self._position_hud()
        _theme.register(self.refresh_theme)

        # 100 ms tick
        self._tick_timer = QTimer(self)
        self._tick_timer.timeout.connect(self._tick)
        self._tick_timer.start(100)

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        self.setFixedSize(320, 150)

        # ── background image (loaded once, drawn subtly in paintEvent) ────────
        _tc_img = os.path.join(os.path.dirname(os.path.abspath(__file__)), "TimeClock.png")
        self._bg_pixmap = QPixmap(_tc_img) if os.path.isfile(_tc_img) else QPixmap()

        # ── title row ────────────────────────────────────────────────────────
        self._lbl_title = QLabel("⏱ TIME CLOCK", self)
        self._lbl_title.setFont(QFont("Consolas", 8, QFont.Weight.Bold))
        _td = _theme.TEXT_DIM
        self._lbl_title.setStyleSheet(
            f"color: rgba({_td.red()},{_td.green()},{_td.blue()},210);"
            " background:transparent; letter-spacing:1px;")
        self._lbl_title.setGeometry(8, 5, 240, 14)

        self._btn_min = QPushButton("–", self)
        self._btn_min.setFont(QFont("Consolas", 9, QFont.Weight.Bold))
        _g = _theme.GLOW
        self._btn_min.setStyleSheet(
            f"QPushButton{{background:transparent;color:rgb(90,90,90);border:none;}}"
            f"QPushButton:hover{{background:rgba({_g.red()},{_g.green()},{_g.blue()},45);color:white;}}")
        self._btn_min.setGeometry(276, 2, 20, 16)
        self._btn_min.setCursor(Qt.PointingHandCursor)
        self._btn_min.clicked.connect(self._collapse)

        self._btn_close = QPushButton("×", self)
        self._btn_close.setFont(QFont("Consolas", 10, QFont.Weight.Bold))
        self._btn_close.setStyleSheet(
            "QPushButton{background:transparent;color:rgb(90,90,90);border:none;}"
            "QPushButton:hover{background:rgba(180,45,45,160);color:white;}")
        self._btn_close.setGeometry(296, 2, 20, 16)
        self._btn_close.setCursor(Qt.PointingHandCursor)
        self._btn_close.clicked.connect(self._on_close)

        # ── green active timer ────────────────────────────────────────────────
        self._lbl_active = QLabel("● 00:00:00", self)
        self._lbl_active.setFont(QFont("Consolas", 16, QFont.Weight.Bold))
        self._lbl_active.setStyleSheet("color: rgb(72,213,107); background:transparent;")
        self._lbl_active.setGeometry(8, 24, 200, 33)

        # ── decimal hours (white, click to copy to clipboard) ─────────────────
        self._lbl_decimal = QLabel("0.00", self)
        self._lbl_decimal.setFont(QFont("Consolas", 14, QFont.Weight.Bold))
        self._lbl_decimal.setStyleSheet("color: white; background:transparent;")
        self._lbl_decimal.setAlignment(Qt.AlignCenter)
        self._lbl_decimal.setGeometry(213, 24, 100, 33)
        self._lbl_decimal.setCursor(Qt.PointingHandCursor)
        self._lbl_decimal.setToolTip("Total active time (decimal hours) | Click to copy")
        self._lbl_decimal.mousePressEvent = lambda e: self._copy_decimal_to_clipboard()

        # ── red locked timer ──────────────────────────────────────────────────
        self._lbl_locked = QLabel("● 00:00:00", self)
        self._lbl_locked.setFont(QFont("Consolas", 16, QFont.Weight.Bold))
        self._lbl_locked.setStyleSheet("color: rgb(220,65,65); background:transparent;")
        self._lbl_locked.setGeometry(8, 60, 200, 33)

        # ── lap decimal hours (orange, click to reset) ────────────────────────
        self._lbl_lap = QLabel("0.00", self)
        self._lbl_lap.setFont(QFont("Consolas", 14, QFont.Weight.Bold))
        self._lbl_lap.setStyleSheet("color: rgb(255,165,30); background:transparent;")
        self._lbl_lap.setAlignment(Qt.AlignCenter)
        self._lbl_lap.setGeometry(213, 60, 100, 33)
        self._lbl_lap.setCursor(Qt.PointingHandCursor)
        self._lbl_lap.setToolTip("Lap active time (decimal hours) | Click to reset")
        self._lbl_lap.mousePressEvent = lambda e: self._reset_lap()
        self._lbl_lap.enterEvent = lambda e: self._lbl_lap.setStyleSheet(
            "color: rgb(255,200,60); background: rgba(255,165,30,40); border-radius: 5px;")
        self._lbl_lap.leaveEvent = lambda e: self._lbl_lap.setStyleSheet(
            "color: rgb(255,165,30); background:transparent;")

        # ── toolbar row ───────────────────────────────────────────────────────
        self._btn_adjust = self._make_toolbar_btn("±", 8, 90)
        self._btn_adjust.setToolTip("Adjust timer (scroll wheel to set ± minutes)")
        self._btn_adjust.clicked.connect(self._show_adjust)

        self._btn_log = self._make_toolbar_btn("≡", 104, 90)
        self._btn_log.setToolTip("View time log")
        self._btn_log.clicked.connect(self._show_log)

        self._lbl_status = QLabel("ACTIVE", self)
        self._lbl_status.setFont(QFont("Consolas", 7))
        self._lbl_status.setStyleSheet("color: rgb(60,160,80); background:transparent;")
        self._lbl_status.setGeometry(200, 122, 112, 14)
        self._lbl_status.setAlignment(Qt.AlignRight)

        # ── context menu ──────────────────────────────────────────────────────
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_ctx_menu)

        # ── collapsed pill (visible only when collapsed) ──────────────────────
        self._btn_restore = QPushButton("▶ 0.00", self)
        self._btn_restore.setFont(QFont("Consolas", 9, QFont.Weight.Bold))
        self._btn_restore.setStyleSheet(
            "QPushButton{background:rgb(15,15,15);color:rgb(72,213,107);"
            "border:1px solid rgb(50,100,50);border-radius:4px;}"
            "QPushButton:hover{background:rgb(20,40,20);}")
        self._btn_restore.setGeometry(0, 0, 86, 24)
        self._btn_restore.setCursor(Qt.PointingHandCursor)
        self._btn_restore.hide()
        self._btn_restore.clicked.connect(self._expand)

    def refresh_theme(self):
        """Re-apply all theme-dependent stylesheets when the theme changes."""
        _td = _theme.TEXT_DIM
        _g  = _theme.GLOW
        _t  = _theme
        self._lbl_title.setStyleSheet(
            f"color: rgba({_td.red()},{_td.green()},{_td.blue()},210);"
            " background:transparent; letter-spacing:1px;")
        self._btn_min.setStyleSheet(
            f"QPushButton{{background:transparent;color:rgb(90,90,90);border:none;}}"
            f"QPushButton:hover{{background:rgba({_g.red()},{_g.green()},{_g.blue()},45);color:white;}}")
        self._btn_adjust.setStyleSheet(
            f"QPushButton{{background:rgba({_t.BG_MID.red()},{_t.BG_MID.green()},{_t.BG_MID.blue()},160);"
            f"color:{_t.TEXT_SECONDARY.name()};"
            f"border:1px solid rgba({_g.red()},{_g.green()},{_g.blue()},45);border-radius:3px;}}"
            f"QPushButton:hover{{background:rgba({_t.TILE_BG_HOVER.red()},{_t.TILE_BG_HOVER.green()},{_t.TILE_BG_HOVER.blue()},200);"
            f"color:{_t.TEXT_PRIMARY.name()};"
            f"border-color:rgba({_g.red()},{_g.green()},{_g.blue()},100);}}")
        self._btn_log.setStyleSheet(
            f"QPushButton{{background:rgba({_t.BG_MID.red()},{_t.BG_MID.green()},{_t.BG_MID.blue()},160);"
            f"color:{_t.TEXT_SECONDARY.name()};"
            f"border:1px solid rgba({_g.red()},{_g.green()},{_g.blue()},45);border-radius:3px;}}"
            f"QPushButton:hover{{background:rgba({_t.TILE_BG_HOVER.red()},{_t.TILE_BG_HOVER.green()},{_t.TILE_BG_HOVER.blue()},200);"
            f"color:{_t.TEXT_PRIMARY.name()};"
            f"border-color:rgba({_g.red()},{_g.green()},{_g.blue()},100);}}")
        self.update()

    def _make_toolbar_btn(self, text: str, x: int, w: int) -> QPushButton:
        b = QPushButton(text, self)
        b.setFont(QFont("Consolas", 9, QFont.Weight.Bold))
        _t = _theme; _g = _t.GLOW
        b.setStyleSheet(
            f"QPushButton{{background:rgba({_t.BG_MID.red()},{_t.BG_MID.green()},{_t.BG_MID.blue()},160);"
            f"color:{_t.TEXT_SECONDARY.name()};"
            f"border:1px solid rgba({_g.red()},{_g.green()},{_g.blue()},45);border-radius:3px;}}"
            f"QPushButton:hover{{background:rgba({_t.TILE_BG_HOVER.red()},{_t.TILE_BG_HOVER.green()},{_t.TILE_BG_HOVER.blue()},200);"
            f"color:{_t.TEXT_PRIMARY.name()};"
            f"border-color:rgba({_g.red()},{_g.green()},{_g.blue()},100);}}")
        b.setGeometry(x, 121, w, 22)
        b.setCursor(Qt.PointingHandCursor)
        return b

    # ── positioning ───────────────────────────────────────────────────────────

    def _position_hud(self):
        """Anchor to bottom-right of primary screen work area."""
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.right() - self.width() - 10,
                  screen.bottom() - self.height() - 10)

    def _anchor_collapsed(self):
        """Pin the collapsed pill firmly to the bottom-right of the primary screen."""
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.right() - self.width() - 10,
                  screen.bottom() - self.height() - 10)

    def _reanchor_if_drifted(self):
        """Deferred check: snap the pill back if Windows/DWM moved it."""
        if self._collapsed and self._drag_pos is None:
            screen = QApplication.primaryScreen().availableGeometry()
            ex = screen.right()  - self.width()  - 10
            ey = screen.bottom() - self.height() - 10
            if self.pos().x() != ex or self.pos().y() != ey:
                self.move(ex, ey)

    # ── collapse / expand ────────────────────────────────────────────────────

    def _collapse(self):
        self._collapsed = True
        for w in (self._lbl_title, self._btn_min, self._btn_close,
                  self._lbl_active, self._lbl_decimal,
                  self._lbl_locked, self._lbl_lap,
                  self._btn_adjust, self._btn_log, self._lbl_status):
            w.hide()
        self._btn_restore.show()
        self.setFixedSize(86, 24)
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.right() - self.width() - 10,
                  screen.bottom() - self.height() - 10)

    def _expand(self):
        self._collapsed = False
        self._btn_restore.hide()
        for w in (self._lbl_title, self._btn_min, self._btn_close,
                  self._lbl_active, self._lbl_decimal,
                  self._lbl_locked, self._lbl_lap,
                  self._btn_adjust, self._btn_log, self._lbl_status):
            w.show()
        self.setFixedSize(320, 150)
        self._position_hud()

    # ── tick (100 ms) ─────────────────────────────────────────────────────────

    def _tick(self):
        now = datetime.now()
        # When collapsed AND locked nothing visible changes (active-time is frozen,
        # pill text is frozen).  Skipping all widget updates also prevents the
        # 100 ms DWM repaints that can drift the pinned translucent pill during lock.
        if self._collapsed and self._is_locked:
            return
        if self._is_locked:
            live_locked = self._total_locked_sec + (now - self._lock_start).total_seconds()
            live_login  = self._total_login_sec
            self._lbl_status.setText("LOCKED")
            self._lbl_status.setStyleSheet("color: rgb(180,55,55); background:transparent;")
            # pulse red
            pulse = max(30, 65 - (now.second % 2) * 15)
            self._lbl_locked.setStyleSheet(f"color: rgb(220,{pulse},65); background:transparent;")
        else:
            live_login  = self._total_login_sec + (now - self._login_start).total_seconds()
            live_locked = self._total_locked_sec
            self._lbl_status.setText("ACTIVE")
            self._lbl_status.setStyleSheet("color: rgb(60,160,80); background:transparent;")
            self._lbl_locked.setStyleSheet("color: rgb(220,65,65); background:transparent;")

        live_lap = (self._lap_base_sec if self._is_locked
                    else self._lap_base_sec + (now - self._lap_seg_start).total_seconds())

        if self._collapsed:
            self._btn_restore.setText(f"▶ {live_login / 3600:.2f}")
            return

        self._lbl_active.setText(f"● {_tc_fmt(live_login)}")
        self._lbl_locked.setText(f"● {_tc_fmt(live_locked)}")
        self._lbl_decimal.setText(f"{live_login / 3600:.2f}")
        self._lbl_lap.setText(f"{live_lap / 3600:.2f}")

    # ── lock / unlock ─────────────────────────────────────────────────────────

    def _on_lock(self):
        if self._is_locked:
            return
        # Cancel any in-progress drag so the window can't drift during lock
        self._drag_pos = None
        now = datetime.now()
        seg = (now - self._login_start).total_seconds()
        self._total_login_sec += seg
        self._lock_start = now
        self._is_locked = True
        _tc_append("LOCK", seg,
                   f"Active for {_tc_fmt(seg)} before lock | "
                   f"total active {_tc_fmt(self._total_login_sec)}")
        self._lap_base_sec += (now - self._lap_seg_start).total_seconds()
        # Pin the collapsed pill tightly before the lock screen activates
        if self._collapsed:
            self._anchor_collapsed()

    def _on_unlock(self):
        if not self._is_locked:
            return
        now = datetime.now()
        seg = (now - self._lock_start).total_seconds()
        self._total_locked_sec += seg
        self._login_start = now
        self._is_locked = False
        _tc_append("UNLOCK", seg,
                   f"Locked for {_tc_fmt(seg)} | "
                   f"total locked {_tc_fmt(self._total_locked_sec)}")
        self._lap_seg_start = now
        # Restore the collapsed pill to its pinned corner after the lock screen
        # clears — Windows/DWM can drift Qt.Tool windows during the session lock.
        if self._collapsed:
            self._anchor_collapsed()

    # ── actions ───────────────────────────────────────────────────────────────

    def _copy_decimal_to_clipboard(self):
        val = self._lbl_decimal.text()
        QApplication.clipboard().setText(val)

    def _reset_lap(self):
        now      = datetime.now()
        prev_sec = (self._lap_base_sec if self._is_locked
                    else self._lap_base_sec + (now - self._lap_seg_start).total_seconds())
        _tc_append("LAP_RESET", prev_sec,
                   f"Lap timer reset | previous value: {prev_sec / 3600:.2f} hrs")
        self._lap_base_sec = 0.0
        if not self._is_locked:
            self._lap_seg_start = now

    def _reset_timers(self):
        now = datetime.now()
        if self._is_locked:
            seg = (now - self._lock_start).total_seconds()
            _tc_append("RESET_WHILE_LOCKED", seg, "Timer reset by user")
            self._lock_start = now
        else:
            seg = (now - self._login_start).total_seconds()
            _tc_append("RESET", seg, "Timer reset by user")
            self._login_start = now
        self._total_login_sec  = 0.0
        self._total_locked_sec = 0.0

    def _show_adjust(self):
        dlg = TimeClockAdjustDialog(self)
        dlg.apply_active.connect(self._apply_adjust_active)
        dlg.apply_locked.connect(self._apply_adjust_locked)
        # position above button
        gpos  = self._btn_adjust.mapToGlobal(QPoint(0, 0))
        dlg.move(gpos.x(), gpos.y() - dlg.height() - 6)
        dlg.exec()

    def _apply_adjust_active(self, minutes: int):
        delta = minutes * 60.0
        self._total_login_sec = max(0.0, self._total_login_sec + delta)
        _tc_append("ADJUST_ACTIVE", delta,
                   f"Manual adjustment: {minutes} min to active timer")

    def _apply_adjust_locked(self, minutes: int):
        delta = minutes * 60.0
        self._total_locked_sec = max(0.0, self._total_locked_sec + delta)
        _tc_append("ADJUST_LOCKED", delta,
                   f"Manual adjustment: {minutes} min to locked timer")

    def _show_log(self):
        if self._log_viewer is None or not self._log_viewer.isVisible():
            self._log_viewer = TimeClockLogViewer(None)
            self._log_viewer.show()
        else:
            self._log_viewer.raise_()
            self._log_viewer._load_data()

    def _show_ctx_menu(self, pos: QPoint):
        menu = QMenu(self)
        menu.setStyleSheet(
            "QMenu{background:rgb(22,22,22);color:rgb(200,200,200);"
            "border:1px solid rgb(50,50,50);border-radius:6px;font-family:Consolas;font-size:8pt;}"
            "QMenu::item{padding:5px 18px;} QMenu::item:selected{background:rgb(40,60,40);}"
            "QMenu::separator{height:1px;background:rgb(40,40,40);margin:3px 8px;}")
        a_reset    = menu.addAction("Reset Timers")
        a_lap      = menu.addAction("Reset Lap Timer")
        menu.addSeparator()
        a_view_log = menu.addAction("Open Log File")
        a_open_dir = menu.addAction("Open Log Folder")
        menu.addSeparator()
        a_exit     = menu.addAction("Close Time Tracker")
        chosen = menu.exec(self.mapToGlobal(pos))
        if chosen == a_reset:
            self._reset_timers()
        elif chosen == a_lap:
            self._reset_lap()
        elif chosen == a_view_log:
            try:
                os.startfile(str(_TIMECLOCK_LOG))
            except Exception:
                pass
        elif chosen == a_open_dir:
            try:
                os.startfile(str(_TIMECLOCK_LOG.parent))
            except Exception:
                pass
        elif chosen == a_exit:
            self._on_close()

    def _on_close(self):
        now = datetime.now()
        if self._is_locked:
            seg = (now - self._lock_start).total_seconds()
            self._total_locked_sec += seg
            _tc_append("SESSION_END_LOCKED", seg,
                       f"App closed while locked | "
                       f"total active {_tc_fmt(self._total_login_sec)} | "
                       f"total locked {_tc_fmt(self._total_locked_sec)}")
        else:
            seg = (now - self._login_start).total_seconds()
            self._total_login_sec += seg
            _tc_append("SESSION_END", seg,
                       f"App closed | "
                       f"total active {_tc_fmt(self._total_login_sec)} | "
                       f"total locked {_tc_fmt(self._total_locked_sec)}")
        _tc_prune()
        self._tick_timer.stop()
        self._wts.cleanup()
        self._wts.close()
        _theme.unregister(self.refresh_theme)
        self.hide()

    # ── drag ─────────────────────────────────────────────────────────────────

    def moveEvent(self, e):
        """Snap the collapsed pill back to its anchor if Windows/DWM drifts it."""
        super().moveEvent(e)
        if self._collapsed and self._drag_pos is None:
            QTimer.singleShot(0, self._reanchor_if_drifted)

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_pos = e.globalPosition().toPoint() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if self._drag_pos and e.buttons() == Qt.LeftButton:
            new_pos = e.globalPosition().toPoint() - self._drag_pos
            # Clamp to available screen area so the HUD can't drift off-screen
            screen = QApplication.primaryScreen().availableGeometry()
            new_pos.setX(max(screen.left(), min(new_pos.x(), screen.right()  - self.width())))
            new_pos.setY(max(screen.top(),  min(new_pos.y(), screen.bottom() - self.height())))
            self.move(new_pos)

    def mouseReleaseEvent(self, e):
        self._drag_pos = None

    def leaveEvent(self, e):
        """Clear drag state when the mouse leaves (e.g. during screen lock)."""
        self._drag_pos = None
        super().leaveEvent(e)

    # ── painting ─────────────────────────────────────────────────────────────

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        r = self.rect().adjusted(0, 0, -1, -1)

        # ── gradient background ──────────────────────────────────────────────
        grad = QLinearGradient(0, 0, 0, r.height())
        tb = _theme.TITLEBAR_BG; bm = _theme.BG_MID
        grad.setColorAt(0.0, QColor(tb.red(), tb.green(), tb.blue(), 245))
        grad.setColorAt(1.0, QColor(bm.red(), bm.green(), bm.blue(), 245))
        g = _theme.GLOW
        p.setBrush(QBrush(grad))
        p.setPen(QPen(QColor(g.red(), g.green(), g.blue(), 55), 1))
        p.drawRoundedRect(r, 6, 6)

        # ── subtle background image ──────────────────────────────────────────
        if not self._collapsed and not self._bg_pixmap.isNull():
            sz = int(self.height() * 1.1)
            scaled = self._bg_pixmap.scaled(
                sz, sz, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            p.setOpacity(0.07)
            p.drawPixmap(self.width() - scaled.width() - 2,
                         (self.height() - scaled.height()) // 2, scaled)
            p.setOpacity(1.0)

        # ── separator lines ──────────────────────────────────────────────────
        if not self._collapsed:
            p.setPen(QPen(QColor(g.red(), g.green(), g.blue(), 35), 1))
            p.drawLine(6, 22, self.width() - 6, 22)
            p.drawLine(6, 117, self.width() - 6, 117)


# ---------------------------------------------------------------------------
# Startup tips
# ---------------------------------------------------------------------------

_TIPS: list[str] = [
    "Drag any tile to rearrange it — hold and drag to move it anywhere on the canvas. Other tiles animate out of the way in real time.",
    "Right-click a tile to access Launch, Edit, Export, Archive, and Delete from a single context menu.",
    "Drop any file from Windows Explorer onto the canvas to instantly create a launch node for it.",
    "Drop a text file (.txt, .md, .py, .js, etc.) onto the canvas to import it as a Notebook note and create a linked Note tile automatically.",
    "Drop an image file (.png, .jpg, .bmp, .gif, .webp) onto the canvas to create a File node with that image as its icon.",
    "Enable 'Auto Launch' in the node wizard to open that app or file automatically every time Command Center starts.",
    "Folders hold multiple 1×1 nodes — open a folder to see and launch everything inside.",
    "Drag a 1×1 node tile onto a folder tile to nest it inside the folder.",
    "Right-click a folder for the full menu: Open, Rename, Empty Folder, Delete Folder (keep nodes), or Delete Folder and Contents.",
    "'Delete Folder (keep nodes)' removes the folder but returns all its nodes to the main canvas. 'Delete Folder and Contents' permanently deletes everything.",
    "Use Ctrl+F (or the 🔍 button in the title bar) to search the canvas by name or tag — matching tiles float to the front instantly.",
    "Search also finds nodes nested inside folders — they surface as temporary tiles on the canvas so you can launch or edit them directly.",
    "Right-click a search result from inside a folder and choose 'Remove From Folder' to move it permanently to the main canvas.",
    "Note tiles (📝) link directly to Notebook pages — hover one to preview the content without opening the window.",
    "The built-in Notebook supports rich text: bold, italic, headings, bullet lists, numbered lists, hyperlinks, code blocks, and embedded images.",
    "Paste an image from the clipboard (Ctrl+V) directly into a Notebook note — a size picker appears so you choose the display width.",
    "Click the 📌 pin button in the Notebook title bar to instantly create a Note tile on the canvas linked to the open note.",
    "The Time Tracker counts unlocked (green) and locked/away (red) time separately, with a lap counter and full session log.",
    "Click the orange decimal number in the Time Tracker HUD to reset the lap timer — great for tracking billable chunks within a session.",
    "Click the white decimal number in the Time Tracker HUD to copy your total active hours to the clipboard.",
    "Use the ≡ button on the Time Tracker to open the full session log — filter by event type or search by keyword.",
    "The Time Tracker log shows day separator rows so you can quickly see which entries belong to which work day.",
    "Press Ctrl+Shift+T to open or bring the Time Tracker HUD to the front from anywhere in the app.",
    "Use the ± button on the Time Tracker or scroll the mouse wheel over it to manually adjust the active time by ± minutes.",
    "The Time Tracker's '–' button collapses it to a compact strip — handy when you need more screen space.",
    "Export any node as a .node file for backup or sharing — drag it back onto the canvas to import it again.",
    "Switch themes instantly in Settings → Appearance — every open window updates live with no restart needed.",
    "Use the UI Brightness slider in Settings → Appearance to scale all theme colors brighter or dimmer. The range is 30%–200%.",
    "Archived nodes are never permanently deleted — restore them any time from Settings → Archived Nodes.",
    "You can set a custom cursor in Settings → Appearance, with adjustable size. Changes apply instantly.",
    "Choose a Canvas Background style (Solid, Dots, Grid, Noise, Gradient) in Settings → Appearance.",
    "The footer toolbar gives you quick access to Task Manager, Notepad++, Calculator, File Explorer, and Lock Screen.",
    "Resize the main window by dragging any edge or corner, or double-click the title bar to maximise/restore.",
    "Nodes support an optional Accent Color — right-click → Edit and pick a color to give any tile its own border tint.",
    "Assign tags to nodes in the node wizard. Tags are searchable with Ctrl+F just like names.",
    "URL nodes automatically add 'https://' if you forget the protocol — just paste the domain.",
    "Drag a .node file onto the canvas at any time to import a previously exported node.",
]


class TipOfDayDialog(QDialog):
    """Full-screen scrim overlay popup shown at startup with a random tip."""

    _HEADER_H = 62   # height of the painted accent header strip

    def __init__(self, parent=None):
        super().__init__(parent, Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        # Cover the parent window to create a modal scrim effect
        if parent:
            self.setGeometry(parent.geometry())
        else:
            self.resize(1000, 680)
        self._build_ui()

    def _build_ui(self):
        t = _theme
        g = t.GLOW

        # Center the card within the full-size dialog using stretch spacers
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addStretch()
        hrow = QHBoxLayout()
        hrow.setContentsMargins(0, 0, 0, 0)
        hrow.addStretch()

        self._card = QWidget()
        self._card.setFixedWidth(520)
        self._card.setStyleSheet("background:transparent;")

        cl = QVBoxLayout(self._card)
        # Top margin = header height so painted header doesn't overlap content
        cl.setContentsMargins(26, self._HEADER_H + 10, 26, 22)
        cl.setSpacing(14)

        # ── Tip text ─────────────────────────────────────────────────────
        tip = random.choice(_TIPS)
        self._tip_lbl = QLabel(tip)
        self._tip_lbl.setFont(QFont("Segoe UI", 10))
        self._tip_lbl.setStyleSheet(
            f"color:{t.TEXT_PRIMARY.name()}; background:transparent;")
        self._tip_lbl.setWordWrap(True)
        self._tip_lbl.setMinimumHeight(50)
        cl.addWidget(self._tip_lbl)

        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(
            f"color:rgba({g.red()},{g.green()},{g.blue()},55); background:transparent;")
        cl.addWidget(sep)

        # ── Footer: checkbox + button ────────────────────────────────────
        foot = QHBoxLayout(); foot.setSpacing(10)
        self._no_tips_chk = QCheckBox("Don't show tips at startup")
        self._no_tips_chk.setFont(FONT_SMALL)
        self._no_tips_chk.setStyleSheet(f"""
            QCheckBox {{ color:{t.TEXT_DIM.name()}; spacing:6px; background:transparent; }}
            QCheckBox::indicator {{
                width:14px; height:14px; border-radius:3px;
                border:1px solid rgba({g.red()},{g.green()},{g.blue()},80);
                background:rgba({t.TILE_BG_BASE.red()},{t.TILE_BG_BASE.green()},{t.TILE_BG_BASE.blue()},200);
            }}
            QCheckBox::indicator:checked {{
                background:rgba({g.red()},{g.green()},{g.blue()},200);
                border-color:{g.name()};
            }}
        """)
        foot.addWidget(self._no_tips_chk)
        foot.addStretch()

        dr, dg, db = int(g.red()*0.65), int(g.green()*0.65), int(g.blue()*0.65)
        got_btn = QPushButton("Got it!")
        got_btn.setFont(QFont("Segoe UI", 10, QFont.Weight.DemiBold))
        got_btn.setFixedHeight(34)
        got_btn.setCursor(Qt.PointingHandCursor)
        got_btn.setStyleSheet(f"""
            QPushButton {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 rgba({g.red()},{g.green()},{g.blue()},220),
                    stop:1 rgba({dr},{dg},{db},220));
                border: none; border-radius: 8px;
                color: #ffffff; padding: 0 24px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 rgba({g.red()},{g.green()},{g.blue()},255),
                    stop:1 rgba({dr},{dg},{db},255));
            }}
        """)
        got_btn.clicked.connect(self._dismiss)
        foot.addWidget(got_btn)
        cl.addLayout(foot)

        hrow.addWidget(self._card)
        hrow.addStretch()
        outer.addLayout(hrow)
        outer.addStretch()

        # Card mouse: top-right ✕ click → dismiss; drag to reposition
        self._card.mousePressEvent = self._on_card_press
        self._card.mouseMoveEvent  = self._on_card_move
        self._card.mouseReleaseEvent = self._on_card_release
        self._drag_pos: Optional[QPoint] = None

    # ── Mouse handling ────────────────────────────────────────────────────

    def _on_card_press(self, ev: QMouseEvent):
        if ev.button() != Qt.LeftButton: return
        pos = ev.position().toPoint()
        # Top-right close zone (painted ✕)
        if pos.y() < self._HEADER_H and pos.x() > self._card.width() - 52:
            self._dismiss(); return
        self._drag_pos = ev.globalPosition().toPoint()

    def _on_card_move(self, ev: QMouseEvent):
        if self._drag_pos:
            delta = ev.globalPosition().toPoint() - self._drag_pos
            self.move(self.pos() + delta)
            self._drag_pos = ev.globalPosition().toPoint()

    def _on_card_release(self, ev: QMouseEvent):
        self._drag_pos = None

    def mousePressEvent(self, ev: QMouseEvent):
        """Click on scrim (outside card) also dismisses."""
        if ev.button() == Qt.LeftButton:
            if hasattr(self, "_card") and not self._card.geometry().contains(
                    ev.position().toPoint()):
                self._dismiss()

    def keyPressEvent(self, ev):
        if ev.key() == Qt.Key_Escape:
            self._dismiss()
        else:
            super().keyPressEvent(ev)

    def _dismiss(self):
        if self._no_tips_chk.isChecked():
            _settings_store.setValue("show_tips", "false")
            _settings_store.sync()
        self.accept()

    def paintEvent(self, e: QPaintEvent):
        t = _theme
        g = t.GLOW
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)

        # ── Dark scrim over the entire dialog ─────────────────────────────
        p.fillRect(self.rect(), QColor(0, 0, 0, 175))

        if not hasattr(self, "_card"):
            p.end(); return
        cr = self._card.geometry()
        if cr.width() == 0:
            p.end(); return

        card_rect = QRectF(cr.x(), cr.y(), cr.width(), cr.height())
        radius = float(BORDER_RADIUS + 4)

        # ── Glow aura (soft rings expanding outward) ──────────────────────
        for i in range(18, 0, -2):
            aura = QPainterPath()
            aura.addRoundedRect(card_rect.adjusted(-i, -i, i, i),
                                radius + i, radius + i)
            p.setPen(Qt.NoPen)
            p.setBrush(QColor(g.red(), g.green(), g.blue(), max(0, 40 - i * 2)))
            p.drawPath(aura)

        # ── Card body ─────────────────────────────────────────────────────
        card_path = QPainterPath()
        card_path.addRoundedRect(card_rect, radius, radius)
        bg = QLinearGradient(card_rect.topLeft(), card_rect.bottomLeft())
        bg.setColorAt(0, QColor(min(255, t.BG_MID.red() + 20),
                                min(255, t.BG_MID.green() + 20),
                                min(255, t.BG_MID.blue() + 28), 255))
        bg.setColorAt(1, QColor(t.BG_DARK.red(), t.BG_DARK.green(),
                                t.BG_DARK.blue(), 255))
        p.setPen(Qt.NoPen)
        p.setBrush(bg)
        p.drawPath(card_path)

        # ── Accent header strip (clipped to card top) ─────────────────────
        p.save()
        p.setClipPath(card_path)
        accent_rect = QRectF(card_rect.x(), card_rect.y(),
                             card_rect.width(), float(self._HEADER_H))
        hg = QLinearGradient(accent_rect.topLeft(), accent_rect.bottomLeft())
        hg.setColorAt(0.0, QColor(g.red(), g.green(), g.blue(), 130))
        hg.setColorAt(0.7, QColor(g.red(), g.green(), g.blue(), 55))
        hg.setColorAt(1.0, QColor(g.red(), g.green(), g.blue(), 0))
        p.setBrush(hg); p.setPen(Qt.NoPen)
        p.drawRect(accent_rect)
        p.restore()

        # ── Header: bulb icon ─────────────────────────────────────────────
        p.setFont(QFont("Segoe UI Emoji", 20))
        p.setPen(QColor(255, 255, 255, 240))
        p.drawText(
            QRectF(card_rect.x() + 18, card_rect.y() + 8, 46, 46),
            Qt.AlignCenter, "💡")

        # ── Header: title text ────────────────────────────────────────────
        p.setFont(QFont("Segoe UI", 13, QFont.Weight.DemiBold))
        p.setPen(QColor(255, 255, 255, 250))
        p.drawText(
            QRectF(card_rect.x() + 70, card_rect.y() + 12,
                   card_rect.width() - 116, 40),
            Qt.AlignVCenter | Qt.AlignLeft, "Tip of the Day")

        # ── Header: ✕ close hint ──────────────────────────────────────────
        p.setFont(QFont("Segoe UI", 12))
        p.setPen(QColor(255, 255, 255, 130))
        p.drawText(
            QRectF(card_rect.right() - 46, card_rect.y() + 8, 36, 36),
            Qt.AlignCenter, "✕")

        # ── Bright glowing border ─────────────────────────────────────────
        p.setPen(QPen(QColor(g.red(), g.green(), g.blue(), 220), 2.0))
        p.setBrush(Qt.NoBrush)
        p.drawPath(card_path)

        p.end()


# ---------------------------------------------------------------------------
# Startup splash screen
# ---------------------------------------------------------------------------

class SplashScreen(QWidget):
    """Frameless splash shown on startup: logo + 'Command Center' decode animation.

    Timeline:
        0 ms  – shown, decode starts immediately
        ~2 s  – all characters resolved (13 non-space chars × 5 flashes × 30 ms)
        ~2.4s – opacity fade-out completes → ``finished`` emitted
    """
    finished = Signal()

    _CHARSET   = "!@#$%^&*<>?/|0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    _FLASHES   = 5      # ticks per character before locking
    _TICK_MS   = 30     # ms per decode step
    _FADE_STEP = 0.04   # opacity decrement per 16 ms frame  (~400 ms total)

    def __init__(self):
        super().__init__(None,
                         Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setWindowOpacity(1.0)

        _icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                  "CommandCenter.png")
        self._logo = QPixmap(_icon_path) if os.path.isfile(_icon_path) else QPixmap()
        if not self._logo.isNull():
            self._logo = self._logo.scaled(
                110, 110, Qt.KeepAspectRatio, Qt.SmoothTransformation)

        self._full_text    = "Command Center"
        # Start fully scrambled
        self._display_text = "".join(
            " " if c == " " else random.choice(self._CHARSET)
            for c in self._full_text)
        self._decode_char    = 0
        self._decode_flashes = 0
        self._decode_done    = False

        # Skip any leading spaces
        while (self._decode_char < len(self._full_text)
               and self._full_text[self._decode_char] == " "):
            self._decode_char += 1

        self.setFixedSize(380, 270)
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.center() - self.rect().center())

        self._decode_timer = QTimer(self)
        self._decode_timer.setInterval(self._TICK_MS)
        self._decode_timer.timeout.connect(self._decode_tick)
        self._decode_timer.start()

        self._fade_timer = QTimer(self)
        self._fade_timer.setInterval(16)
        self._fade_timer.timeout.connect(self._fade_tick)

    # ── decode ────────────────────────────────────────────────────────────────

    def _decode_tick(self):
        full = self._full_text
        if self._decode_char >= len(full):
            self._finish_decode()
            return

        if self._decode_flashes < self._FLASHES:
            self._decode_flashes += 1
        else:
            self._decode_char   += 1
            while (self._decode_char < len(full)
                   and full[self._decode_char] == " "):
                self._decode_char += 1
            self._decode_flashes  = 0
            if self._decode_char >= len(full):
                self._finish_decode()
                return

        idx    = self._decode_char
        result = []
        for i, c in enumerate(full):
            if c == " ":
                result.append(" ")
            elif i < idx:
                result.append(c)
            else:
                result.append(random.choice(self._CHARSET))
        self._display_text = "".join(result)
        self.update()

    def _finish_decode(self):
        self._decode_done    = True
        self._display_text   = self._full_text
        self._decode_timer.stop()
        self.update()
        QTimer.singleShot(80, self._fade_timer.start)

    def _fade_tick(self):
        op = self.windowOpacity() - self._FADE_STEP
        if op <= 0:
            self._fade_timer.stop()
            self.hide()
            self.finished.emit()
            return
        self.setWindowOpacity(op)

    # ── painting ──────────────────────────────────────────────────────────────

    def paintEvent(self, _e: QPaintEvent):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        r = self.rect().adjusted(0, 0, -1, -1)
        t = _theme
        g = t.GLOW

        # Background gradient
        grad = QLinearGradient(0, 0, 0, self.height())
        grad.setColorAt(0.0, QColor(t.BG_DARK.red(), t.BG_DARK.green(), t.BG_DARK.blue(), 248))
        grad.setColorAt(1.0, QColor(t.BG_MID.red(),  t.BG_MID.green(),  t.BG_MID.blue(),  248))
        p.setBrush(QBrush(grad))
        p.setPen(QPen(QColor(g.red(), g.green(), g.blue(), 90), 1.2))
        p.drawRoundedRect(r, 14, 14)

        # Soft radial glow behind logo
        cx = self.width() // 2
        rg = QRadialGradient(cx, 100, 90)
        rg.setColorAt(0.0, QColor(g.red(), g.green(), g.blue(), 38))
        rg.setColorAt(1.0, QColor(g.red(), g.green(), g.blue(),  0))
        p.setBrush(QBrush(rg)); p.setPen(Qt.NoPen)
        p.drawEllipse(cx - 90, 10, 180, 180)

        # Logo
        if not self._logo.isNull():
            lx = (self.width() - self._logo.width()) // 2
            p.drawPixmap(lx, 38, self._logo)

        # Decoded / decoding text
        p.setFont(QFont("Consolas", 22, QFont.Weight.Bold))
        if self._decode_done:
            txt_color = QColor(t.TEXT_PRIMARY)
        else:
            txt_color = QColor(g.red(), g.green(), g.blue(), 215)
        p.setPen(txt_color)
        p.drawText(QRect(0, 170, self.width(), 44), Qt.AlignCenter, self._display_text)

        # Version hint
        p.setFont(QFont("Consolas", 7))
        p.setPen(QColor(t.TEXT_DIM.red(), t.TEXT_DIM.green(), t.TEXT_DIM.blue(), 140))
        p.drawText(QRect(0, self.height() - 28, self.width(), 22),
                   Qt.AlignCenter, f"v{APP_VERSION}")
        p.end()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def _restart_application() -> None:
    """Spawn a fresh Command Center process then quit the current one."""
    try:
        subprocess.Popen(
            [sys.executable] + sys.argv,
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
            close_fds=False,
        )
    except Exception as _exc:
        print(f"[CommandCenter] Restart error: {_exc}", file=sys.stderr)
    finally:
        app = QApplication.instance()
        if app is not None:
            app.quit()


def main():
    # Set Windows taskbar AppUserModelID so the correct icon appears
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("CommandCenter.App")
    except Exception:
        pass

    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv)
    app.setOrganizationName("CommandCenter")
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setStyleSheet("QToolTip { background-color:transparent; border:none; }")

    # Ensure data directories exist now that QApplication is live
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    ICON_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    NOTEBOOK_DIR.mkdir(parents=True, exist_ok=True)
    PLUGIN_DIR.mkdir(parents=True, exist_ok=True)
    for _d in (MEDIA_LIBRARY_DIR, MEDIA_LIBRARY_GIFS, MEDIA_LIBRARY_PICS):
        _d.mkdir(parents=True, exist_ok=True)

    # Load window icon from the same directory as this script
    _icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "CommandCenter.png")
    if os.path.isfile(_icon_path):
        app.setWindowIcon(QIcon(_icon_path))

    # Apply the saved custom cursor (if any)
    _apply_app_cursor()

    store = NodeStore()
    win = MainWindow(store)

    if _settings_store.value("skip_startup_anim", "false") == "true":
        win.show()
        QTimer.singleShot(0, win._maybe_show_tip)
    else:
        splash = SplashScreen()
        def _after_splash():
            win.show()
            QTimer.singleShot(250, win._maybe_show_tip)
        splash.finished.connect(_after_splash)
        splash.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
