"""
Caffeine v2 — CommandCenter keep-alive plugin
==============================================
Prevents Windows from sleeping / locking via three independent, configurable
keep-alive methods:

  1. Mouse Jiggle      — SendInput relative move (up then down), fully
                         invisible when cursor-restore is enabled.
  2. Ghost Keystroke   — SendInput key press/release for a virtual key that
                         is almost never captured by applications (F15 by
                         default). Shift and Ctrl are also available.
  3. Execution State   — SetThreadExecutionState so the OS knows the app is
                         actively running and should not sleep.

Detection-avoidance jitter randomises the interval and pixel-delta each cycle
so automated monitoring tools cannot detect a fixed pattern.

All configuration is persisted via api.settings and exposed in a Settings tab
and a dedicated status dialog opened from the footer button.
"""

# ---------------------------------------------------------------------------
# Standard library imports
# ---------------------------------------------------------------------------
import ctypes
import ctypes.wintypes
import random

# ---------------------------------------------------------------------------
# PySide6 imports  (use exact names that match the guide examples)
# ---------------------------------------------------------------------------
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QGroupBox,
    QLabel, QPushButton, QCheckBox, QSpinBox, QComboBox,
    QFrame, QWidget, QSizePolicy, QScrollArea,
)
from PySide6.QtCore import Qt, QTimer, QObject, Signal
from PySide6.QtGui  import QFont, QColor

# ===========================================================================
# Module-level state  (reset on every activate() call)
# ===========================================================================
_api          = None   # CommandCenterAPI object
_btn          = None   # footer QPushButton handle
_TAB_LABEL    = "Caffeine"

_jiggle_timer = None   # current single-shot QTimer (or None)
_active       = False  # whether the keep-alive loop is running
_sig          = None   # _Sig instance for cross-thread toasts

# ===========================================================================
# Windows API setup
# ===========================================================================
_user32   = ctypes.windll.user32
_kernel32 = ctypes.windll.kernel32

# SetThreadExecutionState flags
_ES_CONTINUOUS        = 0x80000000
_ES_SYSTEM_REQUIRED   = 0x00000001
_ES_DISPLAY_REQUIRED  = 0x00000002

# SendInput type constants
_INPUT_MOUSE    = 0
_INPUT_KEYBOARD = 1

# Mouse event flags
_MOUSEEVENTF_MOVE = 0x0001

# Key event flags
_KEYEVENTF_KEYUP = 0x0002

# Virtual keys (safe ghost keys)
_VK_F15    = 0x7E
_VK_SHIFT  = 0x10
_VK_CTRL   = 0x11
_VK_MAP    = {"F15": _VK_F15, "Shift": _VK_SHIFT, "Ctrl": _VK_CTRL}

# ---------------------------------------------------------------------------
# ctypes structs for SendInput
# ---------------------------------------------------------------------------

class _MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx",          ctypes.c_long),
        ("dy",          ctypes.c_long),
        ("mouseData",   ctypes.c_ulong),
        ("dwFlags",     ctypes.c_ulong),
        ("time",        ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]

class _KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk",         ctypes.c_ushort),
        ("wScan",       ctypes.c_ushort),
        ("dwFlags",     ctypes.c_ulong),
        ("time",        ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]

class _INPUT_UNION(ctypes.Union):
    _fields_ = [
        ("mi", _MOUSEINPUT),
        ("ki", _KEYBDINPUT),
    ]

class _INPUT(ctypes.Structure):
    _fields_ = [
        ("type",   ctypes.c_ulong),
        ("_input", _INPUT_UNION),
    ]

# ===========================================================================
# Default settings
# ===========================================================================
_DEFAULTS = {
    "interval_sec":  30,
    "jitter_pct":    20,       # ± % of interval to randomise timing
    "mouse_on":      True,
    "mouse_px":      10,       # pixels to move up then down
    "mouse_restore": True,     # warp cursor back after jiggle
    "key_on":        False,
    "key_vk":        "F15",    # ghost key choice
    "es_on":         True,     # SetThreadExecutionState
    "es_display":    False,    # also assert ES_DISPLAY_REQUIRED
    "autostart":     False,    # activate on plugin load
    "active":        False,    # persisted run state across reloads
}

# ===========================================================================
# Settings helpers
# ===========================================================================

def _cfg(key):
    """Read a setting, coercing to the correct type from _DEFAULTS."""
    default = _DEFAULTS[key]
    raw = _api.settings.value(key, default)
    if isinstance(default, bool):
        if isinstance(raw, str):
            return raw.lower() in ("true", "1", "yes")
        return bool(raw)
    if isinstance(default, int):
        try:
            return int(raw)
        except (TypeError, ValueError):
            return default
    return raw

# ===========================================================================
# Windows helpers
# ===========================================================================

def _get_cursor_pos():
    """Return (x, y) of the current cursor position."""
    pt = ctypes.wintypes.POINT()
    _user32.GetCursorPos(ctypes.byref(pt))
    return pt.x, pt.y


def _send_mouse_move(dx, dy):
    """Send a relative mouse movement using SendInput (harder to detect than SetCursorPos)."""
    inp = _INPUT()
    inp.type = _INPUT_MOUSE
    inp._input.mi.dx        = dx
    inp._input.mi.dy        = dy
    inp._input.mi.mouseData = 0
    inp._input.mi.dwFlags   = _MOUSEEVENTF_MOVE
    inp._input.mi.time      = 0
    try:
        _user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(_INPUT))
    except Exception as exc:
        _api.log(f"SendInput (mouse) error: {exc}")


def _send_key_tap(vk):
    """Press and release a virtual key using SendInput."""
    try:
        inputs = (_INPUT * 2)()
        for i, flags in enumerate((0, _KEYEVENTF_KEYUP)):
            inputs[i].type             = _INPUT_KEYBOARD
            inputs[i]._input.ki.wVk   = vk
            inputs[i]._input.ki.wScan = 0
            inputs[i]._input.ki.dwFlags = flags
            inputs[i]._input.ki.time  = 0
        _user32.SendInput(2, inputs, ctypes.sizeof(_INPUT))
    except Exception as exc:
        _api.log(f"SendInput (key) error: {exc}")


def _set_execution_state(display=False):
    """Assert ES_SYSTEM_REQUIRED (and optionally ES_DISPLAY_REQUIRED)."""
    try:
        flags = _ES_CONTINUOUS | _ES_SYSTEM_REQUIRED
        if display:
            flags |= _ES_DISPLAY_REQUIRED
        _kernel32.SetThreadExecutionState(flags)
    except Exception as exc:
        _api.log(f"SetThreadExecutionState error: {exc}")


def _clear_execution_state():
    """Release the execution-state assertion."""
    try:
        _kernel32.SetThreadExecutionState(_ES_CONTINUOUS)
    except Exception as exc:
        _api.log(f"ClearExecutionState error: {exc}")

# ===========================================================================
# Keep-alive core
# ===========================================================================

def _jittered_ms():
    """Return an interval (ms) with ±jitter applied."""
    base_ms = _cfg("interval_sec") * 1000
    pct     = _cfg("jitter_pct") / 100.0
    spread  = int(base_ms * pct)
    return base_ms + random.randint(-spread, spread) if spread > 0 else base_ms


def _do_keepalive():
    """Execute all enabled keep-alive methods. Called on the Qt main thread."""
    if not _active:
        return
    try:
        # 1. Execution state
        if _cfg("es_on"):
            _set_execution_state(display=_cfg("es_display"))

        # 2. Mouse jiggle
        if _cfg("mouse_on"):
            px = _cfg("mouse_px")
            ox, oy = _get_cursor_pos()

            # Vary the actual pixel delta slightly per cycle (avoidance)
            up   = px + random.randint(0, max(1, px // 4))
            down = px + random.randint(0, max(1, px // 4))

            # Move up
            _send_mouse_move(0, -up)

            # Move down after a small organic delay
            delay_down = 60 + random.randint(0, 60)
            QTimer.singleShot(delay_down, lambda: _send_mouse_move(0, down))

            # Restore cursor position after both moves complete
            if _cfg("mouse_restore"):
                cx, cy = ox, oy
                def _restore():
                    nx, ny = _get_cursor_pos()
                    # Only restore if nobody moved the mouse significantly
                    if abs(nx - cx) < px * 4 and abs(ny - cy) < px * 4:
                        _user32.SetCursorPos(cx, cy)
                QTimer.singleShot(delay_down + 80, _restore)

        # 3. Ghost keystroke
        if _cfg("key_on"):
            vk = _VK_MAP.get(_cfg("key_vk"), _VK_F15)
            # Slight random delay so key never coincides exactly with mouse
            key_delay = random.randint(40, 280)
            QTimer.singleShot(key_delay, lambda: _send_key_tap(vk))

    except Exception as exc:
        _api.log(f"keep-alive error: {exc}")


def _tick():
    """Called by the single-shot timer: run keep-alive then schedule next tick."""
    _do_keepalive()
    _schedule_next()


def _schedule_next():
    """Cancel any pending timer and create a fresh single-shot with jittered interval."""
    global _jiggle_timer
    if _jiggle_timer is not None:
        try:
            _api.timers.cancel(_jiggle_timer)
        except Exception:
            pass
        _jiggle_timer = None
    if _active:
        _jiggle_timer = _api.timers.create(_jittered_ms(), _tick, single_shot=True)


def _start():
    global _active
    if _active:
        return
    _active = True
    _api.settings.set("active", True)
    _set_execution_state(display=_cfg("es_display"))
    _schedule_next()
    _api.toast("Caffeine activated ☕", "success")
    _refresh_btn_label()


def _stop():
    global _active, _jiggle_timer
    if not _active:
        return
    _active = False
    _api.settings.set("active", False)
    if _jiggle_timer is not None:
        try:
            _api.timers.cancel(_jiggle_timer)
        except Exception:
            pass
        _jiggle_timer = None
    _clear_execution_state()
    _api.toast("Caffeine deactivated.", "info")
    _refresh_btn_label()


def _toggle():
    """Toggle the keep-alive loop on / off."""
    if _active:
        _stop()
    else:
        _start()


def _refresh_btn_label():
    """Update footer button text to reflect current state."""
    if _btn is not None:
        try:
            _btn.setText("☕ ON" if _active else "☕ Caffeine")
        except Exception:
            pass

# ===========================================================================
# Settings tab widget
# ===========================================================================

def _build_settings_widget(on_saved=None):
    """Build and return the QWidget injected into the Settings dialog tab."""
    root = QWidget()
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QFrame.NoFrame)

    inner = QWidget()
    lay = QVBoxLayout(inner)
    lay.setContentsMargins(16, 16, 16, 16)
    lay.setSpacing(14)

    # ── Timing ────────────────────────────────────────────────────────────
    grp_timing = QGroupBox("Timing")
    fl = QFormLayout(grp_timing)
    fl.setSpacing(8)

    spin_interval = QSpinBox()
    spin_interval.setRange(5, 3600)
    spin_interval.setSuffix("  sec")
    spin_interval.setValue(_cfg("interval_sec"))
    spin_interval.setToolTip(
        "Base interval between keep-alive pulses.\n"
        "Actual interval will vary by the Jitter % below."
    )
    fl.addRow("Interval:", spin_interval)

    spin_jitter = QSpinBox()
    spin_jitter.setRange(0, 50)
    spin_jitter.setSuffix("  %")
    spin_jitter.setValue(_cfg("jitter_pct"))
    spin_jitter.setToolTip(
        "Randomise timing by ±this percentage each cycle.\n"
        "20 % means a 30-second interval fires anywhere between 24–36 s.\n"
        "Set to 0 for an exact fixed interval."
    )
    fl.addRow("Avoidance jitter:", spin_jitter)

    lay.addWidget(grp_timing)

    # ── Mouse Jiggle ─────────────────────────────────────────────────────
    grp_mouse = QGroupBox("Mouse Jiggle")
    grp_mouse.setCheckable(True)
    grp_mouse.setChecked(_cfg("mouse_on"))
    ml = QFormLayout(grp_mouse)
    ml.setSpacing(8)

    spin_px = QSpinBox()
    spin_px.setRange(1, 200)
    spin_px.setSuffix("  px")
    spin_px.setValue(_cfg("mouse_px"))
    spin_px.setToolTip(
        "Distance in pixels to move the cursor up then back down.\n"
        "Small values (5–15) are invisible in practice."
    )
    ml.addRow("Jiggle distance:", spin_px)

    chk_restore = QCheckBox("Restore cursor to original position after each jiggle")
    chk_restore.setChecked(_cfg("mouse_restore"))
    chk_restore.setToolTip(
        "After moving up and down, warp the cursor back to exactly where it was.\n"
        "Highly recommended — makes the jiggle completely invisible."
    )
    ml.addRow("", chk_restore)

    lay.addWidget(grp_mouse)

    # ── Ghost Keystroke ───────────────────────────────────────────────────
    grp_key = QGroupBox("Ghost Keystroke")
    grp_key.setCheckable(True)
    grp_key.setChecked(_cfg("key_on"))
    kl = QFormLayout(grp_key)
    kl.setSpacing(8)

    combo_key = QComboBox()
    for opt in _VK_MAP.keys():
        combo_key.addItem(opt)
    idx = combo_key.findText(_cfg("key_vk"))
    if idx >= 0:
        combo_key.setCurrentIndex(idx)
    combo_key.setToolTip(
        "F15  — safest; almost never captured by any application.\n"
        "Shift — riskier but effective at resetting some idle detectors.\n"
        "Ctrl  — same risk level as Shift."
    )
    kl.addRow("Virtual key:", combo_key)

    lay.addWidget(grp_key)

    # ── Execution State ───────────────────────────────────────────────────
    grp_es = QGroupBox("Windows Execution State")
    grp_es.setCheckable(True)
    grp_es.setChecked(_cfg("es_on"))
    el = QFormLayout(grp_es)
    el.setSpacing(8)

    hint_es = QLabel(
        "Calls SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED) "
        "so Windows knows the app is active and will not sleep the CPU."
    )
    hint_es.setWordWrap(True)
    el.addRow(hint_es)

    chk_display = QCheckBox("Also prevent display sleep  (ES_DISPLAY_REQUIRED)")
    chk_display.setChecked(_cfg("es_display"))
    chk_display.setToolTip(
        "Assert ES_DISPLAY_REQUIRED as well.\n"
        "Keeps the monitor on in addition to the system."
    )
    el.addRow("", chk_display)

    lay.addWidget(grp_es)

    # ── Startup ───────────────────────────────────────────────────────────
    grp_startup = QGroupBox("Startup")
    sl = QFormLayout(grp_startup)
    sl.setSpacing(8)

    chk_autostart = QCheckBox("Activate Caffeine automatically when the plugin loads")
    chk_autostart.setChecked(_cfg("autostart"))
    sl.addRow("", chk_autostart)

    lay.addWidget(grp_startup)

    # ── Change tracking ───────────────────────────────────────────────────
    dirty_lbl = QLabel("  ● Unsaved changes")
    dirty_lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
    dirty_lbl.setVisible(False)
    lay.addWidget(dirty_lbl)

    def _mark_dirty():
        dirty_lbl.setVisible(True)

    spin_interval.valueChanged.connect(_mark_dirty)
    spin_jitter.valueChanged.connect(_mark_dirty)
    grp_mouse.toggled.connect(_mark_dirty)
    spin_px.valueChanged.connect(_mark_dirty)
    chk_restore.toggled.connect(_mark_dirty)
    grp_key.toggled.connect(_mark_dirty)
    combo_key.currentIndexChanged.connect(_mark_dirty)
    grp_es.toggled.connect(_mark_dirty)
    chk_display.toggled.connect(_mark_dirty)
    chk_autostart.toggled.connect(_mark_dirty)

    # ── Save ──────────────────────────────────────────────────────────────
    save_row = QHBoxLayout()
    save_row.addStretch()
    btn_save = QPushButton("Save Settings")
    btn_save.setMinimumWidth(160)
    btn_save.setMinimumHeight(34)
    save_row.addWidget(btn_save)
    lay.addLayout(save_row)
    lay.addStretch()

    def _save():
        try:
            _api.settings.set("interval_sec", spin_interval.value())
            _api.settings.set("jitter_pct",   spin_jitter.value())
            _api.settings.set("mouse_on",     grp_mouse.isChecked())
            _api.settings.set("mouse_px",     spin_px.value())
            _api.settings.set("mouse_restore",chk_restore.isChecked())
            _api.settings.set("key_on",       grp_key.isChecked())
            _api.settings.set("key_vk",       combo_key.currentText())
            _api.settings.set("es_on",        grp_es.isChecked())
            _api.settings.set("es_display",   chk_display.isChecked())
            _api.settings.set("autostart",    chk_autostart.isChecked())
            if _active:
                _schedule_next()  # apply new interval immediately
            dirty_lbl.setVisible(False)
            btn_save.setEnabled(False)
            btn_save.setText("\u2713  Saved!")
            try:
                c = _api.theme.colors()
                btn_save.setStyleSheet(
                    f"QPushButton {{ background:{c['accent_teal']}; color:{c['bg_dark']};"
                    f" border:none; border-radius:4px; padding:5px 16px; font-weight:700; }}"
                )
            except Exception:
                pass
            if on_saved is not None:
                QTimer.singleShot(700, on_saved)
            else:
                def _restore_btn():
                    btn_save.setText("Save Settings")
                    btn_save.setEnabled(True)
                    _apply_theme()
                QTimer.singleShot(1800, _restore_btn)
        except Exception as exc:
            _api.log(f"Settings save error: {exc}")
            btn_save.setEnabled(True)
            btn_save.setText("\u2717  Error \u2014 try again")
            try:
                c = _api.theme.colors()
                btn_save.setStyleSheet(
                    f"QPushButton {{ background:{c['accent_red']}; color:{c['bg_dark']};"
                    f" border:none; border-radius:4px; padding:5px 16px; font-weight:700; }}"
                )
            except Exception:
                pass
            QTimer.singleShot(2500, lambda: btn_save.setText("Save Settings"))

    btn_save.clicked.connect(_save)

    # ── Theme support ─────────────────────────────────────────────────────
    def _apply_theme():
        try:
            c = _api.theme.colors()
            common = (
                f"QWidget   {{ background:{c['bg_dark']}; color:{c['text_primary']}; }}"
                f"QGroupBox {{ background:{c['bg_mid']}; color:{c['text_primary']};"
                f"             border:1px solid {c['glow']}; border-radius:6px;"
                f"             margin-top:10px; padding:10px; font-weight:600; }}"
                f"QGroupBox::title {{ subcontrol-origin:margin; left:10px; padding:0 4px; }}"
                f"QGroupBox::indicator {{ width:14px; height:14px; }}"
                f"QGroupBox::indicator:checked   {{ background:{c['accent_teal']};"
                f"  border:1px solid {c['accent_teal']}; border-radius:3px; }}"
                f"QGroupBox::indicator:unchecked {{ background:{c['bg_dark']};"
                f"  border:1px solid {c['glow']}; border-radius:3px; }}"
                f"QSpinBox  {{ background:{c['bg_dark']}; color:{c['text_primary']};"
                f"             border:1px solid {c['glow']}; border-radius:4px; padding:3px 6px; }}"
                f"QComboBox {{ background:{c['bg_dark']}; color:{c['text_primary']};"
                f"             border:1px solid {c['glow']}; border-radius:4px; padding:3px 6px; }}"
                f"QComboBox QAbstractItemView {{ background:{c['bg_mid']}; color:{c['text_primary']};"
                f"  selection-background-color:{c['accent_blue']}; }}"
                f"QCheckBox {{ color:{c['text_primary']}; }}"
                f"QCheckBox::indicator {{ width:14px; height:14px; }}"
                f"QCheckBox::indicator:checked   {{ background:{c['accent_teal']};"
                f"  border:1px solid {c['accent_teal']}; border-radius:3px; }}"
                f"QCheckBox::indicator:unchecked {{ background:{c['bg_dark']};"
                f"  border:1px solid {c['glow']}; border-radius:3px; }}"
                f"QLabel {{ color:{c['text_secondary']}; }}"
                f"QPushButton {{ background:{c['accent_teal']}; color:{c['bg_dark']};"
                f"               border:none; border-radius:4px; padding:5px 16px;"
                f"               font-weight:600; }}"
                f"QPushButton:hover {{ background:{c['accent_blue']}; color:{c['bg_dark']}; }}"
                f"QScrollArea {{ background:{c['bg_dark']}; border:none; }}"
            )
            inner.setStyleSheet(common)
            dirty_lbl.setStyleSheet(
                f"color:{c['accent_amber']}; font-size:11px; font-weight:600;"
                f" background:transparent;"
            )
        except Exception as exc:
            _api.log(f"Settings tab theme error: {exc}")

    _api.theme.register(_apply_theme)
    _apply_theme()

    scroll.setWidget(inner)
    root_lay = QVBoxLayout(root)
    root_lay.setContentsMargins(0, 0, 0, 0)
    root_lay.addWidget(scroll)
    return root

# ===========================================================================
# Status dialog (opened from footer button)
# ===========================================================================

class _CaffeineDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("☕ Caffeine — Keep-Alive")
        self.setMinimumSize(400, 320)
        self.resize(440, 360)
        self._build_ui()
        _api.theme.register(self._apply_theme)
        self._apply_theme()
        self._count_timer = QTimer(self)
        self._count_timer.setInterval(500)
        self._count_timer.timeout.connect(self._update_pulse_lbl)
        self._refresh()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 22, 24, 20)
        lay.setSpacing(16)

        # ── Status row ────────────────────────────────────────────────────
        status_row = QHBoxLayout()
        self._dot = QLabel("●")
        self._dot.setFont(QFont("Segoe UI", 22))
        self._status_lbl = QLabel()
        self._status_lbl.setFont(QFont("Segoe UI", 12))
        status_row.addWidget(self._dot)
        status_row.addSpacing(10)
        status_row.addWidget(self._status_lbl)
        status_row.addStretch()
        lay.addLayout(status_row)

        # ── Divider ───────────────────────────────────────────────────────
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        lay.addWidget(sep)

        # ── Method summary ────────────────────────────────────────────────
        self._methods_lbl = QLabel()
        self._methods_lbl.setWordWrap(True)
        lay.addWidget(self._methods_lbl)

        self._interval_lbl = QLabel()
        lay.addWidget(self._interval_lbl)

        self._pulse_lbl = QLabel()
        self._pulse_lbl.setVisible(False)
        lay.addWidget(self._pulse_lbl)

        lay.addStretch()

        # ── Buttons ───────────────────────────────────────────────────────
        btn_row = QHBoxLayout()

        self._toggle_btn = QPushButton()
        self._toggle_btn.setMinimumHeight(36)
        self._toggle_btn.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        self._toggle_btn.clicked.connect(self._on_toggle)
        btn_row.addWidget(self._toggle_btn, stretch=2)

        btn_row.addSpacing(8)

        btn_settings = QPushButton("Settings →")
        btn_settings.setMinimumHeight(36)
        btn_settings.clicked.connect(self._on_open_settings)
        btn_row.addWidget(btn_settings, stretch=1)

        lay.addLayout(btn_row)

    def _refresh(self):
        """Update all dynamic labels and button states."""
        try:
            # Status indicator
            if _active:
                self._dot.setText("●")
                self._status_lbl.setText("Running  —  system is awake")
            else:
                self._dot.setText("○")
                self._status_lbl.setText("Inactive  —  system may sleep")

            # Active methods summary
            methods = []
            if _cfg("mouse_on"):
                methods.append(f"Mouse jiggle (±{_cfg('mouse_px')} px)")
            if _cfg("key_on"):
                methods.append(f"Ghost key ({_cfg('key_vk')})")
            if _cfg("es_on"):
                label = "Exec state (system"
                if _cfg("es_display"):
                    label += " + display"
                label += ")"
                methods.append(label)
            self._methods_lbl.setText(
                "Methods: " + (", ".join(methods) if methods else "none enabled")
            )

            # Interval
            j = _cfg("jitter_pct")
            self._interval_lbl.setText(
                f"Interval: every {_cfg('interval_sec')} s"
                + (f"  ±{j} % jitter" if j > 0 else "  (no jitter)")
            )
            # Countdown / pulse label
            if _active:
                self._count_timer.start()
                self._update_pulse_lbl()
            else:
                self._count_timer.stop()
                self._pulse_lbl.setVisible(False)
            # Toggle button
            self._toggle_btn.setText(
                "⏹  Deactivate" if _active else "▶  Activate ☕"
            )

            self._apply_theme()
        except Exception as exc:
            _api.log(f"Dialog refresh error: {exc}")

    def _update_pulse_lbl(self):
        try:
            if not _active or _jiggle_timer is None:
                self._pulse_lbl.setVisible(False)
                return
            rem = _jiggle_timer.remainingTime()
            if rem < 0:
                self._pulse_lbl.setVisible(False)
                return
            secs = max(1, (rem + 999) // 1000)
            self._pulse_lbl.setText(f"Next pulse in  {secs} s")
            self._pulse_lbl.setVisible(True)
        except Exception:
            self._pulse_lbl.setVisible(False)

    def _on_toggle(self):
        try:
            _toggle()
            self._refresh()
        except Exception as exc:
            _api.log(f"Dialog toggle error: {exc}")

    def _on_open_settings(self):
        try:
            dlg = QDialog(self)
            dlg.setWindowTitle("☕ Caffeine — Settings")
            dlg.setMinimumSize(480, 560)
            dlg.resize(500, 600)
            lay = QVBoxLayout(dlg)
            lay.setContentsMargins(0, 0, 0, 0)
            lay.addWidget(_build_settings_widget(on_saved=dlg.accept))
            c = _api.theme.colors()
            dlg.setStyleSheet(f"QDialog {{ background:{c['bg_dark']}; }}")
            dlg.exec()
            self._refresh()  # update status dialog with any changed settings
        except Exception as exc:
            _api.log(f"Open settings error: {exc}")
            _api.toast("Could not open Caffeine settings.", "error")

    def _apply_theme(self):
        try:
            c = _api.theme.colors()
            dot_color    = c["accent_teal"] if _active else c["text_dim"]
            label_color  = c["text_primary"] if _active else c["text_secondary"]
            toggle_bg    = c["accent_red"]  if _active else c["accent_teal"]

            self._dot.setStyleSheet(f"color:{dot_color}; background:transparent;")
            self._status_lbl.setStyleSheet(f"color:{label_color}; background:transparent;")
            self._methods_lbl.setStyleSheet(
                f"color:{c['text_secondary']}; font-size:12px; background:transparent;"
            )
            self._interval_lbl.setStyleSheet(
                f"color:{c['text_dim']}; font-size:11px; background:transparent;"
            )
            self._pulse_lbl.setStyleSheet(
                f"color:{c['accent_teal']}; font-size:11px; font-weight:600;"
                f" background:transparent;"
            )
            self.setStyleSheet(
                f"QDialog  {{ background:{c['bg_dark']}; color:{c['text_primary']}; }}"
                f"QFrame   {{ color:{c['glow']}; }}"
                f"QPushButton {{ background:{c['bg_mid']}; color:{c['text_primary']};"
                f"               border:1px solid {c['glow']}; border-radius:5px;"
                f"               padding:5px 14px; }}"
                f"QPushButton:hover {{ border-color:{c['accent_blue']};"
                f"                     color:{c['accent_blue']}; }}"
            )
            # Override colour for the main toggle button
            self._toggle_btn.setStyleSheet(
                f"QPushButton {{ background:{toggle_bg}; color:{c['bg_dark']};"
                f"               border:none; border-radius:5px;"
                f"               padding:5px 14px; font-weight:700; }}"
                f"QPushButton:hover {{ opacity:0.9; }}"
            )
        except Exception as exc:
            _api.log(f"Dialog theme error: {exc}")


def _open_dialog():
    try:
        dlg = _CaffeineDialog(_api.ui.main_window)
        dlg.exec()
    except Exception as exc:
        _api.log(f"Open dialog error: {exc}")
        _api.toast("Could not open Caffeine dialog.", "error")

# ===========================================================================
# Plugin entry points
# ===========================================================================

def activate(api):
    global _api, _btn, _sig, _active

    _api = api

    # Footer button
    _btn = api.ui.add_footer_button("☕ Caffeine", _open_dialog)

    # Global hotkeys
    api.hotkeys.register(
        "Ctrl+Shift+K", _toggle,      "Caffeine: Toggle keep-alive on/off"
    )
    api.hotkeys.register(
        "Ctrl+Shift+J", _open_dialog, "Caffeine: Open status dialog"
    )

    # Settings tab
    api.ui.add_settings_tab(_TAB_LABEL, _build_settings_widget())

    # Auto-start if configured or if it was running before a reload
    if _cfg("autostart") or _cfg("active"):
        # Delay slightly so the main window is fully settled
        QTimer.singleShot(500, _start)

    _refresh_btn_label()
    api.log("Caffeine v2 activated.")


def deactivate():
    global _btn, _jiggle_timer, _active

    # Stop the keep-alive loop (clears execution state, cancels timer)
    try:
        _stop()
    except Exception:
        pass

    # Belt-and-suspenders: clear execution state even if _stop() failed
    try:
        _clear_execution_state()
    except Exception:
        pass

    # Remove footer button
    if _btn is not None:
        try:
            _api.ui.remove_footer_button(_btn)
        except Exception:
            pass
        _btn = None

    # Unregister hotkeys
    try:
        _api.hotkeys.unregister("Ctrl+Shift+K")
    except Exception:
        pass
    try:
        _api.hotkeys.unregister("Ctrl+Shift+J")
    except Exception:
        pass

    # Remove settings tab
    try:
        _api.ui.remove_settings_tab(_TAB_LABEL)
    except Exception:
        pass

    _active = False
    _api.log("Caffeine v2 deactivated.")
