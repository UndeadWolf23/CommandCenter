"""
Credential Manager Plugin for Command Center — v1.1.0
=============================================
Encryption backend priority:
  1. Windows CNG (bcrypt.dll) via ctypes — always available in compiled .exe, no packages needed
  2. `cryptography` library — fallback for non-Windows / explicit installs

Data format (both backends identical):  base64( nonce[12] || ciphertext || auth_tag[16] )
Vault key: 32-byte CSPRNG random, stored in vault.key (file-permission restricted)
Export key derivation: PBKDF2-HMAC-SHA256 (600k iterations) via stdlib hashlib — no packages.

v1.1.0 improvements:
  - Atomic credential saves (temp-file + os.replace) — data safe against crashes
  - Clipboard clear uses QTimer directly (no CC API dependency)
  - Search debounced (200 ms) to avoid layout thrash on fast typing
  - Copy buttons flash "✓ Copied!" for 1.5 s with visual feedback
  - Fixed expiry label layout bug (always in DOM, visibility toggled)
  - Double-click any credential row to edit
  - Category autocomplete in edit dialog
  - Password length spinbox for the generator (8–64 chars)
  - Tab order and Enter-to-save in edit dialog
  - Export requires passphrase confirmation step
  - Window title shows live credential count
  - Status bar messages auto-clear after 5 s
  - Ctrl+F focuses search, improved empty-state, Ctrl+Shift+K opens vault
"""

from __future__ import annotations

import base64
import ctypes
import hashlib
import json
import os
import re
import secrets
import string
import traceback
from datetime import datetime, date
from pathlib import Path
from typing import Optional

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QWidget, QLabel, QPushButton,
    QLineEdit, QTextEdit, QScrollArea, QFrame, QSizePolicy,
    QDateEdit, QFormLayout, QComboBox, QSpinBox, QCheckBox, QInputDialog,
    QCompleter, QSlider, QMessageBox,
)
from PySide6.QtCore import Qt, QDate, QTimer, Signal
from PySide6.QtGui import QFont, QKeySequence, QShortcut


# ── Windows CNG AES-256-GCM ──────────────────────────────────────────────────

class _AuthInfo(ctypes.Structure):
    """BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO — matches Windows SDK layout on x64."""
    _fields_ = [
        ("cbSize",       ctypes.c_ulong),    # offset  0, size 4
        ("dwInfoVersion",ctypes.c_ulong),    # offset  4, size 4
        ("pbNonce",      ctypes.c_void_p),   # offset  8, size 8
        ("cbNonce",      ctypes.c_ulong),    # offset 16, size 4
        # 4-byte pad to align next pointer
        ("pbAuthData",   ctypes.c_void_p),   # offset 24, size 8
        ("cbAuthData",   ctypes.c_ulong),    # offset 32, size 4
        # 4-byte pad
        ("pbTag",        ctypes.c_void_p),   # offset 40, size 8
        ("cbTag",        ctypes.c_ulong),    # offset 48, size 4
        # 4-byte pad
        ("pbMacContext", ctypes.c_void_p),   # offset 56, size 8
        ("cbMacContext", ctypes.c_ulong),    # offset 64, size 4
        ("cbAAD",        ctypes.c_ulong),    # offset 68, size 4
        ("cbData",       ctypes.c_uint64),   # offset 72, size 8  (aligned ✓)
        ("dwFlags",      ctypes.c_ulong),    # offset 80, size 4
        # 4-byte pad → total 88 bytes
    ]


_NTSTATUS_SUCCESS           = 0
_NTSTATUS_AUTH_TAG_MISMATCH = -1073709054   # 0xC000A002 signed int32


def _cng_aes_gcm_encrypt(bcrypt, alg_handle: ctypes.c_void_p,
                          key_obj_len: int, key32: bytes,
                          plaintext: bytes) -> bytes:
    """Encrypt plaintext with AES-256-GCM via Windows CNG.
    Returns nonce(12) || ciphertext || tag(16).
    """
    nonce = secrets.token_bytes(12)

    key_obj    = ctypes.create_string_buffer(key_obj_len)
    key_handle = ctypes.c_void_p()
    key_arr    = (ctypes.c_ubyte * 32)(*key32)

    st = bcrypt.BCryptGenerateSymmetricKey(
        alg_handle, ctypes.byref(key_handle),
        key_obj, key_obj_len,
        key_arr, 32, 0
    )
    if st != _NTSTATUS_SUCCESS:
        raise RuntimeError(f"BCryptGenerateSymmetricKey: 0x{st & 0xFFFFFFFF:08X}")

    try:
        nonce_buf = (ctypes.c_ubyte * 12)(*nonce)
        tag_buf   = (ctypes.c_ubyte * 16)()
        n         = len(plaintext)
        pt_buf    = (ctypes.c_ubyte * max(n, 1))(*plaintext[:max(n, 1)])
        ct_buf    = (ctypes.c_ubyte * max(n, 1))()

        auth = _AuthInfo()
        auth.cbSize        = ctypes.sizeof(_AuthInfo)
        auth.dwInfoVersion = 1
        auth.pbNonce       = ctypes.cast(nonce_buf, ctypes.c_void_p)
        auth.cbNonce       = 12
        auth.pbTag         = ctypes.cast(tag_buf,   ctypes.c_void_p)
        auth.cbTag         = 16

        out_len = ctypes.c_ulong(0)
        st = bcrypt.BCryptEncrypt(
            key_handle, pt_buf, n,
            ctypes.byref(auth),
            None, 0,
            ct_buf, n,
            ctypes.byref(out_len), 0
        )
        if st != _NTSTATUS_SUCCESS:
            raise RuntimeError(f"BCryptEncrypt: 0x{st & 0xFFFFFFFF:08X}")

        return nonce + bytes(ct_buf[:out_len.value]) + bytes(tag_buf)
    finally:
        bcrypt.BCryptDestroyKey(key_handle)


def _cng_aes_gcm_decrypt(bcrypt, alg_handle: ctypes.c_void_p,
                          key_obj_len: int, key32: bytes,
                          data: bytes) -> bytes:
    """Decrypt AES-256-GCM data (nonce||ct||tag) via Windows CNG."""
    if len(data) < 12 + 16:
        raise ValueError("Data too short for AES-256-GCM decryption")

    nonce = data[:12]
    ct    = data[12:-16]
    tag   = data[-16:]

    key_obj    = ctypes.create_string_buffer(key_obj_len)
    key_handle = ctypes.c_void_p()
    key_arr    = (ctypes.c_ubyte * 32)(*key32)

    st = bcrypt.BCryptGenerateSymmetricKey(
        alg_handle, ctypes.byref(key_handle),
        key_obj, key_obj_len,
        key_arr, 32, 0
    )
    if st != _NTSTATUS_SUCCESS:
        raise RuntimeError(f"BCryptGenerateSymmetricKey: 0x{st & 0xFFFFFFFF:08X}")

    try:
        nonce_buf = (ctypes.c_ubyte * 12)(*nonce)
        tag_buf   = (ctypes.c_ubyte * 16)(*tag)
        n         = len(ct)
        ct_buf    = (ctypes.c_ubyte * max(n, 1))(*ct[:max(n, 1)])
        pt_buf    = (ctypes.c_ubyte * max(n, 1))()

        auth = _AuthInfo()
        auth.cbSize        = ctypes.sizeof(_AuthInfo)
        auth.dwInfoVersion = 1
        auth.pbNonce       = ctypes.cast(nonce_buf, ctypes.c_void_p)
        auth.cbNonce       = 12
        auth.pbTag         = ctypes.cast(tag_buf,   ctypes.c_void_p)
        auth.cbTag         = 16

        out_len = ctypes.c_ulong(0)
        st = bcrypt.BCryptDecrypt(
            key_handle, ct_buf, n,
            ctypes.byref(auth),
            None, 0,
            pt_buf, n,
            ctypes.byref(out_len), 0
        )
        if st != _NTSTATUS_SUCCESS:
            if st == _NTSTATUS_AUTH_TAG_MISMATCH:
                raise ValueError(
                    "Authentication tag mismatch — "
                    "data may be corrupted or tampered with"
                )
            raise RuntimeError(f"BCryptDecrypt: 0x{st & 0xFFFFFFFF:08X}")

        return bytes(pt_buf[:out_len.value])
    finally:
        bcrypt.BCryptDestroyKey(key_handle)


# ── Vault ─────────────────────────────────────────────────────────────────────

class _Vault:
    """
    AES-256-GCM credential vault.
    Backend selected at init:
      1. Windows CNG (bcrypt.dll) — always available in compiled .exe, zero packages
      2. cryptography library — fallback
    Both backends produce identical wire format so data is portable between them.
    """

    KEY_LEN  = 32
    KEY_FILE = "vault.key"

    def __init__(self, plugin_dir: Path, log_fn=None):
        self._log      = log_fn or (lambda m: None)
        self._key_path = plugin_dir / self.KEY_FILE
        self._key      = self._load_or_create_key()
        self._backend  = "none"
        self._bcrypt   = None
        self._cng_alg  = None
        self._cng_klen = 0
        self._init_backend()

    def _load_or_create_key(self) -> bytes:
        if self._key_path.exists():
            raw = self._key_path.read_bytes()
            if len(raw) == self.KEY_LEN:
                return raw
        key = secrets.token_bytes(self.KEY_LEN)
        self._key_path.write_bytes(key)
        try:
            import stat
            os.chmod(self._key_path, stat.S_IRUSR | stat.S_IWUSR)
        except Exception:
            pass
        return key

    def _init_backend(self):
        # --- Try Windows CNG ---
        try:
            bcrypt = ctypes.WinDLL("bcrypt.dll")
            alg    = ctypes.c_void_p()

            st = bcrypt.BCryptOpenAlgorithmProvider(
                ctypes.byref(alg), ctypes.c_wchar_p("AES"), None, 0
            )
            if st != _NTSTATUS_SUCCESS:
                raise RuntimeError(f"BCryptOpenAlgorithmProvider: 0x{st & 0xFFFFFFFF:08X}")

            mode_bytes = "ChainingModeGCM\0".encode("utf-16-le")
            st = bcrypt.BCryptSetProperty(
                alg, ctypes.c_wchar_p("ChainingMode"),
                mode_bytes, len(mode_bytes), 0
            )
            if st != _NTSTATUS_SUCCESS:
                bcrypt.BCryptCloseAlgorithmProvider(alg, 0)
                raise RuntimeError(f"BCryptSetProperty: 0x{st & 0xFFFFFFFF:08X}")

            obj_len = ctypes.c_ulong(0)
            res_len = ctypes.c_ulong(0)
            st = bcrypt.BCryptGetProperty(
                alg, ctypes.c_wchar_p("ObjectLength"),
                ctypes.byref(obj_len), 4,
                ctypes.byref(res_len), 0
            )
            if st != _NTSTATUS_SUCCESS:
                bcrypt.BCryptCloseAlgorithmProvider(alg, 0)
                raise RuntimeError(f"BCryptGetProperty: 0x{st & 0xFFFFFFFF:08X}")

            self._bcrypt   = bcrypt
            self._cng_alg  = alg
            self._cng_klen = obj_len.value
            self._backend  = "cng"
            self._log("Vault backend: Windows CNG (AES-256-GCM)")
            return

        except Exception as cng_err:
            self._log(f"CNG unavailable ({cng_err}), trying cryptography library…")

        # --- Try cryptography library ---
        try:
            from cryptography.hazmat.primitives.ciphers.aead import AESGCM  # noqa: F401
            self._backend = "cryptography"
            self._log("Vault backend: cryptography library (AES-256-GCM)")
            return
        except ImportError:
            pass

        raise RuntimeError(
            "No encryption backend available.\n"
            "Windows CNG (bcrypt.dll) failed AND the 'cryptography' package is not importable.\n"
            "On a compiled .exe this usually means bcrypt.dll could not be opened — "
            "please report this with the error in the log."
        )

    def raw_encrypt(self, key32: bytes, plaintext: bytes) -> bytes:
        """Returns nonce(12) || ciphertext || tag(16)."""
        if self._backend == "cng":
            return _cng_aes_gcm_encrypt(
                self._bcrypt, self._cng_alg, self._cng_klen, key32, plaintext
            )
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        nonce  = secrets.token_bytes(12)
        ct_tag = AESGCM(key32).encrypt(nonce, plaintext, None)
        return nonce + ct_tag

    def raw_decrypt(self, key32: bytes, data: bytes) -> bytes:
        """data = nonce(12) || ciphertext || tag(16) → plaintext bytes."""
        if self._backend == "cng":
            return _cng_aes_gcm_decrypt(
                self._bcrypt, self._cng_alg, self._cng_klen, key32, data
            )
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        return AESGCM(key32).decrypt(data[:12], data[12:], None)

    def encrypt(self, plaintext: str) -> str:
        raw = self.raw_encrypt(self._key, plaintext.encode("utf-8"))
        return base64.b64encode(raw).decode("ascii")

    def decrypt(self, encoded: str) -> str:
        raw = base64.b64decode(encoded.encode("ascii"))
        return self.raw_decrypt(self._key, raw).decode("utf-8")

    def close(self):
        try:
            if self._cng_alg is not None and self._bcrypt is not None:
                self._bcrypt.BCryptCloseAlgorithmProvider(self._cng_alg, 0)
                self._cng_alg = None
        except Exception:
            pass


# ── PBKDF2 via stdlib (works in compiled .exe — no packages) ─────────────────

def _pbkdf2_key(passphrase: str, salt: bytes, iterations: int = 600_000) -> bytes:
    return hashlib.pbkdf2_hmac(
        "sha256", passphrase.encode("utf-8"), salt, iterations, dklen=32
    )


# ── Master password helpers ───────────────────────────────────────────────────

def _master_path() -> Path:
    return Path(_api.plugin_dir) / _MASTER_FILE


def _master_load() -> Optional[dict]:
    p = _master_path()
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return None


def _master_save(data: dict) -> None:
    p   = _master_path()
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=2), encoding="utf-8")
    os.replace(tmp, p)


def _master_is_enabled() -> bool:
    d = _master_load()
    return bool(d and d.get("enabled"))


def _master_verify(password: str) -> bool:
    """Return True if password matches the stored master-password hash."""
    d = _master_load()
    if not d or not d.get("enabled"):
        return False
    salt     = bytes.fromhex(d["pwd_salt"])
    expected = bytes.fromhex(d["pwd_hash"])
    computed = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 200_000, 32)
    return secrets.compare_digest(computed, expected)


def _master_answers_key(a1: str, a2: str, a3: str, salt: bytes) -> bytes:
    """Derive a 32-byte AES key from the three case-insensitive security answers."""
    combined = (
        a1.strip().lower() + "|" +
        a2.strip().lower() + "|" +
        a3.strip().lower()
    )
    return hashlib.pbkdf2_hmac("sha256", combined.encode("utf-8"), salt, 100_000, 32)


def _master_recover(a1: str, a2: str, a3: str) -> Optional[str]:
    """Return the master password in plain text if all 3 answers are correct, else None."""
    d = _master_load()
    if not d or not d.get("enabled") or _vault is None:
        return None
    try:
        salt  = bytes.fromhex(d["recovery_salt"])
        key   = _master_answers_key(a1, a2, a3, salt)
        enc   = base64.b64decode(d["encrypted_master"].encode("ascii"))
        plain = _vault.raw_decrypt(key, enc)
        return plain.decode("utf-8")
    except Exception:
        return None


def _master_enable(password: str, a1: str, a2: str, a3: str) -> None:
    """Hash and store the master password plus a recovery-encrypted copy."""
    pwd_salt      = secrets.token_bytes(32)
    pwd_hash      = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), pwd_salt, 200_000, 32)
    recovery_salt = secrets.token_bytes(32)
    rec_key       = _master_answers_key(a1, a2, a3, recovery_salt)
    enc_bytes     = _vault.raw_encrypt(rec_key, password.encode("utf-8"))
    _master_save({
        "enabled":          True,
        "pwd_salt":         pwd_salt.hex(),
        "pwd_hash":         pwd_hash.hex(),
        "recovery_salt":    recovery_salt.hex(),
        "encrypted_master": base64.b64encode(enc_bytes).decode("ascii"),
    })


def _master_disable() -> None:
    p = _master_path()
    if p.exists():
        p.unlink()


# ── Plugin globals ────────────────────────────────────────────────────────────

_api   = None
_btn   = None
_win:                  Optional["CredentialManagerWindow"] = None
_vault:                Optional[_Vault]                    = None

_HOTKEY          = "Ctrl+Shift+K"
_TAB_LABEL       = "Credential Manager"
_DATA_FILE       = "credentials.json"
_MASTER_FILE     = "master.json"
_master_unlocked = False   # True after a successful unlock; reset when the vault window closes


# ── Lifecycle ─────────────────────────────────────────────────────────────────

def activate(api):
    global _api, _btn, _vault
    _api = api
    try:
        _vault = _Vault(api.plugin_dir, log_fn=api.log)
    except Exception as e:
        api.log(f"Vault init failed: {e}\n{traceback.format_exc()}")
        api.toast(f"Credential Manager vault error: {e}", "error")
        return

    _btn = api.ui.add_footer_button("🔐 Vault", _open_window)
    api.hotkeys.register(_HOTKEY, _open_window, "Open Credential Manager")
    api.ui.add_settings_tab(_TAB_LABEL, _build_settings_tab())
    api.settings.on_changed(_on_settings_changed)


def deactivate():
    global _btn, _win, _vault, _master_unlocked
    _master_unlocked = False
    try:
        if _btn is not None:
            _api.ui.remove_footer_button(_btn)
            _btn = None
    except Exception:
        pass
    try:
        _api.hotkeys.unregister(_HOTKEY)
    except Exception:
        pass
    try:
        _api.ui.remove_settings_tab(_TAB_LABEL)
    except Exception:
        pass
    try:
        if _win is not None:
            _win.close()
            _win = None
    except Exception:
        pass
    try:
        if _vault is not None:
            _vault.close()
            _vault = None
    except Exception:
        pass


# ── Open window ───────────────────────────────────────────────────────────────

def _open_window():
    global _win, _master_unlocked
    try:
        if _vault is None:
            _api.toast("Vault not initialised.", "error")
            return
        # If the window is already visible just bring it to front
        if _win is not None and _win.isVisible():
            _win.raise_()
            _win.activateWindow()
            return
        # Show lock screen if master password is enabled and not yet unlocked this session
        if _master_is_enabled() and not _master_unlocked:
            dlg = MasterLockDialog(_api.ui.main_window)
            if dlg.exec() != QDialog.Accepted:
                return
            _master_unlocked = True
        _win = CredentialManagerWindow(_api.ui.main_window)
        _win.show()
    except Exception as e:
        _api.log(f"_open_window error: {e}\n{traceback.format_exc()}")
        _api.toast("Error opening Credential Manager.", "error")


# ── Credential persistence ────────────────────────────────────────────────────

def _data_path() -> Path:
    return _api.plugin_dir / _DATA_FILE


def _load_credentials() -> list[dict]:
    try:
        p = _data_path()
        if not p.exists():
            return []
        data = json.loads(p.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except Exception as e:
        _api.log(f"_load_credentials error: {e}")
        return []


def _save_credentials(creds: list[dict]) -> None:
    """Atomic write via temp file + os.replace() — safe against crashes mid-write."""
    p   = _data_path()
    tmp = p.with_suffix(".tmp")
    try:
        tmp.write_text(
            json.dumps(creds, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )
        os.replace(tmp, p)
    except Exception:
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass
        raise


def _new_id() -> str:
    return secrets.token_hex(16)


# ── Password utilities ────────────────────────────────────────────────────────

# Common weak passwords and dictionary words to penalise heavily.
# Checked against the raw password AND a leet-substitution variant.
_COMMON_PASSWORDS: frozenset = frozenset([
    "password", "passw0rd", "qwerty", "letmein", "welcome", "monkey",
    "dragon", "master", "login", "admin", "user", "superman", "batman",
    "iloveyou", "sunshine", "princess", "football", "baseball", "shadow",
    "soccer", "michael", "jessica", "trustno1", "hello", "charlie",
    "mustang", "access", "hockey", "ranger", "daniel", "starwars",
    "112233", "121212", "696969", "abc123", "secret", "changeme",
    "test", "guest", "server", "default", "temp", "pass",
])

# Map common leet-speak substitutions so "p@ssw0rd" → "password"
_LEET_TABLE = str.maketrans({
    "@": "a", "0": "o", "3": "e", "1": "l",
    "$": "s", "5": "s", "4": "a", "!": "i", "+": "t", "8": "b",
})


def _password_strength(pwd: str) -> tuple[int, str]:
    if not pwd:
        return 0, ""

    lower  = pwd.lower()
    deobf  = lower.translate(_LEET_TABLE)
    length = len(pwd)

    if length < 8:
        return 0, "Very Weak"

    score = 0

    # Length is the dominant factor
    if   length >= 20: score += 3
    elif length >= 14: score += 2
    elif length >= 10: score += 1
    # 8-9 chars: +0

    # Character variety bonus
    variety = sum([
        bool(re.search(r"[A-Z]", pwd)),
        bool(re.search(r"[a-z]", pwd)),
        bool(re.search(r"\d",    pwd)),
        bool(re.search(r"[^A-Za-z0-9]", pwd)),
    ])
    if   variety >= 4: score += 2
    elif variety >= 3: score += 1

    # ── Penalties ──────────────────────────────────────────────────────────
    # 1. Dictionary / common password match (includes leet variants)
    for word in _COMMON_PASSWORDS:
        if word in lower or word in deobf:
            score -= 4
            break

    # 2. Predictable word-then-digits(-then-symbol) pattern
    #    e.g. "dragon99!", "Sunshine2024!", "mydog123"
    if re.search(r"[a-z]{4,}\d{1,6}[^a-z\d]{0,3}$", lower):
        score -= 2

    # 3. All digits
    if re.match(r"^\d+$", pwd):
        score -= 4

    # 4. Keyboard walks and alphabet runs (3+ chars)
    _SEQUENCES = [
        "1234567890", "0987654321",
        "abcdefghijklmnopqrstuvwxyz", "zyxwvutsrqponmlkjihgfedcba",
        "qwertyuiop", "poiuytrewq", "asdfghjkl", "lkjhgfdsa", "zxcvbnm",
    ]
    for seq in _SEQUENCES:
        for run in range(3, len(seq) + 1):
            hit = False
            for i in range(len(seq) - run + 1):
                if seq[i:i + run] in lower:
                    hit = True
                    break
            if hit:
                score -= 1
                break

    # 5. Repeated characters ("aaa", "111", …)
    if re.search(r"(.)\1{2,}", pwd):
        score -= 1

    score = max(0, min(score, 4))
    labels = {0: "Very Weak", 1: "Weak", 2: "Fair", 3: "Strong", 4: "Very Strong"}
    return score, labels[score]


def _strength_color(score: int, c: dict) -> str:
    return [c["accent_red"], c["accent_red"], c["accent_amber"],
            c["accent_teal"], c["accent_teal"]][score]


# ── Security questions (shared constant) ─────────────────────────────────────

_SECURITY_QUESTIONS = [
    "What city were you born in?",
    "What was the model of your first car?",
    "What is your mother's maiden name?",
]


# ── Master password dialogs ───────────────────────────────────────────────────

class MasterLockDialog(QDialog):
    """Lock screen shown before opening the vault when master password is enabled."""

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("\U0001f512 Vault Locked")
        self.setMinimumWidth(420)
        self.setModal(True)
        self._build_ui()
        _api.theme.register(self._apply_theme)
        self._apply_theme()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(32, 28, 32, 28)
        lay.setSpacing(14)

        icon_lbl = QLabel("\U0001f512")
        icon_lbl.setFont(QFont("Segoe UI", 32))
        icon_lbl.setAlignment(Qt.AlignCenter)
        lay.addWidget(icon_lbl)

        title = QLabel("Vault Locked")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignCenter)
        lay.addWidget(title)

        desc = QLabel("Enter your master password to access your credentials.")
        desc.setWordWrap(True)
        desc.setAlignment(Qt.AlignCenter)
        lay.addWidget(desc)
        self._desc_lbl = desc

        # Password field
        pwd_row = QHBoxLayout()
        self._pwd_edit = QLineEdit()
        self._pwd_edit.setEchoMode(QLineEdit.Password)
        self._pwd_edit.setPlaceholderText("Master password\u2026")
        self._pwd_edit.setFixedHeight(36)
        self._pwd_edit.returnPressed.connect(self._try_unlock)
        show_btn = QPushButton("\U0001f441")
        show_btn.setCheckable(True)
        show_btn.setFixedSize(36, 36)
        show_btn.setToolTip("Show / hide password")
        show_btn.toggled.connect(
            lambda on: self._pwd_edit.setEchoMode(
                QLineEdit.Normal if on else QLineEdit.Password
            )
        )
        pwd_row.addWidget(self._pwd_edit, 1)
        pwd_row.addWidget(show_btn)
        lay.addLayout(pwd_row)

        # Error label
        self._err_lbl = QLabel("")
        self._err_lbl.setAlignment(Qt.AlignCenter)
        self._err_lbl.setVisible(False)
        lay.addWidget(self._err_lbl)

        self._unlock_btn = QPushButton("Unlock")
        self._unlock_btn.setFixedHeight(36)
        self._unlock_btn.setDefault(True)
        self._unlock_btn.clicked.connect(self._try_unlock)
        lay.addWidget(self._unlock_btn)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setFixedHeight(32)
        cancel_btn.clicked.connect(self.reject)
        lay.addWidget(cancel_btn)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        lay.addWidget(sep)
        self._sep = sep

        forgot_btn = QPushButton("Forgot your master password?")
        forgot_btn.setFlat(True)
        forgot_btn.setCursor(Qt.PointingHandCursor)
        forgot_btn.clicked.connect(self._open_recovery)
        lay.addWidget(forgot_btn, 0, Qt.AlignCenter)
        self._forgot_btn = forgot_btn

    def _try_unlock(self):
        pwd = self._pwd_edit.text()
        if not pwd:
            self._show_err("Please enter your master password.")
            return
        if _master_verify(pwd):
            self.accept()
        else:
            self._show_err("Incorrect password.  Please try again.")
            self._pwd_edit.clear()
            self._pwd_edit.setFocus()

    def _show_err(self, msg: str):
        self._err_lbl.setText(msg)
        self._err_lbl.setVisible(True)

    def _open_recovery(self):
        dlg = MasterRecoveryDialog(self)
        if dlg.exec() == QDialog.Accepted:
            # Vault was purged — allow opening (empty) vault
            self.accept()

    def _apply_theme(self):
        try:
            c = _api.theme.colors()
            self.setStyleSheet(
                f"QDialog{{background:{c['bg_dark']};color:{c['text_primary']}}}"
                f"QLabel{{color:{c['text_secondary']};background:transparent}}"
                f"QLineEdit{{background:{c['bg_mid']};color:{c['text_primary']};"
                f"  border:1px solid {c['glow']};border-radius:4px;padding:4px 8px}}"
                f"QLineEdit:focus{{border-color:{c['accent_blue']}}}"
                f"QPushButton{{background:{c['bg_mid']};color:{c['text_primary']};"
                f"  border:1px solid {c['glow']};border-radius:4px;padding:4px 10px}}"
                f"QPushButton:hover{{border-color:{c['accent_blue']}}}"
                f"QPushButton:pressed{{background:{c['glow']}}}"
            )
            self._unlock_btn.setStyleSheet(
                f"QPushButton{{background:transparent;color:{c['accent_blue']};"
                f"  border:2px solid {c['accent_blue']};border-radius:4px;padding:4px 18px;font-weight:600}}"
                f"QPushButton:hover{{background:{c['accent_blue']}22;border-color:{c['accent_teal']};color:{c['accent_teal']}}}"
                f"QPushButton:pressed{{background:{c['accent_blue']}44}}"
            )
            self._err_lbl.setStyleSheet(f"color:{c['accent_red']};font-weight:600")
            self._desc_lbl.setStyleSheet(f"color:{c['text_secondary']}")
            self._sep.setStyleSheet(f"background:{c['glow']}")
            self._forgot_btn.setStyleSheet(
                f"QPushButton{{background:transparent;color:{c['accent_blue']};"
                f"  border:none;text-decoration:underline}}"
                f"QPushButton:hover{{color:{c['accent_teal']}}}"
            )
        except Exception:
            pass


class MasterSetupDialog(QDialog):
    """Dialog for enabling or changing the master password."""

    def __init__(self, parent, changing: bool = False):
        super().__init__(parent)
        self._changing = changing
        self.setWindowTitle("Change Master Password" if changing else "Set Master Password")
        self.setMinimumWidth(520)
        self.setModal(True)
        self._build_ui()
        _api.theme.register(self._apply_theme)
        self._apply_theme()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(12)

        title = QLabel(
            "Change Master Password" if self._changing else "Set Master Password"
        )
        title.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        lay.addWidget(title)

        # Warning banner
        warn = QLabel(
            "\u26a0\ufe0f  Your master password gates access to the vault every time you open it.\n"
            "   If you forget it you must answer all 3 security questions below to recover it.\n"
            "   There is NO other way to access your credentials if both are forgotten."
        )
        warn.setWordWrap(True)
        warn.setContentsMargins(10, 8, 10, 8)
        lay.addWidget(warn)
        self._warn_lbl = warn

        sep1 = QFrame(); sep1.setFrameShape(QFrame.HLine)
        lay.addWidget(sep1)
        self._sep1 = sep1

        # Current password row (only when changing)
        if self._changing:
            lay.addWidget(QLabel("Current master password:"))
            self._cur_edit = QLineEdit()
            self._cur_edit.setEchoMode(QLineEdit.Password)
            self._cur_edit.setPlaceholderText("Current password\u2026")
            self._cur_edit.setFixedHeight(34)
            lay.addWidget(self._cur_edit)

        lay.addWidget(QLabel("New master password:"))
        pwd_row = QHBoxLayout()
        self._new_edit = QLineEdit()
        self._new_edit.setEchoMode(QLineEdit.Password)
        self._new_edit.setPlaceholderText("New password\u2026")
        self._new_edit.setFixedHeight(34)
        self._new_edit.textChanged.connect(self._update_strength)
        show_btn = QPushButton("\U0001f441")
        show_btn.setCheckable(True)
        show_btn.setFixedSize(34, 34)
        show_btn.toggled.connect(
            lambda on: self._new_edit.setEchoMode(
                QLineEdit.Normal if on else QLineEdit.Password
            )
        )
        pwd_row.addWidget(self._new_edit, 1)
        pwd_row.addWidget(show_btn)
        lay.addLayout(pwd_row)

        self._strength_lbl = QLabel("")
        self._strength_lbl.setFixedHeight(16)
        lay.addWidget(self._strength_lbl)

        lay.addWidget(QLabel("Confirm new password:"))
        self._conf_edit = QLineEdit()
        self._conf_edit.setEchoMode(QLineEdit.Password)
        self._conf_edit.setPlaceholderText("Confirm password\u2026")
        self._conf_edit.setFixedHeight(34)
        lay.addWidget(self._conf_edit)

        sep2 = QFrame(); sep2.setFrameShape(QFrame.HLine)
        lay.addWidget(sep2)
        self._sep2 = sep2

        sq_title = QLabel("Security Questions  (required for password recovery)")
        sq_title.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        lay.addWidget(sq_title)

        sq_note = QLabel(
            "Answers are case-insensitive.  All 3 must be correct to recover your master password."
        )
        sq_note.setWordWrap(True)
        lay.addWidget(sq_note)
        self._sq_note = sq_note

        self._answer_edits: list[QLineEdit] = []
        for q in _SECURITY_QUESTIONS:
            lay.addWidget(QLabel(q))
            ans = QLineEdit()
            ans.setPlaceholderText("Your answer\u2026")
            ans.setFixedHeight(34)
            lay.addWidget(ans)
            self._answer_edits.append(ans)

        self._err_lbl = QLabel("")
        self._err_lbl.setVisible(False)
        lay.addWidget(self._err_lbl)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setFixedHeight(34)
        cancel_btn.clicked.connect(self.reject)
        self._save_btn = QPushButton("Save")
        self._save_btn.setFixedHeight(34)
        self._save_btn.setDefault(True)
        self._save_btn.clicked.connect(self._save)
        btn_row.addWidget(cancel_btn)
        btn_row.addWidget(self._save_btn)
        lay.addLayout(btn_row)

    def _update_strength(self, text: str):
        score, label = _password_strength(text)
        try:
            c = _api.theme.colors()
            color = _strength_color(score, c)
        except Exception:
            color = "#ffffff"
        self._strength_lbl.setText(f"Strength: {label}")
        self._strength_lbl.setStyleSheet(f"color:{color};font-weight:600")

    def _save(self):
        if self._changing and not _master_verify(self._cur_edit.text()):
            self._show_err("Current password is incorrect.")
            return
        new_pwd  = self._new_edit.text()
        conf_pwd = self._conf_edit.text()
        if len(new_pwd) < 6:
            self._show_err("Password must be at least 6 characters.")
            return
        if new_pwd != conf_pwd:
            self._show_err("Passwords do not match.")
            return
        answers = [e.text().strip() for e in self._answer_edits]
        if any(not a for a in answers):
            self._show_err("All 3 security questions must be answered.")
            return
        _master_enable(new_pwd, answers[0], answers[1], answers[2])
        self.accept()

    def _show_err(self, msg: str):
        self._err_lbl.setText(msg)
        self._err_lbl.setVisible(True)

    def _apply_theme(self):
        try:
            c = _api.theme.colors()
            self.setStyleSheet(
                f"QDialog{{background:{c['bg_dark']};color:{c['text_primary']}}}"
                f"QLabel{{color:{c['text_secondary']};background:transparent}}"
                f"QLineEdit{{background:{c['bg_mid']};color:{c['text_primary']};"
                f"  border:1px solid {c['glow']};border-radius:4px;padding:4px 8px}}"
                f"QLineEdit:focus{{border-color:{c['accent_blue']}}}"
                f"QPushButton{{background:{c['bg_mid']};color:{c['text_primary']};"
                f"  border:1px solid {c['glow']};border-radius:4px;padding:4px 10px}}"
                f"QPushButton:hover{{border-color:{c['accent_blue']}}}"
                f"QPushButton:pressed{{background:{c['glow']}}}"
            )
            self._save_btn.setStyleSheet(
                f"QPushButton{{background:transparent;color:{c['accent_blue']};"
                f"  border:2px solid {c['accent_blue']};border-radius:4px;padding:4px 18px;font-weight:600}}"
                f"QPushButton:hover{{background:{c['accent_blue']}22;border-color:{c['accent_teal']};color:{c['accent_teal']}}}"
                f"QPushButton:pressed{{background:{c['accent_blue']}44}}"
            )
            self._warn_lbl.setStyleSheet(
                f"background:{c['accent_amber']}22;color:{c['accent_amber']};"
                f"border:1px solid {c['accent_amber']}55;border-radius:4px"
            )
            self._err_lbl.setStyleSheet(f"color:{c['accent_red']};font-weight:600")
            self._sq_note.setStyleSheet(f"color:{c['text_dim']}")
            for sep in (self._sep1, self._sep2):
                sep.setStyleSheet(f"background:{c['glow']}")
        except Exception:
            pass


class MasterRecoveryDialog(QDialog):
    """Recover a forgotten master password by answering the 3 security questions."""

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("\U0001f511 Recover Master Password")
        self.setMinimumWidth(520)
        self.setModal(True)
        self._build_ui()
        _api.theme.register(self._apply_theme)
        self._apply_theme()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(12)

        title = QLabel("\U0001f511  Recover Master Password")
        title.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        lay.addWidget(title)

        info = QLabel(
            "Answer all 3 security questions exactly as you entered them (case-insensitive).\n"
            "If all answers are correct, your master password will be revealed in plain text.\n\n"
            "\u26a0\ufe0f  If you cannot answer the questions, you may purge your vault below.\n"
            "   Purging is permanent and irreversible \u2014 ALL credentials will be deleted."
        )
        info.setWordWrap(True)
        lay.addWidget(info)
        self._info_lbl = info

        sep1 = QFrame(); sep1.setFrameShape(QFrame.HLine)
        lay.addWidget(sep1)
        self._sep1 = sep1

        self._answer_edits: list[QLineEdit] = []
        for q in _SECURITY_QUESTIONS:
            lay.addWidget(QLabel(q))
            ans = QLineEdit()
            ans.setPlaceholderText("Your answer\u2026")
            ans.setFixedHeight(34)
            lay.addWidget(ans)
            self._answer_edits.append(ans)

        self._verify_btn = QPushButton("Verify Answers")
        self._verify_btn.setFixedHeight(36)
        self._verify_btn.setDefault(True)
        self._verify_btn.clicked.connect(self._verify)
        lay.addWidget(self._verify_btn)

        # ── Result area (hidden until Verify is clicked) ──────────────────
        self._sep2 = QFrame(); self._sep2.setFrameShape(QFrame.HLine)
        self._sep2.setVisible(False)
        lay.addWidget(self._sep2)

        self._result_lbl = QLabel("")
        self._result_lbl.setWordWrap(True)
        self._result_lbl.setVisible(False)
        lay.addWidget(self._result_lbl)

        # Revealed password row (success)
        self._reveal_widget = QWidget()
        rr = QHBoxLayout(self._reveal_widget)
        rr.setContentsMargins(0, 0, 0, 0)
        rr.setSpacing(6)
        rr.addWidget(QLabel("Your master password:"))
        self._revealed_pwd = QLineEdit()
        self._revealed_pwd.setReadOnly(True)
        self._revealed_pwd.setEchoMode(QLineEdit.Password)
        self._revealed_pwd.setFixedHeight(34)
        self._reveal_show_btn = QPushButton("\U0001f441 Show")
        self._reveal_show_btn.setCheckable(True)
        self._reveal_show_btn.setFixedHeight(34)
        self._reveal_show_btn.toggled.connect(self._toggle_reveal)
        self._copy_pwd_btn = QPushButton("Copy")
        self._copy_pwd_btn.setFixedHeight(34)
        self._copy_pwd_btn.clicked.connect(self._copy_pwd)
        rr.addWidget(self._revealed_pwd, 1)
        rr.addWidget(self._reveal_show_btn)
        rr.addWidget(self._copy_pwd_btn)
        self._reveal_widget.setVisible(False)
        lay.addWidget(self._reveal_widget)

        # Failure action buttons
        self._fail_widget = QWidget()
        fr = QHBoxLayout(self._fail_widget)
        fr.setContentsMargins(0, 0, 0, 0)
        fr.setSpacing(8)
        retry_btn = QPushButton("Try Again")
        retry_btn.setFixedHeight(34)
        retry_btn.clicked.connect(self._reset)
        self._purge_btn = QPushButton("\U0001f5d1 Purge Vault")
        self._purge_btn.setFixedHeight(34)
        self._purge_btn.setToolTip("Permanently delete ALL credentials and remove master password")
        self._purge_btn.clicked.connect(self._purge)
        exit_btn = QPushButton("Exit")
        exit_btn.setFixedHeight(34)
        exit_btn.clicked.connect(self.reject)
        fr.addWidget(retry_btn)
        fr.addWidget(self._purge_btn)
        fr.addStretch()
        fr.addWidget(exit_btn)
        self._fail_widget.setVisible(False)
        lay.addWidget(self._fail_widget)

        # Close button (success path)
        self._close_btn = QPushButton("Close")
        self._close_btn.setFixedHeight(34)
        self._close_btn.setVisible(False)
        self._close_btn.clicked.connect(self.reject)
        lay.addWidget(self._close_btn)

    def _toggle_reveal(self, on: bool):
        self._revealed_pwd.setEchoMode(QLineEdit.Normal if on else QLineEdit.Password)
        self._reveal_show_btn.setText("\U0001f648 Hide" if on else "\U0001f441 Show")

    def _copy_pwd(self):
        try:
            _api.clipboard.set_text(self._revealed_pwd.text())
            self._copy_pwd_btn.setText("\u2713 Copied!")
            QTimer.singleShot(1500, lambda: self._copy_pwd_btn.setText("Copy"))
        except Exception:
            pass

    def _verify(self):
        answers = [e.text() for e in self._answer_edits]
        recovered = _master_recover(answers[0], answers[1], answers[2])
        self._sep2.setVisible(True)
        self._result_lbl.setVisible(True)
        try:
            c = _api.theme.colors()
        except Exception:
            c = {}
        if recovered is not None:
            self._result_lbl.setText("\u2714  Answers correct!  Your master password is shown below.")
            self._result_lbl.setStyleSheet(f"color:{c.get('accent_teal','#00c896')};font-weight:600")
            self._revealed_pwd.setText(recovered)
            self._reveal_widget.setVisible(True)
            self._close_btn.setVisible(True)
            self._verify_btn.setEnabled(False)
            for e in self._answer_edits:
                e.setReadOnly(True)
        else:
            self._result_lbl.setText(
                "\u2718  Incorrect answers.  Check your responses and try again."
            )
            self._result_lbl.setStyleSheet(f"color:{c.get('accent_red','#e05260')};font-weight:600")
            self._fail_widget.setVisible(True)

    def _reset(self):
        for e in self._answer_edits:
            e.clear()
            e.setReadOnly(False)
        self._sep2.setVisible(False)
        self._result_lbl.setVisible(False)
        self._fail_widget.setVisible(False)
        self._verify_btn.setEnabled(True)

    def _purge(self):
        reply = QMessageBox.warning(
            self,
            "Purge Vault \u2014 Point of No Return",
            "This will permanently delete ALL saved credentials and remove the master password.\n\n"
            "This action CANNOT be undone.\n\nAre you absolutely sure?",
            QMessageBox.Yes | QMessageBox.Cancel,
            QMessageBox.Cancel,
        )
        if reply == QMessageBox.Yes:
            try:
                cred_path = Path(_api.plugin_dir) / _DATA_FILE
                if cred_path.exists():
                    cred_path.unlink()
                _master_disable()
                global _master_unlocked
                _master_unlocked = True  # nothing left to lock
            except Exception as e:
                _api.log(f"Vault purge error: {e}")
            try:
                _api.toast("Vault purged. All credentials deleted.", "warning")
            except Exception:
                pass
            self.accept()

    def _apply_theme(self):
        try:
            c = _api.theme.colors()
            self.setStyleSheet(
                f"QDialog{{background:{c['bg_dark']};color:{c['text_primary']}}}"
                f"QLabel{{color:{c['text_secondary']};background:transparent}}"
                f"QLineEdit{{background:{c['bg_mid']};color:{c['text_primary']};"
                f"  border:1px solid {c['glow']};border-radius:4px;padding:4px 8px}}"
                f"QLineEdit:focus{{border-color:{c['accent_blue']}}}"
                f"QPushButton{{background:{c['bg_mid']};color:{c['text_primary']};"
                f"  border:1px solid {c['glow']};border-radius:4px;padding:4px 10px}}"
                f"QPushButton:hover{{border-color:{c['accent_blue']}}}"
                f"QPushButton:pressed{{background:{c['glow']}}}"
            )
            self._verify_btn.setStyleSheet(
                f"QPushButton{{background:transparent;color:{c['accent_blue']};"
                f"  border:2px solid {c['accent_blue']};border-radius:4px;padding:4px 18px;font-weight:600}}"
                f"QPushButton:hover{{background:{c['accent_blue']}22;border-color:{c['accent_teal']};color:{c['accent_teal']}}}"
                f"QPushButton:pressed{{background:{c['accent_blue']}44}}"
                f"QPushButton:disabled{{background:transparent;color:{c['text_dim']};border-color:{c['glow']}}}"
            )
            self._purge_btn.setStyleSheet(
                f"QPushButton{{background:transparent;color:{c['accent_red']};"
                f"  border:1px solid {c['accent_red']};border-radius:4px;padding:4px 10px}}"
                f"QPushButton:hover{{background:{c['accent_red']}22}}"
            )
            self._info_lbl.setStyleSheet(f"color:{c['text_secondary']}")
            for sep in (self._sep1, self._sep2):
                sep.setStyleSheet(f"background:{c['glow']}")
        except Exception:
            pass


# ── Settings tab ──────────────────────────────────────────────────────────────

def _build_settings_tab() -> QWidget:
    w   = QWidget()
    lay = QVBoxLayout(w)
    lay.setContentsMargins(16, 16, 16, 16)
    lay.setSpacing(12)

    form = QFormLayout()
    form.setSpacing(10)

    spin_exp = QSpinBox()
    spin_exp.setRange(0, 3650)
    spin_exp.setValue(int(_api.settings.value("default_expiry_days", 90)))
    spin_exp.setSuffix("  days  (0 = no default)")
    spin_exp.setFixedWidth(210)
    spin_exp.valueChanged.connect(lambda v: _api.settings.set("default_expiry_days", v))
    form.addRow("Default expiry (new entries):", spin_exp)

    backend_lbl = QLabel("")
    if _vault:
        backend_lbl.setText(f"Encryption: {_vault._backend.upper()}  —  AES-256-GCM")
    form.addRow("Backend:", backend_lbl)

    hotkey_lbl = QLabel(_HOTKEY)
    form.addRow("Open vault hotkey:", hotkey_lbl)

    lay.addLayout(form)

    # ── Master Password section ───────────────────────────────────────────────
    mp_sep = QFrame()
    mp_sep.setFrameShape(QFrame.HLine)
    lay.addWidget(mp_sep)

    mp_title = QLabel("Master Password")
    mp_title.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
    lay.addWidget(mp_title)

    mp_info = QLabel(
        "Optionally lock your vault behind a master password.  When enabled, you must "
        "enter it each time you open the vault.  If you forget it, answer the 3 security "
        "questions you set during setup to recover it in plain text."
    )
    mp_info.setWordWrap(True)
    lay.addWidget(mp_info)

    mp_status_lbl = QLabel()
    mp_status_lbl.setFont(QFont("Segoe UI", 9))
    lay.addWidget(mp_status_lbl)

    mp_btn_row = QHBoxLayout()
    mp_btn_row.setSpacing(8)
    lay.addLayout(mp_btn_row)

    def _refresh_mp_ui():
        while mp_btn_row.count():
            item = mp_btn_row.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        if _master_is_enabled():
            mp_status_lbl.setText("\U0001f512  Master password protection is ENABLED")
            chg_btn = QPushButton("Change Password")
            chg_btn.setFixedHeight(30)
            chg_btn.clicked.connect(_on_change_master)
            dis_btn = QPushButton("Disable")
            dis_btn.setFixedHeight(30)
            dis_btn.clicked.connect(_on_disable_master)
            mp_btn_row.addWidget(chg_btn)
            mp_btn_row.addWidget(dis_btn)
            mp_btn_row.addStretch()
        else:
            mp_status_lbl.setText("\u26aa  Not enabled  \u2014  enabling is strongly recommended")
            enb_btn = QPushButton("\U0001f512 Enable Master Password")
            enb_btn.setFixedHeight(30)
            enb_btn.clicked.connect(_on_enable_master)
            mp_btn_row.addWidget(enb_btn)
            mp_btn_row.addStretch()

    def _on_enable_master():
        dlg = MasterSetupDialog(w)
        if dlg.exec() == QDialog.Accepted:
            _api.toast("Master password enabled.  Your vault is now protected.", "info")
            _refresh_mp_ui()

    def _on_change_master():
        dlg = MasterSetupDialog(w, changing=True)
        if dlg.exec() == QDialog.Accepted:
            _api.toast("Master password updated.", "info")
            _refresh_mp_ui()

    def _on_disable_master():
        pwd, ok = QInputDialog.getText(
            w, "Disable Master Password",
            "Enter your current master password to disable vault protection:",
            QLineEdit.Password,
        )
        if not ok:
            return
        if _master_verify(pwd):
            _master_disable()
            _api.toast("Master password disabled.", "info")
            _refresh_mp_ui()
        else:
            _api.toast("Incorrect password \u2014 master protection remains enabled.", "error")

    _refresh_mp_ui()
    lay.addStretch()

    _api.theme.register(lambda: _style_settings_widget(w))
    _style_settings_widget(w)
    return w


def _style_settings_widget(w: QWidget):
    try:
        c = _api.theme.colors()
        w.setStyleSheet(
            f"QWidget{{background:{c['bg_dark']};color:{c['text_primary']}}}"
            f"QSpinBox{{background:{c['bg_mid']};color:{c['text_primary']};"
            f"  border:1px solid {c['glow']};border-radius:4px;padding:3px 6px}}"
            f"QLabel{{color:{c['text_secondary']}}}"
        )
    except Exception:
        pass


def _on_settings_changed():
    pass


# ── Credential edit dialog ────────────────────────────────────────────────────

class CredentialDialog(QDialog):

    def __init__(self, parent, existing: Optional[dict] = None):
        super().__init__(parent)
        self._existing    = existing
        self._result_data: Optional[dict] = None
        self.setWindowTitle("Edit Credential" if existing else "New Credential")
        self.setMinimumSize(660, 590)
        self.setWindowModality(Qt.WindowModal)
        self._build_ui()
        _api.theme.register(self._apply_theme)
        self._apply_theme()
        if existing:
            self._populate(existing)
        else:
            self._title_edit.setFocus()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(10)

        hdr = QLabel("Edit Credential" if self._existing else "New Credential")
        hdr.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        lay.addWidget(hdr)

        # Thin accent separator
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setFixedHeight(1)
        lay.addWidget(sep)
        self._header_sep = sep

        form = QFormLayout()
        form.setSpacing(9)
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        form.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)

        self._title_edit    = QLineEdit()
        self._title_edit.setPlaceholderText("e.g. Gmail, VPN, Server Login…")
        self._user_edit     = QLineEdit()
        self._user_edit.setPlaceholderText("Username or email (optional)")
        self._category_edit = QLineEdit()
        self._category_edit.setPlaceholderText("e.g. Work, Personal, Banking… (optional)")

        # Category autocomplete from existing categories
        try:
            cats = sorted(set(
                c.get("category", "")
                for c in _load_credentials()
                if c.get("category")
            ))
            if cats:
                comp = QCompleter(cats, self._category_edit)
                comp.setCaseSensitivity(Qt.CaseInsensitive)
                comp.setCompletionMode(QCompleter.PopupCompletion)
                self._category_edit.setCompleter(comp)
        except Exception:
            pass

        form.addRow("Title:", self._title_edit)
        form.addRow("Username:", self._user_edit)
        form.addRow("Category:", self._category_edit)

        # Password area: two sub-rows so the generator controls never squish
        self._pwd_edit = QLineEdit()
        self._pwd_edit.setPlaceholderText("Password (required)")
        self._pwd_edit.setEchoMode(QLineEdit.Password)
        self._pwd_edit.textChanged.connect(self._update_strength)

        self._show_btn = QPushButton("👁")
        self._show_btn.setFixedWidth(34)
        self._show_btn.setCheckable(True)
        self._show_btn.setToolTip("Toggle password visibility")
        self._show_btn.toggled.connect(
            lambda on: self._pwd_edit.setEchoMode(
                QLineEdit.Normal if on else QLineEdit.Password
            )
        )

        self._pwd_len = QSlider(Qt.Horizontal)
        self._pwd_len.setRange(8, 64)
        self._pwd_len.setValue(20)
        self._pwd_len.setMinimumWidth(120)
        self._pwd_len.setToolTip("Generated password length (characters)")
        self._len_val_lbl = QLabel("20")
        self._len_val_lbl.setFixedWidth(30)
        self._len_val_lbl.setAlignment(Qt.AlignCenter)
        self._len_val_lbl.setFont(QFont("Segoe UI", 9))
        self._pwd_len.valueChanged.connect(
            lambda v: self._len_val_lbl.setText(str(v))
        )

        self._gen_btn = QPushButton("⚡ Generate")
        self._gen_btn.setMinimumWidth(110)
        self._gen_btn.setToolTip("Generate a strong random password")
        self._gen_btn.clicked.connect(self._generate_password)

        _pwd_container = QWidget()
        _pwd_vlay = QVBoxLayout(_pwd_container)
        _pwd_vlay.setContentsMargins(0, 0, 0, 0)
        _pwd_vlay.setSpacing(4)

        _pwd_top = QHBoxLayout()
        _pwd_top.setSpacing(6)
        _pwd_top.addWidget(self._pwd_edit, 1)
        _pwd_top.addWidget(self._show_btn)

        _pwd_bot = QHBoxLayout()
        _pwd_bot.setSpacing(8)
        _pwd_bot.addStretch()
        _len_lbl = QLabel("Length:")
        _len_lbl.setFont(QFont("Segoe UI", 8))
        _pwd_bot.addWidget(_len_lbl)
        _pwd_bot.addWidget(self._pwd_len)
        _pwd_bot.addWidget(self._len_val_lbl)
        _pwd_bot.addWidget(self._gen_btn)

        _pwd_vlay.addLayout(_pwd_top)
        _pwd_vlay.addLayout(_pwd_bot)
        form.addRow("Password *:", _pwd_container)

        self._strength_lbl = QLabel("")
        self._strength_lbl.setFixedHeight(18)
        form.addRow("Strength:", self._strength_lbl)

        # Expiry date
        exp_row = QHBoxLayout()
        self._exp_date = QDateEdit()
        self._exp_date.setCalendarPopup(True)
        self._exp_date.setDisplayFormat("yyyy-MM-dd")
        default_days = int(_api.settings.value("default_expiry_days", 90))
        self._exp_date.setDate(
            QDate.currentDate().addDays(default_days if default_days > 0 else 90)
        )
        self._no_exp_chk = QCheckBox("No expiry")
        self._no_exp_chk.setChecked(default_days == 0)
        self._no_exp_chk.toggled.connect(lambda on: self._exp_date.setEnabled(not on))
        self._exp_date.setEnabled(not self._no_exp_chk.isChecked())
        exp_row.addWidget(self._exp_date)
        exp_row.addWidget(self._no_exp_chk)
        exp_row.addStretch()
        form.addRow("Expires:", exp_row)

        self._note_edit = QTextEdit()
        self._note_edit.setPlaceholderText("Optional notes…")
        self._note_edit.setFixedHeight(78)
        form.addRow("Note:", self._note_edit)

        lay.addLayout(form)
        lay.addSpacing(6)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self._cancel_btn = QPushButton("Cancel")
        self._cancel_btn.setFixedHeight(34)
        self._save_btn   = QPushButton("💾  Save")
        self._save_btn.setFixedHeight(34)
        self._save_btn.setDefault(True)
        self._save_btn.setAutoDefault(True)
        self._cancel_btn.clicked.connect(self.reject)
        self._save_btn.clicked.connect(self._on_save)
        btn_row.addWidget(self._cancel_btn)
        btn_row.addWidget(self._save_btn)
        lay.addLayout(btn_row)

        # Tab order: title → user → category → password → note → save
        self.setTabOrder(self._title_edit,    self._user_edit)
        self.setTabOrder(self._user_edit,      self._category_edit)
        self.setTabOrder(self._category_edit,  self._pwd_edit)
        self.setTabOrder(self._pwd_edit,   self._note_edit)
        self.setTabOrder(self._note_edit,  self._save_btn)

    def _populate(self, d: dict):
        self._title_edit.setText(d.get("title", ""))
        self._user_edit.setText(d.get("username", ""))
        self._category_edit.setText(d.get("category", ""))
        self._note_edit.setPlainText(d.get("note", ""))
        try:
            self._pwd_edit.setText(_vault.decrypt(d["password_enc"]))
        except Exception as e:
            _api.log(f"Decrypt error on populate: {e}")
            _api.toast("Could not decrypt password for editing — check vault key.", "error")
        exp = d.get("expires")
        if exp:
            self._no_exp_chk.setChecked(False)
            self._exp_date.setEnabled(True)
            qd = QDate.fromString(exp, "yyyy-MM-dd")
            if qd.isValid():
                self._exp_date.setDate(qd)
        else:
            self._no_exp_chk.setChecked(True)
            self._exp_date.setEnabled(False)

    def _update_strength(self, pwd: str):
        score, label = _password_strength(pwd)
        if not label:
            self._strength_lbl.setText("")
            return
        try:
            c   = _api.theme.colors()
            col = _strength_color(score, c)
        except Exception:
            col = "#aaaaaa"
        bar = "█" * (score + 1) + "░" * (4 - score)
        self._strength_lbl.setText(
            f'<span style="color:{col};font-family:Consolas;letter-spacing:1px">{bar}</span>'
            f'&nbsp;&nbsp;<span style="color:{col};font-weight:600">{label}</span>'
            f'&nbsp;<span style="color:#888888;font-size:9px">({len(pwd)} chars)</span>'
        )

    def _generate_password(self):
        try:
            length   = self._pwd_len.value()
            alphabet = string.ascii_letters + string.digits + "!@#$%^&*()-_=+[]{}|;:,.<>?"
            # Guarantee at least one of each required class
            pwd = (
                secrets.choice(string.ascii_uppercase)
                + secrets.choice(string.ascii_lowercase)
                + secrets.choice(string.digits)
                + secrets.choice("!@#$%^&*()-_=+[]{}|;:,.<>?")
                + "".join(secrets.choice(alphabet) for _ in range(max(0, length - 4)))
            )
            # Shuffle so the guaranteed chars aren't always at the start
            pwd_list = list(pwd)
            secrets.SystemRandom().shuffle(pwd_list)
            pwd = "".join(pwd_list)

            self._pwd_edit.setText(pwd)
            self._pwd_edit.setEchoMode(QLineEdit.Normal)
            self._show_btn.setChecked(True)
            _api.clipboard.set_text(pwd)
            _api.toast(f"Password generated ({length} chars) and copied to clipboard.", "success")
        except Exception as e:
            _api.log(f"Generate error: {e}")
            _api.toast("Failed to generate password.", "error")

    def _on_save(self):
        pwd = self._pwd_edit.text()
        if not pwd:
            _api.ui.show_message("Validation", "Password is required.", "warn")
            self._pwd_edit.setFocus()
            return
        try:
            enc = _vault.encrypt(pwd)
        except Exception as e:
            _api.log(f"Encrypt error: {e}")
            _api.ui.show_message("Error", f"Encryption failed:\n{e}", "error")
            return

        now     = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
        expires = (
            None if self._no_exp_chk.isChecked()
            else self._exp_date.date().toString("yyyy-MM-dd")
        )
        self._result_data = {
            "id":           self._existing["id"] if self._existing else _new_id(),
            "title":        self._title_edit.text().strip(),
            "username":     self._user_edit.text().strip(),
            "category":     self._category_edit.text().strip(),
            "password_enc": enc,
            "expires":      expires,
            "note":         self._note_edit.toPlainText().strip(),
            "created_at":   self._existing.get("created_at", now) if self._existing else now,
            "updated_at":   now,
        }
        self.accept()

    def get_data(self) -> Optional[dict]:
        return self._result_data

    def _apply_theme(self):
        try:
            c = _api.theme.colors()
            self.setStyleSheet(
                f"QDialog{{background:{c['bg_dark']};color:{c['text_primary']}}}"
                f"QLabel{{color:{c['text_secondary']}}}"
                f"QLineEdit,QTextEdit{{background:{c['bg_mid']};color:{c['text_primary']};"
                f"  border:1px solid {c['glow']};border-radius:4px;padding:4px 6px}}"
                f"QLineEdit:focus,QTextEdit:focus{{border-color:{c['accent_blue']};"
                f"  background:{c['bg_mid']}}}"
                f"QPushButton{{background:{c['bg_mid']};color:{c['text_primary']};"
                f"  border:1px solid {c['glow']};border-radius:4px;padding:4px 10px}}"
                f"QPushButton:hover{{border-color:{c['accent_blue']};background:{c['bg_dark']}}}"
                f"QPushButton:pressed{{background:{c['glow']}}}"
                f"QSlider::groove:horizontal{{height:4px;background:{c['bg_mid']};"
                f"  border-radius:2px;margin:0 2px}}"
                f"QSlider::sub-page:horizontal{{background:{c['accent_blue']};"
                f"  border-radius:2px}}"
                f"QSlider::handle:horizontal{{background:{c['accent_blue']};"
                f"  border:none;width:14px;height:14px;margin:-5px -1px;"
                f"  border-radius:7px}}"
                f"QSlider::handle:horizontal:hover{{background:{c['accent_teal']}}}"
                f"QDateEdit{{background:{c['bg_mid']};color:{c['text_primary']};"
                f"  border:1px solid {c['glow']};border-radius:4px;padding:3px 6px}}"
                f"QCheckBox{{color:{c['text_secondary']}}}"
                f"QCalendarWidget QWidget{{background:{c['bg_dark']};color:{c['text_primary']}}}"
                f"QCompleter QAbstractItemView{{background:{c['bg_mid']};"
                f"  color:{c['text_primary']};selection-background-color:{c['accent_blue']}}}"
            )
            self._header_sep.setStyleSheet(f"background:{c['glow']}")
            self._save_btn.setStyleSheet(
                f"QPushButton{{background:transparent;color:{c['accent_blue']};"
                f"  border:2px solid {c['accent_blue']};border-radius:4px;padding:4px 18px;font-weight:600}}"
                f"QPushButton:hover{{background:{c['accent_blue']}22;border-color:{c['accent_teal']};color:{c['accent_teal']}}}"
                f"QPushButton:pressed{{background:{c['accent_blue']}44}}"
            )
        except Exception:
            pass


# ── Renew expiry dialog ───────────────────────────────────────────────────────

class RenewDialog(QDialog):
    """Small focused dialog to set a new expiry date or extend an existing one."""

    def __init__(self, parent, cred: dict):
        super().__init__(parent)
        self._cred      = cred
        self._new_date: Optional[str] = None   # None = remove; str = new date
        self.setWindowTitle(f"Renew Expiry — {cred.get('title') or 'Untitled'}")
        self.setMinimumWidth(430)
        self.setWindowModality(Qt.WindowModal)
        self._build_ui()
        _api.theme.register(self._apply_theme)
        self._apply_theme()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(12)

        title_lbl = QLabel(f"Renew expiry for «{self._cred.get('title') or 'Untitled'}»")
        title_lbl.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        title_lbl.setWordWrap(True)
        lay.addWidget(title_lbl)

        self._sep = QFrame()
        self._sep.setFrameShape(QFrame.HLine)
        self._sep.setFixedHeight(1)
        lay.addWidget(self._sep)

        # Current status
        exp = self._cred.get("expires")
        if exp:
            try:
                delta = (datetime.strptime(exp, "%Y-%m-%d").date() - date.today()).days
                if delta < 0:
                    status = f"⚠ Expired {abs(delta)} days ago  ({exp})"
                elif delta == 0:
                    status = f"⚠ Expires today  ({exp})"
                else:
                    status = f"Expires in {delta} days  ({exp})"
            except Exception:
                status = exp
        else:
            status = "No expiry currently set"
        self._curr_lbl = QLabel(f"Current: {status}")
        self._curr_lbl.setFont(QFont("Segoe UI", 9))
        lay.addWidget(self._curr_lbl)

        # Quick-extend buttons
        quick_lbl = QLabel("Quick extend from today:")
        quick_lbl.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        lay.addWidget(quick_lbl)

        quick_row = QHBoxLayout()
        quick_row.setSpacing(6)
        for days, label in [(30, "+30 days"), (60, "+60 days"), (90, "+90 days"),
                            (180, "+6 months"), (365, "+1 year")]:
            btn = QPushButton(label)
            btn.setFixedHeight(28)
            btn.clicked.connect(lambda checked=False, d=days: self._quick_extend(d))
            quick_row.addWidget(btn)
        lay.addLayout(quick_row)

        # Specific date picker
        specific_lbl = QLabel("Or pick a specific date:")
        specific_lbl.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        lay.addWidget(specific_lbl)

        self._date_edit = QDateEdit()
        self._date_edit.setCalendarPopup(True)
        self._date_edit.setDisplayFormat("yyyy-MM-dd")
        self._date_edit.setFixedHeight(30)
        if exp:
            try:
                qd = QDate.fromString(exp, "yyyy-MM-dd")
                self._date_edit.setDate(qd if qd.isValid() else QDate.currentDate().addDays(90))
            except Exception:
                self._date_edit.setDate(QDate.currentDate().addDays(90))
        else:
            self._date_edit.setDate(QDate.currentDate().addDays(90))
        lay.addWidget(self._date_edit)

        # Remove expiry
        self._no_exp_btn = QPushButton("🚫  Remove Expiry")
        self._no_exp_btn.setFixedHeight(28)
        self._no_exp_btn.setToolTip("Remove the expiry date entirely — credential will never expire")
        self._no_exp_btn.clicked.connect(self._on_remove_expiry)
        lay.addWidget(self._no_exp_btn)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self._cancel_btn = QPushButton("Cancel")
        self._cancel_btn.setFixedHeight(32)
        self._cancel_btn.clicked.connect(self.reject)
        self._save_btn = QPushButton("✔  Save Expiry")
        self._save_btn.setFixedHeight(32)
        self._save_btn.setDefault(True)
        self._save_btn.setAutoDefault(True)
        self._save_btn.clicked.connect(self._on_save)
        btn_row.addWidget(self._cancel_btn)
        btn_row.addWidget(self._save_btn)
        lay.addLayout(btn_row)

    def _quick_extend(self, days: int):
        self._date_edit.setDate(QDate.currentDate().addDays(days))

    def _on_remove_expiry(self):
        self._new_date = None
        self.accept()

    def _on_save(self):
        self._new_date = self._date_edit.date().toString("yyyy-MM-dd")
        self.accept()

    def get_new_date(self) -> Optional[str]:
        """Returns the chosen expiry date string, or None if expiry should be removed."""
        return self._new_date

    def _apply_theme(self):
        try:
            c = _api.theme.colors()
            self.setStyleSheet(
                f"QDialog{{background:{c['bg_dark']};color:{c['text_primary']}}}"
                f"QLabel{{color:{c['text_secondary']}}}"
                f"QDateEdit{{background:{c['bg_mid']};color:{c['text_primary']};"
                f"  border:1px solid {c['glow']};border-radius:4px;padding:3px 6px}}"
                f"QCalendarWidget QWidget{{background:{c['bg_dark']};color:{c['text_primary']}}}"
                f"QPushButton{{background:{c['bg_mid']};color:{c['text_primary']};"
                f"  border:1px solid {c['glow']};border-radius:4px;padding:4px 10px}}"
                f"QPushButton:hover{{border-color:{c['accent_blue']};background:{c['bg_dark']}}}"
                f"QPushButton:pressed{{background:{c['glow']}}}"
            )
            self._sep.setStyleSheet(f"background:{c['glow']}")
            self._save_btn.setStyleSheet(
                f"QPushButton{{background:transparent;color:{c['accent_blue']};"
                f"  border:2px solid {c['accent_blue']};border-radius:4px;padding:4px 18px;font-weight:600}}"
                f"QPushButton:hover{{background:{c['accent_blue']}22;border-color:{c['accent_teal']};color:{c['accent_teal']}}}"
                f"QPushButton:pressed{{background:{c['accent_blue']}44}}"
            )
            self._no_exp_btn.setStyleSheet(
                f"QPushButton{{background:transparent;color:{c['accent_red']};"
                f"  border:1px solid {c['accent_red']};border-radius:4px;padding:3px 10px}}"
                f"QPushButton:hover{{background:{c['accent_red']};color:#fff}}"
            )
        except Exception:
            pass


# ── Credential row widget ─────────────────────────────────────────────────────

class CredentialRow(QFrame):

    edit_requested   = Signal(str)
    delete_requested = Signal(str)
    renew_requested  = Signal(str)

    def __init__(self, cred: dict, parent=None):
        super().__init__(parent)
        self._cred      = cred
        self._plain_pwd: Optional[str] = None
        self._build_ui()
        self._apply_theme()

    def mouseDoubleClickEvent(self, event):
        """Double-click anywhere on the row to open the edit dialog."""
        if event.button() == Qt.LeftButton:
            self.edit_requested.emit(self._cred["id"])
        super().mouseDoubleClickEvent(event)

    def _build_ui(self):
        self.setFrameShape(QFrame.StyledPanel)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setCursor(Qt.ArrowCursor)

        main = QVBoxLayout(self)
        main.setContentsMargins(14, 10, 14, 10)
        main.setSpacing(5)

        # ── Row 1: title + category badge + age + action buttons ──────────
        top = QHBoxLayout()
        top.setSpacing(8)

        self._title_lbl = QLabel(f"<b>{self._cred.get('title') or '(Untitled)'}</b>")
        self._title_lbl.setFont(QFont("Segoe UI", 10))

        cat = self._cred.get("category", "")
        self._cat_lbl = QLabel(f"  {cat}  " if cat else "")
        self._cat_lbl.setFont(QFont("Segoe UI", 7, QFont.Weight.Bold))
        self._cat_lbl.setAlignment(Qt.AlignVCenter)

        self._age_lbl = QLabel(self._age_text())
        self._age_lbl.setFont(QFont("Segoe UI", 8))
        self._age_lbl.setToolTip("Time since this credential was last updated")

        top.addWidget(self._title_lbl)
        top.addWidget(self._cat_lbl)
        top.addStretch()
        top.addWidget(self._age_lbl)

        self._copy_user_btn = QPushButton("Copy Username")
        self._copy_user_btn.setToolTip("Copy username to clipboard")
        self._copy_user_btn.setFixedHeight(26)
        self._copy_user_btn.setVisible(bool(self._cred.get("username")))
        self._copy_user_btn.clicked.connect(self._on_copy_username)

        self._copy_pwd_btn = QPushButton("Copy Password")
        self._copy_pwd_btn.setToolTip("Copy password to clipboard")
        self._copy_pwd_btn.setFixedHeight(26)
        self._copy_pwd_btn.clicked.connect(self._on_copy_password)

        self._renew_btn = QPushButton("Renew")
        self._renew_btn.setFixedHeight(26)
        self._renew_btn.setToolTip("Set or extend the expiry date for this credential")
        self._renew_btn.clicked.connect(lambda: self.renew_requested.emit(self._cred["id"]))

        self._edit_btn = QPushButton("Edit")
        self._edit_btn.setFixedHeight(26)
        self._edit_btn.setToolTip("Edit credential  (or double-click the row)")
        self._edit_btn.clicked.connect(lambda: self.edit_requested.emit(self._cred["id"]))

        self._del_btn = QPushButton("🗑")
        self._del_btn.setFixedSize(26, 26)
        self._del_btn.setToolTip("Delete credential")
        self._del_btn.clicked.connect(lambda: self.delete_requested.emit(self._cred["id"]))

        top.addWidget(self._copy_user_btn)
        top.addWidget(self._copy_pwd_btn)
        top.addWidget(self._renew_btn)
        top.addWidget(self._edit_btn)
        top.addWidget(self._del_btn)
        main.addLayout(top)

        # ── Row 2: username / URL + password reveal ───────────────────────
        mid = QHBoxLayout()
        mid.setSpacing(12)

        if self._cred.get("username"):
            u  = self._cred["username"]
            ul = QLabel(f"<span style='opacity:0.55'>User:</span> {u}")
            ul.setFont(QFont("Segoe UI", 9))
            mid.addWidget(ul)

        mid.addStretch()

        self._pwd_display = QLabel("••••••••••••")
        self._pwd_display.setFont(QFont("Consolas", 9))
        self._pwd_display.setTextInteractionFlags(Qt.NoTextInteraction)
        self._pwd_display.setToolTip("Click 👁 to reveal / hide")

        self._reveal_btn = QPushButton("👁")
        self._reveal_btn.setFixedSize(26, 26)
        self._reveal_btn.setCheckable(True)
        self._reveal_btn.setToolTip("Toggle password visibility")
        self._reveal_btn.toggled.connect(self._toggle_reveal)

        mid.addWidget(self._pwd_display)
        mid.addWidget(self._reveal_btn)
        main.addLayout(mid)

        # ── Row 3: expiry + note snippet ──────────────────────────────────
        bot = QHBoxLayout()
        bot.setSpacing(12)

        # Always create and add _exp_lbl; visibility is controlled after
        self._exp_lbl = QLabel(self._expiry_text())
        self._exp_lbl.setFont(QFont("Segoe UI", 8))
        self._exp_lbl.setVisible(bool(self._expiry_text()))
        bot.addWidget(self._exp_lbl)

        self._note_lbl = QLabel("")
        self._note_lbl.setFont(QFont("Segoe UI", 8))
        note = self._cred.get("note", "")
        if note:
            short = note[:92] + "…" if len(note) > 92 else note
            self._note_lbl.setText(f"📝 {short}")
            self._note_lbl.setToolTip(note)
        self._note_lbl.setVisible(bool(note))
        bot.addWidget(self._note_lbl)

        bot.addStretch()
        main.addLayout(bot)

    # ── Helpers ───────────────────────────────────────────────────────────

    def _age_text(self) -> str:
        ts = self._cred.get("updated_at") or self._cred.get("created_at")
        if not ts:
            return ""
        try:
            days = (datetime.utcnow() - datetime.strptime(ts, "%Y-%m-%dT%H:%M:%SZ")).days
            if days == 0:  return "updated today"
            if days == 1:  return "1 day old"
            if days < 30:  return f"{days}d old"
            if days < 365: return f"{days // 30}mo old"
            return f"{days // 365}y old"
        except Exception:
            return ""

    def _expiry_text(self) -> str:
        exp = self._cred.get("expires")
        if not exp:
            return ""
        try:
            delta = (datetime.strptime(exp, "%Y-%m-%d").date() - date.today()).days
            if delta < 0:   return f"⚠ Expired {abs(delta)}d ago"
            if delta == 0:  return "⚠ Expires today"
            if delta <= 7:  return f"⚠ Expires in {delta}d!"
            if delta <= 14: return f"⚠ Expires in {delta}d"
            return f"Expires {exp}"
        except Exception:
            return ""

    def _is_expired(self) -> bool:
        exp = self._cred.get("expires")
        if not exp:
            return False
        try:
            return datetime.strptime(exp, "%Y-%m-%d").date() < date.today()
        except Exception:
            return False

    def _exp_near(self) -> bool:
        exp = self._cred.get("expires")
        if not exp or self._is_expired():
            return False
        try:
            return (datetime.strptime(exp, "%Y-%m-%d").date() - date.today()).days <= 14
        except Exception:
            return False

    # ── Actions ───────────────────────────────────────────────────────────

    def _toggle_reveal(self, on: bool):
        if on:
            if self._plain_pwd is None:
                try:
                    self._plain_pwd = _vault.decrypt(self._cred["password_enc"])
                except Exception as e:
                    _api.log(f"Decrypt error: {e}")
                    _api.toast("Failed to decrypt password.", "error")
                    self._reveal_btn.setChecked(False)
                    return
            self._pwd_display.setText(self._plain_pwd)
        else:
            self._plain_pwd = None
            self._pwd_display.setText("••••••••••••")

    def _on_copy_password(self):
        try:
            plain = _vault.decrypt(self._cred["password_enc"])
            _api.clipboard.set_text(plain)
            _api.toast(
                f"Password copied for «{self._cred.get('title') or 'Untitled'}».",
                "success"
            )
            # Visual flash feedback
            self._copy_pwd_btn.setText("✓ Copied!")
            self._copy_pwd_btn.setEnabled(False)
            QTimer.singleShot(1500, self._reset_copy_pwd_btn)
        except Exception as e:
            _api.log(f"Copy password error: {e}")
            _api.toast("Failed to copy password.", "error")

    def _reset_copy_pwd_btn(self):
        try:
            self._copy_pwd_btn.setText("Copy Password")
            self._copy_pwd_btn.setEnabled(True)
        except Exception:
            pass

    def _on_copy_username(self):
        try:
            user = self._cred.get("username", "")
            if user:
                _api.clipboard.set_text(user)
                _api.toast("Username copied.", "success")
                # Visual flash feedback
                self._copy_user_btn.setText("✓ Copied!")
                self._copy_user_btn.setEnabled(False)
                QTimer.singleShot(1500, self._reset_copy_user_btn)
        except Exception as e:
            _api.log(f"Copy username error: {e}")

    def _reset_copy_user_btn(self):
        try:
            self._copy_user_btn.setText("Copy Username")
            self._copy_user_btn.setEnabled(True)
        except Exception:
            pass

    def update_cred(self, cred: dict):
        self._cred      = cred
        self._plain_pwd = None
        self._reveal_btn.setChecked(False)
        self._pwd_display.setText("••••••••••••")
        self._title_lbl.setText(f"<b>{cred.get('title') or '(Untitled)'}</b>")
        self._age_lbl.setText(self._age_text())
        cat = cred.get("category", "")
        self._cat_lbl.setText(f"  {cat}  " if cat else "")
        exp_text = self._expiry_text()
        self._exp_lbl.setText(exp_text)
        self._exp_lbl.setVisible(bool(exp_text))
        note = cred.get("note", "")
        if note:
            short = note[:92] + "…" if len(note) > 92 else note
            self._note_lbl.setText(f"📝 {short}")
            self._note_lbl.setToolTip(note)
        self._note_lbl.setVisible(bool(note))
        self._copy_user_btn.setVisible(bool(cred.get("username")))
        self._apply_theme()

    # ── Theme ─────────────────────────────────────────────────────────────

    def _apply_theme(self):
        try:
            c       = _api.theme.colors()
            expired = self._is_expired()
            near    = self._exp_near()
            border  = (c["accent_red"] if expired
                       else c["accent_amber"] if near
                       else c["glow"])
            self.setStyleSheet(
                f"CredentialRow{{background:{c['bg_mid']};border:1px solid {border};"
                f"  border-radius:8px;margin:2px 0}}"
                f"CredentialRow:hover{{background:{c['bg_dark']};border-color:{c['accent_blue']}}}"
                f"QLabel{{color:{c['text_primary']};background:transparent}}"
                f"QPushButton{{background:{c['bg_dark']};color:{c['text_primary']};"
                f"  border:1px solid {c['glow']};border-radius:4px;"
                f"  padding:2px 8px;font-size:11px}}"
                f"QPushButton:hover{{border-color:{c['accent_blue']};background:{c['bg_mid']}}}"
                f"QPushButton:pressed{{background:{c['glow']}}}"
                f"QPushButton:disabled{{color:{c['text_dim']};border-color:{c['glow']}}}"
            )
            # Category pill badge
            cat_bg = c.get("accent_teal", "#20b2aa")
            self._cat_lbl.setStyleSheet(
                f"color:#000;background:{cat_bg};border-radius:8px;"
                f"background:transparent;color:{cat_bg};font-style:italic"
            )
            self._age_lbl.setStyleSheet(
                f"color:{c['accent_red'] if expired else c['text_dim']};"
                f"background:transparent"
            )
            self._exp_lbl.setStyleSheet(
                f"color:{c['accent_red'] if (expired or near) else c['text_dim']};"
                f"background:transparent;font-weight:{'600' if (expired or near) else '400'}"
            )
            self._note_lbl.setStyleSheet(
                f"color:{c['text_dim']};background:transparent"
            )
        except Exception:
            pass


# ── Plugin settings popup dialog ─────────────────────────────────────────────

class _PluginSettingsDialog(QDialog):
    """Standalone dialog wrapping the plugin settings UI, openable from the vault."""

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("⚙️  Credential Manager — Settings")
        self.setMinimumWidth(500)
        self.setWindowModality(Qt.WindowModal)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 14)
        lay.setSpacing(0)

        # Embed a fresh settings widget (same content as the Settings-tab)
        self._settings_widget = _build_settings_tab()
        lay.addWidget(self._settings_widget, 1)

        btn_row = QHBoxLayout()
        btn_row.setContentsMargins(16, 0, 16, 0)
        btn_row.addStretch()
        close_btn = QPushButton("Close")
        close_btn.setFixedHeight(34)
        close_btn.setDefault(True)
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)
        lay.addLayout(btn_row)
        self._close_btn = close_btn

        _api.theme.register(self._apply_theme)
        self._apply_theme()

    def _apply_theme(self):
        try:
            c = _api.theme.colors()
            self.setStyleSheet(
                f"QDialog{{background:{c['bg_dark']};color:{c['text_primary']}}}"
            )
            self._close_btn.setStyleSheet(
                f"QPushButton{{background:transparent;color:{c['accent_blue']};"
                f"  border:2px solid {c['accent_blue']};border-radius:4px;padding:4px 18px;font-weight:600}}"
                f"QPushButton:hover{{background:{c['accent_blue']}22;border-color:{c['accent_teal']};color:{c['accent_teal']}}}"
                f"QPushButton:pressed{{background:{c['accent_blue']}44}}"
            )
        except Exception:
            pass


# ── Main vault window ─────────────────────────────────────────────────────────

class CredentialManagerWindow(QDialog):

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("🔐 Credential Manager")
        self.setMinimumSize(780, 560)
        self.resize(900, 680)
        self.setWindowFlags(
            Qt.Window | Qt.WindowCloseButtonHint | Qt.WindowMinMaxButtonsHint
        )
        self._creds: list[dict]               = []
        self._rows:  dict[str, CredentialRow] = {}
        self._search_text      = ""
        self._filter_category  = "All"
        self._sort_key         = "title"
        self._reveal_all_active = False  # True when global reveal-all is toggled on

        # Debounce timer for search input (200 ms)
        self._search_debounce = QTimer(self)
        self._search_debounce.setSingleShot(True)
        self._search_debounce.setInterval(200)
        self._search_debounce.timeout.connect(self._do_rebuild)

        # Auto-clear timer for status messages (5 s)
        self._status_clear_timer = QTimer(self)
        self._status_clear_timer.setSingleShot(True)
        self._status_clear_timer.setInterval(5000)
        self._status_clear_timer.timeout.connect(lambda: self._status_bar.setText("  Ready"))

        self._build_ui()
        _api.theme.register(self._apply_theme)
        self._apply_theme()
        self._load()

    def closeEvent(self, event):
        global _master_unlocked
        _master_unlocked = False
        super().closeEvent(event)

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── Header ────────────────────────────────────────────────────────
        header = QWidget()
        header.setFixedHeight(56)
        hlay = QHBoxLayout(header)
        hlay.setContentsMargins(16, 8, 16, 8)
        hlay.setSpacing(10)

        icon_lbl  = QLabel("🔐")
        icon_lbl.setFont(QFont("Segoe UI", 15))
        title_lbl = QLabel("Credential Manager")
        title_lbl.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        hlay.addWidget(icon_lbl)
        hlay.addWidget(title_lbl)
        hlay.addStretch()

        self._export_btn = QPushButton("📤 Export")
        self._export_btn.setToolTip("Export to passphrase-encrypted .ccvault backup")
        self._export_btn.setFixedHeight(30)
        self._export_btn.clicked.connect(self._export)

        self._import_btn = QPushButton("📥 Import")
        self._import_btn.setToolTip("Import from .ccvault backup")
        self._import_btn.setFixedHeight(30)
        self._import_btn.clicked.connect(self._import)

        self._settings_btn = QPushButton("⚙️")
        self._settings_btn.setFixedSize(46, 30)
        self._settings_btn.setToolTip("Plugin settings  (Ctrl+,)")
        self._settings_btn.clicked.connect(self._open_settings_dialog)

        self._add_btn = QPushButton("＋  New Credential")
        self._add_btn.setFixedHeight(30)
        self._add_btn.setToolTip(f"Add a new credential  ({_HOTKEY} also opens this window)")
        self._add_btn.clicked.connect(self._add_new)

        hlay.addWidget(self._export_btn)
        hlay.addWidget(self._import_btn)
        hlay.addWidget(self._settings_btn)
        hlay.addWidget(self._add_btn)
        root.addWidget(header)
        self._header = header

        # Accent separator below header
        self._header_line = QFrame()
        self._header_line.setFrameShape(QFrame.HLine)
        self._header_line.setFixedHeight(1)
        root.addWidget(self._header_line)

        # ── Toolbar ───────────────────────────────────────────────────────
        toolbar = QWidget()
        toolbar.setFixedHeight(44)
        tlay = QHBoxLayout(toolbar)
        tlay.setContentsMargins(12, 5, 12, 5)
        tlay.setSpacing(8)

        self._search_edit = QLineEdit()
        self._search_edit.setPlaceholderText("🔍  Search title, username, category, note…")
        self._search_edit.setClearButtonEnabled(True)
        self._search_edit.setFixedHeight(32)
        self._search_edit.textChanged.connect(self._on_search)

        self._cat_filter = QComboBox()
        self._cat_filter.setFixedWidth(155)
        self._cat_filter.setFixedHeight(32)
        self._cat_filter.setToolTip("Filter by category")
        self._cat_filter.currentTextChanged.connect(self._on_filter_category)

        self._sort_combo = QComboBox()
        self._sort_combo.addItems(["Sort: Title", "Sort: Oldest", "Sort: Newest", "Sort: Expiry"])
        self._sort_combo.setFixedWidth(135)
        self._sort_combo.setFixedHeight(32)
        self._sort_combo.currentIndexChanged.connect(self._on_sort_changed)

        self._count_lbl = QLabel("")
        self._count_lbl.setFont(QFont("Segoe UI", 8))
        self._count_lbl.setToolTip("Visible / total credentials")

        self._show_all_btn = QPushButton("👁 Show All")
        self._show_all_btn.setFixedHeight(32)
        self._show_all_btn.setCheckable(True)
        self._show_all_btn.setToolTip("Reveal or hide all passwords in the list at once")
        self._show_all_btn.toggled.connect(self._toggle_reveal_all)

        tlay.addWidget(self._search_edit, 1)
        tlay.addWidget(self._cat_filter)
        tlay.addWidget(self._sort_combo)
        tlay.addWidget(self._show_all_btn)
        tlay.addWidget(self._count_lbl)
        root.addWidget(toolbar)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        root.addWidget(sep)

        # ── Scroll area ───────────────────────────────────────────────────
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._scroll.setFrameShape(QFrame.NoFrame)

        self._list_widget = QWidget()
        self._list_layout = QVBoxLayout(self._list_widget)
        self._list_layout.setContentsMargins(12, 10, 12, 10)
        self._list_layout.setSpacing(6)
        self._list_layout.addStretch()

        self._scroll.setWidget(self._list_widget)
        root.addWidget(self._scroll, 1)

        # ── Status bar ────────────────────────────────────────────────────
        self._status_bar = QLabel("  Ready")
        self._status_bar.setFixedHeight(24)
        self._status_bar.setFont(QFont("Segoe UI", 8))
        root.addWidget(self._status_bar)

        # ── Keyboard shortcuts ────────────────────────────────────────────
        try:
            QShortcut(QKeySequence("Ctrl+F"), self).activated.connect(
                self._focus_search
            )
            QShortcut(QKeySequence("Ctrl+N"), self).activated.connect(self._add_new)
            QShortcut(QKeySequence("Ctrl+,"), self).activated.connect(
                self._open_settings_dialog
            )
        except Exception:
            pass

    def _focus_search(self):
        self._search_edit.setFocus()
        self._search_edit.selectAll()

    def _open_settings_dialog(self):
        dlg = _PluginSettingsDialog(self)
        dlg.exec()

    # ── Data ──────────────────────────────────────────────────────────────

    def _load(self):
        try:
            self._creds = _load_credentials()
            self._rebuild_list()
            self._refresh_categories()
            self._update_title()
            n = len(self._creds)
            self._set_status(
                f"  {n} credential{'s' if n != 1 else ''} loaded."
                + (f"  Encryption: {_vault._backend.upper()}" if _vault else ""),
                temp=True
            )
        except Exception as e:
            _api.log(f"_load error: {e}")
            _api.toast("Error loading credentials.", "error")

    def _save(self):
        try:
            _save_credentials(self._creds)
        except Exception as e:
            _api.log(f"_save error: {e}")
            _api.toast("Error saving credentials — data may not have been written.", "error")
            raise

    # ── Rendering ─────────────────────────────────────────────────────────

    def _rebuild_list(self):
        """Remove all existing rows and rebuild from filtered/sorted list."""
        while self._list_layout.count() > 1:
            item = self._list_layout.takeAt(0)
            if item and item.widget():
                item.widget().deleteLater()
        self._rows.clear()

        filtered = self._filtered_sorted()
        total    = len(self._creds)
        vis      = len(filtered)
        self._count_lbl.setText(f"{vis} / {total}")

        if not filtered:
            self._list_layout.insertWidget(0, self._make_empty_state())
            return

        for cred in filtered:
            row = CredentialRow(cred, self._list_widget)
            row.edit_requested.connect(self._edit_credential)
            row.delete_requested.connect(self._delete_credential)
            row.renew_requested.connect(self._renew_credential)
            self._list_layout.insertWidget(self._list_layout.count() - 1, row)
            self._rows[cred["id"]] = row
            if self._reveal_all_active:
                row._reveal_btn.setChecked(True)

    def _make_empty_state(self) -> QWidget:
        """Returns a centered empty-state widget."""
        frame = QWidget()
        lay   = QVBoxLayout(frame)
        lay.setAlignment(Qt.AlignCenter)
        lay.setContentsMargins(40, 40, 40, 40)
        lay.setSpacing(10)

        if self._search_text or self._filter_category != "All":
            icon_lbl = QLabel("🔍")
            icon_lbl.setFont(QFont("Segoe UI", 28))
            icon_lbl.setAlignment(Qt.AlignCenter)
            msg_lbl  = QLabel("No credentials match your search.")
            msg_lbl.setFont(QFont("Segoe UI", 10))
            hint_lbl = QLabel("Try a different search term or category filter.")
            hint_lbl.setFont(QFont("Segoe UI", 8))
        else:
            icon_lbl = QLabel("🔐")
            icon_lbl.setFont(QFont("Segoe UI", 36))
            icon_lbl.setAlignment(Qt.AlignCenter)
            msg_lbl  = QLabel("No credentials yet.")
            msg_lbl.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
            hint_lbl = QLabel(
                "Click  ＋ New Credential  to add your first entry.\n"
                f"You can also press  Ctrl+N  at any time."
            )
            hint_lbl.setFont(QFont("Segoe UI", 9))

        for lbl in (icon_lbl, msg_lbl, hint_lbl):
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setWordWrap(True)
            lay.addWidget(lbl)

        return frame

    def _filtered_sorted(self) -> list[dict]:
        creds = list(self._creds)
        q = self._search_text.lower()
        if q:
            creds = [
                c for c in creds
                if q in (c.get("title")    or "").lower()
                or q in (c.get("username") or "").lower()
                or q in (c.get("category") or "").lower()
                or q in (c.get("note")     or "").lower()
            ]
        if self._filter_category and self._filter_category != "All":
            creds = [c for c in creds
                     if (c.get("category") or "") == self._filter_category]

        if   self._sort_key == "title":
            creds.sort(key=lambda c: (c.get("title") or "").lower())
        elif self._sort_key == "oldest":
            creds.sort(key=lambda c: c.get("updated_at") or c.get("created_at") or "")
        elif self._sort_key == "newest":
            creds.sort(
                key=lambda c: c.get("updated_at") or c.get("created_at") or "",
                reverse=True
            )
        elif self._sort_key == "expiry":
            creds.sort(key=lambda c: c.get("expires") or "9999-12-31")
        return creds

    def _refresh_categories(self):
        cats    = sorted(set(c.get("category", "") for c in self._creds if c.get("category")))
        current = self._cat_filter.currentText()
        self._cat_filter.blockSignals(True)
        self._cat_filter.clear()
        self._cat_filter.addItem("All categories")
        for cat in cats:
            self._cat_filter.addItem(cat)
        idx = self._cat_filter.findText(current)
        self._cat_filter.setCurrentIndex(max(0, idx))
        self._cat_filter.blockSignals(False)

    def _update_title(self):
        total = len(self._creds)
        try:
            expired = sum(
                1 for c in self._creds
                if c.get("expires")
                and datetime.strptime(c["expires"], "%Y-%m-%d").date() < date.today()
            )
        except Exception:
            expired = 0
        if expired:
            self.setWindowTitle(
                f"🔐 Credential Manager  ({total}  •  ⚠ {expired} expired)"
            )
        else:
            self.setWindowTitle(f"🔐 Credential Manager  ({total})")

    # ── Actions ───────────────────────────────────────────────────────────

    def _add_new(self):
        try:
            dlg = CredentialDialog(self)
            if dlg.exec() == QDialog.Accepted:
                data = dlg.get_data()
                if data:
                    self._creds.append(data)
                    self._save()
                    self._rebuild_list()
                    self._refresh_categories()
                    self._update_title()
                    title = data.get("title") or "Untitled"
                    _api.toast(f"Saved «{title}».", "success")
                    self._set_status(f"  Added «{title}».", temp=True)
        except Exception as e:
            _api.log(f"_add_new error: {e}\n{traceback.format_exc()}")
            _api.toast("Error adding credential.", "error")

    def _edit_credential(self, cred_id: str):
        try:
            cred = next((c for c in self._creds if c["id"] == cred_id), None)
            if cred is None:
                return
            dlg = CredentialDialog(self, existing=cred)
            if dlg.exec() == QDialog.Accepted:
                data = dlg.get_data()
                if data:
                    idx = next((i for i, c in enumerate(self._creds)
                                if c["id"] == cred_id), None)
                    if idx is not None:
                        self._creds[idx] = data
                    self._save()
                    self._rebuild_list()
                    self._refresh_categories()
                    self._update_title()
                    title = data.get("title") or "Untitled"
                    _api.toast(f"Updated «{title}».", "success")
                    self._set_status(f"  Updated «{title}».", temp=True)
        except Exception as e:
            _api.log(f"_edit_credential error: {e}")
            _api.toast("Error editing credential.", "error")

    def _delete_credential(self, cred_id: str):
        try:
            cred = next((c for c in self._creds if c["id"] == cred_id), None)
            if cred is None:
                return
            label = cred.get("title") or "Untitled"
            if not _api.ui.ask_confirm(
                "Delete Credential",
                f"Permanently delete «{label}»?\nThis cannot be undone."
            ):
                return
            self._creds = [c for c in self._creds if c["id"] != cred_id]
            self._save()
            self._rebuild_list()
            self._refresh_categories()
            self._update_title()
            _api.toast(f"Deleted «{label}».", "info")
            self._set_status(f"  Deleted «{label}».", temp=True)
        except Exception as e:
            _api.log(f"_delete error: {e}")
            _api.toast("Error deleting credential.", "error")

    # ── Export ────────────────────────────────────────────────────────────

    def _export(self):
        """Export vault as a passphrase-encrypted .ccvault file.
        Key derivation: stdlib hashlib.pbkdf2_hmac — no external packages required.
        Includes a passphrase confirmation step to prevent typos locking you out.
        """
        try:
            passphrase, ok = QInputDialog.getText(
                self, "Export — Set Passphrase",
                "Enter a strong passphrase to protect the backup.\n"
                "You will need this passphrase to import on another machine.",
                QLineEdit.Password
            )
            if not ok or not passphrase:
                return
            if len(passphrase) < 8:
                _api.ui.show_message(
                    "Passphrase Too Short",
                    "Please use at least 8 characters.", "warn"
                )
                return

            # Confirmation step — prevents typos that would lock out the backup
            confirm, ok2 = QInputDialog.getText(
                self, "Export — Confirm Passphrase",
                "Re-enter your passphrase to confirm:",
                QLineEdit.Password
            )
            if not ok2:
                return
            if confirm != passphrase:
                _api.ui.show_message(
                    "Passphrase Mismatch",
                    "The passphrases do not match. Export cancelled.", "warn"
                )
                return

            path = _api.files.save_dialog(
                "Export Credentials",
                "Encrypted Vault Backup (*.ccvault);;All Files (*)"
            )
            if not path:
                return

            self._set_status("  Exporting… (deriving key, please wait)", temp=False)

            salt       = secrets.token_bytes(16)
            export_key = _pbkdf2_key(passphrase, salt)   # ~1 s at 600k iterations

            export_entries = []
            errors = 0
            for c in self._creds:
                try:
                    plain   = _vault.decrypt(c["password_enc"])
                    raw_enc = _vault.raw_encrypt(export_key, plain.encode("utf-8"))
                    entry   = dict(c)
                    entry["password_enc"] = base64.b64encode(raw_enc).decode("ascii")
                    export_entries.append(entry)
                except Exception as ee:
                    _api.log(f"Export: skipping {c.get('id')}: {ee}")
                    errors += 1

            payload = {
                "version":    2,
                "salt":       base64.b64encode(salt).decode("ascii"),
                "iterations": 600_000,
                "count":      len(export_entries),
                "entries":    export_entries,
            }
            Path(path).write_text(json.dumps(payload, indent=2), encoding="utf-8")

            fname = Path(path).name
            msg   = f"Exported {len(export_entries)} credential(s) → {fname}"
            if errors:
                msg += f"  ({errors} failed)"
            _api.toast(msg, "success" if not errors else "warn")
            self._set_status(f"  {msg}", temp=True)

        except Exception as e:
            _api.log(f"_export error: {e}\n{traceback.format_exc()}")
            _api.toast("Export failed.", "error")
            self._set_status("  Export failed — check log.", temp=True)

    # ── Import ────────────────────────────────────────────────────────────

    def _import(self):
        """Import from a .ccvault backup. Uses stdlib PBKDF2 + vault backend for decryption."""
        try:
            path = _api.files.open_dialog(
                "Import Credentials",
                "Encrypted Vault Backup (*.ccvault);;All Files (*)"
            )
            if not path:
                return

            passphrase, ok = QInputDialog.getText(
                self, "Import — Enter Passphrase",
                "Enter the passphrase used when exporting this backup.",
                QLineEdit.Password
            )
            if not ok or not passphrase:
                return

            self._set_status("  Importing… (deriving key, please wait)", temp=False)

            payload    = json.loads(Path(path).read_text(encoding="utf-8"))
            if payload.get("version", 1) < 2:
                _api.ui.show_message(
                    "Unsupported Format",
                    "This .ccvault file uses an older format (version < 2).\n"
                    "Re-export it from the original machine to upgrade.", "warn"
                )
                self._set_status("  Import cancelled — unsupported file version.", temp=True)
                return

            salt       = base64.b64decode(payload["salt"].encode("ascii"))
            iterations = payload.get("iterations", 600_000)
            import_key = _pbkdf2_key(passphrase, salt, iterations)

            imported     = 0
            skipped      = 0
            replaced     = 0
            existing_ids = {c["id"] for c in self._creds}

            for entry in payload.get("entries", []):
                try:
                    raw_enc = base64.b64decode(entry["password_enc"].encode("ascii"))
                    plain   = _vault.raw_decrypt(import_key, raw_enc).decode("utf-8")
                    entry["password_enc"] = _vault.encrypt(plain)
                    if entry["id"] in existing_ids:
                        idx = next(
                            i for i, c in enumerate(self._creds)
                            if c["id"] == entry["id"]
                        )
                        self._creds[idx] = entry
                        replaced += 1
                    else:
                        self._creds.append(entry)
                        existing_ids.add(entry["id"])
                    imported += 1
                except Exception as ee:
                    _api.log(f"Import: failed to decrypt entry: {ee}")
                    skipped += 1

            self._save()
            self._rebuild_list()
            self._refresh_categories()
            self._update_title()

            parts = [f"Imported {imported} credential(s)."]
            if replaced:
                parts.append(f"{replaced} updated.")
            if skipped:
                parts.append(f"{skipped} skipped (wrong passphrase or corrupt).")
            msg = "  " + "  ".join(parts)
            _api.toast(msg.strip(), "success" if imported > 0 else "warn")
            self._set_status(msg, temp=True)

        except json.JSONDecodeError:
            self._set_status("  Import failed — not a valid .ccvault file.", temp=True)
            _api.toast("Import failed: not a valid .ccvault file.", "error")
        except Exception as e:
            _api.log(f"_import error: {e}\n{traceback.format_exc()}")
            _api.toast("Import failed. Check passphrase and file.", "error")
            self._set_status("  Import failed — check log.", temp=True)

    # ── Reveal-all toggle ─────────────────────────────────────────────────

    def _toggle_reveal_all(self, show: bool):
        self._reveal_all_active = show
        self._show_all_btn.setText("🙈 Hide All" if show else "👁 Show All")
        for row in self._rows.values():
            try:
                row._reveal_btn.setChecked(show)
            except Exception:
                pass

    # ── Renew credential expiry ───────────────────────────────────────────

    def _renew_credential(self, cred_id: str):
        try:
            cred = next((c for c in self._creds if c["id"] == cred_id), None)
            if cred is None:
                return
            dlg = RenewDialog(self, cred)
            if dlg.exec() != QDialog.Accepted:
                return
            new_date = dlg.get_new_date()
            idx = next((i for i, c in enumerate(self._creds) if c["id"] == cred_id), None)
            if idx is None:
                return
            self._creds[idx]["expires"]    = new_date
            self._creds[idx]["updated_at"] = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
            self._save()
            self._rebuild_list()
            self._update_title()
            title = cred.get("title") or "Untitled"
            if new_date:
                msg = f"Expiry updated to {new_date} for «{title}»."
            else:
                msg = f"Expiry removed for «{title}»."
            _api.toast(msg, "success")
            self._set_status(f"  {msg}", temp=True)
        except Exception as e:
            _api.log(f"_renew_credential error: {e}\n{traceback.format_exc()}")
            _api.toast("Error renewing credential.", "error")

    # ── Toolbar signals ───────────────────────────────────────────────────

    def _on_search(self, text: str):
        """Debounce search — wait 200 ms of silence before rebuilding."""
        self._search_text = text
        self._search_debounce.start()   # restarts the timer each keystroke

    def _do_rebuild(self):
        self._rebuild_list()

    def _on_filter_category(self, cat: str):
        # Map display text back to stored key
        self._filter_category = "All" if cat.startswith("All") else cat
        self._rebuild_list()

    def _on_sort_changed(self, idx: int):
        self._sort_key = ["title", "oldest", "newest", "expiry"][idx]
        self._rebuild_list()

    def _set_status(self, msg: str, *, temp: bool = False):
        """Update the status bar.  temp=True auto-clears after 5 s."""
        self._status_bar.setText(msg)
        self._status_clear_timer.stop()
        if temp:
            self._status_clear_timer.start()

    # ── Theme ─────────────────────────────────────────────────────────────

    def _apply_theme(self):
        try:
            c = _api.theme.colors()
            self.setStyleSheet(
                f"QDialog,QWidget{{background:{c['bg_dark']};color:{c['text_primary']}}}"
                f"QScrollArea{{background:{c['bg_dark']};border:none}}"
                f"QLineEdit{{background:{c['bg_mid']};color:{c['text_primary']};"
                f"  border:1px solid {c['glow']};border-radius:4px;padding:4px 8px}}"
                f"QLineEdit:focus{{border-color:{c['accent_blue']}}}"
                f"QComboBox{{background:{c['bg_mid']};color:{c['text_primary']};"
                f"  border:1px solid {c['glow']};border-radius:4px;padding:3px 8px}}"
                f"QComboBox QAbstractItemView{{background:{c['bg_mid']};"
                f"  color:{c['text_primary']};"
                f"  selection-background-color:{c['accent_blue']}}}"
                f"QPushButton{{background:{c['bg_mid']};color:{c['text_primary']};"
                f"  border:1px solid {c['glow']};border-radius:4px;padding:3px 10px}}"
                f"QPushButton:hover{{border-color:{c['accent_blue']};background:{c['bg_dark']}}}"
                f"QPushButton:pressed{{background:{c['glow']}}}"
                f"QLabel{{color:{c['text_primary']};background:transparent}}"
                f"QScrollBar:vertical{{background:{c['bg_dark']};width:8px;border:none}}"
                f"QScrollBar::handle:vertical{{background:{c['glow']};"
                f"  border-radius:4px;min-height:20px}}"
                f"QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical{{height:0}}"
                f"QFrame[frameShape='4']{{color:{c['glow']}}}"
            )
            self._add_btn.setStyleSheet(
                f"QPushButton{{background:transparent;color:{c['accent_blue']};"
                f"  border:2px solid {c['accent_blue']};border-radius:4px;padding:4px 14px;font-weight:600}}"
                f"QPushButton:hover{{background:{c['accent_blue']}22;border-color:{c['accent_teal']};color:{c['accent_teal']}}}"
                f"QPushButton:pressed{{background:{c['accent_blue']}44}}"
            )
            self._status_bar.setStyleSheet(
                f"color:{c['text_dim']};background:{c['bg_dark']};"
                f"border-top:1px solid {c['glow']};padding:3px 6px"
            )
            self._header.setStyleSheet(
                f"background:{c['bg_mid']};border-bottom:1px solid {c['glow']}"
            )
            self._header_line.setStyleSheet(
                f"background:{c['accent_blue']}"
            )
            self._count_lbl.setStyleSheet(f"color:{c['text_dim']}")
            self._show_all_btn.setStyleSheet(
                f"QPushButton{{background:{c['bg_mid']};color:{c['text_primary']};"
                f"  border:1px solid {c['glow']};border-radius:4px;padding:3px 10px}}"
                f"QPushButton:checked{{background:{c['accent_teal']};color:#fff;"
                f"  border-color:{c['accent_teal']}}}"
                f"QPushButton:hover{{border-color:{c['accent_blue']}}}"
            )
            for row in self._rows.values():
                row._apply_theme()
        except Exception as e:
            _api.log(f"_apply_theme error: {e}")
