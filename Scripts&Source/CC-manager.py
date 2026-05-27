"""
CC-Manager.py
Manages CommandCenter versioning + compilation, and .ccplug plugin packaging.
"""

import json
import os
import re
import subprocess
import threading
import zipfile
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

# ── Paths ─────────────────────────────────────────────────────────────────────

BASE_DIR     = Path(__file__).parent
CC_PY        = BASE_DIR / "CommandCenter.py"
VER_TXT      = BASE_DIR / "version.txt"
COMPILE_JSON = BASE_DIR / "CC-compile.json"
PLUGINS_DIR  = BASE_DIR / "plugins"

# ── Helpers ───────────────────────────────────────────────────────────────────

def _read_cc_version() -> str:
    text = CC_PY.read_text(encoding="utf-8")
    m = re.search(r'^APP_VERSION\s*=\s*"([^"]+)"', text, re.MULTILINE)
    return m.group(1) if m else "?"


def _version_tuple(v: str) -> tuple:
    parts = v.split(".")
    while len(parts) < 4:
        parts.append("0")
    return tuple(int(x) for x in parts[:4])


def _today_str() -> str:
    d = datetime.now()
    return d.strftime(f"%B {d.day}, %Y")


def _update_cc_version(new_ver: str) -> None:
    text = CC_PY.read_text(encoding="utf-8")
    if not re.search(r'^APP_VERSION\s*=\s*"[^"]+"', text, re.MULTILINE):
        raise ValueError("APP_VERSION not found in CommandCenter.py")
    new_text = re.sub(
        r'^(APP_VERSION\s*=\s*")[^"]+(")',
        rf'\g<1>{new_ver}\g<2>',
        text, count=1, flags=re.MULTILINE,
    )
    if new_text != text:
        CC_PY.write_text(new_text, encoding="utf-8")


def _update_version_txt(new_ver: str) -> None:
    text = VER_TXT.read_text(encoding="utf-8")
    t = _version_tuple(new_ver)
    tup = f"({t[0]}, {t[1]}, {t[2]}, {t[3]})"
    today = _today_str()

    text = re.sub(r"filevers=\([^)]+\)", f"filevers={tup}", text)
    text = re.sub(r"prodvers=\([^)]+\)", f"prodvers={tup}", text)
    text = re.sub(
        r"(StringStruct\(u'FileVersion',\s*u')[^']+(')",
        rf"\g<1>{new_ver}\g<2>", text,
    )
    text = re.sub(
        r"(StringStruct\(u'ProductVersion',\s*u')[^']+(')",
        rf"\g<1>{new_ver}\g<2>", text,
    )
    text = re.sub(
        r"(StringStruct\(u'Comments',\s*u')[^']*(')",
        rf"\g<1>Built on {today}\g<2>", text,
    )
    VER_TXT.write_text(text, encoding="utf-8")


def _build_pyinstaller_cmd() -> list:
    """Parse CC-compile.json and return the pyinstaller argument list."""
    data = json.loads(COMPILE_JSON.read_text(encoding="utf-8"))
    opts = data.get("pyinstallerOptions", [])

    _BOOL_FLAGS = {
        "noconfirm":                   "--noconfirm",
        "onefile":                     "--onefile",
        "clean_build":                 "--clean",
        "strip":                       "--strip",
        "noupx":                       "--noupx",
        "uac_admin":                   "--uac-admin",
        "bootloader_ignore_signals":   "--bootloader-ignore-signals",
        "disable_windowed_traceback":  "--disable-windowed-traceback",
    }
    _VAL_FLAGS = {
        "icon_file":    "--icon",
        "version_file": "--version-file",
        "name":         "--name",
        "distpath":     "--distpath",
        "workpath":     "--workpath",
        "specpath":     "--specpath",
    }

    cmd = ["pyinstaller"]
    script = None
    datas = []

    for opt in opts:
        dest = opt["optionDest"]
        val  = opt["value"]
        if dest == "filenames":
            script = val
        elif dest == "console":
            if val is False:
                cmd.append("--windowed")
        elif dest == "datas":
            datas.append(val)
        elif dest in _BOOL_FLAGS:
            if val:
                cmd.append(_BOOL_FLAGS[dest])
        elif dest in _VAL_FLAGS:
            if val:
                cmd.extend([_VAL_FLAGS[dest], val])

    for d in datas:
        cmd.extend(["--add-data", d])

    if script:
        cmd.append(script)

    return cmd


def _scan_plugins() -> list:
    """Return [(folder_path, display_name, version), ...] for each plugin."""
    results = []
    if not PLUGINS_DIR.exists():
        return results
    for d in sorted(PLUGINS_DIR.iterdir()):
        if not d.is_dir():
            continue
        mf = d / "manifest.json"
        if not mf.exists():
            continue
        try:
            info = json.loads(mf.read_text(encoding="utf-8"))
        except Exception:
            continue
        name    = info.get("name") or d.name
        version = info.get("version", "?")
        results.append((d, name, version))
    return results


def _update_plugin_version(plugin_dir: Path, new_ver: str) -> None:
    mf   = plugin_dir / "manifest.json"
    info = json.loads(mf.read_text(encoding="utf-8"))
    info["version"] = new_ver
    mf.write_text(json.dumps(info, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def _repack_plugin(plugin_dir: Path) -> Path:
    """Compress manifest.json + plugin.py into the existing (or derived) .ccplug."""
    manifest  = plugin_dir / "manifest.json"
    plugin_py = plugin_dir / "plugin.py"
    if not manifest.exists():
        raise FileNotFoundError(f"manifest.json not found in {plugin_dir.name}")
    if not plugin_py.exists():
        raise FileNotFoundError(f"plugin.py not found in {plugin_dir.name}")

    existing = list(plugin_dir.glob("*.ccplug"))
    out_name = existing[0].name if existing else f"{plugin_dir.name}.ccplug"
    out_path = plugin_dir / out_name

    tmp = plugin_dir / f"_tmp_{out_name}"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(manifest,  "manifest.json")
        zf.write(plugin_py, "plugin.py")
    if out_path.exists():
        out_path.unlink()
    tmp.rename(out_path)
    return out_path


def _validate_version(v: str) -> bool:
    return bool(re.fullmatch(r"\d+\.\d+\.\d+\.\d+", v.strip()))


# ── GUI ───────────────────────────────────────────────────────────────────────

FONT_MONO = ("Consolas", 9)
FONT_BOLD = ("Segoe UI", 10, "bold")
CLR_OK    = "#5cb85c"
CLR_ERR   = "#d9534f"
CLR_WARN  = "#e8a838"
CLR_INFO  = "#5bc0de"


class CCManager(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CC Manager")
        self.minsize(740, 520)
        self.resizable(True, True)

        try:
            self.iconbitmap(str(BASE_DIR / "CommandCenter.ico"))
        except Exception:
            pass

        self._notebook = ttk.Notebook(self)
        self._notebook.pack(fill="both", expand=True, padx=8, pady=8)

        self._cc_frame = ttk.Frame(self._notebook)
        self._pl_frame = ttk.Frame(self._notebook)
        self._notebook.add(self._cc_frame, text="  CommandCenter  ")
        self._notebook.add(self._pl_frame, text="  Plugins  ")

        self._build_cc_tab()
        self._build_plugin_tab()

        # status bar
        self._status_var = tk.StringVar(value="Ready.")
        ttk.Label(self, textvariable=self._status_var, anchor="w",
                  relief="sunken", font=("Segoe UI", 9)).pack(
            side="bottom", fill="x", padx=8, pady=(0, 4))

        self.after(100, self._refresh_all)

    # ═══════════════════════════════════════════════════════════════════════════
    # CommandCenter Tab
    # ═══════════════════════════════════════════════════════════════════════════

    def _build_cc_tab(self):
        f   = self._cc_frame
        PAD = {"padx": 10, "pady": 5}

        # ── Version section ──────────────────────────────────────────────────
        vf = ttk.LabelFrame(f, text="Version")
        vf.pack(fill="x", padx=10, pady=(10, 4))

        ttk.Label(vf, text="Current:").grid(row=0, column=0, sticky="w", **PAD)
        self._cc_cur_lbl = ttk.Label(vf, text="…", font=FONT_BOLD)
        self._cc_cur_lbl.grid(row=0, column=1, sticky="w", **PAD)

        ttk.Label(vf, text="New version:").grid(row=1, column=0, sticky="w", **PAD)
        self._cc_new_entry = ttk.Entry(vf, width=16, font=FONT_MONO)
        self._cc_new_entry.grid(row=1, column=1, sticky="w", **PAD)

        ttk.Button(vf, text="Apply Version",
                   command=self._apply_cc_version).grid(row=1, column=2, **PAD)

        ttk.Label(vf,
                  text="Updates APP_VERSION in CommandCenter.py and all fields in version.txt",
                  foreground="gray").grid(row=2, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 4))

        # ── Action buttons ───────────────────────────────────────────────────
        bf = ttk.Frame(f)
        bf.pack(fill="x", padx=10, pady=4)

        ttk.Button(bf, text="Compile",
                   command=self._compile).pack(side="left", padx=(0, 6))
        ttk.Button(bf, text="Apply Version + Compile",
                   command=self._apply_and_compile).pack(side="left", padx=(0, 6))
        ttk.Button(bf, text="Open dist/",
                   command=self._open_dist).pack(side="left", padx=(0, 6))
        ttk.Button(bf, text="Clear Log",
                   command=self._clear_cc_log).pack(side="right")

        # ── Build log ────────────────────────────────────────────────────────
        lf = ttk.LabelFrame(f, text="Build Log")
        lf.pack(fill="both", expand=True, padx=10, pady=(4, 10))

        self._cc_log = scrolledtext.ScrolledText(
            lf, state="disabled", font=FONT_MONO, wrap="word", height=14)
        self._cc_log.pack(fill="both", expand=True, padx=4, pady=4)
        self._cc_log.tag_config("ok",   foreground=CLR_OK)
        self._cc_log.tag_config("err",  foreground=CLR_ERR)
        self._cc_log.tag_config("warn", foreground=CLR_WARN)
        self._cc_log.tag_config("info", foreground=CLR_INFO)

    # ── CC helpers ────────────────────────────────────────────────────────────

    def _refresh_cc(self):
        try:
            v = _read_cc_version()
            self._cc_cur_lbl.config(text=v)
            self._cc_new_entry.delete(0, "end")
            self._cc_new_entry.insert(0, v)
            self._set_status(f"CommandCenter v{v}")
        except Exception as e:
            self._cc_cur_lbl.config(text="error")
            self._cc_log_write(f"[error] Could not read version: {e}\n", "err")

    def _apply_cc_version(self):
        new_ver = self._cc_new_entry.get().strip()
        if not _validate_version(new_ver):
            messagebox.showerror("Invalid Version",
                                 "Version must be in X.X.X.X format  (e.g. 1.2.3.4)")
            return
        try:
            _update_cc_version(new_ver)
            _update_version_txt(new_ver)
            self._cc_log_write(f"[ok] Version updated → {new_ver}  "
                               f"(CommandCenter.py + version.txt)\n", "ok")
            self._refresh_cc()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self._cc_log_write(f"[error] {e}\n", "err")

    def _compile(self):
        try:
            cmd = _build_pyinstaller_cmd()
        except Exception as e:
            messagebox.showerror("Config Error", str(e))
            return
        self._cc_log_write(f"[info] Starting PyInstaller…\n", "info")
        self._cc_log_write(f"[info] Command: {' '.join(cmd[:5])} …\n", "info")
        self._set_status("Building…")
        threading.Thread(target=self._run_build, args=(cmd,), daemon=True).start()

    def _apply_and_compile(self):
        new_ver = self._cc_new_entry.get().strip()
        if not _validate_version(new_ver):
            messagebox.showerror("Invalid Version",
                                 "Version must be in X.X.X.X format  (e.g. 1.2.3.4)")
            return
        try:
            _update_cc_version(new_ver)
            _update_version_txt(new_ver)
            self._cc_log_write(f"[ok] Version updated → {new_ver}\n", "ok")
            self._refresh_cc()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self._cc_log_write(f"[error] {e}\n", "err")
            return
        self._compile()

    def _run_build(self, cmd: list):
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, bufsize=1, cwd=str(BASE_DIR),
            )
            for line in proc.stdout:
                lo = line.lower()
                if "error" in lo:
                    tag = "err"
                elif "warning" in lo or "warn" in lo:
                    tag = "warn"
                else:
                    tag = None
                self._cc_log_write(line, tag)
            proc.wait()
            if proc.returncode == 0:
                self._cc_log_write("[ok] Build completed successfully.\n", "ok")
                self.after(0, lambda: self._set_status("Build succeeded."))
            else:
                self._cc_log_write(
                    f"[error] Build failed  (exit code {proc.returncode}).\n", "err")
                self.after(0, lambda: self._set_status("Build failed."))
        except FileNotFoundError:
            self._cc_log_write(
                "[error] 'pyinstaller' not found — is it installed and on PATH?\n", "err")
            self.after(0, lambda: self._set_status("pyinstaller not found."))
        except Exception as e:
            self._cc_log_write(f"[error] {e}\n", "err")

    def _open_dist(self):
        dist = BASE_DIR / "dist"
        if not dist.exists():
            messagebox.showinfo("Not Found", "No dist/ folder yet — compile first.")
            return
        os.startfile(str(dist))

    def _clear_cc_log(self):
        self._cc_log.config(state="normal")
        self._cc_log.delete("1.0", "end")
        self._cc_log.config(state="disabled")

    def _cc_log_write(self, text: str, tag=None):
        def _w():
            self._cc_log.config(state="normal")
            if tag:
                self._cc_log.insert("end", text, tag)
            else:
                self._cc_log.insert("end", text)
            self._cc_log.see("end")
            self._cc_log.config(state="disabled")
        self.after(0, _w)

    # ═══════════════════════════════════════════════════════════════════════════
    # Plugins Tab
    # ═══════════════════════════════════════════════════════════════════════════

    def _build_plugin_tab(self):
        f   = self._pl_frame
        PAD = {"padx": 10, "pady": 5}

        # ── Left pane: plugin list ────────────────────────────────────────────
        left = ttk.Frame(f)
        left.pack(side="left", fill="y", padx=(10, 4), pady=10)

        ttk.Label(left, text="Plugins").pack(anchor="w", padx=2)

        list_frame = ttk.Frame(left)
        list_frame.pack(fill="y", expand=True)

        self._pl_listbox = tk.Listbox(
            list_frame, width=26, selectmode="single",
            font=("Segoe UI", 10), activestyle="dotbox",
            relief="solid", borderwidth=1,
        )
        sb = ttk.Scrollbar(list_frame, orient="vertical",
                           command=self._pl_listbox.yview)
        self._pl_listbox.config(yscrollcommand=sb.set)
        self._pl_listbox.pack(side="left", fill="y", expand=True)
        sb.pack(side="right", fill="y")

        self._pl_listbox.bind("<<ListboxSelect>>", self._on_plugin_select)
        ttk.Button(left, text="↻  Refresh List",
                   command=self._refresh_plugins).pack(fill="x", pady=(6, 0))

        # ── Right pane: plugin detail ─────────────────────────────────────────
        right = ttk.Frame(f)
        right.pack(side="left", fill="both", expand=True, padx=(0, 10), pady=10)

        info_f = ttk.LabelFrame(right, text="Plugin")
        info_f.pack(fill="x", pady=(0, 6))

        ttk.Label(info_f, text="Folder:").grid(row=0, column=0, sticky="w", **PAD)
        self._pl_folder_lbl = ttk.Label(info_f, text="—", font=FONT_MONO)
        self._pl_folder_lbl.grid(row=0, column=1, sticky="w", **PAD)

        ttk.Label(info_f, text="Current version:").grid(row=1, column=0, sticky="w", **PAD)
        self._pl_cur_lbl = ttk.Label(info_f, text="—", font=FONT_BOLD)
        self._pl_cur_lbl.grid(row=1, column=1, sticky="w", **PAD)

        ttk.Label(info_f, text="New version:").grid(row=2, column=0, sticky="w", **PAD)
        self._pl_new_entry = ttk.Entry(info_f, width=16, font=FONT_MONO)
        self._pl_new_entry.grid(row=2, column=1, sticky="w", **PAD)

        # ── Action buttons ────────────────────────────────────────────────────
        bf = ttk.Frame(right)
        bf.pack(fill="x", pady=(0, 4))

        ttk.Button(bf, text="Update Version",
                   command=self._update_plugin_ver).pack(side="left", padx=(0, 6))
        ttk.Button(bf, text="Repack .ccplug",
                   command=self._repack_current_plugin).pack(side="left", padx=(0, 6))
        ttk.Button(bf, text="Update Version + Repack",
                   command=self._update_and_repack).pack(side="left")

        # ── Log ──────────────────────────────────────────────────────────────
        log_f = ttk.LabelFrame(right, text="Log")
        log_f.pack(fill="both", expand=True)

        self._pl_log = scrolledtext.ScrolledText(
            log_f, state="disabled", font=FONT_MONO, wrap="word", height=8)
        self._pl_log.pack(fill="both", expand=True, padx=4, pady=4)
        self._pl_log.tag_config("ok",   foreground=CLR_OK)
        self._pl_log.tag_config("err",  foreground=CLR_ERR)
        self._pl_log.tag_config("info", foreground=CLR_INFO)

        self._plugin_data: list = []   # [(Path, display_name, version), ...]

    # ── Plugin helpers ────────────────────────────────────────────────────────

    def _refresh_plugins(self):
        self._plugin_data = _scan_plugins()
        self._pl_listbox.delete(0, "end")
        for _, name, ver in self._plugin_data:
            self._pl_listbox.insert("end", f"  {name}  ({ver})")
        self._pl_folder_lbl.config(text="—")
        self._pl_cur_lbl.config(text="—")
        self._pl_new_entry.delete(0, "end")
        self._set_status(f"{len(self._plugin_data)} plugin(s) found.")

    def _on_plugin_select(self, _event=None):
        sel = self._pl_listbox.curselection()
        if not sel:
            return
        path, name, ver = self._plugin_data[sel[0]]
        self._pl_folder_lbl.config(text=path.name)
        self._pl_cur_lbl.config(text=ver)
        self._pl_new_entry.delete(0, "end")
        self._pl_new_entry.insert(0, ver)
        self._set_status(f"Selected: {name}  v{ver}")

    def _selected_plugin(self):
        sel = self._pl_listbox.curselection()
        if not sel:
            messagebox.showwarning("No Selection", "Select a plugin from the list first.")
            return None
        return self._plugin_data[sel[0]]

    def _update_plugin_ver(self):
        entry = self._selected_plugin()
        if not entry:
            return
        path, name, _ = entry
        new_ver = self._pl_new_entry.get().strip()
        if not _validate_version(new_ver):
            messagebox.showerror("Invalid Version",
                                 "Version must be in X.X.X.X format  (e.g. 1.0.0.5)")
            return
        try:
            _update_plugin_version(path, new_ver)
            self._pl_log_write(f"[ok] {name}  →  v{new_ver}  (manifest.json updated)\n", "ok")
            self._refresh_plugins()
        except Exception as e:
            self._pl_log_write(f"[error] {e}\n", "err")

    def _repack_current_plugin(self):
        entry = self._selected_plugin()
        if not entry:
            return
        path, name, _ = entry
        try:
            out  = _repack_plugin(path)
            size = out.stat().st_size
            self._pl_log_write(
                f"[ok] {name}  →  {out.name}  ({size:,} bytes)\n", "ok")
        except Exception as e:
            self._pl_log_write(f"[error] {e}\n", "err")

    def _update_and_repack(self):
        entry = self._selected_plugin()
        if not entry:
            return
        path, name, _ = entry
        new_ver = self._pl_new_entry.get().strip()
        if not _validate_version(new_ver):
            messagebox.showerror("Invalid Version",
                                 "Version must be in X.X.X.X format  (e.g. 1.0.0.5)")
            return
        try:
            _update_plugin_version(path, new_ver)
            out  = _repack_plugin(path)
            size = out.stat().st_size
            self._pl_log_write(
                f"[ok] {name}  →  v{new_ver}  |  {out.name}  ({size:,} bytes)\n", "ok")
            self._refresh_plugins()
        except Exception as e:
            self._pl_log_write(f"[error] {e}\n", "err")

    def _pl_log_write(self, text: str, tag=None):
        self._pl_log.config(state="normal")
        if tag:
            self._pl_log.insert("end", text, tag)
        else:
            self._pl_log.insert("end", text)
        self._pl_log.see("end")
        self._pl_log.config(state="disabled")

    # ═══════════════════════════════════════════════════════════════════════════
    # Shared
    # ═══════════════════════════════════════════════════════════════════════════

    def _set_status(self, msg: str):
        self._status_var.set(msg)

    def _refresh_all(self):
        self._refresh_cc()
        self._refresh_plugins()


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = CCManager()
    app.mainloop()
