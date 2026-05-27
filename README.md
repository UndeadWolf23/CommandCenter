# Command Center

**Version 1.2.6.2**  
A GUI-based workflow hub built with Python + PySide6 for Windows.

https://github.com/UndeadWolf23/CommandCenter/releases/tag/CommandCenter

---

## What It Does

Command Center is a customizable desktop dashboard that keeps your most-used tools, links, files, and notes one click away. Instead of hunting through the taskbar or file explorer, everything lives on a tile grid you arrange yourself.

**Core features:**

- **Node tile grid** — drag-and-drop tiles that open files, folders, URLs, or notebook notes. Tiles come in multiple sizes (1×1, 2×2, 2×4, and more).
- **Clipboard Manager** — tracks clipboard history (text, images, files) with search and one-click re-copy. Automatically skips password-like entries.
- **Notebook** — a rich-text note editor with spell check, image embedding, colored notes, and a media library for GIFs and pictures.
- **Time Tracker** — a compact HUD that floats over your screen and tracks time, pinned to the bottom-right corner.
- **Quick Connect** — fast-access launcher for ScreenConnect remote sessions with auto-complete client search.
- **Plugin system** — extend functionality with `.ccplug` plugins (included: Caffeine, AutoHotKey bridge, Credential Manager, CySec tools).
- **Themes** — built-in themes: Deep Space, Midnight Blue, Forest Night, Crimson Dark, Dark Knight, Spooky, Noir, Slate Light, and Custom.
- **Custom cursors** — optional themed cursor overlays.
- **Windows startup** — optional auto-launch on login via the registry.

---

## Requirements

- Windows 10/11

---

## Running

Run `CommandCenter.exe`. No installation required — all dependencies and assets are bundled in the executable.

---

## What's New in 1.2.6.2

### Paste as Plain Text — `Ctrl+Shift+V`
A global system-wide hotkey that works in **any application**, not just Command Center. When you press `Ctrl+Shift+V`, the current clipboard contents are stripped of all formatting (HTML, rich text, etc.) and pasted as plain text. Useful for pasting into email composers, Word documents, or any app that normally carries over unwanted formatting.

### New Logo
The app icon and title bar logo have been updated with a redesigned logo

### Animated Startup Screen
The optional splash screen that appears on launch now plays a custom animation.
