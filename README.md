## v1.4.2.2 — June 15, 2026

Separation of assets from the main executable to a separate asset cache. This provides for much faster installs and updates to the app. Changes were made to the App, Installer, Uninstaller, and Updater executables.

---

## v1.4.2.1 — June 11, 2026

bug fix for level up animation - added support for animation queue

---

## v1.4.2.0 — June 11, 2026

Upgraded the footer in the main window to be more efficient and clean. Plugins and Windows Tools now sort into dynamic trays that expand when you hover over them. This allows for more real estate as the app continues to grow and as users install more plugins. 

Added several new Windows Tools to help users navigate windows very quickly and efficiently.

Added a Level Up animation that plays when users level up their profiles. 

The profile and customization systems are still fresh and experimental - please feel free to send me any feedback on changes or additions to these systems you'd like to see. You can use the feedback button in the bottom right to queue up a feedback email to me.

Thank you

---

## v1.4.1.1 — June 5, 2026

added new formatting buttons to the notebook tool: indent, change case, and sort notes by.

squashed a few small bugs.

---

## v1.4.1.0 — June 4, 2026

added a little secret...

check out the profile page for a fun challenge and a secret reward

---

## v1.4.0.1 — June 3, 2026

Added animations for Panda, Axolotl, and Dragon pets.

---

## v1.4.0.0 — June 2, 2026

Introducing experimental new features such as: Profiles, Levels & XP, Pets, and Achievements.

New improvements to the Notebook tool such as: Find,  Find & Replace (similar to Notepad++), and List Comparison tools. Additionally, a settings page has been introduced for the notebook to support a few new options and future options. Users can now set a default paste size for images to avoid having to set a size on every paste, as well as a few other small options.

The folder system has been overhauled to include support for nested folders and allowance for nodes of varying size inside folders. Nested folder navigation can be tricky. This system will continue to be improved as time goes on.

This version also implements a few small bug fixes and optimizations, although with the addition of many new features, please remain patient and send constructive feedback so I can continue to improve the tool.

---

## v1.3.0.6 — June 1, 2026

Fixed a bug with the usage of plugin hotkeys when the command center window is out of focus.

---

# Command Center

**Version 1.3.0.0**  
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

## What's New in 1.3.0.0

### Auto-Update System
Command Center now connects to GitHub Releases to check for and install updates automatically. The version number in the bottom-right of the footer bar is now a **clickable button** that opens the Version Info & Updates dialog.

### Version Info & Updates Dialog
Click the version label (e.g. `v1.3.0.0`) in the footer to open this dialog. It displays the full release notes fetched from GitHub and provides a **Check for Updates** button. If a newer version is available, an **Update Now** button appears — clicking it closes the app, downloads and replaces the executable files, and relaunches Command Center automatically. Your data (nodes, settings, notebooks, plugins) is never touched.

### Auto-Check on Startup
Enable **Auto-check for updates on startup** in the Version Info dialog to have Command Center silently check GitHub Releases 2–3 seconds after every launch. If an update is found, a prompt appears in-app.

### Updater
A lightweight companion process `updater.exe` handles all update operations independently. It downloads the new `CommandCenter.exe` and `updater.exe` from GitHub Releases, waits for the main app to exit, replaces the files on disk, and relaunches the app automatically.

### Installer
New users can use `installer.exe` for first-time setup. It downloads `CommandCenter.exe`, `updater.exe`, and the README to a folder you choose (default: your Desktop), and can optionally create a desktop shortcut.

---

## Previous Versions

### 1.2.6.x
- **Paste as Plain Text (`Ctrl+Shift+V`)** — global system-wide hotkey that strips formatting before pasting, works in any application
- **New Logo** — redesigned app icon and title bar logo
- **Animated Startup Screen** — custom sprite animation on the splash screen
