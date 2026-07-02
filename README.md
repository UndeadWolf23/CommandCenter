## v1.6.0.0 — July 2, 2026

CommandCenter v1.6.0.0 Release Notes
===================================

Release Date: July 2, 2026

NEW FEATURES
------------
Operations Board
- Added the new Operations Board system.
- A full node graph environment for activity tracking, planning,
  organization, and workflow automation.
- Supports interconnected nodes for managing projects, tasks,
  notes, reminders, timers, and operational workflows.
- Includes a node menu on the left side for creating and managing
  nodes within a board.
- Includes an editor panel on the right side for configuring node
  properties, settings, and behavior.
- Supports multiple boards simultaneously through a tabbed interface,
  allowing users to switch between separate workflows, projects,
  and activity spaces.
- Integrates directly with existing CommandCenter systems including
  Notebook, Time Tracker, and Reminders.
- Designed to serve as the central command and organization hub for
  CommandCenter.

Reminders Integration
- The Reminders system has been moved into the Operations Board.
- Reminders can now exist as nodes within workflows and planning
  boards for improved organization and automation.

Notebook Improvements
- Added support for bulleted tabbing and nested bullet lists.

- Improved window scaling method and added diagonal scaling support.

GEMS & PROGRESSION
------------------
- Added 5 new Gem Loadout slots.
- Gem Loadout slots can now be purchased using Gem Dust.


- Updated Achievements system:
  - Removed outdated achievements.
  - Added new achievements to support updated systems and progression
    paths.


BUG FIXES
---------
- Fixed a memory leak issue that could cause increased memory
  consumption during extended usage sessions.
- Fixed multiple GEM-related issues.

---

## v1.5.1.0 — June 26, 2026

Command Center 1.5.1.0 Patch Notes
----------------------------------------------
New Features
Added 7 new Gems to discover and upgrade.
Added 2 new Gem Set Bonuses to expand build options and synergies.
Added a new Gem Set Bonuses section to the Gem Codex for learning how set bonuses work and how to activate them.
Added a new Gem Bonus Tracker to the Gem Menu to display currently active bonuses.
Added new animated gem tiles for Max Rank Gems.
Added several new animations throughout the application.
Added a Chest Luck Bonus Tracker.
Added a Chest Unlock Time Bonus Tracker.
Added a new "Smaller Pets" option in the Appearance Settings.
Pets now appear 20% larger by default for improved visibility and presence.
----------------------------------------------
New Feature
Added Notebook Templates for faster note creation and improved organization.
----------------------------------------------
Reworked
Reworked and rebalanced several existing Gems.
----------------------------------------------
Bug Fixes
Fixed several bugs related to Gems and Gem interactions.
General stability improvements and miscellaneous bug fixes.
----------------------------------------------
Secrets and Rewards
Added a new Secret Pet.
Added a new Achievement.
----------------------------------------------
Coming in Version 1.6.0.0

The next major update introduces the Activity Tracker and To-Do List systems.

These new features will integrate directly with:

Time Tracking
Reminders
Notebook Features

The goal is to continue evolving Command Center into the ultimate workflow machine.

---

## v1.5.0.4 — June 24, 2026

bug fixes and tweaks

---

## v1.5.0.2 — June 23, 2026

Bug fixes

---

## v1.5.0.1 — June 22, 2026

bug fixes related to gems and the gem menu

---

## v1.5.0.0 — June 22, 2026

Introducing Gems
--------------------------------
version 1.5.0.0 brings a complete Gem system to the profile menu. As time passes you will unlock gem chests of increasing rarity. Opening these chests give you gems that you can equip in the new Gem Menu. Gems modify the XP and pet systems allowing users to strategically equip gems to maximize XP gain. Gems can be leveled up with Gem Dust to provide stronger effects. Gem also have a Star Rank that can be increased by merging in extra copies of the same gem. Gems can also be converted to Gem Dust by salvaging unwanted gems. You can configure multiple gem loadouts to quickly switch between your favorite gem sets.

Check out the Gem Codex in the Gem Menu for a full list of all 48 Gems currently in the app. As you discover new gems they will brighten in the codex.

Gems are tradeable! Have a Legendary or Mythic gem that you don't need? No problem. export the encrypted gem file and send it to a buddy. They will always be able to see your name on the gem. Once you export a gem, it can no longer be imported or used again on your profile to prevent duplication and reuse.

--------------------------------
Squashed more bugs in the notebook and profile menus
--------------------------------

what's next?

currently in the works is a full To Do List and activity tracker that will combine with the reminders tool and the time tracker tool to create a full workflow tracking system.

Also in the works is Notebook Templates to help users quickly create a fresh templated note reducing time spent on repeated actions and note styles.

---

## v1.4.2.3 — June 15, 2026

bug fixes related to new installer and separation of assets

---

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
