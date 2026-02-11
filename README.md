# QuickPaths

Desktop floating widget for quick path copying. Click the dot, click a path name, done.

## Features
- Always-on-top draggable dot
- Click to expand path list, click path name to copy
- Add folders with multi-select (Ctrl+click)
- Reorder with move-up button
- Position persists across restarts
- Auto-start on boot + watchdog auto-restart
- Multi-monitor support

## Requirements
- Windows 10/11
- PowerShell 5.1 (built-in)

## Install
1. Clone or copy to any directory
2. Run `QuickPaths.ps1` (or double-click `QuickPaths.vbs`)
3. Optional: run `setup_watchdog.ps1` for auto-restart
4. Optional: copy `QuickPaths.vbs` to `shell:startup` for boot auto-start
