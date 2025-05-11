# LaunchBox Shortcut Generator

This PowerShell script scans a directory of extracted PC games and automatically generates `.lnk` shortcuts for use with LaunchBox.

## 🚀 Features
- Scans deeply nested folders (configurable depth)
- Scores `.exe` files to find the correct launcher
- Handles duplicates, naming issues, and timeouts
- Logs results: found, not found, skipped, errors

## ⚙️ Configuration
Edit the script variables:
```
$rootGameFolder = "D:\Arcade\System roms\PC Games 2"
$shortcutOutputFolder = "$rootGameFolder\Shortcuts2"
$maxDepth = 5
$folderTimeout = 60
```

## 🛠 Usage
1. Open PowerShell as Administrator
2. Run:
```
.\Generate-LaunchBoxShortcuts.ps1
```
3. Logs will be generated in the output folder

## 📂 Logs
- `FoundGames.log`
- `NotFoundGames.log`
- `SkippedGames.log`
- `ErrorGames.log`

## 📜 License
MIT
