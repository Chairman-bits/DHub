DHub

Build:
tools\Build-EXE-ONLY.bat

Distribution:
release\DHub.exe only

New features:
- Ctrl+Space global hotkey opens Spotlight-style quick search.
- Quick search supports incremental search and Enter launch.
- Tags/groups have colors.
- Favorite shortcuts are supported from the edit screen, toolbar, right-click menu, quick search, and tray menu.
- Each shortcut can specify the application used to open it.
- Folders default to explorer.exe as the opening application.
- Right-click context menu supports open, open location, favorite toggle, opening application selection, edit, delete, and path copy.
- Launch count and last launch date are recorded.
- Header sorting is supported.
- GitHub version check and self-update are supported.
- Shortcut export/import for PC replacement.
- Settings export/import.
- Column visibility settings.
- Broken link check.
- Duplicate cleanup.
- Tray quick launch.

GitHub update setup:
1. Build release\DHub.exe.
2. Upload DHub.zip and DHubUpdater.zip to the GitHub raw/release location used by version.json.
3. Put release-template/version.json in your repository, or change UpdateVersionUrl in the stored settings JSON.
4. Update version, downloadUrl, updaterUrl, and releaseNotes.

version.json format:
{
  "version": "1.0.2",
  "downloadUrl": "https://raw.githubusercontent.com/Chairman-bits/DHub/main/DHub.zip",
  "updaterUrl": "https://raw.githubusercontent.com/Chairman-bits/DHub/main/DHubUpdater.zip",
  "releaseNotes": "Update notes here."
}

Data location:
%LocalAppData%\DHub
