ShortcutList

Build:
tools\Build-EXE-ONLY.bat

Distribution:
release\ShortcutList.exe only

New features:
- Ctrl+Space global hotkey opens Spotlight-style quick search.
- Quick search supports incremental search and Enter launch.
- Tags/groups have colors.
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
1. Build release\ShortcutList.exe.
2. Upload ShortcutList.exe to GitHub Releases.
3. Put release/version.json in your repository, or change UpdateVersionUrl in the stored settings JSON.
4. version.json example is in release-template/version.json.

version.json format:
{
  "version": "1.0.1",
  "downloadUrl": "https://github.com/YOUR_GITHUB_USER/YOUR_REPOSITORY/releases/download/v1.0.1/ShortcutList.exe",
  "releaseNotes": "Update notes here."
}

Data location:
%LocalAppData%\ShortcutList
