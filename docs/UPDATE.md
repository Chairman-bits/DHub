# GitHub自動アップデート設定

1. `tools\Set-Version.bat` でバージョンを更新
2. `tools\Build-EXE-ONLY.bat` で exe を作成
3. GitHub Releases に `ShortcutList.exe` をアップロード
4. `release-template/version.json` を編集して公開
5. アプリ設定の `UpdateVersionUrl` を GitHub Raw の version.json URL に設定

version.json 例:

```json
{
  "version": "1.0.1",
  "downloadUrl": "https://github.com/YOUR_GITHUB_USER/YOUR_REPOSITORY/releases/download/v1.0.1/ShortcutList.exe",
  "releaseNotes": "Update notes here."
}
```
