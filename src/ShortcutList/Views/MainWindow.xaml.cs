using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using Forms = System.Windows.Forms;
using Drawing = System.Drawing;
using WpfApplication = System.Windows.Application;
using WpfClipboard = System.Windows.Clipboard;
using ShortcutList.Models;
using ShortcutList.Services;

namespace ShortcutList.Views;

public partial class MainWindow : Window
{
    private const int HotkeyId = 0x534C;
    private const uint ModControl = 0x0002;
    private const uint VkSpace = 0x20;
    private const int WmHotkey = 0x0312;

    private const string AllFilter = "すべて";
    private const string FolderFilter = "フォルダのみ";
    private const string UrlFilter = "URLのみ";
    private const string FileFilter = "ファイルのみ";
    private const string BrokenFilter = "未検出のみ";

    private readonly ShortcutStore _store = new();
    private readonly ObservableCollection<ShortcutItem> _visibleItems = new();
    private readonly UpdateService _updateService = new();

    private Forms.NotifyIcon? _notifyIcon;
    private QuickSearchWindow? _quickSearchWindow;
    private HwndSource? _source;
    private string _sortKey = "Manual";
    private bool _sortAscending = true;

    private const double NameColumnDefaultWidth = 240;
    private const double PathColumnDefaultWidth = 460;
    private const double ArgumentsColumnDefaultWidth = 150;
    private const double GroupColumnDefaultWidth = 120;
    private const double LaunchCountColumnDefaultWidth = 90;
    private const double LastLaunchColumnDefaultWidth = 150;
    private const double CreatedAtColumnDefaultWidth = 150;
    private const double UpdatedAtColumnDefaultWidth = 150;
    private const double StatusColumnDefaultWidth = 90;

    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    public MainWindow()
    {
        InitializeComponent();

        ShortcutListView.ItemsSource = _visibleItems;

        _store.Load();

        InitializeFilter();
        ApplyColumnVisibility();
        ApplyFilter();
        InitializeTrayIcon();

        VersionTextBlock.Text = $"v{_updateService.CurrentVersion}";
    }

    private ShortcutItem? SelectedItem => ShortcutListView.SelectedItem as ShortcutItem;

    private async void Window_Loaded(object sender, RoutedEventArgs e)
    {
        SearchTextBox.Focus();
        Keyboard.Focus(SearchTextBox);

        if (_store.Settings.StartMinimizedToTray)
        {
            Hide();
        }

        if (_store.Settings.CheckUpdateOnStartup)
        {
            await CheckUpdateAsync(silent: true);
        }
    }

    private void Window_SourceInitialized(object sender, EventArgs e)
    {
        var helper = new WindowInteropHelper(this);
        _source = HwndSource.FromHwnd(helper.Handle);
        _source?.AddHook(WndProc);
        RegisterHotKey(helper.Handle, HotkeyId, ModControl, VkSpace);
    }

    private void Window_Closed(object? sender, EventArgs e)
    {
        var helper = new WindowInteropHelper(this);
        UnregisterHotKey(helper.Handle, HotkeyId);
        _source?.RemoveHook(WndProc);
        _notifyIcon?.Dispose();
        _notifyIcon = null;
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WmHotkey && wParam.ToInt32() == HotkeyId)
        {
            ShowQuickSearch();
            handled = true;
        }

        return IntPtr.Zero;
    }

    private void InitializeFilter()
    {
        FilterComboBox.Items.Clear();
        FilterComboBox.Items.Add(AllFilter);
        FilterComboBox.Items.Add(FolderFilter);
        FilterComboBox.Items.Add(UrlFilter);
        FilterComboBox.Items.Add(FileFilter);
        FilterComboBox.Items.Add(BrokenFilter);

        foreach (var group in _store.GetGroups())
        {
            FilterComboBox.Items.Add("タグ: " + group);
        }

        FilterComboBox.SelectedItem = AllFilter;
    }

    private void RefreshFilterItemsKeepSelection()
    {
        var current = FilterComboBox.SelectedItem?.ToString() ?? AllFilter;

        FilterComboBox.SelectionChanged -= FilterComboBox_SelectionChanged;
        InitializeFilter();

        if (FilterComboBox.Items.Contains(current))
        {
            FilterComboBox.SelectedItem = current;
        }

        FilterComboBox.SelectionChanged += FilterComboBox_SelectionChanged;
    }

    private void ApplyFilter()
    {
        var selectedId = SelectedItem?.Id;
        var keyword = SearchTextBox.Text?.Trim() ?? string.Empty;

        var words = keyword
            .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(x => x.ToLowerInvariant())
            .ToList();

        var filter = FilterComboBox.SelectedItem?.ToString() ?? AllFilter;

        IEnumerable<ShortcutItem> items = _store.Items;

        if (filter == FolderFilter)
        {
            items = items.Where(x => x.ShortcutType == ShortcutType.Folder);
        }
        else if (filter == UrlFilter)
        {
            items = items.Where(x => x.ShortcutType == ShortcutType.Url);
        }
        else if (filter == FileFilter)
        {
            items = items.Where(x => x.ShortcutType == ShortcutType.File);
        }
        else if (filter == BrokenFilter)
        {
            items = items.Where(x => x.IsBroken);
        }
        else if (filter.StartsWith("タグ: ", StringComparison.Ordinal))
        {
            var group = filter.Replace("タグ: ", string.Empty);
            items = items.Where(x => string.Equals(x.GroupDisplay, group, StringComparison.OrdinalIgnoreCase));
        }

        if (words.Count > 0)
        {
            items = items.Where(x => words.All(word => x.SearchText.Contains(word)));
        }

        items = ApplySort(items);

        _visibleItems.Clear();

        foreach (var item in items)
        {
            _visibleItems.Add(item);
        }

        if (selectedId is not null)
        {
            var selected = _visibleItems.FirstOrDefault(x => x.Id == selectedId);
            if (selected is not null)
            {
                ShortcutListView.SelectedItem = selected;
            }
        }

        CountTextBlock.Text = $"{_visibleItems.Count} 個のショートカット";
        HealthStatusTextBlock.Text = $"未検出 {_store.Items.Count(x => x.IsBroken)} 件";
        RefreshSortHeader();
        RefreshTrayMenu();
    }

    private IEnumerable<ShortcutItem> ApplySort(IEnumerable<ShortcutItem> items)
    {
        return _sortKey switch
        {
            "Name" => _sortAscending
                ? items.OrderBy(x => x.Name, StringComparer.CurrentCultureIgnoreCase)
                : items.OrderByDescending(x => x.Name, StringComparer.CurrentCultureIgnoreCase),

            "Path" => _sortAscending
                ? items.OrderBy(x => x.TargetPath, StringComparer.CurrentCultureIgnoreCase)
                : items.OrderByDescending(x => x.TargetPath, StringComparer.CurrentCultureIgnoreCase),

            "Arguments" => _sortAscending
                ? items.OrderBy(x => x.Arguments, StringComparer.CurrentCultureIgnoreCase)
                : items.OrderByDescending(x => x.Arguments, StringComparer.CurrentCultureIgnoreCase),

            "Group" => _sortAscending
                ? items.OrderBy(x => x.GroupDisplay, StringComparer.CurrentCultureIgnoreCase)
                : items.OrderByDescending(x => x.GroupDisplay, StringComparer.CurrentCultureIgnoreCase),

            "LaunchCount" => _sortAscending
                ? items.OrderBy(x => x.LaunchCount)
                : items.OrderByDescending(x => x.LaunchCount),

            "LastLaunch" => _sortAscending
                ? items.OrderBy(x => x.LastLaunchedAt ?? DateTime.MinValue)
                : items.OrderByDescending(x => x.LastLaunchedAt ?? DateTime.MinValue),

            "CreatedAt" => _sortAscending
                ? items.OrderBy(x => x.CreatedAt)
                : items.OrderByDescending(x => x.CreatedAt),

            "UpdatedAt" => _sortAscending
                ? items.OrderBy(x => x.UpdatedAt)
                : items.OrderByDescending(x => x.UpdatedAt),

            "Status" => _sortAscending
                ? items.OrderBy(x => x.StatusText, StringComparer.CurrentCultureIgnoreCase)
                : items.OrderByDescending(x => x.StatusText, StringComparer.CurrentCultureIgnoreCase),

            _ => items.OrderBy(x => _store.Items.IndexOf(x))
        };
    }

    private void GridViewColumnHeader_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not GridViewColumnHeader header || header.Column is null)
        {
            return;
        }

        var key = GetSortKeyFromHeader(header.Column.Header?.ToString() ?? string.Empty);

        if (string.IsNullOrWhiteSpace(key))
        {
            return;
        }

        if (_sortKey == key)
        {
            _sortAscending = !_sortAscending;
        }
        else
        {
            _sortKey = key;
            _sortAscending = true;
        }

        ApplyFilter();
    }

    private static string GetSortKeyFromHeader(string headerText)
    {
        var clean = headerText.Replace(" ▲", string.Empty).Replace(" ▼", string.Empty);

        return clean switch
        {
            "名前" => "Name",
            "パス" => "Path",
            "引数" => "Arguments",
            "タグ" => "Group",
            "起動回数" => "LaunchCount",
            "最終起動" => "LastLaunch",
            "作成日" => "CreatedAt",
            "更新日" => "UpdatedAt",
            "状態" => "Status",
            _ => string.Empty
        };
    }

    private void RefreshSortHeader()
    {
        NameColumn.Header = "名前";
        PathColumn.Header = "パス";
        ArgumentsColumn.Header = "引数";
        GroupColumn.Header = "タグ";
        LaunchCountColumn.Header = "起動回数";
        LastLaunchColumn.Header = "最終起動";
        CreatedAtColumn.Header = "作成日";
        UpdatedAtColumn.Header = "更新日";
        StatusColumn.Header = "状態";

        var arrow = _sortAscending ? " ▲" : " ▼";
        var label = "手動";

        switch (_sortKey)
        {
            case "Name":
                NameColumn.Header = "名前" + arrow;
                label = "名前";
                break;
            case "Path":
                PathColumn.Header = "パス" + arrow;
                label = "パス";
                break;
            case "Arguments":
                ArgumentsColumn.Header = "引数" + arrow;
                label = "引数";
                break;
            case "Group":
                GroupColumn.Header = "タグ" + arrow;
                label = "タグ";
                break;
            case "LaunchCount":
                LaunchCountColumn.Header = "起動回数" + arrow;
                label = "起動回数";
                break;
            case "LastLaunch":
                LastLaunchColumn.Header = "最終起動" + arrow;
                label = "最終起動";
                break;
            case "CreatedAt":
                CreatedAtColumn.Header = "作成日" + arrow;
                label = "作成日";
                break;
            case "UpdatedAt":
                UpdatedAtColumn.Header = "更新日" + arrow;
                label = "更新日";
                break;
            case "Status":
                StatusColumn.Header = "状態" + arrow;
                label = "状態";
                break;
        }

        SortStatusTextBlock.Text = _sortKey == "Manual"
            ? "並び順: 手動"
            : $"並び順: {label} {(_sortAscending ? "昇順" : "降順")}";
    }

    private void ApplyColumnVisibility()
    {
        NameColumn.Width = _store.Settings.ShowNameColumn ? NameColumnDefaultWidth : 0;
        PathColumn.Width = _store.Settings.ShowPathColumn ? PathColumnDefaultWidth : 0;
        ArgumentsColumn.Width = _store.Settings.ShowArgumentsColumn ? ArgumentsColumnDefaultWidth : 0;
        GroupColumn.Width = _store.Settings.ShowGroupColumn ? GroupColumnDefaultWidth : 0;
        LaunchCountColumn.Width = _store.Settings.ShowLaunchCountColumn ? LaunchCountColumnDefaultWidth : 0;
        LastLaunchColumn.Width = _store.Settings.ShowLastLaunchColumn ? LastLaunchColumnDefaultWidth : 0;
        CreatedAtColumn.Width = _store.Settings.ShowCreatedAtColumn ? CreatedAtColumnDefaultWidth : 0;
        UpdatedAtColumn.Width = _store.Settings.ShowUpdatedAtColumn ? UpdatedAtColumnDefaultWidth : 0;
        StatusColumn.Width = _store.Settings.ShowStatusColumn ? StatusColumnDefaultWidth : 0;
    }

    private void ColumnSettingsButton_Click(object sender, RoutedEventArgs e)
    {
        var menu = new ContextMenu { StaysOpen = true };

        AddColumnMenuItem(menu, "名前", _store.Settings.ShowNameColumn, value => _store.Settings.ShowNameColumn = value);
        AddColumnMenuItem(menu, "パス", _store.Settings.ShowPathColumn, value => _store.Settings.ShowPathColumn = value);
        AddColumnMenuItem(menu, "引数", _store.Settings.ShowArgumentsColumn, value => _store.Settings.ShowArgumentsColumn = value);
        AddColumnMenuItem(menu, "タグ", _store.Settings.ShowGroupColumn, value => _store.Settings.ShowGroupColumn = value);
        AddColumnMenuItem(menu, "起動回数", _store.Settings.ShowLaunchCountColumn, value => _store.Settings.ShowLaunchCountColumn = value);
        AddColumnMenuItem(menu, "最終起動", _store.Settings.ShowLastLaunchColumn, value => _store.Settings.ShowLastLaunchColumn = value);
        AddColumnMenuItem(menu, "作成日", _store.Settings.ShowCreatedAtColumn, value => _store.Settings.ShowCreatedAtColumn = value);
        AddColumnMenuItem(menu, "更新日", _store.Settings.ShowUpdatedAtColumn, value => _store.Settings.ShowUpdatedAtColumn = value);
        AddColumnMenuItem(menu, "状態", _store.Settings.ShowStatusColumn, value => _store.Settings.ShowStatusColumn = value);

        menu.IsOpen = true;
    }

    private void AddColumnMenuItem(ContextMenu menu, string header, bool isChecked, Action<bool> setter)
    {
        var item = new MenuItem
        {
            Header = header,
            IsCheckable = true,
            IsChecked = isChecked,
            StaysOpenOnClick = true
        };

        item.Click += (_, _) =>
        {
            setter(item.IsChecked);
            ApplyColumnVisibility();
            _store.Save();
        };

        menu.Items.Add(item);
    }

    private void SettingsButton_Click(object sender, RoutedEventArgs e)
    {
        var menu = new ContextMenu();

        AddMenuItem(menu, "ショートカットをエクスポート", ExportShortcuts);
        AddMenuItem(menu, "ショートカットをインポート", ImportShortcuts);
        menu.Items.Add(new Separator());
        AddMenuItem(menu, "設定をエクスポート", ExportSettings);
        AddMenuItem(menu, "設定をインポート", ImportSettings);
        menu.Items.Add(new Separator());
        AddMenuItem(menu, "グループ管理", ManageGroups);
        AddMenuItem(menu, "重複を整理", CleanupDuplicates);
        AddMenuItem(menu, "データ保存場所を開く", OpenStorageFolder);
        menu.Items.Add(new Separator());
        AddMenuItem(menu, "更新を確認", () => _ = CheckUpdateAsync(false));

        menu.IsOpen = true;
    }

    private static void AddMenuItem(ContextMenu menu, string header, Action action)
    {
        var item = new MenuItem { Header = header };
        item.Click += (_, _) => action();
        menu.Items.Add(item);
    }

    private void AddButton_Click(object sender, RoutedEventArgs e)
    {
        OpenAddDialog();
    }

    private void OpenAddDialog()
    {
        var dialog = new ShortcutEditWindow { Owner = this };

        if (dialog.ShowDialog() != true || dialog.ResultItem is null)
        {
            return;
        }

        if (_store.ContainsTarget(dialog.ResultItem.TargetPath))
        {
            System.Windows.MessageBox.Show(
                "同じパスまたはURLがすでに登録されています。",
                "DHub",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
            return;
        }

        _store.Add(dialog.ResultItem);
        RefreshFilterItemsKeepSelection();
        ApplyFilter();
    }

    private void EditButton_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedItem is null)
        {
            return;
        }

        var dialog = new ShortcutEditWindow(SelectedItem) { Owner = this };

        if (dialog.ShowDialog() != true || dialog.ResultItem is null)
        {
            return;
        }

        if (_store.ContainsTarget(dialog.ResultItem.TargetPath, SelectedItem.Id))
        {
            System.Windows.MessageBox.Show(
                "同じパスまたはURLがすでに登録されています。",
                "DHub",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
            return;
        }

        SelectedItem.Name = dialog.ResultItem.Name;
        SelectedItem.TargetPath = dialog.ResultItem.TargetPath;
        SelectedItem.Arguments = dialog.ResultItem.Arguments;
        SelectedItem.ShortcutType = dialog.ResultItem.ShortcutType;
        SelectedItem.GroupName = dialog.ResultItem.GroupName;
        SelectedItem.TagColor = dialog.ResultItem.TagColor;
        SelectedItem.TouchUpdated();

        _store.Save();
        RefreshFilterItemsKeepSelection();
        ApplyFilter();
    }

    private void DeleteButton_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedItem is null)
        {
            return;
        }

        if (System.Windows.MessageBox.Show(
                $"「{SelectedItem.Name}」を削除しますか？",
                "DHub",
                System.Windows.MessageBoxButton.YesNo,
                System.Windows.MessageBoxImage.Question) != System.Windows.MessageBoxResult.Yes)
        {
            return;
        }

        _store.Remove(SelectedItem);
        RefreshFilterItemsKeepSelection();
        ApplyFilter();
    }

    private void BackupButton_Click(object sender, RoutedEventArgs e)
    {
        ExportShortcuts();
    }

    private void ExportShortcuts()
    {
        using var dialog = new Forms.SaveFileDialog
        {
            Title = "ショートカットをエクスポート",
            Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
            FileName = $"ShortcutList_Shortcuts_{DateTime.Now:yyyyMMdd_HHmmss}.json"
        };

        if (dialog.ShowDialog() != Forms.DialogResult.OK)
        {
            return;
        }

        _store.ExportShortcutsToFile(dialog.FileName);
        ShowInfo("ショートカットをエクスポートしました。");
    }

    private void ImportShortcuts()
    {
        using var dialog = new Forms.OpenFileDialog
        {
            Title = "ショートカットをインポート",
            Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() != Forms.DialogResult.OK)
        {
            return;
        }

        if (AskYesNo("現在のショートカット一覧をインポートした内容で置き換えます。続行しますか？") != true)
        {
            return;
        }

        _store.ImportShortcutsFromFile(dialog.FileName);
        RefreshFilterItemsKeepSelection();
        ApplyFilter();
        ShowInfo("ショートカットをインポートしました。");
    }

    private void ExportSettings()
    {
        using var dialog = new Forms.SaveFileDialog
        {
            Title = "設定をエクスポート",
            Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
            FileName = $"ShortcutList_Settings_{DateTime.Now:yyyyMMdd_HHmmss}.json"
        };

        if (dialog.ShowDialog() != Forms.DialogResult.OK)
        {
            return;
        }

        _store.ExportSettingsToFile(dialog.FileName);
        ShowInfo("設定をエクスポートしました。");
    }

    private void ImportSettings()
    {
        using var dialog = new Forms.OpenFileDialog
        {
            Title = "設定をインポート",
            Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() != Forms.DialogResult.OK)
        {
            return;
        }

        if (AskYesNo("現在の表示設定・起動設定をインポートした設定で置き換えます。続行しますか？") != true)
        {
            return;
        }

        _store.ImportSettingsFromFile(dialog.FileName);
        ApplyColumnVisibility();
        RefreshTrayMenu();
        ShowInfo("設定をインポートしました。");
    }

    private void ManageGroups()
    {
        var groups = _store.GetGroups().Where(x => x != "未分類").ToList();
        if (groups.Count == 0)
        {
            ShowInfo("管理できるグループはありません。");
            return;
        }

        using var selectDialog = new Forms.Form
        {
            Text = "グループ管理",
            Width = 360,
            Height = 180,
            StartPosition = Forms.FormStartPosition.CenterScreen
        };

        var combo = new Forms.ComboBox { Left = 16, Top = 16, Width = 300, DropDownStyle = Forms.ComboBoxStyle.DropDownList };
        combo.Items.AddRange(groups.Cast<object>().ToArray());
        combo.SelectedIndex = 0;

        var rename = new Forms.Button { Text = "名前変更", Left = 16, Top = 60, Width = 100 };
        var delete = new Forms.Button { Text = "削除", Left = 126, Top = 60, Width = 100 };
        var close = new Forms.Button { Text = "閉じる", Left = 236, Top = 60, Width = 80 };

        rename.Click += (_, _) =>
        {
            var oldName = combo.SelectedItem?.ToString() ?? "";
            var input = Microsoft.VisualBasic.Interaction.InputBox("新しいグループ名を入力してください。", "グループ名変更", oldName);
            if (string.IsNullOrWhiteSpace(input)) return;
            _store.RenameGroup(oldName, input.Trim());
            selectDialog.DialogResult = Forms.DialogResult.OK;
            selectDialog.Close();
        };

        delete.Click += (_, _) =>
        {
            var group = combo.SelectedItem?.ToString() ?? "";
            if (Forms.MessageBox.Show($"「{group}」を未分類に戻しますか？", "DHub", Forms.MessageBoxButtons.YesNo) != Forms.DialogResult.Yes) return;
            _store.DeleteGroupName(group);
            selectDialog.DialogResult = Forms.DialogResult.OK;
            selectDialog.Close();
        };

        close.Click += (_, _) => selectDialog.Close();

        selectDialog.Controls.Add(combo);
        selectDialog.Controls.Add(rename);
        selectDialog.Controls.Add(delete);
        selectDialog.Controls.Add(close);

        if (selectDialog.ShowDialog() == Forms.DialogResult.OK)
        {
            RefreshFilterItemsKeepSelection();
            ApplyFilter();
        }
    }

    private void CleanupDuplicates()
    {
        var groups = _store.GetDuplicateGroups();
        if (groups.Count == 0)
        {
            ShowInfo("重複しているショートカットはありません。");
            return;
        }

        var removeCount = groups.Sum(x => x.Count() - 1);

        if (AskYesNo($"重複しているショートカットが {removeCount} 件あります。2件目以降を削除しますか？") != true)
        {
            return;
        }

        var removed = _store.RemoveDuplicateExtras();
        RefreshFilterItemsKeepSelection();
        ApplyFilter();
        ShowInfo($"{removed} 件の重複を削除しました。");
    }

    private void OpenStorageFolder()
    {
        ShortcutRunner.OpenFolder(_store.StorageFolder);
    }

    private void CheckLinksButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var item in _store.Items)
        {
            item.RefreshStatus();
        }

        var brokenCount = _store.Items.Count(x => x.IsBroken);
        ApplyFilter();

        System.Windows.MessageBox.Show(
            brokenCount == 0
                ? "存在しないファイル・フォルダはありません。"
                : $"存在しないファイル・フォルダが {brokenCount} 件あります。フィルタで「未検出のみ」を選ぶと確認できます。",
            "DHub",
            System.Windows.MessageBoxButton.OK,
            brokenCount == 0 ? System.Windows.MessageBoxImage.Information : System.Windows.MessageBoxImage.Warning);
    }

    private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedItem is null) return;
        ShortcutRunner.OpenFolder(SelectedItem.TargetPath);
    }

    private void RunButton_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedItem is null)
        {
            if (_visibleItems.Count == 1)
            {
                LaunchItem(_visibleItems[0]);
            }
            return;
        }

        LaunchItem(SelectedItem);
    }

    private void LaunchItem(ShortcutItem item)
    {
        ShortcutRunner.Run(item);
        item.TouchLaunched();
        _store.Save();
        ApplyFilter();
    }

    private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e) => ApplyFilter();
    private void FilterComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e) => ApplyFilter();

    private void ShortcutListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (SelectedItem is null) return;
        LaunchItem(SelectedItem);
    }

    private void ShortcutListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
    }

    private void Window_DragOver(object sender, System.Windows.DragEventArgs e)
    {
        e.Effects = CanAcceptDrop(e) ? System.Windows.DragDropEffects.Copy : System.Windows.DragDropEffects.None;
        e.Handled = true;
    }

    private void Window_Drop(object sender, System.Windows.DragEventArgs e)
    {
        var added = false;
        var duplicateCount = 0;

        if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop) &&
            e.Data.GetData(System.Windows.DataFormats.FileDrop) is string[] paths)
        {
            foreach (var path in paths)
            {
                if (_store.ContainsTarget(path))
                {
                    duplicateCount++;
                    continue;
                }

                added |= TryAddFromTarget(path);
            }
        }
        else if (e.Data.GetDataPresent(System.Windows.DataFormats.Text))
        {
            var text = e.Data.GetData(System.Windows.DataFormats.Text)?.ToString();
            if (!string.IsNullOrWhiteSpace(text))
            {
                var target = text.Trim();

                if (_store.ContainsTarget(target))
                {
                    duplicateCount++;
                }
                else
                {
                    added |= TryAddFromTarget(target);
                }
            }
        }

        if (added)
        {
            RefreshFilterItemsKeepSelection();
            ApplyFilter();
        }

        if (duplicateCount > 0)
        {
            ShowInfo($"重複しているため追加しなかった項目が {duplicateCount} 件あります。");
        }

        e.Handled = true;
    }

    private static bool CanAcceptDrop(System.Windows.DragEventArgs e)
    {
        return e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop) ||
               e.Data.GetDataPresent(System.Windows.DataFormats.Text);
    }

    private bool TryAddFromTarget(string target)
    {
        var type = ShortcutDetector.DetectType(target);
        if (type is null) return false;

        var item = new ShortcutItem
        {
            Name = ShortcutDetector.GuessName(target),
            TargetPath = target,
            ShortcutType = type.Value,
            CreatedAt = DateTime.Now,
            UpdatedAt = DateTime.Now
        };

        _store.Add(item);
        return true;
    }

    private void Window_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control && e.Key == Key.F)
        {
            SearchTextBox.Focus();
            SearchTextBox.SelectAll();
            e.Handled = true;
            return;
        }

        if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control && e.Key == Key.N)
        {
            OpenAddDialog();
            e.Handled = true;
            return;
        }

        if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control && e.Key == Key.Space)
        {
            ShowQuickSearch();
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Delete)
        {
            DeleteButton_Click(this, new RoutedEventArgs());
            e.Handled = true;
        }
        else if (e.Key == Key.Enter)
        {
            RunButton_Click(this, new RoutedEventArgs());
            e.Handled = true;
        }
        else if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control &&
                 e.Key == Key.C &&
                 SelectedItem is not null)
        {
            WpfClipboard.SetText(SelectedItem.TargetPath);
            e.Handled = true;
        }
    }

    private void ShowQuickSearch()
    {
        _quickSearchWindow ??= new QuickSearchWindow();
        _quickSearchWindow.LaunchRequested -= QuickSearchWindow_LaunchRequested;
        _quickSearchWindow.LaunchRequested += QuickSearchWindow_LaunchRequested;
        _quickSearchWindow.ShowSearch(_store.Items.ToList());
    }

    private void QuickSearchWindow_LaunchRequested(object? sender, ShortcutItem item)
    {
        LaunchItem(item);
    }

    private void InitializeTrayIcon()
    {
        _notifyIcon = new Forms.NotifyIcon
        {
            Icon = Drawing.SystemIcons.Application,
            Text = "DHub",
            Visible = true
        };

        _notifyIcon.DoubleClick += (_, _) =>
        {
            Show();
            WindowState = WindowState.Normal;
            Activate();
        };

        RefreshTrayMenu();
    }

    private void RefreshTrayMenu()
    {
        if (_notifyIcon is null) return;

        var menu = new Forms.ContextMenuStrip();

        var quickOpen = new Forms.ToolStripMenuItem("クイック起動");
        var launchItems = _store.Items.OrderBy(x => x.Name, StringComparer.CurrentCultureIgnoreCase).ToList();

        if (launchItems.Count == 0)
        {
            quickOpen.DropDownItems.Add(new Forms.ToolStripMenuItem("登録なし") { Enabled = false });
        }
        else
        {
            foreach (var item in launchItems)
            {
                var text = item.Name.Length > 36 ? item.Name[..36] + "..." : item.Name;
                var menuItem = new Forms.ToolStripMenuItem($"{item.IconText} {text}");
                menuItem.Click += (_, _) => Dispatcher.Invoke(() => LaunchItem(item));
                quickOpen.DropDownItems.Add(menuItem);
            }
        }

        menu.Items.Add(quickOpen);
        menu.Items.Add(new Forms.ToolStripSeparator());

        var showItem = new Forms.ToolStripMenuItem("管理画面を開く");
        showItem.Click += (_, _) =>
        {
            Show();
            WindowState = WindowState.Normal;
            Activate();
        };
        menu.Items.Add(showItem);

        var spotlightItem = new Forms.ToolStripMenuItem("クイック検索");
        spotlightItem.Click += (_, _) => Dispatcher.Invoke(ShowQuickSearch);
        menu.Items.Add(spotlightItem);

        var addItem = new Forms.ToolStripMenuItem("追加");
        addItem.Click += (_, _) => Dispatcher.Invoke(OpenAddDialog);
        menu.Items.Add(addItem);

        var exportShortcutsItem = new Forms.ToolStripMenuItem("ショートカットをエクスポート");
        exportShortcutsItem.Click += (_, _) => Dispatcher.Invoke(ExportShortcuts);
        menu.Items.Add(exportShortcutsItem);

        var importShortcutsItem = new Forms.ToolStripMenuItem("ショートカットをインポート");
        importShortcutsItem.Click += (_, _) => Dispatcher.Invoke(ImportShortcuts);
        menu.Items.Add(importShortcutsItem);

        var exportSettingsItem = new Forms.ToolStripMenuItem("設定をエクスポート");
        exportSettingsItem.Click += (_, _) => Dispatcher.Invoke(ExportSettings);
        menu.Items.Add(exportSettingsItem);

        var importSettingsItem = new Forms.ToolStripMenuItem("設定をインポート");
        importSettingsItem.Click += (_, _) => Dispatcher.Invoke(ImportSettings);
        menu.Items.Add(importSettingsItem);

        var storageItem = new Forms.ToolStripMenuItem("データ保存場所を開く");
        storageItem.Click += (_, _) => Dispatcher.Invoke(OpenStorageFolder);
        menu.Items.Add(storageItem);

        var startTrayItem = new Forms.ToolStripMenuItem("起動時はトレイのみ")
        {
            Checked = _store.Settings.StartMinimizedToTray,
            CheckOnClick = true
        };
        startTrayItem.Click += (_, _) =>
        {
            _store.Settings.StartMinimizedToTray = startTrayItem.Checked;
            _store.Save();
        };
        menu.Items.Add(startTrayItem);

        menu.Items.Add(new Forms.ToolStripSeparator());

        var exitItem = new Forms.ToolStripMenuItem("終了");
        exitItem.Click += (_, _) => Dispatcher.Invoke(() => WpfApplication.Current.Shutdown());
        menu.Items.Add(exitItem);

        var oldMenu = _notifyIcon.ContextMenuStrip;
        _notifyIcon.ContextMenuStrip = menu;
        oldMenu?.Dispose();
    }

    private async Task CheckUpdateAsync(bool silent)
    {
        try
        {
            var latest = await _updateService.CheckAsync(_store.Settings.UpdateVersionUrl);
            if (latest is null)
            {
                if (!silent) ShowInfo("更新URLが未設定です。設定ファイルの UpdateVersionUrl をGitHubのversion.jsonに変更してください。");
                return;
            }

            if (!_updateService.IsNewer(latest.Version))
            {
                if (!silent) ShowInfo("最新バージョンです。");
                return;
            }

            var message = $"新しいバージョンがあります。\n\n現在: {_updateService.CurrentVersion}\n最新: {latest.Version}\n\n更新しますか？";
            if (AskYesNo(message) != true) return;

            var downloadedZip = await _updateService.DownloadAppZipAsync(latest.DownloadUrl);
            var updater = await _updateService.DownloadAndExtractUpdaterAsync(latest.UpdaterUrl);
            _updateService.ReplaceCurrentExeAndRestartFromZip(downloadedZip, updater);
        }
        catch (Exception ex)
        {
            if (!silent)
            {
                System.Windows.MessageBox.Show(
                    "更新確認に失敗しました。\n" + ex.Message,
                    "DHub",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
            }
        }
    }

    private bool? AskYesNo(string message)
    {
        return System.Windows.MessageBox.Show(
            message,
            "DHub",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Question) == System.Windows.MessageBoxResult.Yes;
    }

    private static void ShowInfo(string message)
    {
        System.Windows.MessageBox.Show(
            message,
            "DHub",
            System.Windows.MessageBoxButton.OK,
            System.Windows.MessageBoxImage.Information);
    }
}
