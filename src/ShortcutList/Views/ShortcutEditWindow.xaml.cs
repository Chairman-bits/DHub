using System.Windows;
using System.Windows.Controls;
using Forms = System.Windows.Forms;
using ShortcutList.Models;
using ShortcutList.Services;

namespace ShortcutList.Views;

public partial class ShortcutEditWindow : Window
{
    private bool _nameEditedByUser;
    private bool _openApplicationEditedByUser;
    private bool _updatingOpenApplicationText;

    public ShortcutItem? ResultItem { get; private set; }

    public ShortcutEditWindow()
    {
        InitializeComponent();
        NameTextBox.TextChanged += NameTextBox_TextChanged;
        SelectColor("#CBD5E1");
        ApplyDefaultOpenApplicationIfNeeded(force: true);
    }

    public ShortcutEditWindow(ShortcutItem item)
    {
        InitializeComponent();
        NameTextBox.TextChanged += NameTextBox_TextChanged;

        ResultItem = new ShortcutItem
        {
            Id = item.Id,
            Name = item.Name,
            TargetPath = item.TargetPath,
            Arguments = item.Arguments,
            OpenApplicationPath = item.OpenApplicationPath,
            OpenApplicationArguments = item.OpenApplicationArguments,
            ShortcutType = item.ShortcutType,
            IsFavorite = item.IsFavorite,
            GroupName = item.GroupName,
            TagColor = item.TagColor,
            CreatedAt = item.CreatedAt,
            UpdatedAt = item.UpdatedAt,
            LaunchCount = item.LaunchCount,
            LastLaunchedAt = item.LastLaunchedAt
        };

        NameTextBox.Text = item.Name;
        TargetTextBox.Text = item.TargetPath;
        ArgumentsTextBox.Text = item.Arguments;
        SetOpenApplicationText(string.IsNullOrWhiteSpace(item.OpenApplicationPath)
            ? ShortcutRunner.GetDefaultOpenApplicationPath(item.ShortcutType)
            : item.OpenApplicationPath);
        OpenApplicationArgumentsTextBox.Text = item.OpenApplicationArguments;
        FavoriteCheckBox.IsChecked = item.IsFavorite;
        GroupTextBox.Text = item.GroupName;
        SelectColor(item.TagColor);

        _nameEditedByUser = true;
        _openApplicationEditedByUser = !string.IsNullOrWhiteSpace(item.OpenApplicationPath);
    }

    private void BrowseButton_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new Forms.OpenFileDialog
        {
            CheckFileExists = false,
            ValidateNames = false,
            FileName = "フォルダを選択"
        };

        if (dialog.ShowDialog() != Forms.DialogResult.OK)
        {
            return;
        }

        var path = dialog.FileName;

        if (!System.IO.File.Exists(path) &&
            System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(path)))
        {
            path = System.IO.Path.GetDirectoryName(path) ?? path;
        }

        TargetTextBox.Text = path;
    }

    private void BrowseOpenApplicationButton_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new Forms.OpenFileDialog
        {
            Title = "開くアプリを選択",
            Filter = "実行ファイル (*.exe)|*.exe|すべてのファイル (*.*)|*.*",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() != Forms.DialogResult.OK)
        {
            return;
        }

        _openApplicationEditedByUser = true;
        SetOpenApplicationText(dialog.FileName);
    }

    private void ClearOpenApplicationButton_Click(object sender, RoutedEventArgs e)
    {
        _openApplicationEditedByUser = false;
        OpenApplicationArgumentsTextBox.Text = string.Empty;
        ApplyDefaultOpenApplicationIfNeeded(force: true);
    }

    private void TargetTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (!_nameEditedByUser)
        {
            var target = TargetTextBox.Text.Trim();
            var guessed = ShortcutDetector.GuessName(target);

            if (!string.IsNullOrWhiteSpace(guessed))
            {
                NameTextBox.Text = guessed;
            }
        }

        ApplyDefaultOpenApplicationIfNeeded(force: false);
    }

    private void OpenApplicationTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_updatingOpenApplicationText)
        {
            return;
        }

        _openApplicationEditedByUser = true;
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        var target = TargetTextBox.Text.Trim();
        var type = ShortcutDetector.DetectType(target);

        if (type is null)
        {
            System.Windows.MessageBox.Show(
                "存在するフォルダ、ファイル、またはURLを入力してください。",
                "DHub",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Warning);
            return;
        }

        var name = NameTextBox.Text.Trim();

        if (string.IsNullOrWhiteSpace(name))
        {
            name = ShortcutDetector.GuessName(target);
        }

        if (string.IsNullOrWhiteSpace(name))
        {
            System.Windows.MessageBox.Show(
                "名前を入力してください。",
                "DHub",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Warning);
            return;
        }

        ResultItem ??= new ShortcutItem
        {
            CreatedAt = DateTime.Now
        };

        ResultItem.Name = name;
        ResultItem.TargetPath = target;
        ResultItem.Arguments = ArgumentsTextBox.Text.Trim();
        ResultItem.OpenApplicationPath = ResolveOpenApplicationPathForSave(type.Value);
        ResultItem.OpenApplicationArguments = OpenApplicationArgumentsTextBox.Text.Trim();
        ResultItem.ShortcutType = type.Value;
        ResultItem.IsFavorite = FavoriteCheckBox.IsChecked == true;
        ResultItem.GroupName = GroupTextBox.Text.Trim();
        ResultItem.TagColor = GetSelectedColor();
        ResultItem.UpdatedAt = DateTime.Now;

        DialogResult = true;
        Close();
    }

    private string ResolveOpenApplicationPathForSave(ShortcutType shortcutType)
    {
        var openApplicationPath = OpenApplicationTextBox.Text.Trim();

        if (string.IsNullOrWhiteSpace(openApplicationPath))
        {
            return ShortcutRunner.GetDefaultOpenApplicationPath(shortcutType);
        }

        return openApplicationPath;
    }

    private void ApplyDefaultOpenApplicationIfNeeded(bool force)
    {
        if (!force && _openApplicationEditedByUser)
        {
            return;
        }

        var type = ShortcutDetector.DetectType(TargetTextBox.Text.Trim());
        var defaultPath = type.HasValue ? ShortcutRunner.GetDefaultOpenApplicationPath(type.Value) : string.Empty;
        SetOpenApplicationText(defaultPath);
    }

    private void SetOpenApplicationText(string value)
    {
        _updatingOpenApplicationText = true;
        OpenApplicationTextBox.Text = value;
        _updatingOpenApplicationText = false;
    }

    private string GetSelectedColor()
    {
        if (TagColorComboBox.SelectedItem is ComboBoxItem item &&
            item.Tag is string color &&
            !string.IsNullOrWhiteSpace(color))
        {
            return color;
        }

        return "#CBD5E1";
    }

    private void SelectColor(string color)
    {
        foreach (var item in TagColorComboBox.Items.OfType<ComboBoxItem>())
        {
            if (string.Equals(item.Tag?.ToString(), color, StringComparison.OrdinalIgnoreCase))
            {
                TagColorComboBox.SelectedItem = item;
                return;
            }
        }

        if (TagColorComboBox.Items.Count > 0)
        {
            TagColorComboBox.SelectedIndex = 0;
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void NameTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        _nameEditedByUser = !string.IsNullOrWhiteSpace(NameTextBox.Text);
    }
}
