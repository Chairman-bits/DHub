using System.Collections.ObjectModel;
using System.Text.Json;
using ShortcutList.Models;

namespace ShortcutList.Services;

public class ShortcutStore
{
    private readonly JsonSerializerOptions _options = new() { WriteIndented = true };
    private readonly string _storageFolder;
    private readonly string _storagePath;

    public ObservableCollection<ShortcutItem> Items { get; } = new();
    public AppSettings Settings { get; private set; } = new();

    public string StorageFolder => _storageFolder;

    public ShortcutStore()
    {
        _storageFolder = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DHub");

        System.IO.Directory.CreateDirectory(_storageFolder);
        _storagePath = System.IO.Path.Combine(_storageFolder, "shortcuts.json");
    }

    public void Load()
    {
        Items.Clear();

        if (!System.IO.File.Exists(_storagePath))
        {
            Save();
            return;
        }

        try
        {
            var data = ReadDataFromFile(_storagePath);
            Settings = data.Settings ?? new AppSettings();

            foreach (var item in data.Items)
            {
                Items.Add(item);
            }
        }
        catch
        {
            Settings = new AppSettings();
            Save();
        }
    }

    public void Save()
    {
        var data = new AppData
        {
            Items = Items.ToList(),
            Settings = Settings
        };

        System.IO.File.WriteAllText(_storagePath, JsonSerializer.Serialize(data, _options));
    }

    public void Add(ShortcutItem item)
    {
        Items.Add(item);
        Save();
    }

    public void Remove(ShortcutItem item)
    {
        Items.Remove(item);
        Save();
    }

    public bool ContainsTarget(string target, string? ignoreId = null)
    {
        return Items.Any(x =>
            !string.Equals(x.Id, ignoreId, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(NormalizeTarget(x.TargetPath), NormalizeTarget(target), StringComparison.OrdinalIgnoreCase));
    }

    public IReadOnlyList<IGrouping<string, ShortcutItem>> GetDuplicateGroups()
    {
        return Items
            .GroupBy(x => NormalizeTarget(x.TargetPath), StringComparer.OrdinalIgnoreCase)
            .Where(g => g.Count() > 1)
            .ToList();
    }

    public int RemoveDuplicateExtras()
    {
        var duplicates = GetDuplicateGroups();
        var removeTargets = duplicates
            .SelectMany(g => g.Skip(1))
            .ToList();

        foreach (var item in removeTargets)
        {
            Items.Remove(item);
        }

        if (removeTargets.Count > 0)
        {
            Save();
        }

        return removeTargets.Count;
    }

    public void ExportToFile(string path)
    {
        var data = new AppData
        {
            Items = Items.ToList(),
            Settings = Settings
        };

        System.IO.File.WriteAllText(path, JsonSerializer.Serialize(data, _options));
    }

    public void ExportShortcutsToFile(string path)
    {
        var data = new AppData
        {
            Items = Items.ToList(),
            Settings = new AppSettings()
        };

        System.IO.File.WriteAllText(path, JsonSerializer.Serialize(data, _options));
    }

    public void ImportShortcutsFromFile(string path)
    {
        var data = ReadDataFromFile(path);

        Items.Clear();

        foreach (var item in data.Items)
        {
            Items.Add(item);
        }

        Save();
    }

    public void ExportSettingsToFile(string path)
    {
        System.IO.File.WriteAllText(path, JsonSerializer.Serialize(Settings, _options));
    }

    public void ImportSettingsFromFile(string path)
    {
        var json = System.IO.File.ReadAllText(path);
        Settings = JsonSerializer.Deserialize<AppSettings>(json, _options) ?? new AppSettings();
        Save();
    }

    public IEnumerable<string> GetGroups()
    {
        return Items
            .Select(x => x.GroupDisplay)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(x => x);
    }

    public int RenameGroup(string oldName, string newName)
    {
        var count = 0;

        foreach (var item in Items.Where(x => string.Equals(x.GroupDisplay, oldName, StringComparison.OrdinalIgnoreCase)))
        {
            item.GroupName = newName;
            item.TouchUpdated();
            count++;
        }

        if (count > 0)
        {
            Save();
        }

        return count;
    }

    public int DeleteGroupName(string groupName)
    {
        var count = 0;

        foreach (var item in Items.Where(x => string.Equals(x.GroupDisplay, groupName, StringComparison.OrdinalIgnoreCase)))
        {
            item.GroupName = string.Empty;
            item.TouchUpdated();
            count++;
        }

        if (count > 0)
        {
            Save();
        }

        return count;
    }

    private AppData ReadDataFromFile(string path)
    {
        var json = System.IO.File.ReadAllText(path);
        return JsonSerializer.Deserialize<AppData>(json, _options) ?? new AppData();
    }

    private static string NormalizeTarget(string target)
    {
        return (target ?? string.Empty)
            .Trim()
            .TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar);
    }
}
