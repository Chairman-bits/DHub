using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ShortcutList.Models;

public class ShortcutItem : INotifyPropertyChanged
{
    private string _id = Guid.NewGuid().ToString("N");
    private string _name = string.Empty;
    private string _targetPath = string.Empty;
    private string _arguments = string.Empty;
    private string _openApplicationPath = string.Empty;
    private string _openApplicationArguments = string.Empty;
    private string _groupName = string.Empty;
    private string _tagColor = "#CBD5E1";
    private ShortcutType _shortcutType;
    private bool _isFavorite;
    private DateTime _createdAt = DateTime.Now;
    private DateTime _updatedAt = DateTime.Now;
    private int _launchCount;
    private DateTime? _lastLaunchedAt;

    public event PropertyChangedEventHandler? PropertyChanged;

    public string Id { get => _id; set => SetField(ref _id, value); }

    public string Name
    {
        get => _name;
        set
        {
            if (SetField(ref _name, value ?? string.Empty))
            {
                OnPropertyChanged(nameof(SearchText));
            }
        }
    }

    public string TargetPath
    {
        get => _targetPath;
        set
        {
            if (SetField(ref _targetPath, value ?? string.Empty))
            {
                OnPropertyChanged(nameof(SearchText));
                OnPropertyChanged(nameof(TypeText));
                OnPropertyChanged(nameof(IconText));
                OnPropertyChanged(nameof(StatusText));
                OnPropertyChanged(nameof(IsBroken));
            }
        }
    }

    public string Arguments
    {
        get => _arguments;
        set
        {
            if (SetField(ref _arguments, value ?? string.Empty))
            {
                OnPropertyChanged(nameof(SearchText));
            }
        }
    }

    public string OpenApplicationPath
    {
        get => _openApplicationPath;
        set
        {
            if (SetField(ref _openApplicationPath, value?.Trim() ?? string.Empty))
            {
                OnPropertyChanged(nameof(OpenApplicationDisplay));
                OnPropertyChanged(nameof(SearchText));
            }
        }
    }

    public string OpenApplicationArguments
    {
        get => _openApplicationArguments;
        set
        {
            if (SetField(ref _openApplicationArguments, value?.Trim() ?? string.Empty))
            {
                OnPropertyChanged(nameof(OpenApplicationDisplay));
                OnPropertyChanged(nameof(SearchText));
            }
        }
    }

    public ShortcutType ShortcutType
    {
        get => _shortcutType;
        set
        {
            if (SetField(ref _shortcutType, value))
            {
                OnPropertyChanged(nameof(TypeText));
                OnPropertyChanged(nameof(IconText));
                OnPropertyChanged(nameof(StatusText));
                OnPropertyChanged(nameof(IsBroken));
                OnPropertyChanged(nameof(OpenApplicationDisplay));
                OnPropertyChanged(nameof(SearchText));
            }
        }
    }

    public bool IsFavorite
    {
        get => _isFavorite;
        set
        {
            if (SetField(ref _isFavorite, value))
            {
                OnPropertyChanged(nameof(FavoriteText));
                OnPropertyChanged(nameof(FavoriteDisplay));
                OnPropertyChanged(nameof(SearchText));
            }
        }
    }

    public string GroupName
    {
        get => _groupName;
        set
        {
            if (SetField(ref _groupName, value?.Trim() ?? string.Empty))
            {
                OnPropertyChanged(nameof(GroupDisplay));
                OnPropertyChanged(nameof(SearchText));
            }
        }
    }

    public string TagColor
    {
        get => _tagColor;
        set => SetField(ref _tagColor, string.IsNullOrWhiteSpace(value) ? "#CBD5E1" : value);
    }

    public DateTime CreatedAt { get => _createdAt; set => SetField(ref _createdAt, value); }
    public DateTime UpdatedAt { get => _updatedAt; set => SetField(ref _updatedAt, value); }
    public int LaunchCount { get => _launchCount; set => SetField(ref _launchCount, value); }
    public DateTime? LastLaunchedAt { get => _lastLaunchedAt; set => SetField(ref _lastLaunchedAt, value); }

    public string TypeText => ShortcutType switch
    {
        ShortcutType.Folder => "フォルダ",
        ShortcutType.Url => "URL",
        ShortcutType.File => "ファイル",
        _ => "不明"
    };

    public string IconText => ShortcutType switch
    {
        ShortcutType.Folder => "📁",
        ShortcutType.Url => "🌐",
        ShortcutType.File => "📄",
        _ => "□"
    };

    public string FavoriteText => IsFavorite ? "★" : "☆";
    public string FavoriteDisplay => IsFavorite ? "お気に入り" : "通常";

    public string GroupDisplay => string.IsNullOrWhiteSpace(GroupName) ? "未分類" : GroupName;

    public bool IsBroken => ShortcutType switch
    {
        ShortcutType.Folder => !System.IO.Directory.Exists(TargetPath),
        ShortcutType.File => !System.IO.File.Exists(TargetPath),
        _ => false
    };

    public string StatusText => IsBroken ? "未検出" : "OK";

    public string LastLaunchedDisplay => LastLaunchedAt.HasValue ? LastLaunchedAt.Value.ToString("yyyy/MM/dd HH:mm") : "-";

    public string OpenApplicationDisplay
    {
        get
        {
            var path = OpenApplicationPath;

            if (string.IsNullOrWhiteSpace(path) && ShortcutType == ShortcutType.Folder)
            {
                return "エクスプローラー";
            }

            if (string.IsNullOrWhiteSpace(path))
            {
                return "既定のアプリ";
            }

            if (string.Equals(path, "explorer.exe", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(System.IO.Path.GetFileName(path), "explorer.exe", StringComparison.OrdinalIgnoreCase))
            {
                return "エクスプローラー";
            }

            var fileName = System.IO.Path.GetFileNameWithoutExtension(path);
            return string.IsNullOrWhiteSpace(fileName) ? path : fileName;
        }
    }

    public string SearchText => $"{Name} {TargetPath} {Arguments} {OpenApplicationPath} {OpenApplicationArguments} {OpenApplicationDisplay} {FavoriteDisplay} {TypeText} {GroupDisplay} {StatusText}".ToLowerInvariant();

    public void TouchUpdated()
    {
        UpdatedAt = DateTime.Now;
        OnPropertyChanged(nameof(UpdatedAt));
        RefreshStatus();
    }

    public void TouchLaunched()
    {
        LaunchCount++;
        LastLaunchedAt = DateTime.Now;
        OnPropertyChanged(nameof(LaunchCount));
        OnPropertyChanged(nameof(LastLaunchedAt));
        OnPropertyChanged(nameof(LastLaunchedDisplay));
    }

    public void RefreshStatus()
    {
        OnPropertyChanged(nameof(IsBroken));
        OnPropertyChanged(nameof(StatusText));
        OnPropertyChanged(nameof(SearchText));
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}
