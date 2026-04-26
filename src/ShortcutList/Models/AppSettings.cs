namespace ShortcutList.Models;

public class AppSettings
{
    public bool StartMinimizedToTray { get; set; } = false;

    public bool ShowNameColumn { get; set; } = true;
    public bool ShowPathColumn { get; set; } = true;
    public bool ShowArgumentsColumn { get; set; } = false;
    public bool ShowGroupColumn { get; set; } = true;
    public bool ShowLaunchCountColumn { get; set; } = true;
    public bool ShowLastLaunchColumn { get; set; } = true;
    public bool ShowCreatedAtColumn { get; set; } = false;
    public bool ShowUpdatedAtColumn { get; set; } = true;
    public bool ShowStatusColumn { get; set; } = true;

    public string UpdateVersionUrl { get; set; } = "https://raw.githubusercontent.com/Chairman-bits/DHub/main/version.json";
    public bool CheckUpdateOnStartup { get; set; } = true;
}
