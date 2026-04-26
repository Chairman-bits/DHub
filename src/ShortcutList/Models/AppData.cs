namespace ShortcutList.Models;

public class AppData
{
    public List<ShortcutItem> Items { get; set; } = new();
    public AppSettings Settings { get; set; } = new();
}
