using System.Diagnostics;
using ShortcutList.Models;

namespace ShortcutList.Services;

public static class ShortcutRunner
{
    public static void Run(ShortcutItem item)
    {
        if (item.ShortcutType == ShortcutType.Folder)
        {
            OpenFolder(item.TargetPath);
            return;
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = item.TargetPath,
            Arguments = item.Arguments ?? string.Empty,
            UseShellExecute = true
        };

        Process.Start(startInfo);
    }

    public static void OpenFolder(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        var folder = path;

        if (System.IO.File.Exists(path))
        {
            folder = System.IO.Path.GetDirectoryName(path) ?? path;
        }

        if (!System.IO.Directory.Exists(folder))
        {
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = folder,
            UseShellExecute = true
        });
    }

    public static void OpenUrl(string url)
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = url,
            UseShellExecute = true
        });
    }
}
