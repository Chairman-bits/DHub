using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Linq;

internal static class Program
{
    private static int Main(string[] args)
    {
        try
        {
            if (args.Length < 2)
            {
                return 1;
            }

            var appZipPath = args[0];
            var targetExePath = args[1];

            if (!File.Exists(appZipPath))
            {
                return 2;
            }

            var extractDir = Path.Combine(
                Path.GetTempPath(),
                "DHub_Update_Extract_" + Guid.NewGuid().ToString("N"));

            Directory.CreateDirectory(extractDir);
            ZipFile.ExtractToDirectory(appZipPath, extractDir, overwriteFiles: true);

            var newExePath = Directory
                .EnumerateFiles(extractDir, "DHub.exe", SearchOption.AllDirectories)
                .FirstOrDefault();

            if (string.IsNullOrWhiteSpace(newExePath) || !File.Exists(newExePath))
            {
                return 3;
            }

            WaitForTargetToUnlock(targetExePath, TimeSpan.FromSeconds(30));

            var backupPath = targetExePath + ".bak";

            if (File.Exists(backupPath))
            {
                File.Delete(backupPath);
            }

            if (File.Exists(targetExePath))
            {
                File.Copy(targetExePath, backupPath, overwrite: true);
            }

            File.Copy(newExePath, targetExePath, overwrite: true);

            Process.Start(new ProcessStartInfo
            {
                FileName = targetExePath,
                UseShellExecute = true
            });

            return 0;
        }
        catch
        {
            return 9;
        }
    }

    private static void WaitForTargetToUnlock(string targetExePath, TimeSpan timeout)
    {
        var start = DateTime.Now;

        while (DateTime.Now - start < timeout)
        {
            try
            {
                if (!File.Exists(targetExePath))
                {
                    return;
                }

                using var stream = new FileStream(targetExePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                return;
            }
            catch
            {
                Thread.Sleep(500);
            }
        }
    }
}
