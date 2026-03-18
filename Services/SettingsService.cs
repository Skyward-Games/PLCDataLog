using System;
using System.IO;
using System.Text.Json;
using System.Threading;
using PLCDataLog.Models;

namespace PLCDataLog.Services;

public sealed class SettingsService
{
    private static readonly object SyncRoot = new();

    private static readonly JsonSerializerOptions Options = new(JsonSerializerDefaults.Web)
    {
        WriteIndented = true,
    };

    private readonly string _filePath;

    public SettingsService(string filePath)
    {
        _filePath = filePath;
    }

    public AppSettings Load()
    {
        lock (SyncRoot)
        {
            try
            {
                if (!File.Exists(_filePath))
                    return new AppSettings();

                var json = ExecuteWithIoRetry(() => File.ReadAllText(_filePath));
                return JsonSerializer.Deserialize<AppSettings>(json, Options) ?? new AppSettings();
            }
            catch
            {
                return new AppSettings();
            }
        }
    }

    public void Save(AppSettings settings)
    {
        lock (SyncRoot)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_filePath) ?? ".");
            var json = JsonSerializer.Serialize(settings, Options);

            var tempPath = _filePath + ".tmp";

            ExecuteWithIoRetry(() =>
            {
                File.WriteAllText(tempPath, json);
                File.Copy(tempPath, _filePath, overwrite: true);
                File.Delete(tempPath);
            });
        }
    }

    private static void ExecuteWithIoRetry(Action action)
    {
        const int maxAttempts = 5;
        var delay = TimeSpan.FromMilliseconds(40);

        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                action();
                return;
            }
            catch (IOException) when (attempt < maxAttempts)
            {
                Thread.Sleep(delay);
            }
        }

        action();
    }

    private static T ExecuteWithIoRetry<T>(Func<T> action)
    {
        const int maxAttempts = 5;
        var delay = TimeSpan.FromMilliseconds(40);

        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                return action();
            }
            catch (IOException) when (attempt < maxAttempts)
            {
                Thread.Sleep(delay);
            }
        }

        return action();
    }

    public static string GetDefaultSettingsPath()
    {
        return Path.Combine(GetDefaultDataRootPath(), "appsettings.json");
    }

    public static string GetDefaultDataRootPath()
    {
        var root = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (string.IsNullOrWhiteSpace(root))
            root = AppContext.BaseDirectory;

        return Path.Combine(root, "PLCDataLog");
    }
}
