using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;

namespace PLCDataLog.Services;

public sealed class CsvLogService
{
    public string EnsureLogFileForToday()
    {
        var root = Path.Combine(SettingsService.GetDefaultDataRootPath(), "Logs");
        var now = DateTime.Now;
        var yearDir = Path.Combine(root, now.Year.ToString("0000", CultureInfo.InvariantCulture));
        var monthDir = Path.Combine(yearDir, now.Month.ToString("00", CultureInfo.InvariantCulture));
        Directory.CreateDirectory(monthDir);

        var fileName = $"{now:yyyy-MM-dd}.csv";
        return Path.Combine(monthDir, fileName);
    }

    public void AppendSnapshot(string filePath, DateTime timestamp, IReadOnlyList<(string Name, string Address, string Type, string Value)> values, IReadOnlyList<string> changedNames)
    {
        var exists = File.Exists(filePath);

        using var stream = File.Open(filePath, FileMode.Append, FileAccess.Write, FileShare.Read);
        using var writer = new StreamWriter(stream);
        using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);

        if (!exists)
        {
            csv.WriteField("DataHora");
            csv.WriteField("Alteracoes");
            foreach (var v in values)
                csv.WriteField(v.Name);
            csv.NextRecord();
        }

        csv.WriteField(timestamp.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture));
        csv.WriteField(string.Join(";", changedNames));
        foreach (var v in values)
            csv.WriteField(v.Value);
        csv.NextRecord();
    }

    public static IReadOnlyList<string> GetChanges(Dictionary<string, string> previous, IReadOnlyList<(string Key, string Value)> current)
    {
        var changes = new List<string>();

        foreach (var (key, value) in current)
        {
            if (!previous.TryGetValue(key, out var prev) || !string.Equals(prev, value, StringComparison.Ordinal))
                changes.Add(key);
        }

        return changes;
    }

    public static void UpdatePrevious(Dictionary<string, string> previous, IReadOnlyList<(string Key, string Value)> current)
    {
        foreach (var (key, value) in current)
            previous[key] = value;
    }
}
