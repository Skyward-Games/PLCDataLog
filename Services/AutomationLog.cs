using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace PLCDataLog.Services;

internal static class AutomationLog
{
    private static readonly object Sync = new();

    public static void Info(string message) => Write("INFO", message, null);
    public static void Warn(string message) => Write("WARN", message, null);
    public static void Error(string message) => Write("ERROR", message, null);
    public static void Error(Exception ex, string? message = null) => Write("ERROR", message, ex);

    public static void Ui(string action, string? detail = null) => Write("UI", FormatAction(action, detail), null);
    public static void State(string name, string? value) => Write("STATE", $"{name}='{value ?? string.Empty}'", null);

    public static string? TryGetCurrentLogPath()
    {
        try
        {
            var root = GetLogRootPath();
            var file = $"Automation_{DateTime.Now:yyyy-MM-dd}.log";
            return Path.Combine(root, file);
        }
        catch
        {
            return null;
        }
    }

    private static string GetLogRootPath()
    {
        // LocalAppData costuma ser a opção mais confiável para escrita (sem precisar permissões elevadas)
        // e evita falhas quando o app está em "Program Files" ou pasta somente leitura.
        var root = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (string.IsNullOrWhiteSpace(root))
            root = AppContext.BaseDirectory;

        return Path.Combine(root, "PLCDataLog", "Logs");
    }

    private static string FormatAction(string action, string? detail)
    {
        if (string.IsNullOrWhiteSpace(detail))
            return action;

        return $"{action} | {detail}";
    }

    private static void Write(string level, string? message, Exception? ex)
    {
        try
        {
            var root = GetLogRootPath();
            Directory.CreateDirectory(root);

            var file = $"Automation_{DateTime.Now:yyyy-MM-dd}.log";
            var path = Path.Combine(root, file);

            var pid = Environment.ProcessId;
            var tid = Environment.CurrentManagedThreadId;
            var proc = Process.GetCurrentProcess().ProcessName;

            var sb = new StringBuilder();
            sb.Append(DateTimeOffset.Now.ToString("yyyy-MM-dd HH:mm:ss.fff zzz", CultureInfo.InvariantCulture));
            sb.Append(' ');
            sb.Append('[').Append(level).Append(']');
            sb.Append(' ');
            sb.Append(proc).Append('#').Append(pid).Append('/').Append("T").Append(tid);

            if (!string.IsNullOrWhiteSpace(message))
            {
                sb.Append(' ');
                sb.Append(message.Replace("\r", " ").Replace("\n", " "));
            }

            if (ex is not null)
            {
                sb.Append(' ');
                sb.Append("| ");
                sb.Append(ex.GetType().Name);
                sb.Append(": ");
                sb.Append(ex.Message.Replace("\r", " ").Replace("\n", " "));

                if (!string.IsNullOrWhiteSpace(ex.StackTrace))
                {
                    sb.Append(" | Stack: ");
                    sb.Append(ex.StackTrace.Replace("\r", " ").Replace("\n", " | "));
                }

                if (ex.InnerException is not null)
                {
                    sb.Append(" | Inner: ");
                    sb.Append(ex.InnerException.GetType().Name);
                    sb.Append(": ");
                    sb.Append(ex.InnerException.Message.Replace("\r", " ").Replace("\n", " "));
                }
            }

            var line = sb.ToString();

            lock (Sync)
            {
                File.AppendAllText(path, line + Environment.NewLine, Encoding.UTF8);
            }
        }
        catch
        {
            // never throw from logging
        }
    }
}
