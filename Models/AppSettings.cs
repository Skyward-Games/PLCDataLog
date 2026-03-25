using System.Collections.Generic;

namespace PLCDataLog.Models;

public sealed class AppSettings
{
    public EmailSettings Email { get; set; } = new();
    public EmailAutomationSettings EmailAutomation { get; set; } = new();
    public PlcConnectionSettings PlcConnection { get; set; } = new();
    public NetworkBackupSettings NetworkBackup { get; set; } = new();
    public RecipeMonitorSettings RecipeMonitor { get; set; } = new();
    public UiSettings Ui { get; set; } = new();
}

public sealed class UiSettings
{
    public string ThemeMode { get; set; } = "System"; // System | Light | Dark
}

public sealed class EmailAutomationSettings
{
    public bool Enabled { get; set; } = true;

    // horário local (HH:mm) - por padrão meia-noite
    public string DailySendTime { get; set; } = "00:00";

    // fuso horário do agendamento. Quando null/vazio: usa o fuso local desta máquina.
    // exemplos: "E. South America Standard Time", "UTC"
    public string? TimeZoneId { get; set; }

    // controle de retentativa/persistência
    public string? LastSuccessfulSendDate { get; set; }
    public string? LastSuccessfulCsvDate { get; set; }

    public string? LastYesterdayCsvSentDate { get; set; }
    public string? LastYesterdayCsvSentAt { get; set; }

    public string? LastAppOpenDate { get; set; }
    public string? SkipPreviousDayCatchUpForDate { get; set; }
    public int RetryIntervalMinutes { get; set; } = 15;

    // Controle do disparo diário por horário configurado.
    // Ex: "2026-03-17 14:06" (data no fuso selecionado)
    public string? LastScheduleTriggerKey { get; set; }
}

public sealed class EmailSettings
{
    public string SenderEmail { get; set; } = string.Empty;
    public string SmtpHost { get; set; } = string.Empty;
    public int SmtpPort { get; set; } = 587;
    public EmailSecurityType SecurityType { get; set; } = EmailSecurityType.StartTls;

    public string SmtpUsername { get; set; } = string.Empty;
    public string SmtpPassword { get; set; } = string.Empty;

    public int SmtpTimeoutSeconds { get; set; } = 30;
    public int SmtpRetryCount { get; set; } = 2;
    public int SmtpRetryDelaySeconds { get; set; } = 5;

    public List<string> Recipients { get; set; } = new();
}

public enum EmailSecurityType
{
    None = 0,
    SslTls = 1,
    StartTls = 2,
}

public sealed class PlcConnectionSettings
{
    public string Host { get; set; } = "127.0.0.1";
    public int Port { get; set; } = 502;
    public byte UnitId { get; set; } = 1;
    public int PollIntervalMs { get; set; } = 500;
}

public sealed class NetworkBackupSettings
{
    public bool Enabled { get; set; }
    public string TargetFolder { get; set; } = string.Empty;
    public string Username { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
}

public sealed class RecipeMonitorSettings
{
    public bool Enabled { get; set; }
    public string SourceFolder { get; set; } = string.Empty;
    public string TargetFolder { get; set; } = string.Empty;
    public string Username { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
}
