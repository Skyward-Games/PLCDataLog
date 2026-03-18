using System;
using System.Globalization;
using System.IO;
using System.Net.Mail;
using System.Threading;
using System.Threading.Tasks;
using PLCDataLog.Models;

namespace PLCDataLog.Services;

public sealed class DailyEmailScheduler : IDisposable
{
    private readonly SettingsService _settingsService;
    private readonly Func<CancellationToken, Task<bool>> _sendAction;
    private readonly Action<bool>? _onSendingChanged;
    private readonly Action<string>? _onError;

    private CancellationTokenSource? _cts;
    private Task? _loop;

    public DailyEmailScheduler(SettingsService settingsService, Func<CancellationToken, Task<bool>> sendAction, Action<bool>? onSendingChanged = null, Action<string>? onError = null)
    {
        _settingsService = settingsService ?? throw new ArgumentNullException(nameof(settingsService));
        _sendAction = sendAction ?? throw new ArgumentNullException(nameof(sendAction));
        _onSendingChanged = onSendingChanged;
        _onError = onError;
    }

    public void Start()
    {
        if (_loop is not null)
            return;

        _cts = new CancellationTokenSource();
        _loop = RunAsync(_cts.Token);
    }

    public void Dispose()
    {
        try
        {
            _cts?.Cancel();
        }
        catch { }
    }

    private async Task RunAsync(CancellationToken ct)
    {
        DateTimeOffset? nextRetryUtc = null;

        while (!ct.IsCancellationRequested)
        {
            try
            {
                var settings = _settingsService.Load();
                var automation = settings.EmailAutomation;

                var tz = ResolveTimeZone(automation.TimeZoneId);
                var nowUtc = DateTimeOffset.UtcNow;
                var nowInTz = TimeZoneInfo.ConvertTime(nowUtc, tz);

                if (!TryParseTime(automation.DailySendTime, out var scheduledTime))
                    scheduledTime = TimeSpan.Zero;

                AutomationLog.Info($"Tick tz='{tz.Id}' nowInTz={nowInTz:yyyy-MM-dd HH:mm:ss} enabled={automation.Enabled} sendTime='{automation.DailySendTime}' retryMin={automation.RetryIntervalMinutes} lastCsv='{automation.LastSuccessfulCsvDate}' lastSend='{automation.LastSuccessfulSendDate}' lastYDate='{automation.LastYesterdayCsvSentDate}' skipCatchUp='{automation.SkipPreviousDayCatchUpForDate}'.");

                var todayPathDbg = BuildCsvPathForDate(nowInTz.Date);
                var yPathDbg = BuildCsvPathForDate(nowInTz.Date.AddDays(-1));
                AutomationLog.Info($"Paths: today='{todayPathDbg}' exists={File.Exists(todayPathDbg)} | yesterday='{yPathDbg}' exists={File.Exists(yPathDbg)}.");

                if (!automation.Enabled)
                {
                    nextRetryUtc = null;
                    await DelayAlignedToSecond(ct);
                    continue;
                }

                if (!HasValidAutoSendPrerequisites(settings, nowInTz.Date))
                {
                    AutomationLog.Warn("Prerequisites inválidos para envio automático (sender/host/port/recipients).");
                    nextRetryUtc = null;
                    await DelayAlignedToSecond(ct);
                    continue;
                }

                if (!TryGetPendingCsvDate(settings, nowInTz.Date, scheduledTime, out var pendingDate))
                {
                    AutomationLog.Info("Nenhum CSV pendente para envio (hoje/ontem já enviados ou sem arquivo de ontem)." );
                    nextRetryUtc = null;
                    await DelayAlignedToSecond(ct);
                    continue;
                }

                var isPendingToday = pendingDate.Date == nowInTz.Date;
                var isPendingYesterday = pendingDate.Date == nowInTz.Date.AddDays(-1);

                AutomationLog.Info($"PendingDate={pendingDate:yyyy-MM-dd} pendingToday={isPendingToday} pendingYesterday={isPendingYesterday}.");

                if (!isPendingToday && !isPendingYesterday)
                {
                    nextRetryUtc = null;
                    await DelayAlignedToSecond(ct);
                    continue;
                }

                var isTimeReached =
                    nowInTz.Hour > scheduledTime.Hours ||
                    (nowInTz.Hour == scheduledTime.Hours && nowInTz.Minute >= scheduledTime.Minutes);

                // Para evitar popups/envios surpresa ao abrir o app após o horário,
                // só dispara o envio de HOJE durante o minuto exato do agendamento.
                if (isPendingToday)
                {
                    var inScheduledMinute = nowInTz.Hour == scheduledTime.Hours && nowInTz.Minute == scheduledTime.Minutes;
                    if (!inScheduledMinute)
                    {
                        AutomationLog.Info($"Aguardando minuto do agendamento para envio de hoje. now={nowInTz:HH:mm:ss} agendado={scheduledTime:hh\\:mm}." );
                        nextRetryUtc = null;
                        await DelayAlignedToSecond(ct);
                        continue;
                    }
                }

                if (isPendingToday && !isTimeReached)
                {
                    AutomationLog.Info($"Horário ainda não chegou para envio de hoje. now={nowInTz:HH:mm:ss} agendado={scheduledTime:hh\\:mm}." );
                    nextRetryUtc = null;
                    await DelayAlignedToSecond(ct);
                    continue;
                }

                // Se for pendência de hoje, garante que só dispare 1x por (data + HH:mm)
                if (isPendingToday)
                {
                    var scheduleKey = $"{nowInTz:yyyy-MM-dd} {scheduledTime:hh\\:mm}";
                    if (string.Equals(automation.LastScheduleTriggerKey, scheduleKey, StringComparison.Ordinal))
                    {
                        AutomationLog.Info($"Bloqueado: scheduleKey '{scheduleKey}' já executado hoje (LastScheduleTriggerKey)." );
                        nextRetryUtc = null;
                        await DelayAlignedToSecond(ct);
                        continue;
                    }
                }

                if (nextRetryUtc is not null && nowUtc < nextRetryUtc.Value)
                {
                    AutomationLog.Info($"Aguardando janela de retry. nextRetryUtc={nextRetryUtc:O}." );
                    await DelayAlignedToSecond(ct);
                    continue;
                }

                try
                {
                    NotifySendingChangedSafely(true);
                    AutomationLog.Info("Chamando ação de envio...");
                    var sent = await _sendAction(ct);
                    AutomationLog.Info($"Ação de envio finalizou. sent={sent}.");

                    if (sent && isPendingToday)
                    {
                        var scheduleKey = $"{nowInTz:yyyy-MM-dd} {scheduledTime:hh\\:mm}";
                        try
                        {
                            var s2 = _settingsService.Load();
                            s2.EmailAutomation.LastScheduleTriggerKey = scheduleKey;
                            _settingsService.Save(s2);
                            AutomationLog.Info($"Persistido LastScheduleTriggerKey='{scheduleKey}'.");
                        }
                        catch (Exception ex)
                        {
                            AutomationLog.Error(ex, "Falha ao persistir LastScheduleTriggerKey");
                        }
                    }

                    nextRetryUtc = sent ? null : DateTimeOffset.UtcNow.AddSeconds(30);
                }
                catch (OperationCanceledException)
                {
                    AutomationLog.Warn("Envio cancelado.");
                    throw;
                }
                catch (Exception ex)
                {
                    AutomationLog.Error($"Exceção no envio: {ex.Message}");
                    NotifyErrorSafely(ex.Message);
                    var retryMin = Math.Clamp(automation.RetryIntervalMinutes, 1, 24 * 60);
                    nextRetryUtc = DateTimeOffset.UtcNow.AddMinutes(retryMin);
                }
                finally
                {
                    NotifySendingChangedSafely(false);
                }

                await DelayAlignedToSecond(ct);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                AutomationLog.Error($"Exceção no loop do scheduler: {ex.Message}");
                NotifyErrorSafely(ex.Message);
                await Task.Delay(TimeSpan.FromSeconds(5), ct);
            }
        }
    }

    private void NotifySendingChangedSafely(bool isSending)
    {
        if (_onSendingChanged is null)
            return;

        try
        {
            _onSendingChanged.Invoke(isSending);
        }
        catch
        {
            // UI callback nunca deve impedir o envio em background.
        }
    }

    private void NotifyErrorSafely(string message)
    {
        if (_onError is null)
            return;

        try
        {
            _onError.Invoke(message);
        }
        catch
        {
            // Ignora erro de callback para manter o scheduler ativo.
        }
    }

    private static Task DelayAlignedToSecond(CancellationToken ct)
    {
        return Task.Delay(TimeSpan.FromSeconds(15), ct);
    }

    private static TimeZoneInfo ResolveTimeZone(string? timeZoneId)
    {
        if (string.IsNullOrWhiteSpace(timeZoneId))
            return TimeZoneInfo.Local;

        try
        {
            return TimeZoneInfo.FindSystemTimeZoneById(timeZoneId.Trim());
        }
        catch
        {
            return TimeZoneInfo.Local;
        }
    }

    private static bool TryParseTime(string? hhmm, out TimeSpan time)
    {
        time = default;
        if (string.IsNullOrWhiteSpace(hhmm))
            return false;

        return TimeSpan.TryParseExact(hhmm.Trim(), "hh\\:mm", CultureInfo.InvariantCulture, out time);
    }

    private static DateTimeOffset? ParseLastSuccessful(string? value, TimeZoneInfo tz)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;

        if (DateTime.TryParseExact(value, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out var withTime))
        {
            var unspecified = DateTime.SpecifyKind(withTime, DateTimeKind.Unspecified);
            return new DateTimeOffset(unspecified, tz.GetUtcOffset(unspecified));
        }

        if (DateTime.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateOnly))
        {
            var unspecified = DateTime.SpecifyKind(dateOnly, DateTimeKind.Unspecified);
            return new DateTimeOffset(unspecified, tz.GetUtcOffset(unspecified));
        }

        return null;
    }

    private static bool HasValidAutoSendPrerequisites(AppSettings settings, DateTime todayInTz)
    {
        if (!settings.EmailAutomation.Enabled)
            return false;

        if (string.IsNullOrWhiteSpace(settings.Email.SenderEmail) || !IsValidEmail(settings.Email.SenderEmail))
            return false;

        if (string.IsNullOrWhiteSpace(settings.Email.SmtpHost))
            return false;

        if (settings.Email.SmtpPort is <= 0 or > 65535)
            return false;

        if (settings.Email.Recipients is null || settings.Email.Recipients.Count == 0)
            return false;

        return true;
    }

    private static bool TryGetPendingCsvDate(AppSettings settings, DateTime todayInTz, TimeSpan scheduledTime, out DateTime pendingDate)
    {
        pendingDate = default;

        var todayPath = BuildCsvPathForDate(todayInTz);
        var yesterday = todayInTz.AddDays(-1);
        var yesterdayPath = BuildCsvPathForDate(yesterday);
        var hasToday = File.Exists(todayPath);
        var hasYesterday = File.Exists(yesterdayPath);

        DateTime? lastCsvDate = null;
        if (DateTime.TryParseExact(settings.EmailAutomation.LastSuccessfulCsvDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var lastCsv))
            lastCsvDate = lastCsv.Date;

        var skipCatchUpForToday =
            !string.IsNullOrWhiteSpace(settings.EmailAutomation.SkipPreviousDayCatchUpForDate) &&
            DateTime.TryParseExact(settings.EmailAutomation.SkipPreviousDayCatchUpForDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var skipDate) &&
            skipDate.Date == todayInTz;

        var yesterdayAlreadySentByCatchUp =
            !string.IsNullOrWhiteSpace(settings.EmailAutomation.LastYesterdayCsvSentDate) &&
            DateTime.TryParseExact(settings.EmailAutomation.LastYesterdayCsvSentDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var caughtUpDate) &&
            caughtUpDate.Date == yesterday;

        // Ontem: só tenta se o arquivo existe.
        if (!skipCatchUpForToday && hasYesterday && lastCsvDate != yesterday && !yesterdayAlreadySentByCatchUp)
        {
            pendingDate = yesterday;
            return true;
        }

        // Hoje: disparo baseado no horário configurado (1x por data + HH:mm)
        var scheduleKey = $"{todayInTz:yyyy-MM-dd} {scheduledTime:hh\\:mm}";
        var alreadyTriggered = string.Equals(settings.EmailAutomation.LastScheduleTriggerKey, scheduleKey, StringComparison.Ordinal);
        if (!alreadyTriggered)
        {
            if (!hasToday)
                return false;

            pendingDate = todayInTz;
            return true;
        }

        return false;
    }

    // Mantém a assinatura antiga apenas para compatibilidade interna (não usada).
    private static bool TryGetPendingCsvDate(AppSettings settings, DateTime todayInTz, out DateTime pendingDate)
        => TryGetPendingCsvDate(settings, todayInTz, TimeSpan.Zero, out pendingDate);

    private static string BuildCsvPathForDate(DateTime date)
    {
        var root = Path.Combine(SettingsService.GetDefaultDataRootPath(), "Logs");
        var yearDir = Path.Combine(root, date.Year.ToString("0000", CultureInfo.InvariantCulture));
        var monthDir = Path.Combine(yearDir, date.Month.ToString("00", CultureInfo.InvariantCulture));
        return Path.Combine(monthDir, $"{date:yyyy-MM-dd}.csv");
    }

    private static bool IsValidEmail(string email)
    {
        try
        {
            _ = new MailAddress(email);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
