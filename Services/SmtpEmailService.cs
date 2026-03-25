using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Threading;
using System.Threading.Tasks;
using PLCDataLog.Models;

namespace PLCDataLog.Services;

public sealed class SmtpEmailService
{
    public Task SendTestEmailAsync(EmailSettings settings, string recipient, CancellationToken ct)
    {
        ArgumentNullException.ThrowIfNull(settings);

        if (string.IsNullOrWhiteSpace(recipient))
            throw new InvalidOperationException("Destinatário do teste não informado.");

        var trimmedRecipient = recipient.Trim();
        try
        {
            _ = new MailAddress(trimmedRecipient);
        }
        catch
        {
            throw new InvalidOperationException("Destinatário do teste inválido.");
        }

        return SendWithRetryAsync(
            settings,
            "PLCDataLog - Teste SMTP",
            "Teste de envio SMTP realizado pelo PLCDataLog.",
            null,
            [trimmedRecipient],
            ct);
    }

    public async Task SendCsvAsync(EmailSettings settings, string subject, string body, string attachmentPath, IEnumerable<string> recipients, CancellationToken ct)
    {
        ArgumentNullException.ThrowIfNull(settings);

        if (recipients is null)
            throw new InvalidOperationException("Destinatários não configurados.");

        var distinctRecipients = recipients
            .Select(r => (r ?? string.Empty).Trim())
            .Where(r => !string.IsNullOrWhiteSpace(r))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (distinctRecipients.Length == 0)
            throw new InvalidOperationException("Nenhum destinatário configurado.");

        if (!File.Exists(attachmentPath))
            throw new FileNotFoundException("Arquivo CSV não encontrado.", attachmentPath);

        await SendWithRetryAsync(settings, subject, body, attachmentPath, distinctRecipients, ct);
    }

    public async Task SendMessageAsync(EmailSettings settings, string subject, string body, IEnumerable<string> recipients, CancellationToken ct)
    {
        ArgumentNullException.ThrowIfNull(settings);

        if (recipients is null)
            throw new InvalidOperationException("Destinatários não configurados.");

        var distinctRecipients = recipients
            .Select(r => (r ?? string.Empty).Trim())
            .Where(r => !string.IsNullOrWhiteSpace(r))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (distinctRecipients.Length == 0)
            throw new InvalidOperationException("Nenhum destinatário configurado.");

        await SendWithRetryAsync(settings, subject, body, null, distinctRecipients, ct);
    }

    private static void ValidateSettings(EmailSettings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SmtpHost))
            throw new InvalidOperationException("SMTP Host não configurado.");

        if (settings.SmtpPort is <= 0 or > 65535)
            throw new InvalidOperationException("SMTP Port inválido.");

        if (string.IsNullOrWhiteSpace(settings.SenderEmail))
            throw new InvalidOperationException("E-mail remetente não configurado.");

        try
        {
            _ = new MailAddress(settings.SenderEmail);
        }
        catch
        {
            throw new InvalidOperationException("E-mail remetente inválido.");
        }

        if (settings.SecurityType == EmailSecurityType.StartTls && settings.SmtpPort == 465)
            throw new InvalidOperationException("Configuração incompatível: porta 465 normalmente usa SSL/TLS implícito, não STARTTLS.");

        if (settings.SecurityType == EmailSecurityType.SslTls && settings.SmtpPort == 587)
            throw new InvalidOperationException("Configuração incompatível: porta 587 normalmente usa STARTTLS.");
    }

    private async Task SendWithRetryAsync(EmailSettings settings, string subject, string body, string? attachmentPath, IReadOnlyCollection<string> recipients, CancellationToken ct)
    {
        ValidateSettings(settings);

        var retryCount = Math.Clamp(settings.SmtpRetryCount, 0, 10);
        var retryDelay = TimeSpan.FromSeconds(Math.Clamp(settings.SmtpRetryDelaySeconds, 1, 300));

        Exception? lastException = null;

        for (var attempt = 0; attempt <= retryCount; attempt++)
        {
            ct.ThrowIfCancellationRequested();

            try
            {
                using var message = BuildMessage(settings.SenderEmail, subject, body, attachmentPath, recipients);
                using var client = BuildClient(settings);

                await client.SendMailAsync(message, ct);
                return;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (SmtpException ex)
            {
                lastException = ex;

                if (attempt >= retryCount || !IsTransient(ex))
                    throw new InvalidOperationException($"Falha SMTP ({ex.StatusCode}): {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                lastException = ex;

                if (attempt >= retryCount)
                    throw new InvalidOperationException($"Falha no envio de e-mail: {ex.Message}", ex);
            }

            await Task.Delay(retryDelay, ct);
        }

        throw new InvalidOperationException($"Falha no envio de e-mail: {lastException?.Message}", lastException);
    }

    private static MailMessage BuildMessage(string senderEmail, string subject, string body, string? attachmentPath, IReadOnlyCollection<string> recipients)
    {
        var message = new MailMessage
        {
            From = new MailAddress(senderEmail),
            Subject = subject,
            Body = body,
            IsBodyHtml = false,
        };

        foreach (var recipient in recipients)
            message.To.Add(new MailAddress(recipient));

        if (!string.IsNullOrWhiteSpace(attachmentPath))
            message.Attachments.Add(new Attachment(attachmentPath));

        return message;
    }

    private static SmtpClient BuildClient(EmailSettings settings)
    {
        var timeoutMs = Math.Clamp(settings.SmtpTimeoutSeconds, 5, 300) * 1000;

        var client = new SmtpClient(settings.SmtpHost, settings.SmtpPort)
        {
            DeliveryMethod = SmtpDeliveryMethod.Network,
            UseDefaultCredentials = false,
            EnableSsl = settings.SecurityType != EmailSecurityType.None,
            Timeout = timeoutMs,
        };

        if (!string.IsNullOrWhiteSpace(settings.SmtpUsername) && !string.IsNullOrWhiteSpace(settings.SmtpPassword))
            client.Credentials = new NetworkCredential(settings.SmtpUsername, settings.SmtpPassword);

        return client;
    }

    private static bool IsTransient(SmtpException ex)
    {
        return ex.StatusCode is
            SmtpStatusCode.GeneralFailure or
            SmtpStatusCode.InsufficientStorage or
            SmtpStatusCode.LocalErrorInProcessing or
            SmtpStatusCode.MailboxBusy or
            SmtpStatusCode.MailboxUnavailable or
            SmtpStatusCode.TransactionFailed or
            SmtpStatusCode.ServiceNotAvailable;
    }
}
