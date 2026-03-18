using System;
using System.Net.Mail;
using System.Threading;
using System.Windows;
using PLCDataLog.Models;
using PLCDataLog.Services;

namespace PLCDataLog;

public partial class EmailTestDialog : Window
{
    private readonly EmailSettings _settings;
    private readonly SmtpEmailService _smtp;

    public EmailTestDialog(EmailSettings settings, SmtpEmailService smtp)
    {
        InitializeComponent();
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        _smtp = smtp ?? throw new ArgumentNullException(nameof(smtp));
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

    private async void SendButton_Click(object sender, RoutedEventArgs e)
    {
        var recipient = RecipientTextBox.Text.Trim();
        if (!IsValidEmail(recipient))
        {
            StatusText.Text = "E-mail de teste inválido.";
            return;
        }

        if (!IsValidEmail(_settings.SenderEmail))
        {
            StatusText.Text = "E-mail remetente inválido.";
            return;
        }

        if (string.IsNullOrWhiteSpace(_settings.SmtpHost))
        {
            StatusText.Text = "SMTP Host não configurado.";
            return;
        }

        if (_settings.SmtpPort is <= 0 or > 65535)
        {
            StatusText.Text = "SMTP Port inválido.";
            return;
        }

        try
        {
            SendButton.IsEnabled = false;
            Progress.Visibility = Visibility.Visible;
            StatusText.Text = "Enviando e-mail de teste...";

            await _smtp.SendTestEmailAsync(_settings, recipient, CancellationToken.None);

            StatusText.Text = "E-mail de teste enviado com sucesso.";
        }
        catch (SmtpException ex)
        {
            StatusText.Text = $"Falha SMTP ({ex.StatusCode}): {ex.Message}";
        }
        catch (Exception ex)
        {
            if (ex.InnerException is SmtpException smtpEx)
            {
                StatusText.Text = $"Falha SMTP ({smtpEx.StatusCode}): {smtpEx.Message}";
                return;
            }

            StatusText.Text = $"Falha no teste SMTP: {ex.Message}";
        }
        finally
        {
            Progress.Visibility = Visibility.Collapsed;
            SendButton.IsEnabled = true;
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
