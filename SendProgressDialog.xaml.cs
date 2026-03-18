using System;
using System.Windows;
using System.Windows.Threading;

namespace PLCDataLog;

public partial class SendProgressDialog : Window
{
    private DispatcherTimer? _autoCloseTimer;
    private bool _isClosed;

    public SendProgressDialog(string title)
    {
        InitializeComponent();
        TitleText.Text = title;
        StatusText.Text = "Processando...";

        Closed += (_, _) =>
        {
            _isClosed = true;
            _autoCloseTimer?.Stop();
            _autoCloseTimer = null;
        };
    }

    public void SetDone(string message) => SetFinal(message, isError: false);

    public void SetError(string message) => SetFinal(message, isError: true);

    private void SetFinal(string message, bool isError)
    {
        if (_isClosed)
            return;

        Dispatcher.BeginInvoke(() =>
        {
            if (_isClosed)
                return;

            Progress.IsIndeterminate = false;
            Progress.Value = 100;
            StatusText.Text = message;
            CloseButton.IsEnabled = true;

            StartAutoCloseTimer(TimeSpan.FromSeconds(8));
        });
    }

    private void StartAutoCloseTimer(TimeSpan delay)
    {
        _autoCloseTimer?.Stop();
        _autoCloseTimer = new DispatcherTimer(DispatcherPriority.Background)
        {
            Interval = delay,
        };

        _autoCloseTimer.Tick += (_, _) =>
        {
            _autoCloseTimer?.Stop();
            _autoCloseTimer = null;
            SafeClose();
        };

        _autoCloseTimer.Start();
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        _autoCloseTimer?.Stop();
        _autoCloseTimer = null;
        SafeClose();
    }

    private void SafeClose()
    {
        if (_isClosed)
            return;

        try
        {
            Close();
        }
        catch
        {
            // ignore
        }
    }
}
