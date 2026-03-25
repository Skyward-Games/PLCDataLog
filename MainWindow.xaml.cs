using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Concurrent;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using FluentModbus;
using PLCDataLog.Models;
using PLCDataLog.Services;

using DrawingIcon = System.Drawing.Icon;
using DrawingSystemIcons = System.Drawing.SystemIcons;
using WinForms = System.Windows.Forms;
using DrawingFont = System.Drawing.Font;
using DrawingFontStyle = System.Drawing.FontStyle;
using DrawingSystemFonts = System.Drawing.SystemFonts;

namespace PLCDataLog
{
    public partial class MainWindow : Window
    {
        private CancellationTokenSource? _cts;
        private Task? _pollTask;
        private DailyEmailScheduler? _dailyScheduler;
        private readonly object _autoSendDialogSync = new();

        private readonly SettingsService _settingsService = new(SettingsService.GetDefaultSettingsPath());
        private readonly CsvLogService _csvLogService = new();
        private readonly SmtpEmailService _smtpEmailService = new();
        private AppSettings _settings = new();

        private readonly Dictionary<string, string> _previousValues = new(StringComparer.OrdinalIgnoreCase);

        private readonly DispatcherTimer _clockTimer;

        private readonly DispatcherTimer _autoSaveTimer;
        private bool _autoSavePending;
        private bool _isSavingSettings;

        private string? _lastAutoSendDisplayed;
        private string? _lastYesterdayCsvDisplayed;

        private SendProgressDialog? _autoSendDialog;

        private WinForms.NotifyIcon? _trayIcon;
        private bool _allowExit;

        private bool _isPlcConnected;
        private DateTime _lastAutoSendRefreshReadUtc = DateTime.MinValue;

        private FileSystemWatcher? _recipeWatcher;
        private readonly ConcurrentDictionary<string, byte> _recipeCopyInProgress = new(StringComparer.OrdinalIgnoreCase);
        private int? _lastKnownNokCount;

        public ObservableCollection<PlcValueRow> Rows { get; } = new();

        public MainWindow()
        {
            InitializeComponent();

            _autoSaveTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromMilliseconds(800),
            };
            _autoSaveTimer.Tick += (_, _) =>
            {
                _autoSaveTimer.Stop();
                FlushPendingSettingsSave();
            };

            try
            {
                AutomationLog.Info($"App start. BaseDir='{AppContext.BaseDirectory}' SettingsPath='{SettingsService.GetDefaultSettingsPath()}' DataRoot='{SettingsService.GetDefaultDataRootPath()}' LogPath='{AutomationLog.TryGetCurrentLogPath() ?? "(n/a)"}'.");
            }
            catch
            {
                // ignore
            }

            Loaded += MainWindow_Loaded;
            Closing += MainWindow_Closing;

            try
            {
                RegisterUiAuditLogging();
            }
            catch (Exception ex)
            {
                AutomationLog.Error(ex, "Falha ao registrar auditoria de UI");
            }

            Rows.Add(new PlcValueRow("D21634", "Total_OK", PlcValueType.Word, modbusAddress: 21634, registerCount: 1));
            Rows.Add(new PlcValueRow("D21636", "Total_NOK", PlcValueType.Word, modbusAddress: 21636, registerCount: 1));
            Rows.Add(new PlcValueRow("D21668", "Total_Produzidas", PlcValueType.Word, modbusAddress: 21668, registerCount: 1));
            Rows.Add(new PlcValueRow("D21670", "Quantidade_Meta", PlcValueType.Word, modbusAddress: 21670, registerCount: 1));
            Rows.Add(new PlcValueRow("D21674", "SKU_Nome_Produto", PlcValueType.DWord, modbusAddress: 21674, registerCount: 2));
            Rows.Add(new PlcValueRow("D21700", "Ordem_Producao", PlcValueType.AsciiString, modbusAddress: 21700, registerCount: 16));
            Rows.Add(new PlcValueRow("M167", "Broca_Errada", PlcValueType.Coil, modbusAddress: 2215, registerCount: 1));
            Rows.Add(new PlcValueRow("M168", "Confirma_Erro", PlcValueType.Coil, modbusAddress: 2216, registerCount: 1));
            Rows.Add(new PlcValueRow("M0", "Iniciar_IHM", PlcValueType.Coil, modbusAddress: 2048, registerCount: 1));
            Rows.Add(new PlcValueRow("M1", "Interromper_IHM", PlcValueType.Coil, modbusAddress: 2049, registerCount: 1));

            ValuesGrid.ItemsSource = Rows;

            InitializeScheduleUi();
            LoadSettingsToUi();
            SelectDashboard();

            InitializeRecipeMonitor();

            _clockTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromSeconds(1),
            };
            _clockTimer.Tick += (_, _) => RefreshClockWidget();
            _clockTimer.Start();

            SetConnectionState(false, "Desconectado", "Aguardando conexão com PLC");

            InitializeTrayIcon();
        }

        private void RegisterUiAuditLogging()
        {
            // Janela
            StateChanged += (_, _) => AutomationLog.Ui("WindowStateChanged", $"State={WindowState} Visible={IsVisible}");
            Activated += (_, _) => AutomationLog.Ui("Window", "Activated");
            Deactivated += (_, _) => AutomationLog.Ui("Window", "Deactivated");

            // Navegação
            DashboardNavButton.Click += (_, _) => AutomationLog.Ui("Nav", "Dashboard");
            SettingsNavButton.Click += (_, _) => AutomationLog.Ui("Nav", "Settings");

            // PLC
            StartButton.Click += (_, _) => AutomationLog.Ui("PLC", "Start polling");
            StopButton.Click += (_, _) => AutomationLog.Ui("PLC", "Stop polling");

            PlcHostTextBox.LostKeyboardFocus += (_, _) => AutomationLog.Ui("PLC", $"Host={PlcHostTextBox.Text}");
            PlcPortTextBox.LostKeyboardFocus += (_, _) => AutomationLog.Ui("PLC", $"Port={PlcPortTextBox.Text}");
            PollIntervalTextBox.LostKeyboardFocus += (_, _) => AutomationLog.Ui("PLC", $"PollIntervalMs={PollIntervalTextBox.Text}");
            UnitIdTextBox.LostKeyboardFocus += (_, _) => AutomationLog.Ui("PLC", $"UnitId={UnitIdTextBox.Text}");

            // SMTP
            SenderEmailTextBox.LostKeyboardFocus += (_, _) => AutomationLog.Ui("SMTP", $"Sender={SenderEmailTextBox.Text}");
            SmtpHostTextBox.LostKeyboardFocus += (_, _) => AutomationLog.Ui("SMTP", $"Host={SmtpHostTextBox.Text}");
            SmtpPortTextBox.LostKeyboardFocus += (_, _) => AutomationLog.Ui("SMTP", $"Port={SmtpPortTextBox.Text}");
            SmtpUsernameTextBox.LostKeyboardFocus += (_, _) => AutomationLog.Ui("SMTP", $"Username={SmtpUsernameTextBox.Text}");
            SecurityTypeComboBox.SelectionChanged += (_, _) => AutomationLog.Ui("SMTP", $"Security={((SecurityTypeComboBox.SelectedItem as ComboBoxItem)?.Tag as string) ?? string.Empty}");

            SmtpTimeoutTextBox.LostKeyboardFocus += (_, _) => AutomationLog.Ui("SMTP", $"TimeoutSec={SmtpTimeoutTextBox.Text}");
            SmtpRetryCountTextBox.LostKeyboardFocus += (_, _) => AutomationLog.Ui("SMTP", $"RetryCount={SmtpRetryCountTextBox.Text}");
            SmtpRetryDelayTextBox.LostKeyboardFocus += (_, _) => AutomationLog.Ui("SMTP", $"RetryDelaySec={SmtpRetryDelayTextBox.Text}");

            // (não logar conteúdo de senha)
            SmtpPasswordBox.PasswordChanged += (_, _) => AutomationLog.Ui("SMTP", "Password changed");
            SmtpPasswordTextBox.TextChanged += (_, _) => AutomationLog.Ui("SMTP", "Password(text) changed");

            // Destinatários
            NewRecipientTextBox.LostKeyboardFocus += (_, _) => AutomationLog.Ui("Recipients", $"Typing={NewRecipientTextBox.Text}");

            // Botões de ações
            TestEmailButton.Click += (_, _) => AutomationLog.Ui("Email", "Test SMTP");
            SendTodayCsvButton.Click += (_, _) => AutomationLog.Ui("Email", "Send Today CSV (manual)");

            // Automação
            AutoSendEnabledCheckBox.Checked += (_, _) => AutomationLog.Ui("AutoSend", "Enabled checked");
            AutoSendEnabledCheckBox.Unchecked += (_, _) => AutomationLog.Ui("AutoSend", "Enabled unchecked");
            AutoSendHourComboBox.SelectionChanged += (_, _) => AutomationLog.Ui("AutoSend", $"Hour={AutoSendHourComboBox.SelectedItem}");
            AutoSendMinuteComboBox.SelectionChanged += (_, _) => AutomationLog.Ui("AutoSend", $"Minute={AutoSendMinuteComboBox.SelectedItem}");
            AutoSendTimeZoneComboBox.SelectionChanged += (_, _) => AutomationLog.Ui("AutoSend", $"TimeZone={AutoSendTimeZoneComboBox.SelectedValue}");
            AutoSendRetryMinutesTextBox.LostKeyboardFocus += (_, _) => AutomationLog.Ui("AutoSend", $"RetryMinutes={AutoSendRetryMinutesTextBox.Text}");
            ResetLastAutoSendButton.Click += (_, _) => AutomationLog.Ui("AutoSend", "Reset Last Auto Send");
            ResetYesterdayCsvButton.Click += (_, _) => AutomationLog.Ui("AutoSend", "Reset Yesterday CSV");

            // Backup de rede
            NetworkBackupEnabledCheckBox.Checked += (_, _) => AutomationLog.Ui("NetworkBackup", "Enabled checked");
            NetworkBackupEnabledCheckBox.Unchecked += (_, _) => AutomationLog.Ui("NetworkBackup", "Enabled unchecked");
            NetworkBackupFolderTextBox.LostKeyboardFocus += (_, _) => AutomationLog.Ui("NetworkBackup", $"Folder={NetworkBackupFolderTextBox.Text}");
            BrowseNetworkBackupFolderButton.Click += (_, _) => AutomationLog.Ui("NetworkBackup", "Browse folder");
            NetworkBackupUsernameTextBox.LostKeyboardFocus += (_, _) => AutomationLog.Ui("NetworkBackup", $"Username={NetworkBackupUsernameTextBox.Text}");
            NetworkBackupPasswordBox.PasswordChanged += (_, _) => AutomationLog.Ui("NetworkBackup", "Password changed");
            NetworkBackupPasswordTextBox.TextChanged += (_, _) => AutomationLog.Ui("NetworkBackup", "Password(text) changed");

            // O botão "Testar Acesso" não tem x:Name no XAML.
            // O log dele é feito dentro do handler TestNetworkBackup_Click.
        }

        private void HookSettingsAutoSave()
        {
            AutoSendEnabledCheckBox.Checked += (_, _) => RequestSettingsAutoSave();
            AutoSendEnabledCheckBox.Unchecked += (_, _) => RequestSettingsAutoSave();

            AutoSendHourComboBox.SelectionChanged += (_, _) => RequestSettingsAutoSave();
            AutoSendMinuteComboBox.SelectionChanged += (_, _) => RequestSettingsAutoSave();
            AutoSendTimeZoneComboBox.SelectionChanged += (_, _) => RequestSettingsAutoSave();

            AutoSendRetryMinutesTextBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            AutoSendRetryMinutesTextBox.TextChanged += (_, _) => RequestSettingsAutoSave();

            NetworkBackupEnabledCheckBox.Checked += (_, _) => RequestSettingsAutoSave();
            NetworkBackupEnabledCheckBox.Unchecked += (_, _) => RequestSettingsAutoSave();
            NetworkBackupFolderTextBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            NetworkBackupFolderTextBox.TextChanged += (_, _) => RequestSettingsAutoSave();
            NetworkBackupUsernameTextBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            NetworkBackupUsernameTextBox.TextChanged += (_, _) => RequestSettingsAutoSave();
            NetworkBackupPasswordBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            NetworkBackupPasswordBox.PasswordChanged += (_, _) => RequestSettingsAutoSave();
            NetworkBackupPasswordTextBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            NetworkBackupPasswordTextBox.TextChanged += (_, _) => RequestSettingsAutoSave();

            PlcHostTextBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            PlcHostTextBox.TextChanged += (_, _) => RequestSettingsAutoSave();
            PlcPortTextBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            PlcPortTextBox.TextChanged += (_, _) => RequestSettingsAutoSave();
            PollIntervalTextBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            PollIntervalTextBox.TextChanged += (_, _) => RequestSettingsAutoSave();
            UnitIdTextBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            UnitIdTextBox.TextChanged += (_, _) => RequestSettingsAutoSave();

            SenderEmailTextBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            SenderEmailTextBox.TextChanged += (_, _) => RequestSettingsAutoSave();
            SmtpHostTextBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            SmtpHostTextBox.TextChanged += (_, _) => RequestSettingsAutoSave();
            SmtpPortTextBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            SmtpPortTextBox.TextChanged += (_, _) => RequestSettingsAutoSave();
            SmtpUsernameTextBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            SmtpUsernameTextBox.TextChanged += (_, _) => RequestSettingsAutoSave();
            SmtpPasswordBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            SmtpPasswordBox.PasswordChanged += (_, _) => RequestSettingsAutoSave();
            SmtpPasswordTextBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            SmtpPasswordTextBox.TextChanged += (_, _) => RequestSettingsAutoSave();
            SecurityTypeComboBox.SelectionChanged += (_, _) => RequestSettingsAutoSave();
            
            SmtpTimeoutTextBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            SmtpTimeoutTextBox.TextChanged += (_, _) => RequestSettingsAutoSave();
            SmtpRetryCountTextBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            SmtpRetryCountTextBox.TextChanged += (_, _) => RequestSettingsAutoSave();
            SmtpRetryDelayTextBox.LostKeyboardFocus += (_, _) => RequestSettingsAutoSave();
            SmtpRetryDelayTextBox.TextChanged += (_, _) => RequestSettingsAutoSave();
        }

        private void RequestSettingsAutoSave()
        {
            if (!IsLoaded)
                return;

            _autoSavePending = true;
            _autoSaveTimer.Stop();
            _autoSaveTimer.Start();
        }

        private void FlushPendingSettingsSave()
        {
            if (!_autoSavePending)
                return;

            if (_isSavingSettings)
                return;

            try
            {
                _isSavingSettings = true;
                _autoSavePending = false;
                SaveSettingsFromUiAndPersist();
            }
            finally
            {
                _isSavingSettings = false;
            }
        }

        private void SaveSettingsFromUiAndPersist()
        {
            if (!IsLoaded)
            {
                AutomationLog.Info("SaveSettingsFromUiAndPersist: skipped (not loaded).");
                return;
            }

            _autoSaveTimer.Stop();
            _autoSavePending = false;

            AutomationLog.Info($"SaveSettingsFromUiAndPersist: start UI state: Hour={AutoSendHourComboBox.SelectedItem} Minute={AutoSendMinuteComboBox.SelectedItem}.");

            SaveUiToSettings();

            AutomationLog.Info($"SaveSettingsFromUiAndPersist: after SaveUiToSettings DailySendTime='{_settings.EmailAutomation.DailySendTime}'.");

            if (!TryPersistSettings(out var error))
            {
                StatusTextBlock.Text = $"Falha ao salvar configurações: {error}";
                AutomationLog.Error($"SaveSettingsFromUiAndPersist: persist failed. error={error}");
                return;
            }

            AutomationLog.Info($"SaveSettingsFromUiAndPersist: persisted OK. DailySendTime='{_settings.EmailAutomation.DailySendTime}'.");

            UpdateNetworkBackupStatus(_settings.NetworkBackup.Enabled
                ? "Backup de rede habilitado. Novos logs serão copiados automaticamente."
                : "Backup de rede desativado.");
        }

        private bool TryPersistSettings(out string? error)
        {
            try
            {
                _settingsService.Save(_settings);
                AutomationLog.Info("Settings persisted to disk.");
                error = null;
                return true;
            }
            catch (Exception ex)
            {
                AutomationLog.Error(ex, "Falha ao persistir settings");
                error = ex.Message;
                return false;
            }
        }

        private void SyncVisiblePasswordBoxesToPasswordBoxes()
        {
            if (SmtpPasswordTextBox.Visibility == Visibility.Visible)
                SmtpPasswordBox.Password = SmtpPasswordTextBox.Text;

            if (NetworkBackupPasswordTextBox.Visibility == Visibility.Visible)
                NetworkBackupPasswordBox.Password = NetworkBackupPasswordTextBox.Text;
        }

        private void InitializeScheduleUi()
        {
            AutoSendHourComboBox.ItemsSource = Enumerable.Range(0, 24).Select(h => h.ToString("00", CultureInfo.InvariantCulture)).ToArray();
            AutoSendMinuteComboBox.ItemsSource = Enumerable.Range(0, 60).Select(m => m.ToString("00", CultureInfo.InvariantCulture)).ToArray();

            AutoSendHourComboBox.SelectionChanged += (_, _) => RefreshClockWidget();
            AutoSendMinuteComboBox.SelectionChanged += (_, _) => RefreshClockWidget();
            AutoSendTimeZoneComboBox.SelectionChanged += (_, _) => RefreshClockWidget();

            var tzOptions = new List<TimeZoneOption>
            {
                TimeZoneOption.WindowsLocal,
            };

            tzOptions.AddRange(
                TimeZoneInfo.GetSystemTimeZones()
                    .OrderBy(z => z.BaseUtcOffset)
                    .ThenBy(z => z.DisplayName)
                    .Select(z => new TimeZoneOption(z.Id, z.DisplayName)));

            AutoSendTimeZoneComboBox.ItemsSource = tzOptions;
            AutoSendTimeZoneComboBox.DisplayMemberPath = nameof(TimeZoneOption.Display);
            AutoSendTimeZoneComboBox.SelectedValuePath = nameof(TimeZoneOption.Id);
        }

        private void MainWindow_Loaded(object? sender, RoutedEventArgs e)
        {
            var area = SystemParameters.WorkArea;

            // widescreen: ocupa mais tela
            Width = Math.Round(area.Width * 0.78);
            Height = Math.Round(area.Height * 0.78);

            MinWidth = Math.Min(1080, area.Width);
            MinHeight = Math.Min(680, area.Height);

            Dispatcher.BeginInvoke(() =>
            {
                try
                {
                    ValuesGrid.Height = double.NaN; // deixa o layout decidir; scroll fica no host
                }
                catch
                {
                    // ignore
                }
            });

            MarkAppOpenedToday();

            _dailyScheduler ??= new DailyEmailScheduler(
                _settingsService,
                SendDailyCsvAsync,
                onSendingChanged: isSending =>
                {
                    Dispatcher.BeginInvoke(() =>
                    {
                        lock (_autoSendDialogSync)
                        {
                            if (!IsForegroundUiAvailable())
                            {
                                if (isSending)
                                    ShowTrayBalloon("PLCDataLog", "Envio automático em segundo plano iniciado.", WinForms.ToolTipIcon.Info);
                                else
                                    ShowTrayBalloon("PLCDataLog", "Envio automático concluído em segundo plano.", WinForms.ToolTipIcon.Info);

                                _autoSendDialog = null;
                                return;
                            }

                            if (isSending)
                            {
                                if (_autoSendDialog is null)
                                {
                                    _autoSendDialog = new SendProgressDialog("Enviando e-mail diário...")
                                    {
                                        Owner = this,
                                        WindowStartupLocation = WindowStartupLocation.CenterOwner,
                                    };

                                    _autoSendDialog.Show();
                                }
                            }
                            else
                            {
                                if (_autoSendDialog is not null)
                                {
                                    _autoSendDialog.SetDone("Envio automático concluído.");
                                    _autoSendDialog = null;
                                }
                            }
                        }
                    });
                },
                onError: message =>
                {
                    Dispatcher.BeginInvoke(() =>
                    {
                        lock (_autoSendDialogSync)
                        {
                            if (!IsForegroundUiAvailable())
                            {
                                ShowTrayBalloon("PLCDataLog", $"Falha no envio automático: {message}", WinForms.ToolTipIcon.Error);
                                _autoSendDialog = null;
                                return;
                            }

                            if (_autoSendDialog is null)
                            {
                                _autoSendDialog = new SendProgressDialog("Enviando e-mail diário...")
                                {
                                    Owner = this,
                                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                                };
                                _autoSendDialog.Show();
                            }

                            _autoSendDialog.SetError($"Falha no envio automático: {message}");
                            _autoSendDialog = null;
                        }
                    });
                });

            _dailyScheduler.Start();

            RefreshClockWidget();
        }

        private bool IsForegroundUiAvailable() => IsVisible && WindowState != WindowState.Minimized;

        private void MarkAppOpenedToday()
        {
            try
            {
                var tz = ResolveTimeZoneFromSettings(_settings.EmailAutomation.TimeZoneId);
                var today = TimeZoneInfo.ConvertTime(DateTimeOffset.Now, tz).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

                if (string.Equals(_settings.EmailAutomation.LastAppOpenDate, today, StringComparison.Ordinal))
                    return;

                _settings.EmailAutomation.LastAppOpenDate = today;
                _settingsService.Save(_settings);
            }
            catch
            {
                // ignore
            }
        }

        private void InitializeTrayIcon()
        {
            if (_trayIcon is not null)
                return;

            var icon = DrawingSystemIcons.Application;
            try
            {
                // tenta pegar o ícone do executável (se houver um definido no projeto)
                var exePath = Environment.ProcessPath;
                if (!string.IsNullOrWhiteSpace(exePath))
                    icon = DrawingIcon.ExtractAssociatedIcon(exePath) ?? DrawingSystemIcons.Application;
            }
            catch
            {
                // ignore
            }

            var menu = new WinForms.ContextMenuStrip();

            var showItem = new WinForms.ToolStripMenuItem("Mostrar")
            {
                Font = new DrawingFont(DrawingSystemFonts.DefaultFont, DrawingFontStyle.Bold),
            };
            showItem.Click += (_, _) => Dispatcher.Invoke(ShowFromTray);

            var hideItem = new WinForms.ToolStripMenuItem("Ocultar");
            hideItem.Click += (_, _) => Dispatcher.Invoke(HideToTray);

            var exitItem = new WinForms.ToolStripMenuItem("Sair");
            exitItem.Click += (_, _) => Dispatcher.Invoke(ExitFromTray);

            menu.Items.Add(showItem);
            menu.Items.Add(hideItem);
            menu.Items.Add(new WinForms.ToolStripSeparator());
            menu.Items.Add(exitItem);

            _trayIcon = new WinForms.NotifyIcon
            {
                Icon = icon,
                Text = "PLCDataLog",
                Visible = true,
                ContextMenuStrip = menu,
            };

            _trayIcon.DoubleClick += (_, _) => Dispatcher.Invoke(ShowFromTray);

            StateChanged += (_, _) =>
            {
                if (WindowState == WindowState.Minimized)
                    HideToTray();
            };
        }

        private void ShowFromTray()
        {
            AutomationLog.Ui("Tray", "Show");
            Show();
            if (WindowState == WindowState.Minimized)
                WindowState = WindowState.Normal;
            Activate();
            Topmost = true; // traz para frente
            Topmost = false;
            Focus();
        }

        private void HideToTray()
        {
            AutomationLog.Ui("Tray", "Hide");
            FlushPendingSettingsSave();
            WindowState = WindowState.Minimized;
            Hide();
        }

        private void ExitFromTray()
        {
            AutomationLog.Ui("Tray", "Exit");
            _allowExit = true;
            Close();
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            // Em ambiente industrial, normalmente o app roda em segundo plano.
            // Ao fechar: manda para o tray. Para encerrar de fato, use o menu "Sair" do tray.
            if (!_allowExit)
            {
                FlushPendingSettingsSave();
                e.Cancel = true;
                HideToTray();
                return;
            }

            FlushPendingSettingsSave();

            _dailyScheduler?.Dispose();
            _dailyScheduler = null;

            if (_recipeWatcher is not null)
            {
                try
                {
                    _recipeWatcher.EnableRaisingEvents = false;
                    _recipeWatcher.Dispose();
                }
                catch { }

                _recipeWatcher = null;
            }

            _cts?.Cancel();

            _clockTimer.Stop();

            if (_trayIcon is not null)
            {
                try
                {
                    _trayIcon.Visible = false;
                    _trayIcon.Dispose();
                }
                catch { }

                _trayIcon = null;
            }
        }

        private void DashboardNavButton_Click(object sender, RoutedEventArgs e) => SelectDashboard();

        private void SettingsNavButton_Click(object sender, RoutedEventArgs e) => SelectSettings();

        private void SelectDashboard()
        {
            FlushPendingSettingsSave();
            DashboardPanel.Visibility = Visibility.Visible;
            SettingsPanel.Visibility = Visibility.Collapsed;

            DashboardNavButton.Tag = "Selected";
            SettingsNavButton.Tag = null;

            PageTitleTextBlock.Text = (string)FindResource("MenuDashboard");
        }

        private void SelectSettings()
        {
            FlushPendingSettingsSave();
            DashboardPanel.Visibility = Visibility.Collapsed;
            SettingsPanel.Visibility = Visibility.Visible;

            DashboardNavButton.Tag = null;
            SettingsNavButton.Tag = "Selected";

            PageTitleTextBlock.Text = (string)FindResource("MenuSettings");
        }

        private void LoadSettingsToUi()
        {
            _settings = _settingsService.Load();

            SenderEmailTextBox.Text = _settings.Email.SenderEmail;
            SmtpHostTextBox.Text = _settings.Email.SmtpHost;
            SmtpPortTextBox.Text = _settings.Email.SmtpPort.ToString(CultureInfo.InvariantCulture);
            SmtpUsernameTextBox.Text = _settings.Email.SmtpUsername;
            SmtpPasswordBox.Password = _settings.Email.SmtpPassword;

            SmtpTimeoutTextBox.Text = _settings.Email.SmtpTimeoutSeconds.ToString(CultureInfo.InvariantCulture);
            SmtpRetryCountTextBox.Text = _settings.Email.SmtpRetryCount.ToString(CultureInfo.InvariantCulture);
            SmtpRetryDelayTextBox.Text = _settings.Email.SmtpRetryDelaySeconds.ToString(CultureInfo.InvariantCulture);

            App.ApplyThemeMode(_settings.Ui.ThemeMode);

            AutoSendEnabledCheckBox.IsChecked = _settings.EmailAutomation.Enabled;

            var hhmm = string.IsNullOrWhiteSpace(_settings.EmailAutomation.DailySendTime) ? "00:00" : _settings.EmailAutomation.DailySendTime.Trim();
            if (!TryParseHhMm(hhmm, out var hour, out var minute))
            {
                hour = 0;
                minute = 0;
                hhmm = "00:00";
            }

            AutoSendHourComboBox.SelectedItem = hour.ToString("00", CultureInfo.InvariantCulture);
            AutoSendMinuteComboBox.SelectedItem = minute.ToString("00", CultureInfo.InvariantCulture);

            AutoSendRetryMinutesTextBox.Text = _settings.EmailAutomation.RetryIntervalMinutes.ToString(CultureInfo.InvariantCulture);

            LastAutoSendTextBlock.Text = string.IsNullOrWhiteSpace(_settings.EmailAutomation.LastSuccessfulSendDate)
                ? "Último envio: (nunca)"
                : $"Último envio: {_settings.EmailAutomation.LastSuccessfulSendDate}";

            var tzId = string.IsNullOrWhiteSpace(_settings.EmailAutomation.TimeZoneId)
                ? TimeZoneOption.WindowsLocal.Id
                : _settings.EmailAutomation.TimeZoneId.Trim();

            AutoSendTimeZoneComboBox.SelectedValue = tzId;
            if (AutoSendTimeZoneComboBox.SelectedItem is null)
                AutoSendTimeZoneComboBox.SelectedValue = TimeZoneOption.WindowsLocal.Id;

            PlcHostTextBox.Text = string.IsNullOrWhiteSpace(_settings.PlcConnection.Host) ? "127.0.0.1" : _settings.PlcConnection.Host;
            PlcPortTextBox.Text = _settings.PlcConnection.Port <= 0 ? "502" : _settings.PlcConnection.Port.ToString(CultureInfo.InvariantCulture);
            PollIntervalTextBox.Text = _settings.PlcConnection.PollIntervalMs < 50 ? "500" : _settings.PlcConnection.PollIntervalMs.ToString(CultureInfo.InvariantCulture);
            UnitIdTextBox.Text = _settings.PlcConnection.UnitId.ToString(CultureInfo.InvariantCulture);

            NetworkBackupEnabledCheckBox.IsChecked = _settings.NetworkBackup.Enabled;
            NetworkBackupFolderTextBox.Text = _settings.NetworkBackup.TargetFolder;
            NetworkBackupUsernameTextBox.Text = _settings.NetworkBackup.Username;
            NetworkBackupPasswordBox.Password = _settings.NetworkBackup.Password;

            RecipeMonitorEnabledCheckBox.IsChecked = _settings.RecipeMonitor.Enabled;
            RecipeSourceFolderTextBox.Text = _settings.RecipeMonitor.SourceFolder;
            RecipeTargetFolderTextBox.Text = _settings.RecipeMonitor.TargetFolder;
            RecipeNetworkUsernameTextBox.Text = _settings.RecipeMonitor.Username;
            RecipeNetworkPasswordBox.Password = _settings.RecipeMonitor.Password;

            var securityTag = _settings.Email.SecurityType.ToString();
            foreach (var item in SecurityTypeComboBox.Items.OfType<ComboBoxItem>())
            {
                if (string.Equals(item.Tag as string, securityTag, StringComparison.OrdinalIgnoreCase))
                {
                    SecurityTypeComboBox.SelectedItem = item;
                    break;
                }
            }

            RecipientsListBox.ItemsSource = _settings.Email.Recipients;

            RefreshClockWidget();
            UpdateNetworkBackupStatus("Backup de rede pronto para uso.");
            UpdateRecipeMonitorStatus(_settings.RecipeMonitor.Enabled
                ? "Monitoramento de receitas habilitado."
                : "Monitoramento de receitas desativado.");
        }

        private void SaveUiToSettings()
        {
            SyncVisiblePasswordBoxesToPasswordBoxes();

            _settings.Email.SenderEmail = SenderEmailTextBox.Text.Trim();
            _settings.Email.SmtpHost = SmtpHostTextBox.Text.Trim();

            if (int.TryParse(SmtpPortTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var port))
                _settings.Email.SmtpPort = port;

            _settings.Email.SmtpUsername = SmtpUsernameTextBox.Text.Trim();
            _settings.Email.SmtpPassword = SmtpPasswordBox.Password;

            if (int.TryParse(SmtpTimeoutTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var timeout))
                _settings.Email.SmtpTimeoutSeconds = timeout;

            if (int.TryParse(SmtpRetryCountTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var retryCount))
                _settings.Email.SmtpRetryCount = retryCount;

            if (int.TryParse(SmtpRetryDelayTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var retryDelay))
                _settings.Email.SmtpRetryDelaySeconds = retryDelay;

            if (string.IsNullOrWhiteSpace(_settings.Ui.ThemeMode))
                _settings.Ui.ThemeMode = "System";

            App.ApplyThemeMode(_settings.Ui.ThemeMode);

            _settings.EmailAutomation.Enabled = AutoSendEnabledCheckBox.IsChecked == true;

            var hour = (AutoSendHourComboBox.SelectedItem as string) ?? "00";
            var minute = (AutoSendMinuteComboBox.SelectedItem as string) ?? "00";
            var timeText = $"{hour}:{minute}";
            if (!TimeSpan.TryParseExact(timeText, "hh\\:mm", CultureInfo.InvariantCulture, out _))
                timeText = "00:00";
            _settings.EmailAutomation.DailySendTime = timeText;

            AutomationLog.Info($"SaveUiToSettings: DailySendTime set from UI values Hour='{hour}' Minute='{minute}' -> '{timeText}'.");

            var tzId = (AutoSendTimeZoneComboBox.SelectedValue as string) ?? TimeZoneOption.WindowsLocal.Id;
            _settings.EmailAutomation.TimeZoneId = string.Equals(tzId, TimeZoneOption.WindowsLocal.Id, StringComparison.OrdinalIgnoreCase) ? null : tzId;

            if (int.TryParse(AutoSendRetryMinutesTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var retryMin))
                _settings.EmailAutomation.RetryIntervalMinutes = Math.Clamp(retryMin, 1, 24 * 60);

            if (int.TryParse(PlcPortTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var plcPort))
                _settings.PlcConnection.Port = Math.Clamp(plcPort, 1, 65535);

            _settings.PlcConnection.Host = (PlcHostTextBox.Text ?? string.Empty).Trim();

            if (int.TryParse(PollIntervalTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var pollMs))
                _settings.PlcConnection.PollIntervalMs = Math.Max(50, pollMs);

            if (byte.TryParse(UnitIdTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var plcUnitId))
                _settings.PlcConnection.UnitId = plcUnitId;

            _settings.NetworkBackup.Enabled = NetworkBackupEnabledCheckBox.IsChecked == true;
            _settings.NetworkBackup.TargetFolder = (NetworkBackupFolderTextBox.Text ?? string.Empty).Trim();
            _settings.NetworkBackup.Username = (NetworkBackupUsernameTextBox.Text ?? string.Empty).Trim();
            _settings.NetworkBackup.Password = NetworkBackupPasswordBox.Password;

            _settings.RecipeMonitor.Enabled = RecipeMonitorEnabledCheckBox.IsChecked == true;
            _settings.RecipeMonitor.SourceFolder = (RecipeSourceFolderTextBox.Text ?? string.Empty).Trim();
            _settings.RecipeMonitor.TargetFolder = (RecipeTargetFolderTextBox.Text ?? string.Empty).Trim();
            _settings.RecipeMonitor.Username = (RecipeNetworkUsernameTextBox.Text ?? string.Empty).Trim();
            _settings.RecipeMonitor.Password = RecipeNetworkPasswordBox.Password;

            LastAutoSendTextBlock.Text = string.IsNullOrWhiteSpace(_settings.EmailAutomation.LastSuccessfulSendDate)
                ? "Último envio: (nunca)"
                : $"Último envio: {_settings.EmailAutomation.LastSuccessfulSendDate}";

            if (SecurityTypeComboBox.SelectedItem is ComboBoxItem sec)
            {
                var tag = (sec.Tag as string) ?? "StartTls";
                if (Enum.TryParse<EmailSecurityType>(tag, ignoreCase: true, out var parsed))
                    _settings.Email.SecurityType = parsed;
            }

            RefreshClockWidget();
        }

        private void TestEmailButton_Click(object sender, RoutedEventArgs e)
        {
            AutomationLog.Ui("Email", "Test SMTP dialog open");
            SaveUiToSettings();

            var dialog = new EmailTestDialog(_settings.Email, _smtpEmailService)
            {
                Owner = this,
            };

            dialog.ShowDialog();
        }

        private void SaveSmtpSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            SaveUiToSettings();
            if (!TryPersistSettings(out var error))
            {
                StatusTextBlock.Text = $"Falha ao salvar configurações SMTP: {error}";
                return;
            }

            System.Windows.MessageBox.Show(this, "Configurações SMTP salvas com sucesso.", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void SaveAutomationSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            SaveUiToSettings();
            if (!TryPersistSettings(out var error))
            {
                StatusTextBlock.Text = $"Falha ao salvar automação: {error}";
                return;
            }

            System.Windows.MessageBox.Show(this, "Configurações de automação salvas com sucesso.", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void SaveNetworkSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            SaveUiToSettings();
            if (!TryPersistSettings(out var error))
            {
                StatusTextBlock.Text = $"Falha ao salvar backup de rede: {error}";
                return;
            }

            System.Windows.MessageBox.Show(this, "Configurações de backup de rede salvas com sucesso.", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void SaveRecipeSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            SaveUiToSettings();
            if (!TryPersistSettings(out var error))
            {
                StatusTextBlock.Text = $"Falha ao salvar monitoramento de receita: {error}";
                return;
            }

            InitializeRecipeMonitor();
            System.Windows.MessageBox.Show(this, "Configurações de receita salvas com sucesso.", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void TestRecipeDestination_Click(object sender, RoutedEventArgs e)
        {
            SaveUiToSettings();

            var target = _settings.RecipeMonitor.TargetFolder;
            if (string.IsNullOrWhiteSpace(target))
            {
                System.Windows.MessageBox.Show(this, "Informe a pasta destino da receita.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                using var _ = NetworkShareConnection.ConnectIfNeeded(
                    target,
                    _settings.RecipeMonitor.Username,
                    _settings.RecipeMonitor.Password);

                Directory.CreateDirectory(target);
                var testFile = Path.Combine(target, $"recipe_test_{Guid.NewGuid():N}.tmp");
                File.WriteAllText(testFile, "test");
                File.Delete(testFile);

                System.Windows.MessageBox.Show(this, "Destino de receita validado com sucesso.", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(this, $"Falha ao acessar destino da receita.\n\n{ex.Message}", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddRecipientButton_Click(object sender, RoutedEventArgs e)
        {
            var email = NewRecipientTextBox.Text.Trim();
            AutomationLog.Ui("Recipients", $"Add clicked value='{email}'");
            if (string.IsNullOrWhiteSpace(email))
                return;

            if (!IsValidEmail(email))
                return;

            if (_settings.Email.Recipients.Any(r => string.Equals(r, email, StringComparison.OrdinalIgnoreCase)))
                return;

            _settings.Email.Recipients.Add(email);

            RecipientsListBox.ItemsSource = null;
            RecipientsListBox.ItemsSource = _settings.Email.Recipients;

            NewRecipientTextBox.Text = string.Empty;

            if (!TryPersistSettings(out _))
                StatusTextBlock.Text = "Falha ao salvar destinatários.";
        }

        private void RemoveRecipientButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedRecipient = RecipientsListBox.SelectedItem as string;
            AutomationLog.Ui("Recipients", $"Remove clicked selected='{selectedRecipient ?? string.Empty}'");

            var removed = false;

            // remove via botão dentro do item (DataTemplate)
            if (sender is System.Windows.Controls.Button btn)
            {
                var emailFromContext = btn.DataContext as string;
                if (!string.IsNullOrWhiteSpace(emailFromContext))
                {
                    var toRemove = _settings.Email.Recipients
                        .FirstOrDefault(r => string.Equals(r, emailFromContext, StringComparison.OrdinalIgnoreCase));

                    if (toRemove is not null)
                    {
                        _settings.Email.Recipients.Remove(toRemove);
                        removed = true;
                    }
                }
            }

            if (!removed)
            {
                // fallback: remove item selecionado
                if (RecipientsListBox.SelectedItem is not string selected)
                    return;

                removed = _settings.Email.Recipients.Remove(selected);
            }

            if (!removed)
                return;

            RecipientsListBox.ItemsSource = null;
            RecipientsListBox.ItemsSource = _settings.Email.Recipients;

            if (!TryPersistSettings(out _))
                StatusTextBlock.Text = "Falha ao salvar destinatários.";
        }

        private void RefreshClockWidget()
        {
            if (!IsLoaded)
                return;

            MarkAppOpenedToday();

            var tz = GetSelectedTimeZone();
            var now = TimeZoneInfo.ConvertTime(DateTimeOffset.Now, tz);

            HeaderClockTextBlock.Text = now.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
            HeaderTimeZoneTextBlock.Text = GetSelectedTimeZoneLabel(tz);

            var dailyTime = GetSelectedDailySendTimeOrDefault();
            var next = NextOccurrence(dailyTime, tz, DateTimeOffset.Now);
            var nextInTz = TimeZoneInfo.ConvertTime(next, tz);
            NextSendTextBlock.Text = nextInTz.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

            // Atualiza automaticamente o "Último envie" quando o scheduler persistir no disco
            // Evita I/O a cada segundo.
            if ((DateTime.UtcNow - _lastAutoSendRefreshReadUtc) >= TimeSpan.FromSeconds(15))
            {
                _lastAutoSendRefreshReadUtc = DateTime.UtcNow;

                try
                {
                    var s = _settingsService.Load();

                    var last = s.EmailAutomation.LastSuccessfulSendDate;
                    if (!string.Equals(_lastAutoSendDisplayed, last, StringComparison.Ordinal))
                    {
                        _lastAutoSendDisplayed = last;
                        LastAutoSendTextBlock.Text = string.IsNullOrWhiteSpace(last)
                            ? "Último envio: (nunca)"
                            : $"Último envio: {last}";
                    }

                    var yesterdayStamp = (s.EmailAutomation.LastYesterdayCsvSentDate ?? string.Empty) + "|" + (s.EmailAutomation.LastYesterdayCsvSentAt ?? string.Empty);
                    if (!string.Equals(_lastYesterdayCsvDisplayed, yesterdayStamp, StringComparison.Ordinal))
                    {
                        _lastYesterdayCsvDisplayed = yesterdayStamp;
                        YesterdayCsvStatusTextBlock.Text = BuildYesterdayCsvStatusText(s);
                    }
                }
                catch
                {
                    // ignore
                }
            }
        }

        private static string BuildYesterdayCsvStatusText(AppSettings settings)
        {
            var tz = ResolveTimeZoneFromSettings(settings.EmailAutomation.TimeZoneId);
            var todayInTz = TimeZoneInfo.ConvertTime(DateTimeOffset.Now, tz).Date;
            var yesterdayInTz = todayInTz.AddDays(-1).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

            if (string.Equals(settings.EmailAutomation.LastYesterdayCsvSentDate, yesterdayInTz, StringComparison.Ordinal))
            {
                if (string.IsNullOrWhiteSpace(settings.EmailAutomation.LastYesterdayCsvSentAt))
                    return $"CSV de ontem ({yesterdayInTz}): enviado";

                return $"CSV de ontem ({yesterdayInTz}): enviado às {settings.EmailAutomation.LastYesterdayCsvSentAt}";
            }

            if (DateTime.TryParseExact(settings.EmailAutomation.LastSuccessfulSendDate, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out var lastTodaySent)
                && lastTodaySent.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) == yesterdayInTz)
            {
                return $"CSV de ontem ({yesterdayInTz}): enviado às {lastTodaySent:HH:mm:ss}";
            }

            return $"CSV de ontem ({yesterdayInTz}): pendente";
        }

        private TimeZoneInfo GetSelectedTimeZone()
        {
            var selected = (AutoSendTimeZoneComboBox.SelectedValue as string) ?? TimeZoneOption.WindowsLocal.Id;
            if (string.Equals(selected, TimeZoneOption.WindowsLocal.Id, StringComparison.OrdinalIgnoreCase))
                return TimeZoneInfo.Local;

            try
            {
                return TimeZoneInfo.FindSystemTimeZoneById(selected);
            }
            catch
            {
                return TimeZoneInfo.Local;
            }
        }

        private string GetSelectedTimeZoneLabel(TimeZoneInfo tz)
        {
            var selected = (AutoSendTimeZoneComboBox.SelectedValue as string) ?? TimeZoneOption.WindowsLocal.Id;
            if (string.Equals(selected, TimeZoneOption.WindowsLocal.Id, StringComparison.OrdinalIgnoreCase))
                return "Windows (Local)";

            return tz.StandardName;
        }

        private string GetSelectedDailySendTimeOrDefault()
        {
            var hour = (AutoSendHourComboBox.SelectedItem as string) ?? "00";
            var minute = (AutoSendMinuteComboBox.SelectedItem as string) ?? "00";
            var hhmm = $"{hour}:{minute}";
            return TimeSpan.TryParseExact(hhmm, "hh\\:mm", CultureInfo.InvariantCulture, out _) ? hhmm : "00:00";
        }

        private static bool TryParseHhMm(string hhmm, out int hour, out int minute)
        {
            hour = 0;
            minute = 0;

            if (!TimeSpan.TryParseExact(hhmm, "hh\\:mm", CultureInfo.InvariantCulture, out var ts))
                return false;

            hour = ts.Hours;
            minute = ts.Minutes;
            return true;
        }

        private static DateTimeOffset NextOccurrence(string hhmm, TimeZoneInfo tz, DateTimeOffset now)
        {
            if (!TimeSpan.TryParseExact(hhmm.Trim(), "hh\\:mm", CultureInfo.InvariantCulture, out var ts))
                ts = TimeSpan.Zero;

            var nowInTz = TimeZoneInfo.ConvertTime(now, tz);
            var candidateLocal = new DateTime(nowInTz.Year, nowInTz.Month, nowInTz.Day, 0, 0, 0, DateTimeKind.Unspecified).Add(ts);
            if (candidateLocal <= nowInTz.DateTime)
                candidateLocal = candidateLocal.AddDays(1);

            return new DateTimeOffset(candidateLocal, tz.GetUtcOffset(candidateLocal));
        }

        private async void StartButton_Click(object sender, RoutedEventArgs e)
        {
            AutomationLog.Ui("PLC", $"Start requested host='{PlcHostTextBox.Text}' port='{PlcPortTextBox.Text}' pollMs='{PollIntervalTextBox.Text}' unitId='{UnitIdTextBox.Text}'");

            if (_pollTask is not null)
                return;

            var host = (PlcHostTextBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(host))
            {
                StatusTextBlock.Text = "Informe o IP/host do PLC.";
                return;
            }

            if (!int.TryParse(PlcPortTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var plcPort) || plcPort is <= 0 or > 65535)
            {
                StatusTextBlock.Text = "Porta do PLC inválida.";
                return;
            }

            if (!int.TryParse(PollIntervalTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var pollMs) || pollMs < 50)
            {
                StatusTextBlock.Text = (string)FindResource("InvalidPoll");
                return;
            }

            if (!byte.TryParse(UnitIdTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var unitId))
            {
                StatusTextBlock.Text = (string)FindResource("InvalidUnitId");
                return;
            }

            SaveUiToSettings();

            _previousValues.Clear();
            _lastKnownNokCount = null;

            StartButton.IsEnabled = false;
            StopButton.IsEnabled = true;
            StatusTextBlock.Text = (string)FindResource("Starting");
            SetConnectionState(false, "Conectando...", $"Tentando {host}:{plcPort} | UnitId {unitId}");

            _cts = new CancellationTokenSource();
            _pollTask = PollLoopAsync(host, plcPort, unitId, TimeSpan.FromMilliseconds(pollMs), _cts.Token);

            try
            {
                await _pollTask;
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = string.Format((string)FindResource("StoppedWithError"), ex.Message);
                SetConnectionState(false, "Desconectado", ex.Message);
            }
            finally
            {
                _pollTask = null;
                _cts?.Dispose();
                _cts = null;
                StartButton.IsEnabled = true;
                StopButton.IsEnabled = false;
                SetConnectionState(false, "Desconectado", $"Último endpoint: {host}:{plcPort} | UnitId {unitId}");
            }
        }

        private void StopButton_Click(object sender, RoutedEventArgs e)
        {
            AutomationLog.Ui("PLC", "Stop requested");
            SetConnectionState(false, "Desconectando...", "Solicitando parada do polling");
            _cts?.Cancel();
        }

        private async Task PollLoopAsync(string host, int port, byte unitId, TimeSpan interval, CancellationToken ct)
        {
            var reconnectAttempt = 0;
            while (!ct.IsCancellationRequested)
            {
                ModbusTcpClient? client = null;
                try
                {
                    client = new ModbusTcpClient();
                    client.Connect(ResolveEndpoint(host, port), ModbusEndianness.BigEndian);

                    reconnectAttempt = 0;
                    SetConnectionState(true, "Conectado", $"{host}:{port} | UnitId {unitId}");

                    while (!ct.IsCancellationRequested)
                    {
                        var changed = PollOnceAndLog(client, unitId, ct);
                        Dispatcher.Invoke(() => StatusTextBlock.Text = string.Format((string)FindResource("PollingLastOk"), DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture))
                                              + (changed.Count > 0 ? $" | Changes: {string.Join(",", changed)}" : string.Empty));
                        await Task.Delay(interval, ct);
                    }
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    reconnectAttempt++;
                    var delaySeconds = Math.Min(60, 5 * reconnectAttempt);
                    SetConnectionState(false, "Desconectado", $"Tentando reconectar... {ex.Message}");
                    await Task.Delay(TimeSpan.FromSeconds(delaySeconds), ct);
                }
                finally
                {
                    client?.Dispose();
                }
            }

            SetConnectionState(false, "Desconectado", $"{host}:{port} | UnitId {unitId}");
        }

        private IReadOnlyList<string> PollOnceAndLog(ModbusTcpClient client, byte unitId, CancellationToken ct)
        {
            PollOnce(client, unitId, ct);
            var snapshot = Rows.Select(r => (Key: (string)r.Name, Value: (string)(r.Value ?? string.Empty))).ToArray();

            HandleImmediateNokAlert(snapshot);

            var changes = CsvLogService.GetChanges(_previousValues, snapshot);
            if (changes.Count > 0)
            {
                var values = Rows.Select(r => ((string)r.Name, (string)r.Address, (string)r.ValueType.ToString(), (string)(r.Value ?? string.Empty))).ToArray();
                var filePath = _csvLogService.EnsureLogFileForToday();
                _csvLogService.AppendSnapshot(filePath, DateTime.Now, values, changes);
                BackupLogToNetworkIfEnabled(filePath);
            }

            CsvLogService.UpdatePrevious(_previousValues, snapshot);
            return changes;
        }

        private void HandleImmediateNokAlert(IReadOnlyList<(string Key, string Value)> snapshot)
        {
            var nok = snapshot.FirstOrDefault(s => string.Equals(s.Key, "Total_NOK", StringComparison.OrdinalIgnoreCase));
            if (!int.TryParse(nok.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var currentNok))
                return;

            if (_lastKnownNokCount is null)
            {
                _lastKnownNokCount = currentNok;
                return;
            }

            if (currentNok <= _lastKnownNokCount.Value)
                return;

            var increment = currentNok - _lastKnownNokCount.Value;
            _lastKnownNokCount = currentNok;

            _ = Task.Run(async () =>
            {
                try
                {
                    var settings = _settingsService.Load();
                    if (!IsValidEmail(settings.Email.SenderEmail) ||
                        string.IsNullOrWhiteSpace(settings.Email.SmtpHost) ||
                        settings.Email.SmtpPort is <= 0 or > 65535 ||
                        settings.Email.Recipients.Count == 0)
                    {
                        return;
                    }

                    var now = DateTime.Now;
                    var changedSection = string.Join(Environment.NewLine,
                        snapshot.Where(s => s.Key.Contains("NOK", StringComparison.OrdinalIgnoreCase)
                                            || s.Key.Contains("Erro", StringComparison.OrdinalIgnoreCase)
                                            || s.Key.Contains("Broca", StringComparison.OrdinalIgnoreCase)
                                            || s.Key.Contains("Confirma", StringComparison.OrdinalIgnoreCase))
                                .Select(s => $"- {s.Key}: {s.Value}"));

                    if (string.IsNullOrWhiteSpace(changedSection))
                        changedSection = "- Não foi possível identificar campo específico da falha.";

                    var fullLog = string.Join(Environment.NewLine, snapshot.Select(s => $"- {s.Key}: {s.Value}"));

                    var body =
                        $"Alerta de NOK detectado às {now:yyyy-MM-dd HH:mm:ss}." + Environment.NewLine +
                        $"Incremento detectado: +{increment}. Total NOK atual: {currentNok}." + Environment.NewLine + Environment.NewLine +
                        "Possível origem da falha:" + Environment.NewLine +
                        changedSection + Environment.NewLine + Environment.NewLine +
                        "Log atual do ciclo:" + Environment.NewLine +
                        fullLog;

                    await _smtpEmailService.SendMessageAsync(
                        settings.Email,
                        $"PLCDataLog - ALERTA NOK {now:yyyy-MM-dd HH:mm:ss}",
                        body,
                        settings.Email.Recipients,
                        CancellationToken.None);

                    AutomationLog.Info($"Alerta NOK enviado. incremento={increment} total={currentNok}.");
                }
                catch (Exception ex)
                {
                    AutomationLog.Error(ex, "Falha ao enviar alerta imediato de NOK");
                }
            });
        }

        private string EnsureTodayCsvExists()
        {
            var filePath = _csvLogService.EnsureLogFileForToday();
            if (File.Exists(filePath))
                return filePath;

            var values = Rows.Select(r => ((string)r.Name, (string)r.Address, (string)r.ValueType.ToString(), (string)(r.Value ?? string.Empty))).ToArray();
            _csvLogService.AppendSnapshot(filePath, DateTime.Now, values, Array.Empty<string>());
            return filePath;
        }

        private string EnsureCsvExistsForDate(DateTime dateInTz, TimeZoneInfo tz)
        {
            var filePath = BuildCsvPathForDate(dateInTz);
            if (File.Exists(filePath))
                return filePath;

            var values = Rows.Select(r => ((string)r.Name, (string)r.Address, (string)r.ValueType.ToString(), (string)(r.Value ?? string.Empty))).ToArray();
            var nowInTz = TimeZoneInfo.ConvertTime(DateTimeOffset.UtcNow, tz).DateTime;
            var timestamp = nowInTz.Date == dateInTz.Date ? nowInTz : dateInTz.AddHours(12);
            _csvLogService.AppendSnapshot(filePath, timestamp, values, Array.Empty<string>());
            return filePath;
        }

        private async void SendTodayCsvButton_Click(object sender, RoutedEventArgs e)
        {
            AutomationLog.Ui("Email", "Manual send Today CSV clicked");
            var progress = new SendProgressDialog("Enviando CSV do dia (teste/manual)...") { Owner = this };
            string? attachmentSnapshotPath = null;

            try
            {
                SaveUiToSettings();

                progress.Show();

                if (!IsValidEmail(_settings.Email.SenderEmail))
                    throw new InvalidOperationException("E-mail de envio inválido.");
                if (string.IsNullOrWhiteSpace(_settings.Email.SmtpHost))
                    throw new InvalidOperationException("SMTP Host não configurado.");
                if (_settings.Email.SmtpPort is <= 0 or > 65535)
                    throw new InvalidOperationException("SMTP Port inválido.");
                if (_settings.Email.Recipients.Count == 0)
                    throw new InvalidOperationException("Nenhum destinatário configurado.");

                var csvPath = EnsureTodayCsvExists();
                attachmentSnapshotPath = CreateAttachmentSnapshot(csvPath);

                await _smtpEmailService.SendCsvAsync(
                    _settings.Email,
                    $"PLCDataLog - CSV {DateTime.Now:yyyy-MM-dd} (teste)",
                    $"Segue anexado o CSV do dia ({DateTime.Now:yyyy-MM-dd}). Envio manual para teste.",
                    attachmentSnapshotPath,
                    _settings.Email.Recipients,
                    CancellationToken.None);

                progress.SetDone("CSV enviado com sucesso (teste/manual).\nIsso não altera o agendamento automático.");
            }
            catch (Exception ex)
            {
                progress.SetError($"Falha ao enviar CSV (teste/manual): {ex.Message}");
            }
            finally
            {
                TryDeleteFile(attachmentSnapshotPath);
            }
        }

        private void ResetLastAutoSendButton_Click(object sender, RoutedEventArgs e)
        {
            AutomationLog.Ui("AutoSend", "Reset last auto send clicked");
            try
            {
                _settings = _settingsService.Load();
                var tz = ResolveTimeZoneFromSettings(_settings.EmailAutomation.TimeZoneId);
                var todayInTz = TimeZoneInfo.ConvertTime(DateTimeOffset.Now, tz).Date;

                _settings.EmailAutomation.LastSuccessfulSendDate = null;
                _settings.EmailAutomation.SkipPreviousDayCatchUpForDate = todayInTz.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

                // Reset não deve reabrir envio de ontem. Mantemos LastSuccessfulCsvDate.
                if (!TryPersistSettings(out var error))
                {
                    StatusTextBlock.Text = $"Falha ao salvar configurações: {error}";
                    return;
                }

                LastAutoSendTextBlock.Text = "Último envio: (nunca)";
                RefreshClockWidget();
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = $"Falha ao resetar envio automático: {ex.Message}";
            }
        }

        private void ResetYesterdayCsvButton_Click(object sender, RoutedEventArgs e)
        {
            AutomationLog.Ui("AutoSend", "Reset yesterday CSV clicked");
            try
            {
                _settings = _settingsService.Load();
                var tz = ResolveTimeZoneFromSettings(_settings.EmailAutomation.TimeZoneId);
                var todayInTz = TimeZoneInfo.ConvertTime(DateTimeOffset.Now, tz).Date;
                var yesterdayInTz = todayInTz.AddDays(-1);

                _settings.EmailAutomation.LastYesterdayCsvSentDate = null;
                _settings.EmailAutomation.LastYesterdayCsvSentAt = null;

                if (DateTime.TryParseExact(_settings.EmailAutomation.LastSuccessfulCsvDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var lastCsvDate)
                    && lastCsvDate.Date == yesterdayInTz)
                {
                    _settings.EmailAutomation.LastSuccessfulCsvDate = null;
                }

                _settings.EmailAutomation.SkipPreviousDayCatchUpForDate = null;
                _settingsService.Save(_settings);

                YesterdayCsvStatusTextBlock.Text = BuildYesterdayCsvStatusText(_settings);
                RefreshClockWidget();
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = $"Falha ao resetar envio do CSV de ontem: {ex.Message}";
            }
        }

        private void SetConnectionState(bool connected, string title, string details)
        {
            AutomationLog.State("PLC.Connected", connected ? "1" : "0");
            AutomationLog.State("PLC.Status", title);
            AutomationLog.State("PLC.Details", details);

            Dispatcher.Invoke(() =>
            {
                _isPlcConnected = connected;
                ConnectionStatusTextBlock.Text = title;
                ConnectionInfoTextBlock.Text = details;

                DataContentPanel.Visibility = connected ? Visibility.Visible : Visibility.Collapsed;
                DataLockedPanel.Visibility = connected ? Visibility.Collapsed : Visibility.Visible;

                StatusIndicator.Fill = connected ? new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0x10, 0xB9, 0x81)) : System.Windows.Media.Brushes.Gray;

                PlcHostTextBox.IsEnabled = !connected;
                PlcPortTextBox.IsEnabled = !connected;
                PollIntervalTextBox.IsEnabled = !connected;
                UnitIdTextBox.IsEnabled = !connected;
            });
        }

        private void BackupLogToNetworkIfEnabled(string sourceFilePath)
        {
            try
            {
                if (!_settings.NetworkBackup.Enabled)
                {
                    UpdateNetworkBackupStatus("Backup de rede desativado.");
                    return;
                }

                if (string.IsNullOrWhiteSpace(_settings.NetworkBackup.TargetFolder))
                {
                    UpdateNetworkBackupStatus("Pasta de backup de rede não configurada.");
                    return;
                }

                var sourceRoot = Path.Combine(SettingsService.GetDefaultDataRootPath(), "Logs");
                var relative = Path.GetRelativePath(sourceRoot, sourceFilePath);
                if (relative.StartsWith("..", StringComparison.Ordinal) || Path.IsPathRooted(relative))
                    relative = Path.GetFileName(sourceFilePath);
                var targetFilePath = Path.Combine(_settings.NetworkBackup.TargetFolder, relative);
                var targetDir = Path.GetDirectoryName(targetFilePath) ?? _settings.NetworkBackup.TargetFolder;

                using var _ = NetworkShareConnection.ConnectIfNeeded(
                    _settings.NetworkBackup.TargetFolder,
                    _settings.NetworkBackup.Username,
                    _settings.NetworkBackup.Password);

                Directory.CreateDirectory(targetDir);
                File.Copy(sourceFilePath, targetFilePath, overwrite: true);

                UpdateNetworkBackupStatus($"Backup OK: {targetFilePath}");
            }
            catch (Exception ex)
            {
                UpdateNetworkBackupStatus($"Falha no backup de rede: {ex.Message}");
            }
        }

        private void UpdateNetworkBackupStatus(string message)
        {
            Dispatcher.Invoke(() => NetworkBackupStatusTextBlock.Text = message);
        }

        private void UpdateRecipeMonitorStatus(string message)
        {
            Dispatcher.Invoke(() => RecipeMonitorStatusTextBlock.Text = message);
        }

        private void InitializeRecipeMonitor()
        {
            try
            {
                _recipeWatcher?.Dispose();
                _recipeWatcher = null;

                var cfg = _settings.RecipeMonitor;
                if (!cfg.Enabled)
                {
                    UpdateRecipeMonitorStatus("Monitoramento de receitas desativado.");
                    return;
                }

                if (string.IsNullOrWhiteSpace(cfg.SourceFolder) || !Directory.Exists(cfg.SourceFolder))
                {
                    UpdateRecipeMonitorStatus("Pasta origem da receita inválida ou inexistente.");
                    return;
                }

                _recipeWatcher = new FileSystemWatcher(cfg.SourceFolder)
                {
                    IncludeSubdirectories = false,
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime,
                    Filter = "*.*",
                    EnableRaisingEvents = true,
                };

                _recipeWatcher.Created += RecipeWatcher_OnCreatedOrRenamed;
                _recipeWatcher.Renamed += RecipeWatcher_OnCreatedOrRenamed;

                UpdateRecipeMonitorStatus($"Monitorando receitas em: {cfg.SourceFolder}");
            }
            catch (Exception ex)
            {
                UpdateRecipeMonitorStatus($"Falha no monitoramento de receitas: {ex.Message}");
            }
        }

        private void RecipeWatcher_OnCreatedOrRenamed(object sender, FileSystemEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(e.FullPath))
                return;

            if (!_recipeCopyInProgress.TryAdd(e.FullPath, 0))
                return;

            _ = Task.Run(async () =>
            {
                try
                {
                    await CopyRecipeFileWithRetryAsync(e.FullPath, CancellationToken.None);
                }
                finally
                {
                    _recipeCopyInProgress.TryRemove(e.FullPath, out _);
                }
            });
        }

        private async Task CopyRecipeFileWithRetryAsync(string sourceFilePath, CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                try
                {
                    var cfg = _settings.RecipeMonitor;
                    if (string.IsNullOrWhiteSpace(cfg.TargetFolder))
                    {
                        UpdateRecipeMonitorStatus("Destino da receita não configurado.");
                        await Task.Delay(TimeSpan.FromSeconds(2), ct);
                        continue;
                    }

                    if (!File.Exists(sourceFilePath))
                    {
                        await Task.Delay(TimeSpan.FromMilliseconds(500), ct);
                        continue;
                    }

                    using var _ = NetworkShareConnection.ConnectIfNeeded(
                        cfg.TargetFolder,
                        cfg.Username,
                        cfg.Password);

                    var fileName = Path.GetFileName(sourceFilePath);
                    var targetPath = Path.Combine(cfg.TargetFolder, fileName);
                    Directory.CreateDirectory(cfg.TargetFolder);

                    using var source = new FileStream(sourceFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using var target = new FileStream(targetPath, FileMode.Create, FileAccess.Write, FileShare.None);

                    var total = source.Length <= 0 ? 1 : source.Length;
                    var copied = 0L;
                    var buffer = new byte[1024 * 64];

                    int read;
                    while ((read = await source.ReadAsync(buffer.AsMemory(0, buffer.Length), ct)) > 0)
                    {
                        await target.WriteAsync(buffer.AsMemory(0, read), ct);
                        copied += read;

                        var pct = (int)Math.Clamp((copied * 100.0) / total, 0, 100);
                        Dispatcher.Invoke(() =>
                        {
                            RecipeCopyProgressBar.Value = pct;
                            RecipeCopyProgressTextBlock.Text = $"Cópia receita: {pct}%";
                        });
                    }

                    Dispatcher.Invoke(() =>
                    {
                        RecipeCopyProgressBar.Value = 100;
                        RecipeCopyProgressTextBlock.Text = "Cópia receita: 100%";
                    });

                    UpdateRecipeMonitorStatus($"Receita copiada: {targetPath}");
                    return;
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    UpdateRecipeMonitorStatus($"Falha ao copiar receita. Tentando novamente... {ex.Message}");
                    await Task.Delay(TimeSpan.FromSeconds(2), ct);
                }
            }
        }

        private void ShowTrayBalloon(string title, string text, WinForms.ToolTipIcon icon)
        {
            try
            {
                if (_trayIcon is null)
                    return;

                _trayIcon.BalloonTipTitle = title;
                _trayIcon.BalloonTipText = text;
                _trayIcon.BalloonTipIcon = icon;
                _trayIcon.ShowBalloonTip(4000);
            }
            catch
            {
                // ignore
            }
        }

        private void BrowseNetworkBackupFolderButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFolderDialog
            {
                Title = "Selecione a pasta de backup de rede",
                Multiselect = false,
            };

            if (dialog.ShowDialog() == true)
                NetworkBackupFolderTextBox.Text = dialog.FolderName;
        }

        private void NumericOnly_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = !System.Text.RegularExpressions.Regex.IsMatch(e.Text, "^[0-9]+$");
        }

        private void SmtpShowPwd_Checked(object sender, RoutedEventArgs e)
        {
            SmtpPasswordTextBox.Text = SmtpPasswordBox.Password;
            SmtpPasswordTextBox.Visibility = Visibility.Visible;
            SmtpPasswordBox.Visibility = Visibility.Collapsed;
        }

        private void SmtpShowPwd_Unchecked(object sender, RoutedEventArgs e)
        {
            SmtpPasswordBox.Password = SmtpPasswordTextBox.Text;
            SmtpPasswordTextBox.Visibility = Visibility.Collapsed;
            SmtpPasswordBox.Visibility = Visibility.Visible;
        }

        private void NetworkShowPwd_Checked(object sender, RoutedEventArgs e)
        {
            NetworkBackupPasswordTextBox.Text = NetworkBackupPasswordBox.Password;
            NetworkBackupPasswordTextBox.Visibility = Visibility.Visible;
            NetworkBackupPasswordBox.Visibility = Visibility.Collapsed;
        }

        private void NetworkShowPwd_Unchecked(object sender, RoutedEventArgs e)
        {
            NetworkBackupPasswordBox.Password = NetworkBackupPasswordTextBox.Text;
            NetworkBackupPasswordTextBox.Visibility = Visibility.Collapsed;
            NetworkBackupPasswordBox.Visibility = Visibility.Visible;
        }

        private void TestNetworkBackup_Click(object sender, RoutedEventArgs e)
        {
            AutomationLog.Ui("NetworkBackup", "Test access clicked");
            SaveUiToSettings();

            var path = _settings.NetworkBackup.TargetFolder;
            if (string.IsNullOrWhiteSpace(path))
            {
                System.Windows.MessageBox.Show(this, "Por favor, configure a pasta de destino primeiro.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                using var _ = NetworkShareConnection.ConnectIfNeeded(
                    path,
                    _settings.NetworkBackup.Username,
                    _settings.NetworkBackup.Password);

                Directory.CreateDirectory(path);
                var testFile = Path.Combine(path, $"test_{Guid.NewGuid():N}.tmp");
                File.WriteAllText(testFile, "test");
                File.Delete(testFile);

                System.Windows.MessageBox.Show(this, "Acesso à pasta de rede efetuado com sucesso (leitura/escrita confirmadas).", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(this, $"Falha ao acessar pasta de rede.\n\n{ex.Message}", "Erro de Acesso", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Reintroduz tipos e helpers que são referenciados no arquivo.
        // (Algumas edições anteriores removeram parte do arquivo acidentalmente.)

        public enum PlcValueType
        {
            Word = 0,
            DWord = 1,
            AsciiString = 2,
            Coil = 3,
        }

        public sealed class PlcValueRow : DependencyObject
        {
            public PlcValueRow(string address, string name, PlcValueType valueType, int modbusAddress, int registerCount)
            {
                Address = address;
                Name = name;
                ValueType = valueType;
                ModbusAddress = modbusAddress;
                RegisterCount = registerCount;
            }

            public string Address { get; }
            public string Name { get; }
            public PlcValueType ValueType { get; }
            public int ModbusAddress { get; }
            public int RegisterCount { get; }

            public string Value
            {
                get => (string)GetValue(ValueProperty);
                set => SetValue(ValueProperty, value);
            }

            public static readonly DependencyProperty ValueProperty =
                DependencyProperty.Register(nameof(Value), typeof(string), typeof(PlcValueRow), new PropertyMetadata(string.Empty));

            public string LastUpdate
            {
                get => (string)GetValue(LastUpdateProperty);
                set => SetValue(LastUpdateProperty, value);
            }

            public static readonly DependencyProperty LastUpdateProperty =
                DependencyProperty.Register(nameof(LastUpdate), typeof(string), typeof(PlcValueRow), new PropertyMetadata(string.Empty));
        }

        private sealed record TimeZoneOption(string Id, string Display)
        {
            public static TimeZoneOption WindowsLocal { get; } = new("", "Windows (Local)");
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

        private static TimeZoneInfo ResolveTimeZoneFromSettings(string? timeZoneId)
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

        private static string BuildCsvPathForDate(DateTime date)
        {
            var root = Path.Combine(SettingsService.GetDefaultDataRootPath(), "Logs");
            var yearDir = Path.Combine(root, date.Year.ToString("0000", CultureInfo.InvariantCulture));
            var monthDir = Path.Combine(yearDir, date.Month.ToString("00", CultureInfo.InvariantCulture));
            var fileName = $"{date:yyyy-MM-dd}.csv";
            return Path.Combine(monthDir, fileName);
        }

        private static string CreateAttachmentSnapshot(string sourceCsvPath)
        {
            if (!File.Exists(sourceCsvPath))
                throw new FileNotFoundException("Arquivo CSV não encontrado para envio.", sourceCsvPath);

            var tempDir = Path.Combine(SettingsService.GetDefaultDataRootPath(), "Temp");
            Directory.CreateDirectory(tempDir);

            var snapshotPath = Path.Combine(tempDir, $"mail_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}_{Guid.NewGuid():N}.csv");
            File.Copy(sourceCsvPath, snapshotPath, overwrite: false);
            return snapshotPath;
        }

        private static void TryDeleteFile(string? filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return;

            try
            {
                if (File.Exists(filePath))
                    File.Delete(filePath);
            }
            catch
            {
                // ignore
            }
        }

        private static IPEndPoint ResolveEndpoint(string host, int port)
        {
            if (IPAddress.TryParse(host, out var ip))
                return new IPEndPoint(ip, port);

            var addresses = Dns.GetHostAddresses(host);
            var selected = addresses.FirstOrDefault(a => a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                           ?? addresses.FirstOrDefault()
                           ?? throw new InvalidOperationException($"Host inválido ou não resolvido: {host}");

            return new IPEndPoint(selected, port);
        }

        private void PollOnce(ModbusTcpClient client, byte unitId, CancellationToken ct)
        {
            static ushort ReadU16BE(ReadOnlySpan<byte> buffer, int offset)
                => (ushort)((buffer[offset] << 8) | buffer[offset + 1]);

            static uint ReadU32BE(ReadOnlySpan<byte> buffer, int offset)
                => (uint)((buffer[offset] << 24) | (buffer[offset + 1] << 16) | (buffer[offset + 2] << 8) | buffer[offset + 3]);

            ct.ThrowIfCancellationRequested();

            PlcValueRow[] rows = Array.Empty<PlcValueRow>();
            Dispatcher.Invoke(() => rows = Rows.ToArray());

            var now = DateTime.Now;
            var lastUpdate = now.ToString("HH:mm:ss", CultureInfo.InvariantCulture);

            var updates = new (PlcValueRow Row, string Value, string LastUpdate)[rows.Length];

            for (var i = 0; i < rows.Length; i++)
            {
                ct.ThrowIfCancellationRequested();

                var row = rows[i];
                string value;

                switch (row.ValueType)
                {
                    case PlcValueType.Coil:
                    {
                        var raw = client.ReadCoils(unitId, row.ModbusAddress, row.RegisterCount);
                        var isOn = raw.Length > 0 && (raw[0] & 0b_0000_0001) != 0;
                        value = isOn ? "1" : "0";
                        break;
                    }

                    case PlcValueType.AsciiString:
                    {
                        var raw = client.ReadHoldingRegisters(unitId, (ushort)row.ModbusAddress, (ushort)row.RegisterCount);
                        value = Encoding.ASCII.GetString(raw).TrimEnd('\0', ' ');
                        break;
                    }

                    case PlcValueType.DWord:
                    {
                        var raw = client.ReadHoldingRegisters(unitId, (ushort)row.ModbusAddress, (ushort)row.RegisterCount);
                        value = raw.Length >= 4
                            ? ReadU32BE(raw, 0).ToString(CultureInfo.InvariantCulture)
                            : string.Empty;
                        break;
                    }

                    case PlcValueType.Word:
                    default:
                    {
                        var raw = client.ReadHoldingRegisters(unitId, (ushort)row.ModbusAddress, (ushort)row.RegisterCount);
                        value = raw.Length >= 2
                            ? ReadU16BE(raw, 0).ToString(CultureInfo.InvariantCulture)
                            : string.Empty;
                        break;
                    }
                }

                updates[i] = (row, value, lastUpdate);
            }

            Dispatcher.Invoke(() =>
            {
                foreach (var (row, value, updateTs) in updates)
                {
                    row.Value = value;
                    row.LastUpdate = updateTs;
                }
            });
        }

        private async Task<bool> SendDailyCsvAsync(CancellationToken ct)
        {
            string? attachmentSnapshotPath = null;
            try
            {
                var settings = _settingsService.Load();
                var automation = settings.EmailAutomation;

                var tz = ResolveTimeZoneFromSettings(automation.TimeZoneId);
                var nowInTz = TimeZoneInfo.ConvertTime(DateTimeOffset.UtcNow, tz);
                var todayInTz = nowInTz.Date;
                var yesterdayInTz = todayInTz.AddDays(-1);

                AutomationLog.Info($"SendDailyCsvAsync: start tz='{tz.Id}' nowInTz={nowInTz:yyyy-MM-dd HH:mm:ss} enabled={automation.Enabled} dailyTime='{automation.DailySendTime}' lastCsv='{automation.LastSuccessfulCsvDate}' lastSend='{automation.LastSuccessfulSendDate}' lastYDate='{automation.LastYesterdayCsvSentDate}' skipCatchUp='{automation.SkipPreviousDayCatchUpForDate}'.");

                if (!automation.Enabled)
                {
                    AutomationLog.Info("SendDailyCsvAsync: disabled.");
                    return false;
                }

                if (!IsValidEmail(settings.Email.SenderEmail))
                    throw new InvalidOperationException("E-mail remetente inválido.");
                if (string.IsNullOrWhiteSpace(settings.Email.SmtpHost))
                    throw new InvalidOperationException("SMTP Host não configurado.");
                if (settings.Email.SmtpPort is <= 0 or > 65535)
                    throw new InvalidOperationException("SMTP Port inválido.");
                if (settings.Email.Recipients is null || settings.Email.Recipients.Count == 0)
                    throw new InvalidOperationException("Nenhum destinatário configurado.");

                ct.ThrowIfCancellationRequested();

                var skipCatchUpForToday =
                    !string.IsNullOrWhiteSpace(automation.SkipPreviousDayCatchUpForDate) &&
                    DateTime.TryParseExact(automation.SkipPreviousDayCatchUpForDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var skipDate) &&
                    skipDate.Date == todayInTz;

                var yesterdayAlreadySentByCatchUp =
                    !string.IsNullOrWhiteSpace(automation.LastYesterdayCsvSentDate) &&
                    DateTime.TryParseExact(automation.LastYesterdayCsvSentDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var ySent) &&
                    ySent.Date == yesterdayInTz;

                DateTime? lastCsvDate = null;
                if (DateTime.TryParseExact(automation.LastSuccessfulCsvDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var lastCsv))
                    lastCsvDate = lastCsv.Date;

                var sendDateInTz = todayInTz;
                var isYesterdayCatchUp = false;

                var yesterdayCsvPathExisting = BuildCsvPathForDate(yesterdayInTz);
                var hasYesterday = File.Exists(yesterdayCsvPathExisting);

                if (!skipCatchUpForToday && hasYesterday && lastCsvDate != yesterdayInTz && !yesterdayAlreadySentByCatchUp)
                {
                    sendDateInTz = yesterdayInTz;
                    isYesterdayCatchUp = true;
                }

                var csvPath = EnsureCsvExistsForDate(sendDateInTz, tz);
                attachmentSnapshotPath = CreateAttachmentSnapshot(csvPath);

                var subject = isYesterdayCatchUp
                    ? $"PLCDataLog - CSV {sendDateInTz:yyyy-MM-dd} (ontem)"
                    : $"PLCDataLog - CSV {sendDateInTz:yyyy-MM-dd}";

                var body = isYesterdayCatchUp
                    ? $"Segue anexado o CSV de ontem ({sendDateInTz:yyyy-MM-dd})."
                    : $"Segue anexado o CSV do dia ({sendDateInTz:yyyy-MM-dd}).";

                AutomationLog.Info($"SendDailyCsvAsync: sending csvDate={sendDateInTz:yyyy-MM-dd} path='{csvPath}' recipients={settings.Email.Recipients.Count}.");

                await _smtpEmailService.SendCsvAsync(
                    settings.Email,
                    subject,
                    body,
                    attachmentSnapshotPath,
                    settings.Email.Recipients,
                    ct);

                var sentAt = TimeZoneInfo.ConvertTime(DateTimeOffset.UtcNow, tz);
                settings = _settingsService.Load();

                settings.EmailAutomation.LastSuccessfulCsvDate = sendDateInTz.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                settings.EmailAutomation.LastSuccessfulSendDate = sentAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

                if (isYesterdayCatchUp)
                {
                    settings.EmailAutomation.LastYesterdayCsvSentDate = sendDateInTz.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                    settings.EmailAutomation.LastYesterdayCsvSentAt = sentAt.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
                }

                try
                {
                    _settingsService.Save(settings);
                    AutomationLog.Info("SendDailyCsvAsync: success persisted.");
                }
                catch (Exception ex)
                {
                    AutomationLog.Error(ex, "SendDailyCsvAsync: failed to persist success markers");
                }

                return true;
            }
            catch (Exception ex)
            {
                AutomationLog.Error(ex, "SendDailyCsvAsync: failed");
                throw;
            }
            finally
            {
                TryDeleteFile(attachmentSnapshotPath);
            }
        }

        private void ValuesGrid_PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            e.Handled = true;
            var ev = new System.Windows.Input.MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta)
            {
                RoutedEvent = UIElement.MouseWheelEvent,
                Source = sender,
            };
            DashboardPanel.RaiseEvent(ev);
        }
    }
}