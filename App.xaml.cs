using System;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Effects;
using Microsoft.Win32;
using PLCDataLog.Services;

namespace PLCDataLog
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : System.Windows.Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;

            try
            {
                var settings = new SettingsService(SettingsService.GetDefaultSettingsPath()).Load();
                ApplyThemeMode(settings.Ui.ThemeMode);
            }
            catch
            {
                ApplyThemeMode("System");
            }
        }

        private static void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            try
            {
                AutomationLog.Error(e.Exception, "Exceção não tratada na UI");
            }
            catch
            {
            }

            e.Handled = true;
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
                AutomationLog.Error(ex, "Exceção não tratada no AppDomain");
            else
                AutomationLog.Error($"Exceção não tratada no AppDomain: {e.ExceptionObject}");
        }

        private static void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
        {
            try
            {
                AutomationLog.Error(e.Exception, "Exceção não observada em Task");
            }
            catch
            {
            }

            e.SetObserved();
        }

        public static void ApplyThemeMode(string? mode)
        {
            var app = Current;
            if (app is null)
                return;

            var normalized = (mode ?? "System").Trim();
            if (string.Equals(normalized, "Dark", StringComparison.OrdinalIgnoreCase))
            {
                ApplyDarkPalette(app.Resources);
                return;
            }

            if (string.Equals(normalized, "System", StringComparison.OrdinalIgnoreCase))
            {
                var sysDark = IsSystemDarkMode();
                if (sysDark)
                {
                    ApplyDarkPalette(app.Resources);
                    return;
                }
            }

            ApplyLightPalette(app.Resources);
        }

        private static void ApplyLightPalette(ResourceDictionary resources)
        {
            SetBrush(resources, "BackgroundColor", "#F5F5F5");
            SetBrush(resources, "CardBackground", "#FCFCFC");
            SetBrush(resources, "LayerBackground", "#FFFFFF");
            SetBrush(resources, "TextPrimary", "#1B1B1B");
            SetBrush(resources, "TextSecondary", "#5E5E5E");
            SetBrush(resources, "BorderSoft", "#1F000000");
            SetBrush(resources, "FocusStroke", "#005FA3");

            SetBrush(resources, "ButtonPrimaryBackground", "#005FA3");
            SetBrush(resources, "ButtonPrimaryBackgroundHover", "#004F89");
            SetBrush(resources, "ButtonPrimaryBackgroundPressed", "#003B66");
            SetBrush(resources, "ButtonPrimaryForeground", "#FFFFFF");

            SetBrush(resources, "ButtonSecondaryBackground", "#F6F6F6");
            SetBrush(resources, "ButtonSecondaryBackgroundHover", "#F0F0F0");
            SetBrush(resources, "ButtonSecondaryBackgroundPressed", "#E6E6E6");
            SetBrush(resources, "ButtonSecondaryForeground", "#1B1B1B");
            SetBrush(resources, "ButtonSecondaryBorder", "#24000000");

            SetBrush(resources, "InputBackground", "#FFFFFF");
            SetBrush(resources, "InputBackgroundHover", "#F8F8F8");
            SetBrush(resources, "InputBackgroundFocused", "#FFFFFF");

            SetBrush(resources, "DataGridHeaderBackground", "#F3F3F3");
            SetBrush(resources, "DataGridRowBackground", "#FFFFFF");
            SetBrush(resources, "DataGridAltRowBackground", "#FAFAFA");

            SetSystemBrush(resources, System.Windows.SystemColors.WindowBrushKey, "#FFFFFF");
            SetSystemBrush(resources, System.Windows.SystemColors.WindowTextBrushKey, "#1B1B1B");
            SetSystemBrush(resources, System.Windows.SystemColors.ControlBrushKey, "#F6F6F6");
            SetSystemBrush(resources, System.Windows.SystemColors.ControlTextBrushKey, "#1B1B1B");
            SetSystemBrush(resources, System.Windows.SystemColors.HighlightBrushKey, "#005FA3");
            SetSystemBrush(resources, System.Windows.SystemColors.HighlightTextBrushKey, "#FFFFFF");

            SetShadow(resources, 0.08);
        }

        private static void ApplyDarkPalette(ResourceDictionary resources)
        {
            SetBrush(resources, "BackgroundColor", "#202020");
            SetBrush(resources, "CardBackground", "#2B2B2B");
            SetBrush(resources, "LayerBackground", "#262626");
            SetBrush(resources, "TextPrimary", "#F3F3F3");
            SetBrush(resources, "TextSecondary", "#C8C8C8");
            SetBrush(resources, "BorderSoft", "#40FFFFFF");
            SetBrush(resources, "FocusStroke", "#4CC2FF");

            SetBrush(resources, "ButtonPrimaryBackground", "#005FA3");
            SetBrush(resources, "ButtonPrimaryBackgroundHover", "#2078C2");
            SetBrush(resources, "ButtonPrimaryBackgroundPressed", "#00497D");
            SetBrush(resources, "ButtonPrimaryForeground", "#FFFFFF");

            SetBrush(resources, "ButtonSecondaryBackground", "#333333");
            SetBrush(resources, "ButtonSecondaryBackgroundHover", "#3B3B3B");
            SetBrush(resources, "ButtonSecondaryBackgroundPressed", "#444444");
            SetBrush(resources, "ButtonSecondaryForeground", "#F3F3F3");
            SetBrush(resources, "ButtonSecondaryBorder", "#66FFFFFF");

            SetBrush(resources, "InputBackground", "#1F1F1F");
            SetBrush(resources, "InputBackgroundHover", "#262626");
            SetBrush(resources, "InputBackgroundFocused", "#1F1F1F");

            SetBrush(resources, "DataGridHeaderBackground", "#303030");
            SetBrush(resources, "DataGridRowBackground", "#262626");
            SetBrush(resources, "DataGridAltRowBackground", "#2B2B2B");

            SetSystemBrush(resources, System.Windows.SystemColors.WindowBrushKey, "#1F1F1F");
            SetSystemBrush(resources, System.Windows.SystemColors.WindowTextBrushKey, "#F3F3F3");
            SetSystemBrush(resources, System.Windows.SystemColors.ControlBrushKey, "#333333");
            SetSystemBrush(resources, System.Windows.SystemColors.ControlTextBrushKey, "#F3F3F3");
            SetSystemBrush(resources, System.Windows.SystemColors.HighlightBrushKey, "#005FA3");
            SetSystemBrush(resources, System.Windows.SystemColors.HighlightTextBrushKey, "#FFFFFF");

            SetShadow(resources, 0.24);
        }

        private static bool IsSystemDarkMode()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
                var value = key?.GetValue("AppsUseLightTheme");
                return value is int i && i == 0;
            }
            catch
            {
                return false;
            }
        }

        private static void SetBrush(ResourceDictionary resources, string key, string hex)
        {
            var color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(hex)!;

            if (resources.Contains(key))
            {
                if (resources[key] is SolidColorBrush brush && !brush.IsFrozen)
                {
                    brush.Color = color;
                    return;
                }

                resources[key] = new SolidColorBrush(color);
                return;
            }

            resources.Add(key, new SolidColorBrush(color));
        }

        private static void SetSystemBrush(ResourceDictionary resources, object key, string hex)
        {
            var color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(hex)!;

            if (resources.Contains(key))
            {
                if (resources[key] is SolidColorBrush brush && !brush.IsFrozen)
                {
                    brush.Color = color;
                    return;
                }

                resources[key] = new SolidColorBrush(color);
                return;
            }

            resources.Add(key, new SolidColorBrush(color));
        }

        private static void SetShadow(ResourceDictionary resources, double opacity)
        {
            if (!resources.Contains("CardShadow"))
                return;

            if (resources["CardShadow"] is DropShadowEffect shadow && !shadow.IsFrozen)
            {
                shadow.Opacity = opacity;
                return;
            }

            resources["CardShadow"] = new DropShadowEffect
            {
                BlurRadius = 16,
                ShadowDepth = 1,
                Opacity = opacity,
            };
        }
    }
}
