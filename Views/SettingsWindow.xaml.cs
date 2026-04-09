using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using InventorAutoSave.Models;
using InventorAutoSave.Services;
using InventorAutoSave.ViewModels;

namespace InventorAutoSave.Views
{
    public partial class SettingsWindow : Window
    {
        private readonly MainViewModel _viewModel;

        public SettingsWindow(MainViewModel viewModel)
        {
            try
            {
                InitializeComponent();
                _viewModel = viewModel;
                DataContext = viewModel;

                // Synchroniser l'etat initial des toggles
                RefreshToggleStates();

                // Ecouter les changements de settings pour mettre a jour l'affichage
                viewModel.PropertyChanged += (s, e) =>
                {
                    try
                    {
                        if (e.PropertyName == nameof(viewModel.Settings))
                            Dispatcher.Invoke(RefreshToggleStates);
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"[!] Settings PropertyChanged: {ex.Message}", Logger.LogLevel.WARNING);
                    }
                };

                Logger.Log("[+] SettingsWindow initialisee", Logger.LogLevel.DEBUG);
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] SettingsWindow constructeur: {ex.Message}\n{ex.StackTrace}", Logger.LogLevel.ERROR);
                throw; // Rethrow pour que SafeOpenSettings le capture
            }
        }

        private void RefreshToggleStates()
        {
            try
            {
                var s = _viewModel.Settings;

                // Mode de sauvegarde
                TglSaveActive.IsChecked = (s.SaveMode == SaveMode.SaveActive);
                TglSaveAll.IsChecked = (s.SaveMode == SaveMode.SaveAll);

                // Options - Toggle switches
                TglNotifications.IsChecked = s.ShowNotifications;
                TxtNotifStatus.Text = s.ShowNotifications ? "ON" : "OFF";
                TxtNotifStatus.Foreground = s.ShowNotifications
                    ? new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#00D26A"))
                    : new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#888888"));

                TglSafetyChecks.IsChecked = s.SafetyChecks;
                TxtSafetyStatus.Text = s.SafetyChecks ? "ON" : "OFF";
                TxtSafetyStatus.Foreground = s.SafetyChecks
                    ? new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#00D26A"))
                    : new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#888888"));

                // Intervalle actuel
                if (TxtCurrentInterval != null)
                    TxtCurrentInterval.Text = AutoSaveTimerService.FormatInterval(s.SaveIntervalSeconds);
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] RefreshToggleStates: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        // ── Boutons intervalles ──
        private void IntervalButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Button btn && btn.Tag is string tagStr && int.TryParse(tagStr, out int seconds))
                {
                    _viewModel.ChangeIntervalCommand.Execute(seconds);
                    if (TxtCurrentInterval != null)
                        TxtCurrentInterval.Text = AutoSaveTimerService.FormatInterval(seconds);
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] IntervalButton_Click: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        // ── Mode sauvegarde ──
        private void TglSaveActive_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _viewModel.ChangeSaveModeCommand.Execute(SaveMode.SaveActive);
                TglSaveActive.IsChecked = true;
                TglSaveAll.IsChecked = false;
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] TglSaveActive_Click: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        private void TglSaveAll_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _viewModel.ChangeSaveModeCommand.Execute(SaveMode.SaveAll);
                TglSaveAll.IsChecked = true;
                TglSaveActive.IsChecked = false;
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] TglSaveAll_Click: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        // ── Notifications ──
        private void TglNotifications_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _viewModel.ToggleNotificationsCommand.Execute(null);
                bool val = _viewModel.Settings.ShowNotifications;
                TglNotifications.IsChecked = val;
                TxtNotifStatus.Text = val ? "ON" : "OFF";
                TxtNotifStatus.Foreground = val
                    ? new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#00D26A"))
                    : new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#888888"));
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] TglNotifications_Click: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        // ── Protection calculs ──
        private void TglSafetyChecks_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _viewModel.ToggleSafetyChecksCommand.Execute(null);
                bool val = _viewModel.Settings.SafetyChecks;
                TglSafetyChecks.IsChecked = val;
                TxtSafetyStatus.Text = val ? "ON" : "OFF";
                TxtSafetyStatus.Foreground = val
                    ? new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#00D26A"))
                    : new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#888888"));
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] TglSafetyChecks_Click: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }
    }

    /// <summary>
    /// Converter bool -> "ON" / "OFF" pour les boutons toggle
    /// </summary>
    public class BoolToOnOffConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is bool b && b ? "ON" : "OFF";

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => DependencyProperty.UnsetValue;
    }
}
