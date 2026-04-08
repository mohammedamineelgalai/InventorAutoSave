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
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = viewModel;

            // Synchroniser l'etat initial des toggles
            RefreshToggleStates();

            // Ecouter les changements de settings pour mettre a jour l'affichage
            viewModel.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(viewModel.Settings))
                    RefreshToggleStates();
            };
        }

        private void RefreshToggleStates()
        {
            var s = _viewModel.Settings;

            // Mode de sauvegarde
            TglSaveActive.IsChecked = (s.SaveMode == SaveMode.SaveActive);
            TglSaveAll.IsChecked    = (s.SaveMode == SaveMode.SaveAll);

            // Options
            TglNotifications.IsChecked = s.ShowNotifications;
            TglNotifications.Content   = s.ShowNotifications ? "ON" : "OFF";

            TglSafetyChecks.IsChecked = s.SafetyChecks;
            TglSafetyChecks.Content   = s.SafetyChecks ? "ON" : "OFF";

            // Intervalle actuel
            if (TxtCurrentInterval != null)
                TxtCurrentInterval.Text = AutoSaveTimerService.FormatInterval(s.SaveIntervalSeconds);
        }

        // ── Boutons intervalles ──
        private void IntervalButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string tagStr && int.TryParse(tagStr, out int seconds))
            {
                _viewModel.ChangeIntervalCommand.Execute(seconds);
                if (TxtCurrentInterval != null)
                    TxtCurrentInterval.Text = AutoSaveTimerService.FormatInterval(seconds);
            }
        }

        // ── Mode sauvegarde ──
        private void TglSaveActive_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.ChangeSaveModeCommand.Execute(SaveMode.SaveActive);
            TglSaveActive.IsChecked = true;
            TglSaveAll.IsChecked    = false;
        }

        private void TglSaveAll_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.ChangeSaveModeCommand.Execute(SaveMode.SaveAll);
            TglSaveAll.IsChecked    = true;
            TglSaveActive.IsChecked = false;
        }

        // ── Notifications ──
        private void TglNotifications_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.ToggleNotificationsCommand.Execute(null);
            bool val = _viewModel.Settings.ShowNotifications;
            TglNotifications.IsChecked = val;
            TglNotifications.Content   = val ? "ON" : "OFF";
        }

        // ── Protection calculs ──
        private void TglSafetyChecks_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.ToggleSafetyChecksCommand.Execute(null);
            bool val = _viewModel.Settings.SafetyChecks;
            TglSafetyChecks.IsChecked = val;
            TglSafetyChecks.Content   = val ? "ON" : "OFF";
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
