using System.Windows;
using System.Windows.Controls;
using Hardcodet.Wpf.TaskbarNotification;
using InventorAutoSave.Models;
using InventorAutoSave.Services;
using InventorAutoSave.ViewModels;
using InventorAutoSave.Views;
namespace InventorAutoSave
{
    public partial class App : Application
    {
        // Services (singleton pour toute la duree de vie de l'app)
        private SettingsService _settingsService = null!;
        private InventorSaveService _inventorService = null!;
        private AutoSaveTimerService _timerService = null!;
        private MainViewModel _viewModel = null!;

        // UI
        private TaskbarIcon? _trayIcon;
        private SettingsWindow? _settingsWindow;

        // Menus tray (references pour mise a jour dynamique)
        private MenuItem? _miToggleAutoSave;
        private MenuItem? _miIntervalGroup;
        private MenuItem? _miSaveMode;
        private MenuItem? _miNotifications;
        private MenuItem? _miSafetyChecks;
        private MenuItem? _miStartup;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Demarrage silencieux: pas de fenetre principale
            Logger.Log("[+] InventorAutoSave v2.0 demarre", Logger.LogLevel.INFO);

            // Initialiser les services
            _settingsService = new SettingsService();
            _inventorService = new InventorSaveService();
            _timerService = new AutoSaveTimerService(_inventorService, _settingsService);
            _viewModel = new MainViewModel(_inventorService, _timerService, _settingsService);

            // S'abonner aux evenements pour les notifications et le menu
            _timerService.SaveCompleted += OnSaveCompleted;
            _timerService.StatusChanged += OnStatusChanged;
            _inventorService.Connected += OnInventorConnected;
            _inventorService.Disconnected += OnInventorDisconnected;

            // Creer l'icone systray
            CreateTrayIcon();

            // Synchroniser le raccourci Startup avec le setting persiste
            SyncStartupShortcut();
        }

        // ═══════════════════════════════════════════════════════════════
        // TRAY ICON + MENU CONTEXTUEL
        // ═══════════════════════════════════════════════════════════════

        private void CreateTrayIcon()
        {
            _trayIcon = new TaskbarIcon
            {
                ToolTipText = "Inventor AutoSave v2.0",
                Visibility = Visibility.Visible
            };

            // Charger l'icone
            try
            {
                string iconPath = System.IO.Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory, "Resources", "InventorAutoSave.ico");
                if (System.IO.File.Exists(iconPath))
                {
                    _trayIcon.Icon = new System.Drawing.Icon(iconPath);
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Icone tray non trouvee: {ex.Message}", Logger.LogLevel.WARNING);
            }

            // Double-clic => ouvrir Settings
            _trayIcon.TrayMouseDoubleClick += (s, e) => OpenSettings();

            // Menu contextuel
            _trayIcon.ContextMenu = BuildContextMenu();
        }

        private ContextMenu BuildContextMenu()
        {
            var menu = new ContextMenu
            {
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#1E1E2E")),
                Foreground = System.Windows.Media.Brushes.White
            };

            // Style commun pour les items
            Style menuItemStyle = CreateMenuItemStyle();

            // ── Toggle AutoSave ──
            _miToggleAutoSave = new MenuItem
            {
                Style = menuItemStyle
            };
            UpdateToggleMenuText();
            _miToggleAutoSave.Click += (s, e) => _viewModel.ToggleAutoSaveCommand.Execute(null);

            // ── Sauvegarder maintenant ──
            var miSaveNow = new MenuItem { Header = "💾  Sauvegarder maintenant", Style = menuItemStyle };
            miSaveNow.Click += (s, e) => _viewModel.SaveNowCommand.Execute(null);

            // ── Separateur ──
            var sep1 = new Separator { Background = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#333344")) };

            // ── Mode de sauvegarde (sous-menu) ──
            _miSaveMode = new MenuItem { Header = "💡  Mode de sauvegarde", Style = menuItemStyle };
            var miSaveActive = new MenuItem { Header = "💾  Document actif (recommande)", Style = menuItemStyle, IsCheckable = true };
            var miSaveAll    = new MenuItem { Header = "💾  Tous les documents", Style = menuItemStyle, IsCheckable = true };
            miSaveActive.IsChecked = _settingsService.Current.SaveMode == SaveMode.SaveActive;
            miSaveAll.IsChecked    = _settingsService.Current.SaveMode == SaveMode.SaveAll;
            miSaveActive.Click += (s, e) => { _viewModel.ChangeSaveModeCommand.Execute(SaveMode.SaveActive); miSaveActive.IsChecked = true; miSaveAll.IsChecked = false; };
            miSaveAll.Click    += (s, e) => { _viewModel.ChangeSaveModeCommand.Execute(SaveMode.SaveAll);    miSaveAll.IsChecked = true; miSaveActive.IsChecked = false; };
            _miSaveMode.Items.Add(miSaveActive);
            _miSaveMode.Items.Add(miSaveAll);

            // ── Intervalle (sous-menu) ──
            _miIntervalGroup = new MenuItem { Header = "⏱️  Intervalle de sauvegarde", Style = menuItemStyle };
            int[] intervals = { 10, 30, 60, 120, 180, 300, 600, 900, 1200, 1800 };
            foreach (int sec in intervals)
            {
                string label = AutoSaveTimerService.FormatInterval(sec);
                var mi = new MenuItem
                {
                    Header = label,
                    Style = menuItemStyle,
                    IsCheckable = true,
                    IsChecked = _settingsService.Current.SaveIntervalSeconds == sec,
                    Tag = sec
                };
                mi.Click += (s, e) =>
                {
                    if (mi.Tag is int selectedSec)
                    {
                        _viewModel.ChangeIntervalCommand.Execute(selectedSec);
                        // Mettre a jour les checks
                        foreach (MenuItem item in _miIntervalGroup!.Items)
                            item.IsChecked = (item.Tag is int t && t == selectedSec);
                    }
                };
                _miIntervalGroup.Items.Add(mi);
            }

            // ── Notifications (sous-menu) ──
            _miNotifications = new MenuItem { Header = "🔔  Notifications", Style = menuItemStyle };
            var miNotifOn  = new MenuItem { Header = "Activer", Style = menuItemStyle, IsCheckable = true };
            var miNotifOff = new MenuItem { Header = "Desactiver", Style = menuItemStyle, IsCheckable = true };
            miNotifOn.IsChecked  = _settingsService.Current.ShowNotifications;
            miNotifOff.IsChecked = !_settingsService.Current.ShowNotifications;
            miNotifOn.Click  += (s, e) => { _settingsService.Update(x => x.ShowNotifications = true);  miNotifOn.IsChecked = true;  miNotifOff.IsChecked = false; };
            miNotifOff.Click += (s, e) => { _settingsService.Update(x => x.ShowNotifications = false); miNotifOff.IsChecked = true; miNotifOn.IsChecked  = false; };
            _miNotifications.Items.Add(miNotifOn);
            _miNotifications.Items.Add(miNotifOff);

            // ── Protection calculs ──
            _miSafetyChecks = new MenuItem { Header = "🛡️  Protection calculs", Style = menuItemStyle };
            var miSafeOn  = new MenuItem { Header = "Activer", Style = menuItemStyle, IsCheckable = true };
            var miSafeOff = new MenuItem { Header = "Desactiver", Style = menuItemStyle, IsCheckable = true };
            miSafeOn.IsChecked  = _settingsService.Current.SafetyChecks;
            miSafeOff.IsChecked = !_settingsService.Current.SafetyChecks;
            miSafeOn.Click  += (s, e) => { _settingsService.Update(x => x.SafetyChecks = true);  miSafeOn.IsChecked = true;  miSafeOff.IsChecked = false; };
            miSafeOff.Click += (s, e) => { _settingsService.Update(x => x.SafetyChecks = false); miSafeOff.IsChecked = true; miSafeOn.IsChecked  = false; };
            _miSafetyChecks.Items.Add(miSafeOn);
            _miSafetyChecks.Items.Add(miSafeOff);

            // ── Demarrer avec Windows ──
            _miStartup = new MenuItem
            {
                Header = StartupManager.IsStartupEnabled
                    ? "🚀  Demarrer avec Windows  ✅"
                    : "🚀  Demarrer avec Windows",
                Style = menuItemStyle,
                IsCheckable = true,
                IsChecked = StartupManager.IsStartupEnabled
            };
            _miStartup.Click += (s, e) =>
            {
                if (StartupManager.IsStartupEnabled)
                {
                    StartupManager.DisableStartup();
                    _settingsService.Update(x => x.StartWithWindows = false);
                    _miStartup!.IsChecked = false;
                    _miStartup.Header = "🚀  Demarrer avec Windows";
                }
                else
                {
                    StartupManager.EnableStartup();
                    _settingsService.Update(x => x.StartWithWindows = true);
                    _miStartup!.IsChecked = true;
                    _miStartup.Header = "🚀  Demarrer avec Windows  ✅";
                }
            };

            // ── Separateur ──
            var sep2 = new Separator { Background = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#333344")) };

            // ── Ouvrir configuration ──
            var miOpenSettings = new MenuItem { Header = "⚙️  Ouvrir la configuration", Style = menuItemStyle };
            miOpenSettings.Click += (s, e) => OpenSettings();

            // ── Separateur ──
            var sep3 = new Separator { Background = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#333344")) };

            // ── Quitter ──
            var miQuit = new MenuItem { Header = "❌  Quitter", Style = menuItemStyle };
            miQuit.Click += (s, e) => Shutdown();

            // Assembler le menu
            menu.Items.Add(_miToggleAutoSave);
            menu.Items.Add(miSaveNow);
            menu.Items.Add(sep1);
            menu.Items.Add(_miSaveMode);
            menu.Items.Add(_miIntervalGroup);
            menu.Items.Add(_miNotifications);
            menu.Items.Add(_miSafetyChecks);
            menu.Items.Add(_miStartup);
            menu.Items.Add(sep2);
            menu.Items.Add(miOpenSettings);
            menu.Items.Add(sep3);
            menu.Items.Add(miQuit);

            return menu;
        }

        private Style CreateMenuItemStyle()
        {
            var style = new Style(typeof(MenuItem));
            style.Setters.Add(new Setter(MenuItem.ForegroundProperty, System.Windows.Media.Brushes.White));
            style.Setters.Add(new Setter(MenuItem.FontSizeProperty, 12.0));
            style.Setters.Add(new Setter(MenuItem.BackgroundProperty,
                new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#1E1E2E"))));
            return style;
        }

        private void UpdateToggleMenuText()
        {
            if (_miToggleAutoSave == null) return;
            bool enabled = _settingsService.Current.EnableAutoSave;
            _miToggleAutoSave.Header = enabled ? "⏸️  Desactiver AutoSave" : "▶️  Activer AutoSave";
        }

        // ═══════════════════════════════════════════════════════════════
        // FENETRE SETTINGS
        // ═══════════════════════════════════════════════════════════════

        private void OpenSettings()
        {
            if (_settingsWindow == null || !_settingsWindow.IsLoaded)
            {
                _settingsWindow = new SettingsWindow(_viewModel);
                _settingsWindow.Closed += (s, e) => _settingsWindow = null;
                _settingsWindow.Show();
            }
            else
            {
                _settingsWindow.Activate();
                _settingsWindow.WindowState = WindowState.Normal;
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // NOTIFICATIONS SYSTRAY
        // ═══════════════════════════════════════════════════════════════

        private void OnSaveCompleted(object? sender, SaveResult result)
        {
            if (!_settingsService.Current.ShowNotifications) return;
            if (!result.Success) return;
            if (result.DocumentsSaved == 0) return; // Rien de sauvegarde -> pas de notif

            Dispatcher.Invoke(() =>
            {
                string message = result.Mode == SaveMode.SaveActive
                    ? $"Document actif sauvegarde"
                    : $"{result.DocumentsSaved} document(s) sauvegarde(s)";

                _trayIcon?.ShowBalloonTip(
                    "Inventor AutoSave",
                    $"✅ {message}",
                    BalloonIcon.Info);
            });
        }

        private void OnStatusChanged(object? sender, string status)
        {
            Dispatcher.Invoke(() =>
            {
                if (_trayIcon != null)
                    _trayIcon.ToolTipText = $"Inventor AutoSave - {status}";
            });
        }

        private void OnInventorConnected(object? sender, EventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                if (_trayIcon != null)
                    _trayIcon.ToolTipText = "Inventor AutoSave - Connecte";

                if (_settingsService.Current.ShowNotifications)
                {
                    _trayIcon?.ShowBalloonTip(
                        "Inventor AutoSave",
                        "⚙️ Connecte a Inventor",
                        BalloonIcon.Info);
                }
            });
        }

        private void OnInventorDisconnected(object? sender, EventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                if (_trayIcon != null)
                    _trayIcon.ToolTipText = "Inventor AutoSave - Inventor non connecte";
            });
        }

        // ═══════════════════════════════════════════════════════════════
        // STARTUP WINDOWS
        // ═══════════════════════════════════════════════════════════════

        private void SyncStartupShortcut()
        {
            bool settingEnabled = _settingsService.Current.StartWithWindows;
            bool shortcutExists = StartupManager.IsStartupEnabled;

            // Synchroniser l'etat reel avec le setting persiste
            if (settingEnabled && !shortcutExists)
                StartupManager.EnableStartup();
            else if (!settingEnabled && shortcutExists)
            {
                // Ne pas supprimer automatiquement: l'utilisateur l'a peut-etre
                // active via l'installateur -> mettre a jour le setting
                _settingsService.Update(x => x.StartWithWindows = true);
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // SHUTDOWN
        // ═══════════════════════════════════════════════════════════════

        protected override void OnExit(ExitEventArgs e)
        {
            Logger.Log("[i] Arret de l'application", Logger.LogLevel.INFO);

            _trayIcon?.Dispose();
            _timerService?.Dispose();
            _inventorService?.Disconnect();

            base.OnExit(e);
        }
    }
}
