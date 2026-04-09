using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Media.Effects;
using Hardcodet.Wpf.TaskbarNotification;
using InventorAutoSave.Models;
using InventorAutoSave.Services;
using InventorAutoSave.ViewModels;
using InventorAutoSave.Views;

namespace InventorAutoSave
{
    public partial class App : Application
    {
        // ═══════════════════════════════════════════════════════════════
        // SERVICES (singleton pour toute la duree de vie de l'app)
        // ═══════════════════════════════════════════════════════════════
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

        // Sous-items checkables (pour sync bidirectionnelle avec SettingsWindow)
        private MenuItem? _miSaveActive;
        private MenuItem? _miSaveAll;
        private MenuItem? _miNotifOn;
        private MenuItem? _miNotifOff;
        private MenuItem? _miSafeOn;
        private MenuItem? _miSafeOff;

        // Info documents (mis a jour dynamiquement)
        private MenuItem? _miDocumentsInfo;

        // ═══════════════════════════════════════════════════════════════
        // STARTUP - avec global exception handlers
        // ═══════════════════════════════════════════════════════════════

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // ── GLOBAL EXCEPTION HANDLERS ──
            // Empêche l'app de planter silencieusement
            AppDomain.CurrentDomain.UnhandledException += (s, args) =>
            {
                var ex = args.ExceptionObject as Exception;
                Logger.Log($"[FATAL] UnhandledException: {ex?.Message}\n{ex?.StackTrace}", Logger.LogLevel.ERROR);
            };

            DispatcherUnhandledException += (s, args) =>
            {
                Logger.Log($"[FATAL] DispatcherUnhandledException: {args.Exception.Message}\n{args.Exception.StackTrace}", Logger.LogLevel.ERROR);
                args.Handled = true; // Empeche le crash total
            };

            TaskScheduler.UnobservedTaskException += (s, args) =>
            {
                Logger.Log($"[FATAL] UnobservedTaskException: {args.Exception.Message}\n{args.Exception.StackTrace}", Logger.LogLevel.ERROR);
                args.SetObserved(); // Empeche le crash total
            };

            // ── DEMARRAGE ──
            Logger.Log("[+] InventorAutoSave v1.0.0 demarre", Logger.LogLevel.INFO);

            try
            {
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

                Logger.Log("[+] Initialisation terminee", Logger.LogLevel.INFO);
            }
            catch (Exception ex)
            {
                Logger.Log($"[FATAL] Erreur initialisation: {ex.Message}\n{ex.StackTrace}", Logger.LogLevel.ERROR);
                MessageBox.Show(
                    $"Erreur d'initialisation:\n{ex.Message}",
                    "InventorAutoSave - Erreur",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                Shutdown(1);
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // TRAY ICON + MENU CONTEXTUEL DARK COMPLET
        // ═══════════════════════════════════════════════════════════════

        private void CreateTrayIcon()
        {
            _trayIcon = new TaskbarIcon
            {
                ToolTipText = "Inventor AutoSave v1.0.0",
                Visibility = Visibility.Visible
            };

            // Charger l'icone depuis les ressources WPF embarquees
            try
            {
                var iconUri = new Uri("pack://application:,,,/Resources/InventorAutoSave.ico", UriKind.Absolute);
                using var stream = Application.GetResourceStream(iconUri)?.Stream;
                if (stream != null)
                {
                    _trayIcon.Icon = new System.Drawing.Icon(stream);
                }
                else
                {
                    Logger.Log("[!] Icone non trouvee dans les ressources embarquees", Logger.LogLevel.WARNING);
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Icone tray non trouvee: {ex.Message}", Logger.LogLevel.WARNING);
            }

            // Double-clic => ouvrir Settings
            _trayIcon.TrayMouseDoubleClick += (s, e) => SafeOpenSettings();

            // Menu contextuel DARK
            _trayIcon.ContextMenu = BuildDarkContextMenu();
        }

        /// <summary>
        /// Construit le menu contextuel avec un thème sombre complet.
        /// Utilise les styles du ResourceDictionary DarkTheme.xaml
        /// pour éviter les artéfacts blancs du rendu par défaut Windows.
        /// </summary>
        private ContextMenu BuildDarkContextMenu()
        {
            // Le style DarkContextMenu a un ControlTemplate complet qui supprime
            // le rendu par défaut de Windows (lignes blanches, fond clair, etc.)
            var menu = new ContextMenu();
            if (FindResource("DarkContextMenu") is Style darkMenuStyle)
                menu.Style = darkMenuStyle;

            var darkItemStyle = FindResource("DarkMenuItem") as Style;
            var darkSepStyle = FindResource("DarkSeparator") as Style;

            // ── Toggle AutoSave ──
            _miToggleAutoSave = CreateDarkMenuItem("", darkItemStyle);
            UpdateToggleMenuText();
            _miToggleAutoSave.Click += (s, e) =>
            {
                try { _viewModel.ToggleAutoSaveCommand.Execute(null); UpdateToggleMenuText(); }
                catch (Exception ex) { Logger.Log($"[!] Toggle AutoSave: {ex.Message}", Logger.LogLevel.ERROR); }
            };

            // ── Sauvegarder maintenant ──
            var miSaveNow = CreateDarkMenuItem("💾  Sauvegarder maintenant", darkItemStyle);
            miSaveNow.Click += (s, e) =>
            {
                try { _viewModel.SaveNowCommand.Execute(null); }
                catch (Exception ex) { Logger.Log($"[!] SaveNow: {ex.Message}", Logger.LogLevel.ERROR); }
            };

            // ── Info documents (lecture seule, mis a jour dynamiquement) ──
            _miDocumentsInfo = CreateDarkMenuItem("📄  Documents: --", darkItemStyle);
            _miDocumentsInfo.IsEnabled = false;

            // ── Mode de sauvegarde (sous-menu) ──
            _miSaveMode = CreateDarkMenuItem("💡  Mode de sauvegarde", darkItemStyle);
            _miSaveActive = CreateDarkMenuItem("💾  Document actif (recommande)", darkItemStyle);
            _miSaveActive.IsCheckable = true;
            _miSaveAll = CreateDarkMenuItem("💾  Tous les documents", darkItemStyle);
            _miSaveAll.IsCheckable = true;
            _miSaveActive.IsChecked = _settingsService.Current.SaveMode == SaveMode.SaveActive;
            _miSaveAll.IsChecked = _settingsService.Current.SaveMode == SaveMode.SaveAll;
            _miSaveActive.Click += (s, e) =>
            {
                _viewModel.ChangeSaveModeCommand.Execute(SaveMode.SaveActive);
                _miSaveActive.IsChecked = true; _miSaveAll.IsChecked = false;
            };
            _miSaveAll.Click += (s, e) =>
            {
                _viewModel.ChangeSaveModeCommand.Execute(SaveMode.SaveAll);
                _miSaveAll.IsChecked = true; _miSaveActive.IsChecked = false;
            };
            _miSaveMode.Items.Add(_miSaveActive);
            _miSaveMode.Items.Add(_miSaveAll);

            // ── Intervalle (sous-menu) ──
            _miIntervalGroup = CreateDarkMenuItem("⏱️  Intervalle de sauvegarde", darkItemStyle);
            int[] intervals = [10, 30, 60, 120, 180, 300, 600, 900, 1200, 1800];
            foreach (int sec in intervals)
            {
                string label = AutoSaveTimerService.FormatInterval(sec);
                var mi = CreateDarkMenuItem(label, darkItemStyle);
                mi.IsCheckable = true;
                mi.IsChecked = _settingsService.Current.SaveIntervalSeconds == sec;
                mi.Tag = sec;
                mi.Click += (s, e) =>
                {
                    if (mi.Tag is int selectedSec)
                    {
                        _viewModel.ChangeIntervalCommand.Execute(selectedSec);
                        foreach (MenuItem item in _miIntervalGroup!.Items)
                            item.IsChecked = (item.Tag is int t && t == selectedSec);
                    }
                };
                _miIntervalGroup.Items.Add(mi);
            }

            // ── Notifications (sous-menu) ──
            _miNotifications = CreateDarkMenuItem("🔔  Notifications", darkItemStyle);
            _miNotifOn = CreateDarkMenuItem("✅  Activer", darkItemStyle);
            _miNotifOn.IsCheckable = true;
            _miNotifOff = CreateDarkMenuItem("❌  Desactiver", darkItemStyle);
            _miNotifOff.IsCheckable = true;
            _miNotifOn.IsChecked = _settingsService.Current.ShowNotifications;
            _miNotifOff.IsChecked = !_settingsService.Current.ShowNotifications;
            _miNotifOn.Click += (s, e) =>
            {
                _settingsService.Update(x => x.ShowNotifications = true);
                _miNotifOn.IsChecked = true; _miNotifOff.IsChecked = false;
                _viewModel.OnSettingsChangedExternally();
            };
            _miNotifOff.Click += (s, e) =>
            {
                _settingsService.Update(x => x.ShowNotifications = false);
                _miNotifOff.IsChecked = true; _miNotifOn.IsChecked = false;
                _viewModel.OnSettingsChangedExternally();
            };
            _miNotifications.Items.Add(_miNotifOn);
            _miNotifications.Items.Add(_miNotifOff);

            // ── Protection calculs ──
            _miSafetyChecks = CreateDarkMenuItem("🛡️  Protection calculs", darkItemStyle);
            _miSafeOn = CreateDarkMenuItem("✅  Activer", darkItemStyle);
            _miSafeOn.IsCheckable = true;
            _miSafeOff = CreateDarkMenuItem("❌  Desactiver", darkItemStyle);
            _miSafeOff.IsCheckable = true;
            _miSafeOn.IsChecked = _settingsService.Current.SafetyChecks;
            _miSafeOff.IsChecked = !_settingsService.Current.SafetyChecks;
            _miSafeOn.Click += (s, e) =>
            {
                _settingsService.Update(x => x.SafetyChecks = true);
                _miSafeOn.IsChecked = true; _miSafeOff.IsChecked = false;
                _viewModel.OnSettingsChangedExternally();
            };
            _miSafeOff.Click += (s, e) =>
            {
                _settingsService.Update(x => x.SafetyChecks = false);
                _miSafeOff.IsChecked = true; _miSafeOn.IsChecked = false;
                _viewModel.OnSettingsChangedExternally();
            };
            _miSafetyChecks.Items.Add(_miSafeOn);
            _miSafetyChecks.Items.Add(_miSafeOff);

            // ── Demarrer avec Windows ──
            _miStartup = CreateDarkMenuItem("", darkItemStyle);
            _miStartup.IsCheckable = true;
            _miStartup.IsChecked = StartupManager.IsStartupEnabled;
            UpdateStartupMenuText();
            _miStartup.Click += (s, e) =>
            {
                try
                {
                    if (_miStartup.IsChecked)
                    {
                        StartupManager.EnableStartup();
                        _settingsService.Update(x => x.StartWithWindows = true);
                    }
                    else
                    {
                        StartupManager.DisableStartup();
                        _settingsService.Update(x => x.StartWithWindows = false);
                    }
                    UpdateStartupMenuText();
                }
                catch (Exception ex)
                {
                    Logger.Log($"[!] Startup toggle: {ex.Message}", Logger.LogLevel.ERROR);
                }
            };

            // ── Ouvrir configuration ──
            var miOpenSettings = CreateDarkMenuItem("⚙️  Ouvrir la configuration", darkItemStyle);
            miOpenSettings.Click += (s, e) => SafeOpenSettings();

            // ── Quitter ──
            var miQuit = CreateDarkMenuItem("❌  Quitter", darkItemStyle);
            miQuit.Click += (s, e) => Shutdown();

            // Separateurs
            var sep1 = new Separator();
            if (darkSepStyle != null) sep1.Style = darkSepStyle;
            var sep2 = new Separator();
            if (darkSepStyle != null) sep2.Style = darkSepStyle;
            var sep3 = new Separator();
            if (darkSepStyle != null) sep3.Style = darkSepStyle;

            // Assembler le menu
            menu.Items.Add(_miToggleAutoSave);
            menu.Items.Add(miSaveNow);
            menu.Items.Add(_miDocumentsInfo);
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

            // Synchroniser l'etat du menu a chaque ouverture
            // (rattrape les changements faits via la fenetre Settings)
            menu.Opened += (s, e) => RefreshContextMenuState();

            return menu;
        }

        /// <summary>
        /// Cree un MenuItem avec le style dark complet (ControlTemplate).
        /// </summary>
        private static MenuItem CreateDarkMenuItem(string header, Style? style)
        {
            var mi = new MenuItem { Header = header };
            if (style != null) mi.Style = style;
            return mi;
        }

        private void UpdateToggleMenuText()
        {
            if (_miToggleAutoSave == null) return;
            bool enabled = _settingsService.Current.EnableAutoSave;
            _miToggleAutoSave.Header = enabled ? "⏸️  Desactiver AutoSave" : "▶️  Activer AutoSave";
        }

        private void UpdateStartupMenuText()
        {
            if (_miStartup == null) return;
            _miStartup.Header = _miStartup.IsChecked
                ? "✅  Demarrer avec Windows"
                : "⚙️  Demarrer avec Windows";
        }

        /// <summary>
        /// Synchronise l'état de tous les éléments du menu contextuel
        /// avec les settings actuels. Appelée quand le menu s'ouvre
        /// et quand la fenêtre Settings se ferme.
        /// </summary>
        private void RefreshContextMenuState()
        {
            try
            {
                var s = _settingsService.Current;

                // Toggle AutoSave
                UpdateToggleMenuText();

                // Mode de sauvegarde
                if (_miSaveActive != null) _miSaveActive.IsChecked = (s.SaveMode == SaveMode.SaveActive);
                if (_miSaveAll != null) _miSaveAll.IsChecked = (s.SaveMode == SaveMode.SaveAll);

                // Intervalle
                if (_miIntervalGroup != null)
                {
                    foreach (MenuItem item in _miIntervalGroup.Items)
                    {
                        if (item.Tag is int t)
                            item.IsChecked = (t == s.SaveIntervalSeconds);
                    }
                }

                // Notifications
                if (_miNotifOn != null) _miNotifOn.IsChecked = s.ShowNotifications;
                if (_miNotifOff != null) _miNotifOff.IsChecked = !s.ShowNotifications;

                // Protection calculs
                if (_miSafeOn != null) _miSafeOn.IsChecked = s.SafetyChecks;
                if (_miSafeOff != null) _miSafeOff.IsChecked = !s.SafetyChecks;

                // Demarrage avec Windows
                if (_miStartup != null)
                {
                    _miStartup.IsChecked = s.StartWithWindows;
                    UpdateStartupMenuText();
                }

                // Info documents
                if (_miDocumentsInfo != null)
                {
                    if (_viewModel.IsInventorConnected)
                    {
                        _miDocumentsInfo.Header = $"📄  Documents: {_viewModel.TotalDocuments} ouvert(s), {_viewModel.DirtyDocuments} modifie(s)";
                    }
                    else
                    {
                        _miDocumentsInfo.Header = "📄  Documents: Inventor non connecte";
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] RefreshContextMenuState: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // FENETRE SETTINGS - avec protection contre les crashes
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Ouvre la fenêtre Settings de manière sécurisée.
        /// Le crash d'origine venait du fait que la fenêtre pouvait être créée
        /// alors que l'ancienne n'était pas correctement disposed, ou qu'une
        /// exception non gérée dans le constructeur tuait l'app.
        /// </summary>
        private void SafeOpenSettings()
        {
            try
            {
                // S'assurer qu'on est sur le thread UI
                if (!Dispatcher.CheckAccess())
                {
                    Dispatcher.Invoke(SafeOpenSettings);
                    return;
                }

                if (_settingsWindow != null && _settingsWindow.IsLoaded)
                {
                    // Fenêtre déjà ouverte -> activer
                    _settingsWindow.Activate();
                    if (_settingsWindow.WindowState == WindowState.Minimized)
                        _settingsWindow.WindowState = WindowState.Normal;
                    _settingsWindow.Focus();
                    return;
                }

                // Nettoyer l'ancienne référence si elle existe encore
                _settingsWindow = null;

                // Créer la nouvelle fenêtre
                _settingsWindow = new SettingsWindow(_viewModel);
                _settingsWindow.Closed += (s, e) =>
                {
                    _settingsWindow = null;
                    RefreshContextMenuState(); // Sync menu contextuel avec les changements
                    Logger.Log("[i] Fenetre Settings fermee", Logger.LogLevel.DEBUG);
                };

                _settingsWindow.Show();
                _settingsWindow.Activate();

                Logger.Log("[+] Fenetre Settings ouverte", Logger.LogLevel.DEBUG);
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur ouverture Settings: {ex.Message}\n{ex.StackTrace}", Logger.LogLevel.ERROR);
                _settingsWindow = null;

                // Ne PAS laisser l'exception remonter et tuer l'app
                try
                {
                    MessageBox.Show(
                        $"Impossible d'ouvrir la configuration:\n{ex.Message}",
                        "InventorAutoSave",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
                catch { }
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // NOTIFICATIONS SYSTRAY
        // ═══════════════════════════════════════════════════════════════

        private void OnSaveCompleted(object? sender, SaveResult result)
        {
            try
            {
                if (!_settingsService.Current.ShowNotifications) return;
                if (!result.Success) return;
                if (result.DocumentsSaved == 0) return;

                Dispatcher.Invoke(() =>
                {
                    string message = result.Mode == SaveMode.SaveActive
                        ? "Document actif sauvegarde"
                        : $"{result.DocumentsSaved} document(s) sauvegarde(s)";

                    _trayIcon?.ShowBalloonTip(
                        "Inventor AutoSave",
                        $"✅ {message}",
                        BalloonIcon.Info);
                });
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Notification SaveCompleted: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        private void OnStatusChanged(object? sender, string status)
        {
            try
            {
                Dispatcher.Invoke(() =>
                {
                    if (_trayIcon != null)
                        _trayIcon.ToolTipText = $"Inventor AutoSave - {status}";
                });
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] StatusChanged: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        private void OnInventorConnected(object? sender, EventArgs e)
        {
            try
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
            catch (Exception ex)
            {
                Logger.Log($"[!] InventorConnected: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        private void OnInventorDisconnected(object? sender, EventArgs e)
        {
            try
            {
                Dispatcher.Invoke(() =>
                {
                    if (_trayIcon != null)
                        _trayIcon.ToolTipText = "Inventor AutoSave - Inventor non connecte";
                });
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] InventorDisconnected: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // STARTUP WINDOWS
        // ═══════════════════════════════════════════════════════════════

        private void SyncStartupShortcut()
        {
            try
            {
                bool settingEnabled = _settingsService.Current.StartWithWindows;
                bool shortcutExists = StartupManager.IsStartupEnabled;

                if (settingEnabled && !shortcutExists)
                    StartupManager.EnableStartup();
                else if (!settingEnabled && shortcutExists)
                {
                    _settingsService.Update(x => x.StartWithWindows = true);
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] SyncStartupShortcut: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // SHUTDOWN
        // ═══════════════════════════════════════════════════════════════

        protected override void OnExit(ExitEventArgs e)
        {
            Logger.Log("[i] Arret de l'application", Logger.LogLevel.INFO);

            try { _settingsWindow?.Close(); } catch { }
            try { _trayIcon?.Dispose(); } catch { }
            try { _timerService?.Dispose(); } catch { }
            try { _inventorService?.Disconnect(); } catch { }

            base.OnExit(e);
        }
    }
}
