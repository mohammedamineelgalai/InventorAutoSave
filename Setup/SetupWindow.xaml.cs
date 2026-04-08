using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;

namespace InventorAutoSave.Setup
{
    public partial class SetupWindow : Window
    {
        // ═══════════════════════════════════════════════════════════════
        // CONFIGURATION INSTALLATION
        // ═══════════════════════════════════════════════════════════════

        private const string APP_NAME    = "InventorAutoSave";
        private const string APP_VERSION = "2.0.0";
        private const string EXE_NAME    = "InventorAutoSave.exe";
        private const string SHORTCUT_NAME = "InventorAutoSave.lnk";

        // Dossier d'installation cible: %APPDATA%\XNRGY\InventorAutoSave\
        private static readonly string InstallDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "XNRGY", APP_NAME);

        // Dossier Startup Windows: demarrage automatique avec Windows
        private static readonly string StartupDir =
            Environment.GetFolderPath(Environment.SpecialFolder.Startup);

        // Menu Demarrer
        private static readonly string StartMenuDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.StartMenu),
            "Programs", "XNRGY");

        // Dossier de l'exe Setup (contient InventorAutoSave.exe a cote)
        private static readonly string SetupDir =
            AppDomain.CurrentDomain.BaseDirectory;

        // Etat
        private bool _installDone = false;
#pragma warning disable CS0414
        private double _progressWidth = 0;
#pragma warning restore CS0414

        public SetupWindow()
        {
            InitializeComponent();
            TxtInstallPath.Text = InstallDir;

            // Verifier si l'app est deja installee
            string targetExe = Path.Combine(InstallDir, EXE_NAME);
            if (File.Exists(targetExe))
            {
                BtnInstall.Content = "Reinstaller  ➜";
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // BARRE TITRE CUSTOM (draggable)
        // ═══════════════════════════════════════════════════════════════

        private void TitleBar_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        // ═══════════════════════════════════════════════════════════════
        // BOUTONS
        // ═══════════════════════════════════════════════════════════════

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            if (_installDone)
            {
                Application.Current.Shutdown();
            }
            else
            {
                var result = MessageBox.Show(
                    "Annuler l'installation ?",
                    "Confirmation",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                    Application.Current.Shutdown();
            }
        }

        private async void BtnInstall_Click(object sender, RoutedEventArgs e)
        {
            if (_installDone)
            {
                // Bouton "Terminer" apres succes
                if (ChkLaunch != null && ChkLaunch.IsChecked == true)
                    LaunchApp();
                Application.Current.Shutdown();
                return;
            }

            // Lancer l'installation
            BtnInstall.IsEnabled = false;
            BtnCancel.IsEnabled  = false;
            BtnInstall.Content   = "Installation...";

            ShowPanel("progress");

            bool success = await Task.Run(() => RunInstallation());

            if (success)
            {
                ShowPanel("success");
                _installDone = true;
                BtnInstall.Content  = "Terminer  ✅";
                BtnInstall.IsEnabled = true;
                BtnCancel.Content   = "Fermer";
                BtnCancel.IsEnabled  = true;
            }
            else
            {
                ShowPanel("error");
                BtnInstall.IsEnabled = false;
                BtnCancel.Content    = "Fermer";
                BtnCancel.IsEnabled  = true;
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // LOGIQUE D'INSTALLATION (sur thread background)
        // ═══════════════════════════════════════════════════════════════

        private bool RunInstallation()
        {
            try
            {
                // ── ETAPE 1: Creer le dossier d'installation ──
                SetStep(1, "en cours");
                SetProgress(10);

                Directory.CreateDirectory(InstallDir);
                Log("[+] Dossier cree: " + InstallDir);
                SetStep(1, "ok");

                // ── ETAPE 2: Copier les fichiers ──
                SetStep(2, "en cours");
                SetProgress(30);

                string sourceExe = FindSourceExe();
                if (string.IsNullOrEmpty(sourceExe) || !File.Exists(sourceExe))
                {
                    Log("[-] InventorAutoSave.exe introuvable. Cherche dans: " + SetupDir);
                    Log("[!] Assurez-vous que InventorAutoSave.exe est a cote de ce Setup.exe");
                    ShowError($"Fichier source introuvable: {EXE_NAME}\n\nEmplacement attendu: {SetupDir}");
                    return false;
                }

                // Copier l'exe principal
                string destExe = Path.Combine(InstallDir, EXE_NAME);
                Log("[>] Copie: " + EXE_NAME);
                File.Copy(sourceExe, destExe, overwrite: true);

                // Copier config.json si present (settings par defaut)
                string sourceConfig = Path.Combine(Path.GetDirectoryName(sourceExe)!, "config.json");
                string destConfig   = Path.Combine(InstallDir, "config.json");
                if (File.Exists(sourceConfig) && !File.Exists(destConfig))
                {
                    File.Copy(sourceConfig, destConfig);
                    Log("[>] Copie: config.json (settings par defaut)");
                }

                // Copier Resources/InventorAutoSave.ico si present
                string sourceIcoDir = Path.Combine(Path.GetDirectoryName(sourceExe)!, "Resources");
                string destIcoDir   = Path.Combine(InstallDir, "Resources");
                if (Directory.Exists(sourceIcoDir))
                {
                    Directory.CreateDirectory(destIcoDir);
                    foreach (var ico in Directory.GetFiles(sourceIcoDir, "*.ico"))
                    {
                        File.Copy(ico, Path.Combine(destIcoDir, Path.GetFileName(ico)), overwrite: true);
                        Log("[>] Copie: Resources\\" + Path.GetFileName(ico));
                    }
                }

                Log("[+] Fichiers copies dans: " + InstallDir);
                SetStep(2, "ok");
                SetProgress(55);

                // ── ETAPE 3: Demarrage Windows (Startup folder) ──
                SetStep(3, "en cours");
                SetProgress(65);

                bool addStartup = false;
                Dispatcher.Invoke(() => addStartup = ChkStartup.IsChecked == true);

                string startupLnk = Path.Combine(StartupDir, SHORTCUT_NAME);

                if (addStartup)
                {
                    CreateShortcut(startupLnk, destExe, "InventorAutoSave - Demarrage automatique");
                    Log("[+] Raccourci demarrage: " + startupLnk);
                }
                else if (File.Exists(startupLnk))
                {
                    File.Delete(startupLnk);
                    Log("[i] Raccourci demarrage supprime (option desactivee)");
                }

                SetStep(3, "ok");
                SetProgress(75);

                // ── ETAPE 4: Raccourci Menu Demarrer ──
                SetStep(4, "en cours");
                SetProgress(85);

                bool addShortcut = false;
                Dispatcher.Invoke(() => addShortcut = ChkShortcut.IsChecked == true);

                if (addShortcut)
                {
                    Directory.CreateDirectory(StartMenuDir);
                    string menuLnk = Path.Combine(StartMenuDir, SHORTCUT_NAME);
                    CreateShortcut(menuLnk, destExe, "InventorAutoSave - Sauvegarde automatique Inventor");
                    Log("[+] Raccourci menu Demarrer: " + menuLnk);
                }

                SetStep(4, "ok");
                SetProgress(95);

                // ── ETAPE 5: Finalisation ──
                SetStep(5, "en cours");

                // Mettre a jour l'affichage succes
                bool willLaunch = false;
                Dispatcher.Invoke(() => willLaunch = ChkLaunch.IsChecked == true);

                Dispatcher.Invoke(() =>
                {
                    TxtSuccessInfo1.Text = "[+] Installe dans: " + InstallDir;
                    TxtSuccessInfo2.Text = addStartup
                        ? "[+] Demarrage Windows: active"
                        : "[i] Demarrage Windows: non configure";
                    TxtSuccessInfo3.Text = willLaunch
                        ? "[>] Lancement de l'application..."
                        : "[i] Lancer manuellement depuis le menu Demarrer";
                });

                SetStep(5, "ok");
                SetProgress(100);
                Log("[+] Installation terminee avec succes!");

                return true;
            }
            catch (Exception ex)
            {
                Log("[-] ERREUR: " + ex.Message);
                ShowError(ex.Message);
                return false;
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // LANCEMENT DE L'APPLICATION APRES INSTALLATION
        // ═══════════════════════════════════════════════════════════════

        private void LaunchApp()
        {
            try
            {
                string exePath = Path.Combine(InstallDir, EXE_NAME);
                if (File.Exists(exePath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = exePath,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Impossible de lancer l'application:\n{ex.Message}",
                    "Avertissement", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // HELPERS: RACCOURCIS .LNK (WScript.Shell via COM)
        // ═══════════════════════════════════════════════════════════════

        private static void CreateShortcut(string lnkPath, string targetPath, string description)
        {
            // Utilise WScript.Shell via COM (toujours disponible sur Windows)
            Type? wshType = Type.GetTypeFromProgID("WScript.Shell");
            if (wshType == null)
                throw new InvalidOperationException("WScript.Shell non disponible");

            object? shell = Activator.CreateInstance(wshType);
            if (shell == null) return;

            try
            {
                object? shortcut = wshType.InvokeMember(
                    "CreateShortcut",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null, shell, new object[] { lnkPath });

                if (shortcut == null) return;

                Type scType = shortcut.GetType();
                scType.InvokeMember("TargetPath", System.Reflection.BindingFlags.SetProperty,
                    null, shortcut, new object[] { targetPath });
                scType.InvokeMember("WorkingDirectory", System.Reflection.BindingFlags.SetProperty,
                    null, shortcut, new object[] { Path.GetDirectoryName(targetPath) ?? "" });
                scType.InvokeMember("Description", System.Reflection.BindingFlags.SetProperty,
                    null, shortcut, new object[] { description });

                // Icone
                string icoPath = Path.Combine(Path.GetDirectoryName(targetPath)!, "Resources", "InventorAutoSave.ico");
                if (File.Exists(icoPath))
                {
                    scType.InvokeMember("IconLocation", System.Reflection.BindingFlags.SetProperty,
                        null, shortcut, new object[] { icoPath });
                }

                scType.InvokeMember("Save", System.Reflection.BindingFlags.InvokeMethod,
                    null, shortcut, null);
            }
            finally
            {
                Marshal.ReleaseComObject(shell);
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // HELPER: TROUVER L'EXE SOURCE
        // ═══════════════════════════════════════════════════════════════

        private static string FindSourceExe()
        {
            // Chercher a cote du Setup.exe
            string candidate1 = Path.Combine(SetupDir, EXE_NAME);
            if (File.Exists(candidate1)) return candidate1;

            // Chercher dans le dossier parent (structure dist\)
            string? parentDir = Path.GetDirectoryName(SetupDir.TrimEnd('\\', '/'));
            if (parentDir != null)
            {
                string candidate2 = Path.Combine(parentDir, EXE_NAME);
                if (File.Exists(candidate2)) return candidate2;
            }

            return string.Empty;
        }

        // ═══════════════════════════════════════════════════════════════
        // UI HELPERS
        // ═══════════════════════════════════════════════════════════════

        private void ShowPanel(string panel)
        {
            Dispatcher.Invoke(() =>
            {
                PanelWelcome.Visibility  = panel == "welcome"  ? Visibility.Visible : Visibility.Collapsed;
                PanelProgress.Visibility = panel == "progress" ? Visibility.Visible : Visibility.Collapsed;
                PanelSuccess.Visibility  = panel == "success"  ? Visibility.Visible : Visibility.Collapsed;
                PanelError.Visibility    = panel == "error"    ? Visibility.Visible : Visibility.Collapsed;
            });
        }

        private void SetProgress(double percent)
        {
            Dispatcher.Invoke(() =>
            {
                double maxWidth = ProgressBar.ActualWidth > 0
                    ? ((Border)ProgressBar.Parent).ActualWidth
                    : 480;
                double targetWidth = maxWidth * percent / 100.0;
                ProgressBar.Width = targetWidth;
                TxtProgressStep.Text = $"Progression: {(int)percent}%";
            });
        }

        private void SetStep(int step, string state)
        {
            Dispatcher.Invoke(() =>
            {
                string icon = state switch { "ok" => "✅", "en cours" => "🔄", _ => "⬜" };
                System.Windows.Media.Brush color = state switch
                {
                    "ok"       => System.Windows.Media.Brushes.LightGreen,
                    "en cours" => System.Windows.Media.Brushes.White,
                    _          => new System.Windows.Media.SolidColorBrush(
                                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter
                                        .ConvertFromString("#666677"))
                };

                switch (step)
                {
                    case 1: Step1Icon.Text = icon; Step1Text.Foreground = color; break;
                    case 2: Step2Icon.Text = icon; Step2Text.Foreground = color; break;
                    case 3: Step3Icon.Text = icon; Step3Text.Foreground = color; break;
                    case 4: Step4Icon.Text = icon; Step4Text.Foreground = color; break;
                    case 5: Step5Icon.Text = icon; Step5Text.Foreground = color; break;
                }
            });
        }

        private void Log(string message)
        {
            Dispatcher.Invoke(() =>
            {
                var tb = new TextBlock
                {
                    Text = message,
                    Foreground = message.StartsWith("[-")
                        ? System.Windows.Media.Brushes.OrangeRed
                        : message.StartsWith("[+")
                            ? System.Windows.Media.Brushes.LightGreen
                            : new System.Windows.Media.SolidColorBrush(
                                (System.Windows.Media.Color)System.Windows.Media.ColorConverter
                                    .ConvertFromString("#888899")),
                    FontSize   = 11,
                    FontFamily = new System.Windows.Media.FontFamily("Consolas"),
                    Margin     = new Thickness(0, 1, 0, 1)
                };
                LogPanel.Children.Add(tb);
                LogScroller.ScrollToBottom();
            });
        }

        private void ShowError(string message)
        {
            Dispatcher.Invoke(() =>
            {
                TxtErrorMessage.Text = message;
                ShowPanel("error");
            });
        }
    }
}
