using System.IO;
using System.Runtime.InteropServices;

namespace InventorAutoSave.Services
{
    /// <summary>
    /// Gere le demarrage automatique de l'application avec Windows.
    /// Utilise le dossier Startup de l'utilisateur (pas le registre) pour eviter
    /// les restrictions de droits administrateur.
    /// 
    /// Chemin Startup: %APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\
    /// </summary>
    public static class StartupManager
    {
        private const string SHORTCUT_NAME = "InventorAutoSave.lnk";

        private static readonly string StartupDir =
            Environment.GetFolderPath(Environment.SpecialFolder.Startup);

        private static readonly string ShortcutPath =
            Path.Combine(StartupDir, SHORTCUT_NAME);

        /// <summary>
        /// Verifie si le raccourci de demarrage existe
        /// </summary>
        public static bool IsStartupEnabled => File.Exists(ShortcutPath);

        /// <summary>
        /// Active le demarrage automatique avec Windows
        /// </summary>
        public static void EnableStartup()
        {
            try
            {
                string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                // Si publie en SingleFile, .Location retourne "" -> utiliser le process
                if (string.IsNullOrEmpty(exePath))
                    exePath = Environment.ProcessPath ?? "";

                if (string.IsNullOrEmpty(exePath) || !File.Exists(exePath))
                {
                    Logger.Log("[!] StartupManager: impossible de determiner le chemin exe", Logger.LogLevel.WARNING);
                    return;
                }

                CreateShortcut(ShortcutPath, exePath, "InventorAutoSave - Sauvegarde automatique Inventor");
                Logger.Log("[+] Demarrage Windows active: " + ShortcutPath, Logger.LogLevel.INFO);
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] StartupManager.EnableStartup: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        /// <summary>
        /// Desactive le demarrage automatique avec Windows
        /// </summary>
        public static void DisableStartup()
        {
            try
            {
                if (File.Exists(ShortcutPath))
                {
                    File.Delete(ShortcutPath);
                    Logger.Log("[i] Demarrage Windows desactive", Logger.LogLevel.INFO);
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] StartupManager.DisableStartup: {ex.Message}", Logger.LogLevel.WARNING);
            }
        }

        /// <summary>
        /// Cree un raccourci .lnk via WScript.Shell (COM, toujours disponible sur Windows)
        /// </summary>
        private static void CreateShortcut(string lnkPath, string targetPath, string description)
        {
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
                scType.InvokeMember("TargetPath",
                    System.Reflection.BindingFlags.SetProperty, null, shortcut,
                    new object[] { targetPath });
                scType.InvokeMember("WorkingDirectory",
                    System.Reflection.BindingFlags.SetProperty, null, shortcut,
                    new object[] { Path.GetDirectoryName(targetPath) ?? "" });
                scType.InvokeMember("Description",
                    System.Reflection.BindingFlags.SetProperty, null, shortcut,
                    new object[] { description });

                // Icone (chercher InventorAutoSave.ico a cote de l'exe)
                string icoPath = Path.Combine(
                    Path.GetDirectoryName(targetPath) ?? "", "Resources", "InventorAutoSave.ico");
                if (File.Exists(icoPath))
                {
                    scType.InvokeMember("IconLocation",
                        System.Reflection.BindingFlags.SetProperty, null, shortcut,
                        new object[] { icoPath });
                }

                scType.InvokeMember("Save",
                    System.Reflection.BindingFlags.InvokeMethod, null, shortcut, null);
            }
            finally
            {
                Marshal.ReleaseComObject(shell);
            }
        }
    }
}
