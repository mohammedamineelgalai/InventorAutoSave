using System.IO;

namespace InventorAutoSave.Services
{
    /// <summary>
    /// Logger simple vers fichier texte (meme pattern que XEAT)
    /// </summary>
    public static class Logger
    {
        public enum LogLevel { DEBUG, INFO, WARNING, ERROR }

        private static readonly string LogDir = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "Logs");

        private static readonly string LogFile = Path.Combine(
            LogDir, $"InventorAutoSave_{DateTime.Now:yyyyMMdd}.log");

        private static readonly object _lock = new();

        static Logger()
        {
            try { Directory.CreateDirectory(LogDir); } catch { }
        }

        public static void Log(string message, LogLevel level = LogLevel.INFO)
        {
            string formatted = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level,-7}] {message}";

            lock (_lock)
            {
                try { File.AppendAllText(LogFile, formatted + Environment.NewLine); } catch { }
            }

            // Debug output
            System.Diagnostics.Debug.WriteLine(formatted);
        }
    }
}
