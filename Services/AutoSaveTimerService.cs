using System.Timers;
using System.Windows;
using InventorAutoSave.Models;

namespace InventorAutoSave.Services
{
    /// <summary>
    /// Gestion du timer AutoSave avec intervalle configurable.
    ///
    /// CORRECTION BUG AHK: Dans le script AHK, les settings n'etaient pas persistes.
    /// Chaque restart du script remettait l'intervalle a la valeur hardcodee (180s).
    /// Ici, l'intervalle est lu depuis config.json AU DEMARRAGE et persiste a chaque changement.
    ///
    /// CORRECTION BUG COM THREADING: Les appels COM vers Inventor (serveur STA) doivent
    /// etre effectues depuis un thread STA. System.Timers.Timer fire sur le ThreadPool (MTA)
    /// ce qui cause des MissingMethodException sur les appels Type.InvokeMember.
    /// Solution: marshaller TriggerSave sur le thread UI (STA) via Dispatcher.Invoke.
    /// </summary>
    public class AutoSaveTimerService : IDisposable
    {
        private readonly InventorSaveService _inventorService;
        private readonly SettingsService _settingsService;
        private System.Timers.Timer? _timer;
        private bool _isSaving;
        private DateTime _lastSaveTime = DateTime.MinValue;
        private int _pendingRetryCount;
        private System.Timers.Timer? _retryTimer;

        // Evenements pour l'UI
        public event EventHandler<SaveResult>? SaveCompleted;
        public event EventHandler<string>? StatusChanged;
        public event EventHandler? TimerTick;

        public bool IsRunning => _timer?.Enabled ?? false;
        public DateTime LastSaveTime => _lastSaveTime;
        public int IntervalSeconds => _settingsService.Current.SaveIntervalSeconds;

        public AutoSaveTimerService(InventorSaveService inventorService, SettingsService settingsService)
        {
            _inventorService = inventorService;
            _settingsService = settingsService;

            // S'abonner aux evenements Inventor
            _inventorService.Connected += OnInventorConnected;
            _inventorService.Disconnected += OnInventorDisconnected;
        }

        // ═══════════════════════════════════════════════════════════════
        // GESTION DU TIMER
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Demarre le timer AutoSave avec l'intervalle de config.json
        /// </summary>
        public void Start()
        {
            // Arreter l'ancien timer proprement
            Stop();

            int intervalMs = _settingsService.Current.SaveIntervalSeconds * 1000;
            if (intervalMs <= 0) intervalMs = 180_000; // Securite: minimum 3 min

            _timer = new System.Timers.Timer(intervalMs);
            _timer.Elapsed += OnTimerElapsed;
            _timer.AutoReset = true;
            _timer.Start();

            Logger.Log($"[+] AutoSave demarre - intervalle: {_settingsService.Current.SaveIntervalSeconds}s", Logger.LogLevel.INFO);
            StatusChanged?.Invoke(this, $"AutoSave actif ({FormatInterval(_settingsService.Current.SaveIntervalSeconds)})");
        }

        /// <summary>
        /// Arrete le timer
        /// </summary>
        public void Stop()
        {
            if (_timer != null)
            {
                _timer.Stop();
                _timer.Elapsed -= OnTimerElapsed;
                _timer.Dispose();
                _timer = null;
            }
            StopRetryTimer();
            Logger.Log("[i] AutoSave arrete", Logger.LogLevel.INFO);
        }

        /// <summary>
        /// Change l'intervalle et redémarre le timer immédiatement avec le nouvel intervalle.
        /// CORRECTION BUG AHK: ici le changement est effectif IMMEDIATEMENT et persiste.
        /// </summary>
        public void ChangeInterval(int newIntervalSeconds)
        {
            _settingsService.Update(s => s.SaveIntervalSeconds = newIntervalSeconds);

            if (IsRunning)
            {
                Start(); // Redemarre avec le nouvel intervalle
            }

            Logger.Log($"[>] Intervalle modifie: {newIntervalSeconds}s", Logger.LogLevel.INFO);
            StatusChanged?.Invoke(this, $"Intervalle: {FormatInterval(newIntervalSeconds)}");
        }

        // ═══════════════════════════════════════════════════════════════
        // SAUVEGARDE DECLENCHEE PAR LE TIMER
        // ═══════════════════════════════════════════════════════════════

        private void OnTimerElapsed(object? sender, ElapsedEventArgs e)
        {
            try
            {
                TimerTick?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] TimerTick event erreur: {ex.Message}", Logger.LogLevel.DEBUG);
            }

            // Eviter les sauvegardes concurrentes
            if (_isSaving)
            {
                Logger.Log("[~] Sauvegarde en cours, tick ignore", Logger.LogLevel.DEBUG);
                return;
            }

            try
            {
                // IMPORTANT: Les appels COM vers Inventor (STA) doivent etre sur le thread UI (STA).
                // System.Timers.Timer fire sur le ThreadPool (MTA), donc on marshalle via Dispatcher.
                var dispatcher = Application.Current?.Dispatcher;
                if (dispatcher != null && !dispatcher.CheckAccess())
                {
                    dispatcher.Invoke(() => TriggerSave());
                }
                else
                {
                    TriggerSave();
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] OnTimerElapsed erreur non geree: {ex.Message}", Logger.LogLevel.ERROR);
            }
        }

        /// <summary>
        /// Declenche une sauvegarde (appelee par le timer OU par SaveNow)
        /// </summary>
        public SaveResult TriggerSave()
        {
            if (_isSaving)
            {
                return new SaveResult { Success = false, ErrorMessage = "Sauvegarde deja en cours" };
            }

            _isSaving = true;
            try
            {
                // Reconnexion si necessaire
                if (!_inventorService.IsConnected)
                {
                    _inventorService.TryConnect();
                }

                if (!_inventorService.IsConnected)
                {
                    Logger.Log("[~] AutoSave: Inventor non connecte, sauvegarde ignoree", Logger.LogLevel.DEBUG);
                    return new SaveResult { Success = false, ErrorMessage = "Inventor non connecte" };
                }

                // Verification de securite: Inventor en calcul?
                if (_settingsService.Current.SafetyChecks && _inventorService.IsInventorCalculating())
                {
                    Logger.Log("[~] AutoSave: Inventor en calcul, sauvegarde reportee", Logger.LogLevel.INFO);
                    StatusChanged?.Invoke(this, "Calcul en cours - sauvegarde reportee");
                    StartRetryTimer();
                    return new SaveResult { Success = false, ErrorMessage = "Inventor en calcul" };
                }

                // SAUVEGARDE SILENCIEUSE VIA API COM
                var result = _inventorService.Save(_settingsService.Current.SaveMode);

                if (result.Success && result.DocumentsSaved > 0)
                {
                    _lastSaveTime = DateTime.Now;
                    StopRetryTimer();
                    _pendingRetryCount = 0;
                }
                else if (result.Success && result.DocumentsSaved == 0)
                {
                    Logger.Log("[~] AutoSave: rien a sauvegarder", Logger.LogLevel.DEBUG);
                }

                SaveCompleted?.Invoke(this, result);
                return result;
            }
            finally
            {
                _isSaving = false;
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // RETRY TIMER (si Inventor en calcul)
        // ═══════════════════════════════════════════════════════════════

        private void StartRetryTimer()
        {
            StopRetryTimer();
            int retryMs = _settingsService.Current.RetryDelaySeconds * 1000;
            _retryTimer = new System.Timers.Timer(retryMs);
            _retryTimer.Elapsed += OnRetryElapsed;
            _retryTimer.AutoReset = false;
            _retryTimer.Start();
        }

        private void StopRetryTimer()
        {
            if (_retryTimer != null)
            {
                _retryTimer.Stop();
                _retryTimer.Elapsed -= OnRetryElapsed;
                _retryTimer.Dispose();
                _retryTimer = null;
            }
        }

        private void OnRetryElapsed(object? sender, ElapsedEventArgs e)
        {
            _pendingRetryCount++;

            if (!_inventorService.IsConnected) return;

            // Marshal sur le thread UI (STA) pour les appels COM
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher != null && !dispatcher.CheckAccess())
            {
                dispatcher.Invoke(() => OnRetryElapsedCore());
            }
            else
            {
                OnRetryElapsedCore();
            }
        }

        private void OnRetryElapsedCore()
        {
            if (_inventorService.IsInventorCalculating())
            {
                // Max 60 retries (~5 min), puis forcer quand meme
                if (_pendingRetryCount < 60)
                {
                    Logger.Log($"[~] Retry {_pendingRetryCount}: Inventor encore en calcul", Logger.LogLevel.DEBUG);
                    StartRetryTimer(); // Retry suivant
                    return;
                }
                Logger.Log("[!] Sauvegarde forcee apres 5 min d'attente calcul", Logger.LogLevel.WARNING);
            }

            TriggerSave();
        }

        // ═══════════════════════════════════════════════════════════════
        // HELPERS
        // ═══════════════════════════════════════════════════════════════

        private void OnInventorConnected(object? sender, EventArgs e)
        {
            StatusChanged?.Invoke(this, "Inventor connecte");
        }

        private void OnInventorDisconnected(object? sender, EventArgs e)
        {
            StatusChanged?.Invoke(this, "Inventor deconnecte");
        }

        public static string FormatInterval(int seconds)
        {
            if (seconds < 60) return $"{seconds}s";
            if (seconds < 3600) return $"{seconds / 60} min";
            return $"{seconds / 3600}h {(seconds % 3600) / 60}min";
        }

        public void Dispose()
        {
            _inventorService.Connected -= OnInventorConnected;
            _inventorService.Disconnected -= OnInventorDisconnected;
            Stop();
            _inventorService.Disconnect();
        }
    }
}
