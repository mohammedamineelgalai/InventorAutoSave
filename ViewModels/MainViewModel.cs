using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using InventorAutoSave.Models;
using InventorAutoSave.Services;

namespace InventorAutoSave.ViewModels
{
    /// <summary>
    /// ViewModel principal - coordonne tous les services
    /// </summary>
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly InventorSaveService _inventorService;
        private readonly AutoSaveTimerService _timerService;
        private readonly SettingsService _settingsService;
        private readonly System.Windows.Threading.DispatcherTimer _uiRefreshTimer;

        // ═══════════════════════════════════════════════════════════════
        // PROPRIETES OBSERVABLES
        // ═══════════════════════════════════════════════════════════════

        private bool _isInventorConnected;
        public bool IsInventorConnected
        {
            get => _isInventorConnected;
            set { _isInventorConnected = value; OnPropertyChanged(); OnPropertyChanged(nameof(InventorStatusText)); OnPropertyChanged(nameof(InventorStatusColor)); }
        }

        private bool _isAutoSaveEnabled;
        public bool IsAutoSaveEnabled
        {
            get => _isAutoSaveEnabled;
            set { _isAutoSaveEnabled = value; OnPropertyChanged(); OnPropertyChanged(nameof(AutoSaveButtonText)); OnPropertyChanged(nameof(AutoSaveButtonColor)); }
        }

        private string _lastSaveText = "Jamais";
        public string LastSaveText
        {
            get => _lastSaveText;
            set { _lastSaveText = value; OnPropertyChanged(); }
        }

        private string _statusMessage = "Demarrage...";
        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value; OnPropertyChanged(); }
        }

        private string _activeDocumentName = "Aucun document";
        public string ActiveDocumentName
        {
            get => _activeDocumentName;
            set { _activeDocumentName = value; OnPropertyChanged(); }
        }

        private int _totalDocuments;
        public int TotalDocuments
        {
            get => _totalDocuments;
            set { _totalDocuments = value; OnPropertyChanged(); OnPropertyChanged(nameof(DocumentsStatusText)); }
        }

        private int _dirtyDocuments;
        public int DirtyDocuments
        {
            get => _dirtyDocuments;
            set { _dirtyDocuments = value; OnPropertyChanged(); OnPropertyChanged(nameof(DocumentsStatusText)); }
        }

        private string _nextSaveText = "--:--";
        public string NextSaveText
        {
            get => _nextSaveText;
            set { _nextSaveText = value; OnPropertyChanged(); }
        }

        // Proprietes calculees (pas de setter, juste OnPropertyChanged)
        public string InventorStatusText => IsInventorConnected ? "Inventor : Connecte" : "Inventor : Deconnecte";
        public string InventorStatusColor => IsInventorConnected ? "#107C10" : "#E81123";
        public string AutoSaveButtonText => IsAutoSaveEnabled ? "Desactiver AutoSave" : "Activer AutoSave";
        public string AutoSaveButtonColor => IsAutoSaveEnabled ? "#C62828" : "#107C10";
        public string DocumentsStatusText => IsInventorConnected
            ? $"{TotalDocuments} doc(s) ouvert(s), {DirtyDocuments} modifie(s)"
            : "Inventor non connecte";

        // Acces aux settings
        public AppSettings Settings => _settingsService.Current;

        // ═══════════════════════════════════════════════════════════════
        // TIMER DE PROCHAIN SAVE (compte a rebours)
        // ═══════════════════════════════════════════════════════════════
        private DateTime _timerStartedAt = DateTime.MinValue;

        // ═══════════════════════════════════════════════════════════════
        // CONSTRUCTEUR
        // ═══════════════════════════════════════════════════════════════

        public MainViewModel(InventorSaveService inventorService,
                             AutoSaveTimerService timerService,
                             SettingsService settingsService)
        {
            _inventorService = inventorService;
            _timerService = timerService;
            _settingsService = settingsService;

            // Etat initial depuis la config
            _isAutoSaveEnabled = settingsService.Current.EnableAutoSave;

            // S'abonner aux evenements
            _inventorService.Connected += (s, e) =>
                App.Current.Dispatcher.Invoke(() => IsInventorConnected = true);
            _inventorService.Disconnected += (s, e) =>
                App.Current.Dispatcher.Invoke(() => { IsInventorConnected = false; ActiveDocumentName = "Aucun document"; TotalDocuments = 0; DirtyDocuments = 0; });

            _timerService.SaveCompleted += (s, e) =>
                App.Current.Dispatcher.BeginInvoke(() => OnSaveCompleted(e));
            _timerService.StatusChanged += (s, msg) =>
                App.Current.Dispatcher.BeginInvoke(() => StatusMessage = msg);
            _timerService.TimerTick += (s, e) =>
                App.Current.Dispatcher.BeginInvoke(() => _timerStartedAt = DateTime.Now);

            // Timer UI: refresh toutes les secondes (compte a rebours, nom doc actif, etc.)
            _uiRefreshTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _uiRefreshTimer.Tick += OnUiRefreshTick;
            _uiRefreshTimer.Start();

            // Demarrage initial
            _ = Task.Run(() => InitializeAsync());
        }

        private async Task InitializeAsync()
        {
            await Task.Delay(500); // Laisser l'UI s'initialiser

            // Tentative de connexion initiale - DOIT etre sur le thread UI (STA)
            // car les objets COM Inventor sont STA et doivent etre crees/accedes
            // depuis un thread STA pour que Type.InvokeMember fonctionne.
            bool connected = false;
            await App.Current.Dispatcher.InvokeAsync(() =>
            {
                connected = _inventorService.TryConnect();
            });

            App.Current.Dispatcher.Invoke(() =>
            {
                IsInventorConnected = connected;
                StatusMessage = connected ? "Connecte a Inventor" : "En attente d'Inventor...";
            });

            // Demarrer AutoSave si active
            if (_settingsService.Current.EnableAutoSave)
            {
                _timerService.Start();
                _timerStartedAt = DateTime.Now;
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // REFRESH UI (toutes les secondes)
        // ═══════════════════════════════════════════════════════════════

        private void OnUiRefreshTick(object? sender, EventArgs e)
        {
            // Reconnexion automatique silencieuse
            if (!_inventorService.IsConnected)
            {
                bool reconnected = _inventorService.TryConnect();
                if (reconnected != IsInventorConnected)
                    IsInventorConnected = reconnected;
            }
            else
            {
                if (!IsInventorConnected) IsInventorConnected = true;

                // Mettre a jour nom doc actif et compteurs
                string? docName = _inventorService.GetActiveDocumentName();
                ActiveDocumentName = string.IsNullOrEmpty(docName) ? "Aucun document actif" : docName;

                var (total, dirty) = _inventorService.GetDocumentCounts();
                TotalDocuments = total;
                DirtyDocuments = dirty;
            }

            // Compte a rebours prochain save
            if (_timerService.IsRunning && _timerStartedAt != DateTime.MinValue)
            {
                int intervalSec = _settingsService.Current.SaveIntervalSeconds;
                double elapsed = (DateTime.Now - _timerStartedAt).TotalSeconds;
                double remaining = intervalSec - elapsed;

                if (remaining > 0)
                {
                    int min = (int)(remaining / 60);
                    int sec = (int)(remaining % 60);
                    NextSaveText = min > 0 ? $"{min:D2}:{sec:D2}" : $"00:{sec:D2}";
                }
                else
                {
                    NextSaveText = "Maintenant...";
                    _timerStartedAt = DateTime.Now; // Reset pour le prochain cycle
                }
            }
            else
            {
                NextSaveText = _timerService.IsRunning ? "--:--" : "Arrete";
            }

            // Derniere sauvegarde
            if (_timerService.LastSaveTime != DateTime.MinValue)
            {
                var diff = DateTime.Now - _timerService.LastSaveTime;
                if (diff.TotalSeconds < 60)
                    LastSaveText = $"Il y a {(int)diff.TotalSeconds}s";
                else if (diff.TotalMinutes < 60)
                    LastSaveText = $"Il y a {(int)diff.TotalMinutes} min";
                else
                    LastSaveText = _timerService.LastSaveTime.ToString("HH:mm:ss");
            }
        }

        // ═══════════════════════════════════════════════════════════════
        // COMMANDES
        // ═══════════════════════════════════════════════════════════════

        private ICommand? _saveNowCommand;
        public ICommand SaveNowCommand => _saveNowCommand ??= new RelayCommand(_ =>
        {
            // SaveNow est deja sur le thread UI (appele depuis l'UI),
            // donc les appels COM seront sur le thread STA.
            var result = _timerService.TriggerSave();
            StatusMessage = result.Success
                ? $"[+] {result.Summary}"
                : $"[-] {result.ErrorMessage ?? "Erreur inconnue"}";
        });

        private ICommand? _toggleAutoSaveCommand;
        public ICommand ToggleAutoSaveCommand => _toggleAutoSaveCommand ??= new RelayCommand(_ =>
        {
            IsAutoSaveEnabled = !IsAutoSaveEnabled;
            _settingsService.Update(s => s.EnableAutoSave = IsAutoSaveEnabled);

            if (IsAutoSaveEnabled)
            {
                _timerService.Start();
                _timerStartedAt = DateTime.Now;
                StatusMessage = $"AutoSave active ({AutoSaveTimerService.FormatInterval(_settingsService.Current.SaveIntervalSeconds)})";
            }
            else
            {
                _timerService.Stop();
                NextSaveText = "Arrete";
                StatusMessage = "AutoSave desactive";
            }
        });

        private ICommand? _changeIntervalCommand;
        public ICommand ChangeIntervalCommand => _changeIntervalCommand ??= new RelayCommand(param =>
        {
            if (param is int seconds)
            {
                _timerService.ChangeInterval(seconds);
                _timerStartedAt = DateTime.Now;
                StatusMessage = $"Intervalle: {AutoSaveTimerService.FormatInterval(seconds)}";
            }
        });

        private ICommand? _changeSaveModeCommand;
        public ICommand ChangeSaveModeCommand => _changeSaveModeCommand ??= new RelayCommand(param =>
        {
            if (param is SaveMode mode)
            {
                _settingsService.Update(s => s.SaveMode = mode);
                OnPropertyChanged(nameof(Settings));
                StatusMessage = mode == SaveMode.SaveActive ? "Mode: Sauvegarder document actif" : "Mode: Sauvegarder tout";
            }
        });

        private ICommand? _toggleNotificationsCommand;
        public ICommand ToggleNotificationsCommand => _toggleNotificationsCommand ??= new RelayCommand(_ =>
        {
            _settingsService.Update(s => s.ShowNotifications = !s.ShowNotifications);
            OnPropertyChanged(nameof(Settings));
        });

        private ICommand? _toggleSafetyChecksCommand;
        public ICommand ToggleSafetyChecksCommand => _toggleSafetyChecksCommand ??= new RelayCommand(_ =>
        {
            _settingsService.Update(s => s.SafetyChecks = !s.SafetyChecks);
            OnPropertyChanged(nameof(Settings));
        });

        // ═══════════════════════════════════════════════════════════════
        // HELPERS
        // ═══════════════════════════════════════════════════════════════

        private void OnSaveCompleted(SaveResult result)
        {
            if (result.Success && result.DocumentsSaved > 0)
            {
                StatusMessage = $"[+] {result.DocumentsSaved} doc(s) sauvegarde(s)";
            }
            else if (!result.Success)
            {
                StatusMessage = $"[-] {result.ErrorMessage ?? "Echec sauvegarde"}";
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>
    /// Commande MVVM simple
    /// </summary>
    public class RelayCommand : ICommand
    {
        private readonly Action<object?> _execute;
        private readonly Func<object?, bool>? _canExecute;

        public RelayCommand(Action<object?> execute, Func<object?, bool>? canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged;
        public bool CanExecute(object? parameter) => _canExecute?.Invoke(parameter) ?? true;
        public void Execute(object? parameter) => _execute(parameter);
        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }
}
