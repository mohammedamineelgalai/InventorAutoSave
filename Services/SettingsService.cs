using System;
using System.IO;
using System.Text.Json;
using InventorAutoSave.Models;

namespace InventorAutoSave.Services
{
    /// <summary>
    /// Gestion de la persistance des parametres via config.json
    /// Corrige le bug AHK: les settings etaient ignores a chaque restart
    /// </summary>
    public class SettingsService
    {
        private static readonly string ConfigPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "config.json");

        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true,
            Converters = { new System.Text.Json.Serialization.JsonStringEnumConverter() }
        };

        private AppSettings _current;

        public AppSettings Current => _current;

        public SettingsService()
        {
            _current = Load();
        }

        /// <summary>
        /// Charge la configuration depuis config.json (cree le fichier si absent)
        /// </summary>
        public AppSettings Load()
        {
            try
            {
                if (File.Exists(ConfigPath))
                {
                    string json = File.ReadAllText(ConfigPath);
                    var settings = JsonSerializer.Deserialize<AppSettings>(json, JsonOptions);
                    if (settings != null)
                    {
                        _current = settings;
                        Logger.Log("[+] Configuration chargee depuis config.json", Logger.LogLevel.INFO);
                        return _current;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur lecture config.json: {ex.Message} - utilisation valeurs par defaut", Logger.LogLevel.WARNING);
            }

            // Valeurs par defaut si fichier absent/corrompu
            _current = new AppSettings();
            Save(_current);
            return _current;
        }

        /// <summary>
        /// Persiste la configuration dans config.json
        /// </summary>
        public void Save(AppSettings settings)
        {
            try
            {
                _current = settings;
                string json = JsonSerializer.Serialize(settings, JsonOptions);
                File.WriteAllText(ConfigPath, json);
                Logger.Log("[+] Configuration sauvegardee dans config.json", Logger.LogLevel.INFO);
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur ecriture config.json: {ex.Message}", Logger.LogLevel.ERROR);
            }
        }

        /// <summary>
        /// Met a jour un seul parametre et persiste
        /// </summary>
        public void Update(Action<AppSettings> updateAction)
        {
            updateAction(_current);
            Save(_current);
        }
    }
}
