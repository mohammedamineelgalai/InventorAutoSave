using System.Text.Json.Serialization;

namespace InventorAutoSave.Models
{
    /// <summary>
    /// Mode de sauvegarde Inventor (via API COM silencieuse)
    /// </summary>
    public enum SaveMode
    {
        /// <summary>
        /// Sauvegarde uniquement le document actif (pas de popup si seul ce doc est modifie)
        /// </summary>
        SaveActive,

        /// <summary>
        /// Sauvegarde tous les documents ouverts (Parts -> Assemblies -> Drawings)
        /// </summary>
        SaveAll
    }

    /// <summary>
    /// Configuration persistee dans config.json
    /// </summary>
    public class AppSettings
    {
        [JsonPropertyName("SaveIntervalSeconds")]
        public int SaveIntervalSeconds { get; set; } = 180;

        [JsonPropertyName("SaveMode")]
        [JsonConverter(typeof(JsonStringEnumConverter))]
        public SaveMode SaveMode { get; set; } = SaveMode.SaveActive;

        [JsonPropertyName("EnableAutoSave")]
        public bool EnableAutoSave { get; set; } = true;

        [JsonPropertyName("ShowNotifications")]
        public bool ShowNotifications { get; set; } = true;

        [JsonPropertyName("NotificationDurationSeconds")]
        public int NotificationDurationSeconds { get; set; } = 3;

        [JsonPropertyName("StartWithWindows")]
        public bool StartWithWindows { get; set; } = false;

        [JsonPropertyName("SafetyChecks")]
        public bool SafetyChecks { get; set; } = true;

        [JsonPropertyName("RetryDelaySeconds")]
        public int RetryDelaySeconds { get; set; } = 5;
    }

    /// <summary>
    /// Resultat d'une tentative de sauvegarde
    /// </summary>
    public class SaveResult
    {
        public bool Success { get; set; }
        public int DocumentsSaved { get; set; }
        public int DocumentsSkipped { get; set; }
        public string? ErrorMessage { get; set; }
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public SaveMode Mode { get; set; }

        public string Summary => Success
            ? $"{DocumentsSaved} doc(s) sauvegarde(s)"
            : $"Erreur: {ErrorMessage}";
    }
}
