using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using InventorAutoSave.Models;

namespace InventorAutoSave.Services
{
    // Helper interne: evite les tuples avec dynamic (non deconstructibles dans LINQ)
    internal sealed class DocEntry
    {
        public required object Doc { get; init; }
        public required string Name { get; init; }
        public required int Order { get; init; }
    }

    /// <summary>
    /// Service de connexion COM et sauvegarde silencieuse via l'API Inventor.
    /// 
    /// PRINCIPE CLE - Sauvegarde silencieuse:
    ///   1. SilentOperation = true  => Inventor supprime TOUTES les popups (confirmation, erreurs)
    ///   2. doc.Save() direct via COM => pas de SendKeys, pas de raccourci clavier
    ///   3. Verification doc.Dirty   => ne sauvegarde que les docs modifies
    ///
    /// Ce service resout le probleme principal du script AHK:
    ///   - AHK utilisait SendKeys (Ctrl+D) qui declenchait la popup "Save multiple files?"
    ///   - Ici, l'API COM sauvegarde directement, SANS aucune popup
    /// </summary>
    public class InventorSaveService
    {
        // COM Interop P/Invoke
        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID(
            [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
            out Guid pclsid);

        [DllImport("ole32.dll")]
        private static extern int GetActiveObject(
            ref Guid rclsid,
            IntPtr pvReserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        private dynamic? _inventorApp;
        private bool _isConnected;
        private DateTime _lastConnectionAttempt = DateTime.MinValue;
        private int _consecutiveFailures;
        private const int MIN_RETRY_INTERVAL_MS = 3000;

        // Evenements
        public event EventHandler<SaveResult>? SaveCompleted;
        public event EventHandler? Connected;
        public event EventHandler? Disconnected;

        public bool IsConnected => _isConnected && _inventorApp != null;

        // ═══════════════════════════════════════════════════════════════
        // CONNEXION COM
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Tente de se connecter a Inventor via COM (avec throttling)
        /// </summary>
        public bool TryConnect()
        {
            var now = DateTime.Now;
            if ((now - _lastConnectionAttempt).TotalMilliseconds < MIN_RETRY_INTERVAL_MS
                && _lastConnectionAttempt != DateTime.MinValue)
            {
                return _isConnected; // Trop tot, retourner etat actuel
            }

            _lastConnectionAttempt = now;

            // Verifier que le processus Inventor tourne
            var processes = Process.GetProcessesByName("Inventor");
            if (processes.Length == 0)
            {
                if (_isConnected) SetDisconnected();
                return false;
            }

            // Verifier que la fenetre principale est prete (pas splash screen)
            bool windowReady = processes.Any(p =>
                p.MainWindowHandle != IntPtr.Zero && !string.IsNullOrEmpty(p.MainWindowTitle));

            if (!windowReady) return false;

            // Methode 1: P/Invoke via ole32 (Marshal.GetActiveObject retire dans .NET 5+)
            try
            {
                int hr = CLSIDFromProgID("Inventor.Application", out Guid clsid);
                if (hr == 0)
                {
                    GetActiveObject(ref clsid, IntPtr.Zero, out object inventorObj);
                    if (inventorObj != null)
                    {
                        bool wasConnected = _isConnected;
                        _inventorApp = inventorObj;
                        _isConnected = true;
                        _consecutiveFailures = 0;
                        if (!wasConnected)
                        {
                            Logger.Log("[+] Connecte a Inventor via COM (P/Invoke)", Logger.LogLevel.INFO);
                            Connected?.Invoke(this, EventArgs.Empty);
                        }
                        return true;
                    }
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                Logger.Log($"[!] Erreur connexion COM: {ex.Message}", Logger.LogLevel.DEBUG);
            }

            _consecutiveFailures++;
            if (_isConnected) SetDisconnected();
            return false;
        }

        /// <summary>
        /// Force une reconnexion en reinitialisant l'etat COM
        /// </summary>
        public bool ForceReconnect()
        {
            Disconnect();
            _consecutiveFailures = 0;
            _lastConnectionAttempt = DateTime.MinValue;
            return TryConnect();
        }

        private void SetDisconnected()
        {
            _isConnected = false;
            if (_inventorApp != null)
            {
                try { Marshal.ReleaseComObject(_inventorApp); } catch { }
                _inventorApp = null;
            }
            Logger.Log("[!] Connexion Inventor perdue", Logger.LogLevel.WARNING);
            Disconnected?.Invoke(this, EventArgs.Empty);
        }

        public void Disconnect()
        {
            if (_inventorApp != null)
            {
                try { Marshal.ReleaseComObject(_inventorApp); } catch { }
                _inventorApp = null;
            }
            _isConnected = false;
        }

        // ═══════════════════════════════════════════════════════════════
        // INFORMATIONS INVENTOR
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Retourne le nom du document actif (ou null si aucun)
        /// </summary>
        public string? GetActiveDocumentName()
        {
            try
            {
                if (_inventorApp != null)
                {
                    dynamic doc = _inventorApp.ActiveDocument;
                    return doc?.DisplayName;
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// Verifie si Inventor est en train de faire un calcul (rebuild, FEA, etc.)
        /// </summary>
        public bool IsInventorCalculating()
        {
            if (!IsConnected) return false;
            try
            {
                // Verifier via l'API si Inventor est occupe
                if (_inventorApp != null)
                {
                    // CommandIsRunning retourne true si une commande est en cours
                    // On verifie aussi le titre de la fenetre active
                    string? title = null;
                    try
                    {
                        title = _inventorApp.Caption as string;
                    }
                    catch { }

                    if (title != null)
                    {
                        string[] busyKeywords = { "Calculating", "Calcul", "Processing", "Rebuilding",
                                                   "Generating", "Computing", "Working", "Loading" };
                        foreach (var kw in busyKeywords)
                        {
                            if (title.Contains(kw, StringComparison.OrdinalIgnoreCase))
                                return true;
                        }
                    }
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// Compte les documents ouverts et modifies
        /// </summary>
        public (int total, int dirty) GetDocumentCounts()
        {
            if (!IsConnected || _inventorApp == null) return (0, 0);
            try
            {
                dynamic docs = _inventorApp!.Documents;
                int total = (int)docs.Count;
                int dirty = 0;
                for (int i = 1; i <= total; i++)
                {
                    try
                    {
                        dynamic doc = docs.Item(i);
                        if ((bool)doc.Dirty) dirty++;
                    }
                    catch { }
                }
                return (total, dirty);
            }
            catch { return (0, 0); }
        }

        // ═══════════════════════════════════════════════════════════════
        // SAUVEGARDE SILENCIEUSE VIA API COM
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Sauvegarde via l'API COM selon le mode configure.
        /// SILENCIEUX: aucune popup, aucun raccourci clavier.
        /// </summary>
        public SaveResult Save(SaveMode mode)
        {
            if (!IsConnected || _inventorApp == null)
            {
                // Tenter reconnexion automatique
                if (!TryConnect())
                {
                    return new SaveResult
                    {
                        Success = false,
                        ErrorMessage = "Inventor non connecte",
                        Mode = mode
                    };
                }
            }

            return mode == SaveMode.SaveAll
                ? SaveAllDocuments()
                : SaveActiveDocument();
        }

        /// <summary>
        /// Sauvegarde le document actif ET tous ses composants modifies (enfants dirty).
        ///
        /// COMPORTEMENT "SaveActive intelligent":
        ///   - Si le doc actif est une PIECE (.ipt): sauvegarde uniquement cette piece
        ///   - Si le doc actif est un ASSEMBLAGE (.iam): sauvegarde recursivement
        ///     tous les sous-assemblages et pieces modifies (dirty), puis l'assemblage lui-meme
        ///   - Si le doc actif est un DESSIN (.idw/.dwg): sauvegarde le dessin uniquement
        ///
        /// AVANTAGE vs SaveAll:
        ///   - SaveAll sauvegarde TOUS les docs ouverts (meme ceux d'autres projets)
        ///   - SaveActive sauvegarde uniquement ce qui touche le doc actif -> PLUS RAPIDE
        ///
        /// SILENCIEUX: SilentOperation=true supprime TOUTES les popups Inventor,
        ///   y compris "Save modifications to referenced documents?" qui apparait
        ///   normalement quand on sauvegarde un assemblage avec des composants modifies.
        /// </summary>
        private SaveResult SaveActiveDocument()
        {
            try
            {
                dynamic? activeDoc = null;
                try { activeDoc = _inventorApp!.ActiveDocument; } catch { }

                if (activeDoc == null)
                {
                    return new SaveResult
                    {
                        Success = true,
                        DocumentsSaved = 0,
                        DocumentsSkipped = 0,
                        ErrorMessage = null,
                        Mode = SaveMode.SaveActive
                    };
                }

                string docName = "unknown";
                try { docName = (string)activeDoc.DisplayName; } catch { }

                // Verifier chemin valide sur disque (pas un nouveau doc sans Save As)
                bool hasPath = false;
                try
                {
                    string? fullPath = activeDoc.FullFileName as string;
                    hasPath = !string.IsNullOrWhiteSpace(fullPath) && fullPath != "." && fullPath.Length > 3;
                }
                catch { }

                if (!hasPath)
                {
                    Logger.Log($"[i] {docName}: nouveau document sans chemin, sauvegarde ignoree", Logger.LogLevel.DEBUG);
                    return new SaveResult
                    {
                        Success = true,
                        DocumentsSaved = 0,
                        DocumentsSkipped = 1,
                        Mode = SaveMode.SaveActive
                    };
                }

                // SILENCIEUX: SilentOperation = true AVANT toute operation de sauvegarde.
                // Cela supprime:
                //   - "Save modifications to referenced documents?"
                //   - "Multiple files have been updated..."
                //   - Toute autre popup de confirmation Inventor
                bool wasSilent = false;
                try { wasSilent = (bool)_inventorApp!.SilentOperation; } catch { }

                try
                {
                    try { _inventorApp!.SilentOperation = true; } catch { }

                    // Collecter les documents a sauvegarder (doc actif + enfants dirty)
                    // Cast explicite car activeDoc est dynamic (evite que le compilateur
                    // infere toSave comme dynamic aussi)
                    List<DocEntry> toSave = CollectDirtyDocuments(activeDoc);

                    if (toSave.Count == 0)
                    {
                        Logger.Log($"[i] {docName}: aucune modification (ni enfants), sauvegarde ignoree", Logger.LogLevel.DEBUG);
                        return new SaveResult
                        {
                            Success = true,
                            DocumentsSaved = 0,
                            DocumentsSkipped = 1,
                            Mode = SaveMode.SaveActive
                        };
                    }

                    int savedCount = 0;
                    int skippedCount = 0;

                    // Sauvegarder dans l'ordre: ipt(0) -> iam(1) -> idw/dwg(2)
                    foreach (var entry in toSave.OrderBy(x => x.Order))
                    {
                        try
                        {
                            ((dynamic)entry.Doc).Save();
                            savedCount++;
                            Logger.Log($"[+] SaveActive: {entry.Name} sauvegarde", Logger.LogLevel.DEBUG);
                        }
                        catch (Exception ex)
                        {
                            Logger.Log($"[!] Erreur sauvegarde {entry.Name}: {ex.Message}", Logger.LogLevel.WARNING);
                            skippedCount++;
                        }
                    }

                    string logMsg = savedCount == 1
                        ? $"[+] SaveActive: {docName} sauvegarde"
                        : $"[+] SaveActive: {savedCount} doc(s) sauvegarde(s) (doc actif + composants)";
                    Logger.Log(logMsg, Logger.LogLevel.INFO);

                    var result = new SaveResult
                    {
                        Success = true,
                        DocumentsSaved = savedCount,
                        DocumentsSkipped = skippedCount,
                        Mode = SaveMode.SaveActive
                    };
                    SaveCompleted?.Invoke(this, result);
                    return result;
                }
                finally
                {
                    try { _inventorApp!.SilentOperation = wasSilent; } catch { }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur SaveActive: {ex.Message}", Logger.LogLevel.ERROR);

                // Verifier si c'est une erreur de connexion COM
                if (ex is COMException || ex.Message.Contains("RPC"))
                {
                    SetDisconnected();
                }

                var result = new SaveResult
                {
                    Success = false,
                    ErrorMessage = ex.Message,
                    Mode = SaveMode.SaveActive
                };
                SaveCompleted?.Invoke(this, result);
                return result;
            }
        }

        /// <summary>
        /// Collecte recursivement tous les documents dirty lies au document donne.
        /// Pour un assemblage: inclut les pieces et sous-assemblages modifies.
        /// Pour une piece ou un dessin: retourne uniquement ce document s'il est dirty.
        /// Retourne une liste triee (ipt=0, iam=1, idw=2) prete pour la sauvegarde.
        /// </summary>
        private List<DocEntry> CollectDirtyDocuments(dynamic rootDoc)
        {
            var result = new List<DocEntry>();
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            CollectRecursive(rootDoc, result, visited);
            return result;
        }

        private void CollectRecursive(dynamic doc, List<DocEntry> result, HashSet<string> visited)
        {
            try
            {
                string? fullPath = null;
                try { fullPath = doc.FullFileName as string; } catch { }

                // Ignorer les documents sans chemin (nouveaux non enregistres)
                if (string.IsNullOrWhiteSpace(fullPath) || fullPath == "." || fullPath.Length <= 3)
                    return;

                // Anti-boucle infinie (cas d'assemblages circulaires ou references croisees)
                if (!visited.Add(fullPath)) return;

                // Verifier si modifiable
                bool isModifiable = true;
                try { isModifiable = (bool)doc.IsModifiable; } catch { }
                if (!isModifiable) return;

                string name = Path.GetFileName(fullPath);
                string ext = Path.GetExtension(fullPath).ToLowerInvariant();

                // Pour les assemblages: parcourir recursivement les composants
                // DocumentType: 12291=Assembly, 12289=Part, 12292=Drawing, 12290=Presentation
                bool isAssembly = false;
                try
                {
                    int docType = (int)doc.DocumentType;
                    isAssembly = (docType == 12291); // kAssemblyDocumentObject
                }
                catch { isAssembly = ext == ".iam"; }

                if (isAssembly)
                {
                    // Parcourir les ReferencedDocuments (references directes de cet assemblage)
                    try
                    {
                        dynamic referencedDocs = doc.ReferencedDocuments;
                        int refCount = (int)referencedDocs.Count;
                        for (int i = 1; i <= refCount; i++)
                        {
                            try
                            {
                                dynamic refDoc = referencedDocs.Item(i);
                                CollectRecursive(refDoc, result, visited);
                            }
                            catch { }
                        }
                    }
                    catch { }
                }

                // Ajouter ce document uniquement s'il est modifie (dirty)
                bool isDirty = true;
                try { isDirty = (bool)doc.Dirty; } catch { }

                if (isDirty)
                {
                    int order = ext switch
                    {
                        ".ipt" => 0,
                        ".iam" => 1,
                        ".idw" or ".dwg" => 2,
                        _ => 3
                    };
                    result.Add(new DocEntry { Doc = doc, Name = name, Order = order });
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[!] CollectRecursive: {ex.Message}", Logger.LogLevel.DEBUG);
            }
        }

        /// <summary>
        /// Sauvegarde TOUS les documents ouverts de facon silencieuse.
        ///
        /// ORDRE DE SAUVEGARDE (important pour Inventor):
        ///   1. Parts (.ipt)      => composants de base
        ///   2. Assemblies (.iam) => referencent les parts
        ///   3. Drawings (.idw/.dwg) => referencent les assemblages
        ///
        /// SILENCIEUX: SilentOperation=true supprime la popup 
        ///   "Multiple files have been updated. Do you want to save?" 
        ///   que tu vois dans la capture d'ecran.
        /// </summary>
        private SaveResult SaveAllDocuments()
        {
            int savedCount = 0;
            int skippedCount = 0;

            bool wasSilent = false;
            try { wasSilent = (bool)_inventorApp!.SilentOperation; } catch { }

            try
            {
                // CRITIQUE: SilentOperation AVANT d'iterer les documents
                try { _inventorApp!.SilentOperation = true; } catch { }

                dynamic docs = _inventorApp!.Documents;
                int docCount = (int)docs.Count;

                if (docCount == 0)
                {
                    return new SaveResult
                    {
                        Success = true,
                        DocumentsSaved = 0,
                        DocumentsSkipped = 0,
                        Mode = SaveMode.SaveAll
                    };
                }

                // Construire la liste des documents a sauvegarder
                var toSave = new List<DocEntry>();

                for (int i = 1; i <= docCount; i++)
                {
                    try
                    {
                        dynamic doc = docs.Item(i);

                        // Verifier modification
                        bool isDirty = false;
                        try { isDirty = (bool)doc.Dirty; } catch { isDirty = true; }
                        if (!isDirty) { skippedCount++; continue; }

                        // Verifier chemin valide
                        string? fullPath = null;
                        try { fullPath = doc.FullFileName as string; } catch { }
                        if (string.IsNullOrWhiteSpace(fullPath) || fullPath == "." || fullPath.Length <= 3)
                        {
                            skippedCount++;
                            continue;
                        }

                        // Verifier modifiable
                        bool isModifiable = true;
                        try { isModifiable = (bool)doc.IsModifiable; } catch { }
                        if (!isModifiable) { skippedCount++; continue; }

                        string name = "unknown";
                        try { name = (string)doc.DisplayName; } catch { }

                        // Ordre de tri: ipt=0, iam=1, idw/dwg=2, autres=3
                        string ext = Path.GetExtension(fullPath).ToLowerInvariant();
                        int order = ext switch
                        {
                            ".ipt" => 0,
                            ".iam" => 1,
                            ".idw" or ".dwg" => 2,
                            _ => 3
                        };

                        toSave.Add(new DocEntry { Doc = doc, Name = name, Order = order });
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"[!] Erreur lecture document {i}: {ex.Message}", Logger.LogLevel.DEBUG);
                    }
                }

                // Trier et sauvegarder
                foreach (var entry in toSave.OrderBy(x => x.Order))
                {
                    try
                    {
                        ((dynamic)entry.Doc).Save();
                        savedCount++;
                        Logger.Log($"[+] Sauvegarde: {entry.Name}", Logger.LogLevel.DEBUG);
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"[!] Erreur sauvegarde {entry.Name}: {ex.Message}", Logger.LogLevel.WARNING);
                        skippedCount++;
                    }
                }

                Logger.Log($"[+] SaveAll: {savedCount} doc(s) sauvegarde(s), {skippedCount} ignore(s)", Logger.LogLevel.INFO);

                var result = new SaveResult
                {
                    Success = true,
                    DocumentsSaved = savedCount,
                    DocumentsSkipped = skippedCount,
                    Mode = SaveMode.SaveAll
                };
                SaveCompleted?.Invoke(this, result);
                return result;
            }
            catch (Exception ex)
            {
                Logger.Log($"[-] Erreur SaveAll: {ex.Message}", Logger.LogLevel.ERROR);

                if (ex is COMException || ex.Message.Contains("RPC"))
                {
                    SetDisconnected();
                }

                var result = new SaveResult
                {
                    Success = false,
                    ErrorMessage = ex.Message,
                    Mode = SaveMode.SaveAll
                };
                SaveCompleted?.Invoke(this, result);
                return result;
            }
            finally
            {
                try { _inventorApp!.SilentOperation = wasSilent; } catch { }
            }
        }
    }
}
