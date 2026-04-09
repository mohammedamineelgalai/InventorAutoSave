using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using InventorAutoSave.Models;

namespace InventorAutoSave.Services
{
    /// <summary>
    /// Helper pour appels COM late-bound via IDispatch.
    /// En .NET 8, dynamic sur un objet COM retourne souvent "Specified cast is not valid"
    /// quand la propriete retourne un autre objet COM (ex: ActiveDocument).
    /// Type.InvokeMember utilise IDispatch::Invoke correctement.
    ///
    /// IMPORTANT: Tous les appels doivent etre effectues depuis un thread STA
    /// (thread UI WPF). Les appels depuis des threads MTA (ThreadPool, Task.Run)
    /// causent des MissingMethodException.
    /// </summary>
    internal static class ComInvoke
    {
        private const BindingFlags GETPROP = BindingFlags.GetProperty;
        private const BindingFlags SETPROP = BindingFlags.SetProperty;
        private const BindingFlags INVOKE = BindingFlags.InvokeMethod;

        /// <summary>
        /// Lire une propriete COM via IDispatch.
        /// Utilise GetProperty car c'est le seul flag qui fonctionne de maniere fiable
        /// pour les proprietes Inventor retournant des objets COM (ActiveDocument, Documents, etc.)
        /// </summary>
        public static object? GetProp(object comObj, string name)
        {
            return comObj.GetType().InvokeMember(name, GETPROP, null, comObj, null);
        }

        /// <summary>Ecrire une propriete COM</summary>
        public static void SetProp(object comObj, string name, object value)
        {
            comObj.GetType().InvokeMember(name, SETPROP, null, comObj, new[] { value });
        }

        /// <summary>Appeler une methode COM sans arguments</summary>
        public static object? Call(object comObj, string name)
        {
            return comObj.GetType().InvokeMember(name, INVOKE, null, comObj, null);
        }

        /// <summary>Appeler une methode COM avec arguments</summary>
        public static object? Call(object comObj, string name, params object[] args)
        {
            return comObj.GetType().InvokeMember(name, INVOKE, null, comObj, args);
        }

        /// <summary>Lire une propriete string COM</summary>
        public static string? GetString(object comObj, string name)
        {
            var result = GetProp(comObj, name);
            return result?.ToString();
        }

        /// <summary>Lire une propriete bool COM</summary>
        public static bool GetBool(object comObj, string name, bool defaultValue = false)
        {
            try
            {
                var result = GetProp(comObj, name);
                return Convert.ToBoolean(result);
            }
            catch { return defaultValue; }
        }

        /// <summary>Lire une propriete int COM</summary>
        public static int GetInt(object comObj, string name, int defaultValue = 0)
        {
            try
            {
                var result = GetProp(comObj, name);
                return Convert.ToInt32(result);
            }
            catch { return defaultValue; }
        }

        /// <summary>Acceder a un element indexe (collection COM, 1-based)</summary>
        public static object? GetItem(object collection, int index)
        {
            var type = collection.GetType();
            // Item sur les collections Inventor peut etre une propriete ou methode
            try { return type.InvokeMember("Item", GETPROP, null, collection, new object[] { index }); }
            catch (MissingMethodException) { }
            return type.InvokeMember("Item", INVOKE, null, collection, new object[] { index });
        }
    }

    internal sealed class DocEntry
    {
        public required object Doc { get; init; }
        public required string Name { get; init; }
        public required int Order { get; init; }
    }

    /// <summary>
    /// Service de connexion COM et sauvegarde silencieuse via l'API Inventor.
    /// Utilise ComInvoke (Type.InvokeMember) au lieu de dynamic pour les appels COM
    /// car en .NET 8 dynamic echoue avec "Specified cast is not valid" sur les
    /// proprietes qui retournent des objets COM (ActiveDocument, Documents, etc.).
    /// </summary>
    public class InventorSaveService
    {
        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID(
            [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
            out Guid pclsid);

        [DllImport("oleaut32.dll", EntryPoint = "GetActiveObject", PreserveSig = false)]
        private static extern void GetActiveObjectFromOleAut32(
            ref Guid rclsid,
            IntPtr pvReserved,
            [MarshalAs(UnmanagedType.IDispatch)] out object ppunk);

        [DllImport("oleaut32.dll", EntryPoint = "GetActiveObject")]
        private static extern int GetActiveObjectHResult(
            ref Guid rclsid,
            IntPtr pvReserved,
            [MarshalAs(UnmanagedType.IDispatch)] out object ppunk);

        private object? _inventorApp;
        private bool _isConnected;
        private DateTime _lastConnectionAttempt = DateTime.MinValue;
        private int _consecutiveFailures;
        private const int MIN_RETRY_INTERVAL_MS = 3000;

        public event EventHandler<SaveResult>? SaveCompleted;
        public event EventHandler? Connected;
        public event EventHandler? Disconnected;

        public bool IsConnected => _isConnected && _inventorApp != null;

        public bool TryConnect()
        {
            var now = DateTime.Now;
            if ((now - _lastConnectionAttempt).TotalMilliseconds < MIN_RETRY_INTERVAL_MS
                && _lastConnectionAttempt != DateTime.MinValue)
                return _isConnected;
            _lastConnectionAttempt = now;

            if (!Process.GetProcessesByName("Inventor").Any())
            {
                if (_isConnected) SetDisconnected();
                return false;
            }

            if (TryConnectMethod1() || TryConnectMethod2() || TryConnectMethod3())
            {
                _consecutiveFailures = 0;
                return true;
            }

            _consecutiveFailures++;
            if (_isConnected) SetDisconnected();

            if (_consecutiveFailures % 10 == 0)
                Logger.Log(string.Format("[~] Connexion Inventor: {0} tentatives...", _consecutiveFailures), Logger.LogLevel.DEBUG);

            return false;
        }

        private bool TryConnectMethod1()
        {
            try
            {
                int hr = CLSIDFromProgID("Inventor.Application", out Guid clsid);
                if (hr != 0) return false;

                GetActiveObjectFromOleAut32(ref clsid, IntPtr.Zero, out object inventorObj);
                if (inventorObj != null)
                {
                    bool wasConnected = _isConnected;
                    _inventorApp = inventorObj;
                    _isConnected = true;
                    if (!wasConnected)
                    {
                        var tid = Thread.CurrentThread.ManagedThreadId;
                        var apt = Thread.CurrentThread.GetApartmentState();
                        Logger.Log(string.Format("[+] Connecte a Inventor via oleaut32 IDispatch (Thread={0}, Apt={1})", tid, apt), Logger.LogLevel.INFO);
                        try
                        {
                            string? caption = ComInvoke.GetString(_inventorApp, "Caption");
                            Logger.Log(string.Format("[+] ComInvoke OK: Caption = \"{0}\"", caption), Logger.LogLevel.INFO);
                        }
                        catch (Exception ex)
                        {
                            Logger.Log(string.Format("[!] ComInvoke Caption echoue: {0}: {1}", ex.GetType().Name, ex.Message), Logger.LogLevel.WARNING);
                        }

                        Connected?.Invoke(this, EventArgs.Empty);
                    }
                    return true;
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                if (_consecutiveFailures < 3)
                    Logger.Log(string.Format("[~] Methode 1 erreur: {0}", ex.Message), Logger.LogLevel.DEBUG);
            }
            return false;
        }

        private bool TryConnectMethod2()
        {
            try
            {
                int hr = CLSIDFromProgID("Inventor.Application", out Guid clsid);
                if (hr != 0) return false;

                int hrActive = GetActiveObjectHResult(ref clsid, IntPtr.Zero, out object inventorObj);
                if (hrActive == 0 && inventorObj != null)
                {
                    bool wasConnected = _isConnected;
                    _inventorApp = inventorObj;
                    _isConnected = true;
                    if (!wasConnected)
                    {
                        Logger.Log("[+] Connecte a Inventor via oleaut32 IDispatch (HRESULT)", Logger.LogLevel.INFO);
                        Connected?.Invoke(this, EventArgs.Empty);
                    }
                    return true;
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                if (_consecutiveFailures < 3)
                    Logger.Log(string.Format("[~] Methode 2 erreur: {0}", ex.Message), Logger.LogLevel.DEBUG);
            }
            return false;
        }

        private bool TryConnectMethod3()
        {
            try
            {
                Type? inventorType = Type.GetTypeFromProgID("Inventor.Application", false);
                if (inventorType == null) return false;

                object? inventorObj = null;
                try
                {
                    inventorObj = Marshal.BindToMoniker("!{" + inventorType.GUID.ToString() + "}");
                }
                catch (COMException)
                {
                    try
                    {
                        inventorObj = Activator.CreateInstance(inventorType);
                        if (inventorObj != null)
                        {
                            try
                            {
                                bool visible = ComInvoke.GetBool(inventorObj, "Visible");
                                if (!visible)
                                {
                                    try { ComInvoke.Call(inventorObj, "Quit"); } catch { }
                                    Marshal.ReleaseComObject(inventorObj);
                                    return false;
                                }
                            }
                            catch { }
                        }
                    }
                    catch (COMException) { return false; }
                }

                if (inventorObj != null)
                {
                    _inventorApp = inventorObj;
                    _isConnected = true;
                    Logger.Log("[+] Connecte a Inventor via Type/BindToMoniker", Logger.LogLevel.INFO);
                    Connected?.Invoke(this, EventArgs.Empty);
                    return true;
                }
            }
            catch (COMException) { }
            catch (Exception ex)
            {
                if (_consecutiveFailures < 3)
                    Logger.Log(string.Format("[~] Methode 3 erreur: {0}", ex.Message), Logger.LogLevel.DEBUG);
            }
            return false;
        }

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

        public string? GetActiveDocumentName()
        {
            try
            {
                if (_inventorApp != null)
                {
                    object? doc = ComInvoke.GetProp(_inventorApp, "ActiveDocument");
                    if (doc != null)
                    {
                        try
                        {
                            return ComInvoke.GetString(doc, "DisplayName");
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(doc);
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        public bool IsInventorCalculating()
        {
            if (!IsConnected) return false;
            try
            {
                if (_inventorApp != null)
                {
                    string? title = null;
                    try { title = ComInvoke.GetString(_inventorApp, "Caption"); } catch { }
                    if (title != null)
                    {
                        string[] busyKw = { "Calculating", "Calcul", "Processing", "Rebuilding",
                                             "Generating", "Computing", "Working", "Loading" };
                        foreach (var kw in busyKw)
                            if (title.Contains(kw, StringComparison.OrdinalIgnoreCase)) return true;
                    }
                }
            }
            catch { }
            return false;
        }

        public (int total, int dirty) GetDocumentCounts()
        {
            if (!IsConnected || _inventorApp == null) return (0, 0);
            try
            {
                object? docs = ComInvoke.GetProp(_inventorApp, "Documents");
                if (docs == null) return (0, 0);

                try
                {
                    int total = ComInvoke.GetInt(docs, "Count");
                    int dirty = 0;
                    for (int i = 1; i <= total; i++)
                    {
                        object? doc = null;
                        try
                        {
                            doc = ComInvoke.GetItem(docs, i);
                            if (doc != null && ComInvoke.GetBool(doc, "Dirty")) dirty++;
                        }
                        catch { }
                        finally
                        {
                            if (doc != null) try { Marshal.ReleaseComObject(doc); } catch { }
                        }
                    }
                    return (total, dirty);
                }
                finally
                {
                    Marshal.ReleaseComObject(docs);
                }
            }
            catch { return (0, 0); }
        }

        public SaveResult Save(SaveMode mode)
        {
            if (!IsConnected || _inventorApp == null)
            {
                if (!TryConnect())
                    return new SaveResult { Success = false, ErrorMessage = "Inventor non connecte", Mode = mode };
            }
            return mode == SaveMode.SaveAll ? SaveAllDocuments() : SaveActiveDocument();
        }

        private SaveResult SaveActiveDocument()
        {
            try
            {
                var tid = Thread.CurrentThread.ManagedThreadId;
                var apt = Thread.CurrentThread.GetApartmentState();
                Logger.Log(string.Format("[>] SaveActiveDocument: debut (Thread={0}, Apt={1})", tid, apt), Logger.LogLevel.DEBUG);

                // Diagnostic: tester si l'objet COM est encore valide
                try
                {
                    string? caption = ComInvoke.GetString(_inventorApp!, "Caption");
                    Logger.Log(string.Format("[>] SaveActiveDocument: Caption = \"{0}\"", caption), Logger.LogLevel.DEBUG);
                }
                catch (Exception ex)
                {
                    Logger.Log(string.Format("[!] SaveActiveDocument: Caption ECHOUE = {0}", ex.Message), Logger.LogLevel.WARNING);
                    // L'objet COM est invalide, forcer une reconnexion
                    SetDisconnected();
                    if (!TryConnect())
                        return new SaveResult { Success = false, ErrorMessage = "Inventor deconnecte (COM invalide)", Mode = SaveMode.SaveActive };
                    Logger.Log("[+] SaveActiveDocument: reconnexion reussie", Logger.LogLevel.INFO);
                }

                object? activeDoc = null;
                try { activeDoc = ComInvoke.GetProp(_inventorApp!, "ActiveDocument"); }
                catch (Exception ex) { Logger.Log(string.Format("[!] ActiveDocument exception: {0}", ex.Message), Logger.LogLevel.WARNING); }

                if (activeDoc == null)
                {
                    Logger.Log("[i] SaveActiveDocument: ActiveDocument est null", Logger.LogLevel.DEBUG);
                    return new SaveResult { Success = true, DocumentsSaved = 0, DocumentsSkipped = 0, Mode = SaveMode.SaveActive };
                }

                string docName = "unknown";
                try { docName = ComInvoke.GetString(activeDoc, "DisplayName") ?? "unknown"; } catch { }
                Logger.Log(string.Format("[>] SaveActiveDocument: doc actif = '{0}'", docName), Logger.LogLevel.DEBUG);

                string fullPath = "";
                try { fullPath = ComInvoke.GetString(activeDoc, "FullFileName") ?? ""; }
                catch (Exception ex) { Logger.Log(string.Format("[!] FullFileName exception: {0}", ex.Message), Logger.LogLevel.WARNING); }
                Logger.Log(string.Format("[>] SaveActiveDocument: FullFileName = '{0}'", fullPath), Logger.LogLevel.DEBUG);

                bool hasPath = !string.IsNullOrWhiteSpace(fullPath) && fullPath != "." && fullPath.Length > 3;
                if (!hasPath)
                {
                    Logger.Log(string.Format("[i] '{0}' n'a pas de chemin valide (jamais sauvegarde?)", docName), Logger.LogLevel.INFO);
                    return new SaveResult { Success = true, DocumentsSaved = 0, DocumentsSkipped = 1, Mode = SaveMode.SaveActive };
                }

                bool activeDocDirty = false;
                try { activeDocDirty = ComInvoke.GetBool(activeDoc, "Dirty"); }
                catch (Exception ex) { Logger.Log(string.Format("[!] Dirty exception: {0}", ex.Message), Logger.LogLevel.WARNING); activeDocDirty = true; }
                Logger.Log(string.Format("[>] SaveActiveDocument: Dirty = {0}", activeDocDirty), Logger.LogLevel.DEBUG);

                int docType = 0;
                try { docType = ComInvoke.GetInt(activeDoc, "DocumentType"); }
                catch (Exception ex) { Logger.Log(string.Format("[!] DocumentType exception: {0}", ex.Message), Logger.LogLevel.WARNING); }
                Logger.Log(string.Format("[>] SaveActiveDocument: DocumentType = {0}", docType), Logger.LogLevel.DEBUG);

                bool wasSilent = false;
                try { wasSilent = ComInvoke.GetBool(_inventorApp!, "SilentOperation"); } catch { }

                try
                {
                    try { ComInvoke.SetProp(_inventorApp!, "SilentOperation", true); }
                    catch (Exception ex) { Logger.Log(string.Format("[!] Set SilentOperation: {0}", ex.Message), Logger.LogLevel.WARNING); }

                    bool silentNow = false;
                    try { silentNow = ComInvoke.GetBool(_inventorApp!, "SilentOperation"); } catch { }
                    Logger.Log(string.Format("[>] SilentOperation apres set = {0}", silentNow), Logger.LogLevel.DEBUG);

                    List<DocEntry> toSave = CollectDirtyDocuments(activeDoc);
                    Logger.Log(string.Format("[>] CollectDirtyDocuments retourne {0} doc(s)", toSave.Count), Logger.LogLevel.DEBUG);

                    if (toSave.Count == 0 && activeDocDirty)
                    {
                        Logger.Log(string.Format("[!] CollectDirty=0 mais Dirty=true, sauvegarde directe de '{0}'", docName), Logger.LogLevel.WARNING);
                        try
                        {
                            ComInvoke.Call(activeDoc, "Save");
                            Logger.Log(string.Format("[+] Sauvegarde directe reussie: {0}", docName), Logger.LogLevel.INFO);
                            var directResult = new SaveResult { Success = true, DocumentsSaved = 1, DocumentsSkipped = 0, Mode = SaveMode.SaveActive };
                            SaveCompleted?.Invoke(this, directResult);
                            return directResult;
                        }
                        catch (Exception ex)
                        {
                            Logger.Log(string.Format("[!] Sauvegarde directe echouee: {0}", ex.Message), Logger.LogLevel.ERROR);
                        }
                    }

                    if (toSave.Count == 0)
                    {
                        Logger.Log(string.Format("[i] {0}: aucune modification", docName), Logger.LogLevel.DEBUG);
                        return new SaveResult { Success = true, DocumentsSaved = 0, DocumentsSkipped = 1, Mode = SaveMode.SaveActive };
                    }

                    int savedCount = 0, skippedCount = 0;
                    foreach (var entry in toSave.OrderBy(x => x.Order))
                    {
                        try
                        {
                            Logger.Log(string.Format("[>] Sauvegarde: {0} (order={1})", entry.Name, entry.Order), Logger.LogLevel.DEBUG);
                            ComInvoke.Call(entry.Doc, "Save");
                            savedCount++;
                            Logger.Log(string.Format("[+] OK: {0}", entry.Name), Logger.LogLevel.DEBUG);
                        }
                        catch (Exception ex)
                        {
                            Logger.Log(string.Format("[!] Erreur Save {0}: {1}", entry.Name, ex.Message), Logger.LogLevel.WARNING);
                            skippedCount++;
                        }
                    }

                    Logger.Log(string.Format("[+] SaveActive: {0} doc(s) sauvegardes", savedCount), Logger.LogLevel.INFO);
                    var result = new SaveResult { Success = true, DocumentsSaved = savedCount, DocumentsSkipped = skippedCount, Mode = SaveMode.SaveActive };
                    SaveCompleted?.Invoke(this, result);
                    return result;
                }
                finally
                {
                    try { ComInvoke.SetProp(_inventorApp!, "SilentOperation", wasSilent); } catch { }
                }
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("[-] Erreur SaveActive: {0}\n{1}", ex.Message, ex.StackTrace), Logger.LogLevel.ERROR);
                if (ex is COMException || ex.Message.Contains("RPC")) SetDisconnected();
                var result = new SaveResult { Success = false, ErrorMessage = ex.Message, Mode = SaveMode.SaveActive };
                SaveCompleted?.Invoke(this, result);
                return result;
            }
        }

        private List<DocEntry> CollectDirtyDocuments(object rootDoc)
        {
            var result = new List<DocEntry>();
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            CollectRecursive(rootDoc, result, visited);
            return result;
        }

        private void CollectRecursive(object doc, List<DocEntry> result, HashSet<string> visited)
        {
            try
            {
                string? fullPath = null;
                try { fullPath = ComInvoke.GetString(doc, "FullFileName"); }
                catch (Exception ex) { Logger.Log(string.Format("[!] CollectRecursive FullFileName: {0}", ex.Message), Logger.LogLevel.DEBUG); }

                if (string.IsNullOrWhiteSpace(fullPath) || fullPath == "." || fullPath.Length <= 3) return;
                if (!visited.Add(fullPath)) return;

                bool isModifiable = ComInvoke.GetBool(doc, "IsModifiable", true);
                if (!isModifiable)
                {
                    Logger.Log(string.Format("[i] '{0}' non modifiable", Path.GetFileName(fullPath)), Logger.LogLevel.DEBUG);
                    return;
                }

                string name = Path.GetFileName(fullPath);
                string ext = Path.GetExtension(fullPath).ToLowerInvariant();

                bool isAssembly = false;
                try { isAssembly = ComInvoke.GetInt(doc, "DocumentType") == 12291; }
                catch { isAssembly = ext == ".iam"; }

                if (isAssembly)
                {
                    try
                    {
                        object? refDocs = ComInvoke.GetProp(doc, "ReferencedDocuments");
                        if (refDocs != null)
                        {
                            int refCount = ComInvoke.GetInt(refDocs, "Count");
                            Logger.Log(string.Format("[>] '{0}' assemblage, {1} ref(s)", name, refCount), Logger.LogLevel.DEBUG);
                            for (int i = 1; i <= refCount; i++)
                            {
                                try
                                {
                                    object? refDoc = ComInvoke.GetItem(refDocs, i);
                                    if (refDoc != null) CollectRecursive(refDoc, result, visited);
                                }
                                catch { }
                            }
                        }
                    }
                    catch { }
                }

                bool isDirty = false;
                try { isDirty = ComInvoke.GetBool(doc, "Dirty"); }
                catch { isDirty = true; }

                Logger.Log(string.Format("[>] CollectRecursive: '{0}' Dirty={1}", name, isDirty), Logger.LogLevel.DEBUG);

                if (isDirty)
                {
                    int order = ext switch { ".ipt" => 0, ".iam" => 1, ".idw" or ".dwg" => 2, _ => 3 };
                    result.Add(new DocEntry { Doc = doc, Name = name, Order = order });
                }
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("[!] CollectRecursive erreur: {0}", ex.Message), Logger.LogLevel.DEBUG);
            }
        }

        private SaveResult SaveAllDocuments()
        {
            int savedCount = 0, skippedCount = 0;
            bool wasSilent = false;
            try { wasSilent = ComInvoke.GetBool(_inventorApp!, "SilentOperation"); } catch { }

            try
            {
                try { ComInvoke.SetProp(_inventorApp!, "SilentOperation", true); } catch { }

                object? docs = ComInvoke.GetProp(_inventorApp!, "Documents");
                if (docs == null)
                    return new SaveResult { Success = true, DocumentsSaved = 0, DocumentsSkipped = 0, Mode = SaveMode.SaveAll };

                int docCount = ComInvoke.GetInt(docs, "Count");
                Logger.Log(string.Format("[>] SaveAll: {0} document(s) ouverts", docCount), Logger.LogLevel.DEBUG);

                if (docCount == 0)
                    return new SaveResult { Success = true, DocumentsSaved = 0, DocumentsSkipped = 0, Mode = SaveMode.SaveAll };

                var toSave = new List<DocEntry>();
                for (int i = 1; i <= docCount; i++)
                {
                    try
                    {
                        object? doc = ComInvoke.GetItem(docs, i);
                        if (doc == null) continue;

                        string docName = ComInvoke.GetString(doc, "DisplayName") ?? "unknown";
                        bool isDirty = ComInvoke.GetBool(doc, "Dirty");
                        Logger.Log(string.Format("[>] SaveAll doc[{0}]: '{1}' Dirty={2}", i, docName, isDirty), Logger.LogLevel.DEBUG);

                        if (!isDirty) { skippedCount++; continue; }

                        string? fp = ComInvoke.GetString(doc, "FullFileName");
                        if (string.IsNullOrWhiteSpace(fp) || fp == "." || fp.Length <= 3) { skippedCount++; continue; }

                        bool isMod = ComInvoke.GetBool(doc, "IsModifiable", true);
                        if (!isMod) { skippedCount++; continue; }

                        string ext = Path.GetExtension(fp).ToLowerInvariant();
                        int order = ext switch { ".ipt" => 0, ".iam" => 1, ".idw" or ".dwg" => 2, _ => 3 };
                        toSave.Add(new DocEntry { Doc = doc, Name = docName, Order = order });
                    }
                    catch (Exception ex)
                    {
                        Logger.Log(string.Format("[!] Erreur doc {0}: {1}", i, ex.Message), Logger.LogLevel.DEBUG);
                    }
                }

                Logger.Log(string.Format("[>] SaveAll: {0} doc(s) a sauvegarder", toSave.Count), Logger.LogLevel.DEBUG);

                foreach (var entry in toSave.OrderBy(x => x.Order))
                {
                    try
                    {
                        ComInvoke.Call(entry.Doc, "Save");
                        savedCount++;
                        Logger.Log(string.Format("[+] Sauvegarde OK: {0}", entry.Name), Logger.LogLevel.DEBUG);
                    }
                    catch (Exception ex)
                    {
                        Logger.Log(string.Format("[!] Erreur Save {0}: {1}", entry.Name, ex.Message), Logger.LogLevel.WARNING);
                        skippedCount++;
                    }
                }

                Logger.Log(string.Format("[+] SaveAll: {0} doc(s), {1} ignore(s)", savedCount, skippedCount), Logger.LogLevel.INFO);
                var result = new SaveResult { Success = true, DocumentsSaved = savedCount, DocumentsSkipped = skippedCount, Mode = SaveMode.SaveAll };
                SaveCompleted?.Invoke(this, result);
                return result;
            }
            catch (Exception ex)
            {
                Logger.Log(string.Format("[-] Erreur SaveAll: {0}\n{1}", ex.Message, ex.StackTrace), Logger.LogLevel.ERROR);
                if (ex is COMException || ex.Message.Contains("RPC")) SetDisconnected();
                var result = new SaveResult { Success = false, ErrorMessage = ex.Message, Mode = SaveMode.SaveAll };
                SaveCompleted?.Invoke(this, result);
                return result;
            }
            finally
            {
                try { ComInvoke.SetProp(_inventorApp!, "SilentOperation", wasSilent); } catch { }
            }
        }
    }
}
