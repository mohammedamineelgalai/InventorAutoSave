# 🔄 HANDOFF — Fix Segment Corruption pour `InventorAutoSave`

> **Date**: 25 avril 2026  
> **Source**: Projet XEAT, Phase 18 — `Modules/OpenVaultProject/Services/VaultDownloadService.cs` → `SaveAllAndCloseAllDocuments`  
> **Cible**: `C:\Users\mohammedamine.elgala\source\repos\InventorAutoSave`  
> **Framework cible**: **.NET 8** (PAS .NET Framework 4.8 — l'autre projet utilise late-bound COM via `Type.InvokeMember`)

---

## 🎯 OBJECTIF

Porter la stratégie de sauvegarde 4-phases (validée par doc Autodesk officielle) dans `InventorAutoSave` pour éliminer les **erreurs de segment BREP** sur les sub-assemblies.

---

## 🧠 PRINCIPE DIRECTEUR — INTELLIGENCE CONTEXTUELLE (RÈGLE ABSOLUE)

> **Peu importe ce qui est ouvert dans Inventor, la stratégie doit s'adapter intelligemment.**

Le **document actif** (`ActiveDocument`) peut être de n'importe quel type. Pour chaque cas, on détermine **un "racine effectif"** et on remonte sa hiérarchie pour collecter UNIQUEMENT les Dirty :

| `ActiveDocument` | Racine effectif | Enfants/petits-enfants à scanner | Stratégie save |
|---|---|---|---|
| **Top Assembly (.iam)** | Lui-même = top | Toute sa hiérarchie `ReferencedDocuments` (récursif) | Phase 1 IPT Dirty → Phase 2 Sub-IAM Dirty → **Update + Save Top** |
| **Sub-Assembly (.iam)** ouvert seul ou comme racine de fenêtre | Lui-même = "top de session" | Ses propres `ReferencedDocuments` (récursif) | Phase 1 IPT Dirty (ses parts) → Phase 2 Sub-IAM Dirty (ses petits-enfants) → **Update + Save lui-même** |
| **Part (.ipt)** | Lui-même | Aucun enfant (pas de référence) | Phase 1 unique : Save direct si Dirty |
| **Drawing (.idw/.dwg)** | Lui-même | `ReferencedDocuments` = modèles 3D source | Skip drawing (drafter rebuild) MAIS si modèles 3D Dirty → Save IPT → Save Sub-IAM → Update + Save Top-IAM des modèles. Le drawing lui-même n'est PAS sauvé. |
| **Presentation (.ipn)** | Lui-même | `ReferencedDocuments` = assemblage source | Idem drawing : skip .ipn, traiter les modèles Dirty en 4-phases |

### 🔑 Règles de détection

1. **`ActiveDocument` = ancrage de la session** → c'est lui qui définit "le top à updater" si c'est un .iam.
2. **Récursivité via `ReferencedDocuments`** → descendre TOUS les niveaux (enfants, petits-enfants, arrière-petits-enfants).
3. **Filtre `Dirty=true`** → ne JAMAIS sauver un doc non modifié (évite réécritures inutiles + corruption).
4. **Filtre `IsModifiable=true`** → respecter les fichiers en lecture seule (Vault checked-in, etc.).
5. **Déduplication via `FullFileName`** (HashSet) → un même part référencé 10× = traité 1 seule fois.
6. **Top = racine de la session** (pas seulement `ActiveDocument` global) :
   - Si `SaveAll` global → top = `ActiveDocument` (fenêtre au premier plan).
   - Si `SaveActive` ciblé → top = le doc passé en paramètre (peut être un sub-asm ouvert dans sa propre fenêtre).
7. **Drawings/.ipn ne sont JAMAIS le "top" pour Update** → ils n'ont pas de segments BREP à régénérer. Si actif, on traite uniquement leurs modèles 3D Dirty référencés.

### 🎯 Implémentation : factoriser dans `CollectAndClassify(rootDoc)`

```csharp
// Pseudo-code de la méthode unique réutilisable
private (List<DocEntry> toSave, int skipped) CollectAndClassify(object rootDoc)
{
    string? rootPath = ComInvoke.GetString(rootDoc, "FullFileName");
    string rootExt = Path.GetExtension(rootPath ?? "").ToLowerInvariant();
    
    // Si racine = drawing/ipn, le "top effectif" pour Update sera le 1er .iam référencé
    // Sinon, le "top effectif" = la racine elle-même
    string? effectiveTopPath = (rootExt == ".iam" || rootExt == ".ipt") ? rootPath : null;
    
    var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    var collected = new List<DocEntry>();
    int skipped = 0;
    
    // Récursion: descend dans ReferencedDocuments si .iam, skip drawings du collect
    CollectRecursive(rootDoc, collected, visited, effectiveTopPath, ref skipped);
    
    // Si racine était un drawing/ipn et qu'on n'a pas de top détecté,
    // promouvoir le 1er .iam Dirty trouvé comme top effectif
    if (effectiveTopPath == null)
    {
        var firstIam = collected.FirstOrDefault(e => e.Ext == ".iam");
        if (firstIam != null)
        {
            // Marquer comme top + recalculer son order
            // (création d'un nouveau DocEntry avec IsTopAssembly=true)
        }
    }
    
    return (collected, skipped);
}
```

**Cette méthode est appelée :**
- Depuis `SaveActiveDocument()` avec `rootDoc = ActiveDocument`
- Depuis `SaveAllDocuments()` avec `rootDoc = ActiveDocument` aussi (la fenêtre au premier plan définit le contexte)
- OU optionnel : itérer sur toutes les fenêtres ouvertes (`Application.Documents`) et appeler `CollectAndClassify` pour chaque document de niveau racine, puis fusionner — mais attention à la déduplication globale via `visited` partagé.

---

---

## ✅ VALIDATION OFFICIELLE (déjà effectuée)

3 sources autoritatives confirment la stratégie :

1. **Autodesk DevBlog** — [Tips for creating large assemblies using Inventor API](https://blog.autodesk.io/tips-for-creating-large-assemblies-using-inventor-api/) — Tip #8 : sauvegarde **progressive**, pas en un seul bloc final.
2. **Jean-Sébastien Hould (MVP)** — [Update and Save All](https://jshould.ca/blog/2019/01/02/inventor-macro-update-and-save-all/) — pattern `Update2 → Save2`.
3. **API officielle** `Document.Save2(SaveDependents=True)` → sauve bottom-up automatiquement (parts → sub-asm → top).

**Verdict**: la stratégie manuelle est l'équivalent **instrumenté** de `Save2(True)`, avec en plus telemetry, try/catch par doc, skip drawings.

---

## 🔬 DIAGNOSTIC du fichier `Services/InventorSaveService.cs` (655 lignes, .NET 8)

Lu intégralement. **4 lacunes critiques** qui causent les segment errors :

| # | Lacune | Localisation | Risque |
|---|---|---|---|
| 1 | **Tous les `.iam` ont `Order=1`** → sub-asm et top mélangés, ordre indéterministe | `CollectRecursive` (~ligne 540) ET `SaveAllDocuments` (~ligne 600) : `int order = ext switch { ".ipt" => 0, ".iam" => 1, ".idw" or ".dwg" => 2, _ => 3 };` | ❌ **Cause racine** : top peut être sauvé AVANT ses subs |
| 2 | **Pas d'`Update()` sur le top avant `Save()`** | Boucle finale `foreach (var entry in toSave.OrderBy(x => x.Order))` | ❌ **Cause racine** : segments BREP non régénérés |
| 3 | Drawings `.idw/.dwg` sauvegardés (Order=2) | Même endroit | ⚠️ Lent + risqué (drafter rebuild possible) |
| 4 | Pas de telemetry par phase | Logs en vrac | ⚠️ Diagnostic difficile |

---

## 🛠️ FIX À APPLIQUER (modifications additives — ne casse pas l'API)

### Étape 1 — Étendre `DocEntry`

```csharp
internal sealed class DocEntry
{
    public required object Doc { get; init; }
    public required string Name { get; init; }
    public required int Order { get; init; }
    public string Ext { get; init; } = "";          // NOUVEAU
    public bool IsTopAssembly { get; init; }        // NOUVEAU
}
```

### Étape 2 — Identifier le top via `ActiveDocument.FullFileName`

Dans `SaveAllDocuments()` (et `CollectDirtyDocuments()` pour `SaveActive`) :

```csharp
// Avant la boucle de classification, identifier le top
string? topFullPath = null;
try
{
    object? topDoc = ComInvoke.GetProp(_inventorApp!, "ActiveDocument");
    if (topDoc != null)
    {
        try { topFullPath = ComInvoke.GetString(topDoc, "FullFileName"); }
        finally { Marshal.ReleaseComObject(topDoc); }
    }
}
catch { }
```

### Étape 3 — Nouvel ordering 4-phases

```csharp
// Order: 0=IPT, 1=Sub-IAM, 2=Top-IAM, 3=skip drawings (NE PAS ajouter)
int order;
bool isTop = false;
switch (ext)
{
    case ".ipt":
        order = 0;
        break;
    case ".iam":
        isTop = !string.IsNullOrEmpty(topFullPath)
                && string.Equals(fp, topFullPath, StringComparison.OrdinalIgnoreCase);
        order = isTop ? 2 : 1;
        break;
    case ".idw":
    case ".dwg":
    case ".ipn":
        skippedCount++;
        Logger.Log(string.Format("[i] Skip drawing/presentation: {0}", docName), Logger.LogLevel.DEBUG);
        continue;  // Skip TOTAL
    default:
        order = 3;
        break;
}
toSave.Add(new DocEntry { Doc = doc, Name = docName, Order = order, Ext = ext, IsTopAssembly = isTop });
```

### Étape 4 — Update + Save sur le Top (boucle de save)

Remplacer la boucle simple `foreach (var entry in toSave.OrderBy(x => x.Order))` par :

```csharp
var sw = System.Diagnostics.Stopwatch.StartNew();
int iptCount = 0, subIamCount = 0, topUpdated = 0, topSaved = 0;

// Phase 1+2: IPT puis Sub-IAM (Order 0 et 1)
foreach (var entry in toSave.Where(e => e.Order < 2).OrderBy(e => e.Order))
{
    try
    {
        ComInvoke.Call(entry.Doc, "Save");
        savedCount++;
        if (entry.Ext == ".ipt") iptCount++;
        else if (entry.Ext == ".iam") subIamCount++;
        Logger.Log(string.Format("[+] Sauvegarde OK: {0}", entry.Name), Logger.LogLevel.DEBUG);
    }
    catch (Exception ex)
    {
        Logger.Log(string.Format("[!] Erreur Save {0}: {1}", entry.Name, ex.Message), Logger.LogLevel.WARNING);
        skippedCount++;
    }
}

// Phase 3: Update + Save Top Assembly
var topEntry = toSave.FirstOrDefault(e => e.IsTopAssembly);
if (topEntry != null)
{
    try
    {
        Logger.Log(string.Format("[>] Update top assembly: {0}", topEntry.Name), Logger.LogLevel.DEBUG);
        try { ComInvoke.Call(topEntry.Doc, "Update"); topUpdated = 1; }
        catch (Exception exU) { Logger.Log(string.Format("[!] Update top echoue: {0}", exU.Message), Logger.LogLevel.WARNING); }

        ComInvoke.Call(topEntry.Doc, "Save");
        topSaved = 1;
        savedCount++;
        Logger.Log(string.Format("[+] Top sauvegarde: {0}", topEntry.Name), Logger.LogLevel.INFO);
    }
    catch (Exception ex)
    {
        Logger.Log(string.Format("[!] Erreur Save top {0}: {1}", topEntry.Name, ex.Message), Logger.LogLevel.WARNING);
        skippedCount++;
    }
}

sw.Stop();
Logger.Log(string.Format(
    "[+] SaveAll ordonne: {0} doc(s) | IPT={1}, Sub-IAM={2}, Top-IAM update={3}/save={4} | skip={5} | {6} ms",
    savedCount, iptCount, subIamCount, topUpdated, topSaved, skippedCount, sw.ElapsedMilliseconds),
    Logger.LogLevel.INFO);
```

### Étape 5 — Appliquer la même logique à `CollectRecursive` (pour `SaveActive`)

Même remplacement du `int order = ext switch { ... }` + ajout `IsTopAssembly` (le root passé à `CollectDirtyDocuments` est le top).

---

## 📊 TELEMETRY ATTENDUE

```
[+] SaveAll ordonne: 8 doc(s) | IPT=5, Sub-IAM=2, Top-IAM update=1/save=1 | skip=3 | 1247 ms
```

Lecture :
- 5 parts sauvées en premier
- 2 sub-assemblies sauvés ensuite
- 1 top assembly : Update OK + Save OK
- 3 docs skip (drawings/presentations)
- Durée totale 1.2s

---

## ⚠️ POINTS DE VIGILANCE .NET 8

1. **STA Thread obligatoire** — déjà respecté dans le projet (commentaires en tête du fichier).
2. **`ComInvoke.Call(doc, "Update")`** — la méthode COM s'appelle bien `Update` (pas `Update2`) sur `Document`. `Update2` existe aussi mais `Update` est la version standard et synchrone.
3. **Marshal.ReleaseComObject** — bien libérer `topDoc` après lecture du `FullFileName`.
4. **`required` properties** — C# 11+ OK en .NET 8.

---

## 🧪 TEST DE VALIDATION

Après le fix :
1. Ouvrir un assemblage avec sub-assemblies dans Inventor.
2. Modifier 1 part + 1 sub-asm + le top.
3. Déclencher `SaveAll`.
4. Vérifier dans le log : `IPT=1, Sub-IAM=1, Top-IAM update=1/save=1`.
5. Fermer Inventor + rouvrir le sub-asm seul → **doit ouvrir sans erreur de segment**.

---

## 📁 FICHIERS À MODIFIER (1 seul)

`C:\Users\mohammedamine.elgala\source\repos\InventorAutoSave\Services\InventorSaveService.cs`

Localisations exactes :
- `internal sealed class DocEntry` (~ligne 99) → ajouter 2 properties
- `CollectRecursive` (~ligne 540) → identifier top + nouveau ordering + skip drawings
- `SaveAllDocuments` (~ligne 580-655) → identifier top + nouveau ordering + boucle 2-phases + Update Top + telemetry

---

## 🔧 BUILD

```powershell
cd C:\Users\mohammedamine.elgala\source\repos\InventorAutoSave
dotnet build -c Release
```

(Pas de `build-and-run.ps1` style XEAT — c'est `dotnet build` standard pour .NET 8.)

---

## 🚀 PROMPT À UTILISER DANS L'AUTRE PROJET

```
Lis DevDocs/HANDOFF_InventorAutoSave_SegmentFix.md (copie depuis le projet XEAT) 
et applique le fix 4-phases dans Services/InventorSaveService.cs.

Stratégie validée par doc Autodesk officielle :
- Phase 1: IPT Dirty
- Phase 2: Sub-IAM Dirty (non-top)
- Phase 3: Update() + Save() sur Top Assembly (identifié via ActiveDocument.FullFileName)
- Skip TOTAL des .idw/.dwg/.ipn

Modifications additives uniquement, ne pas casser l'API publique.
Build avec: dotnet build -c Release
```

---

✅ **Fin du handoff.**
