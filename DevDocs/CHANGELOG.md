# Changelog — InventorAutoSave

> Historique des versions et changements
>
> **Auteur**: Mohammed Amine Elgalai — XNRGY Climate Systems ULC

---

## v1.0.0 (2026-04-26) — Fix Segment Corruption + Code Cleanup

### 🛡️ Fix critique : Sauvegarde 4-phases anti-segment-corruption

**Problème résolu** : Sur les assemblages avec sub-assemblies, l'ancienne stratégie sauvegardait
tous les `.iam` avec `Order=1` (mélangés), sans `Update()` sur le top → **erreurs de segment BREP**
au prochain ouvrage du sub-assembly seul (corruption silencieuse des références).

**Solution** : Stratégie 4-phases validée par documentation Autodesk officielle
([DevBlog Tip #8](https://blog.autodesk.io/tips-for-creating-large-assemblies-using-inventor-api/) +
[Jean-Sébastien Hould MVP](https://jshould.ca/blog/2019/01/02/inventor-macro-update-and-save-all/) +
API `Document.Save2(SaveDependents=True)`).

| Phase | Action | Order |
|-------|--------|-------|
| **1** | `Save()` sur **toutes les parts (.ipt) Dirty** | 0 |
| **2** | `Save()` sur **tous les sub-assemblies (.iam) Dirty** (non-top) | 1 |
| **3** | `Update()` puis `Save()` sur le **Top Assembly** identifié via `ActiveDocument` | 2 |
| **Skip** | **Total** des `.idw / .dwg / .ipn` (drafter rebuild risqué + segments BREP absents) | — |

#### 🧠 Intelligence contextuelle (RÈGLE ABSOLUE)

Le top n'est jamais figé : il s'adapte à `ActiveDocument` :

| `ActiveDocument` | Top effectif | Stratégie |
|------------------|--------------|-----------|
| **Top Assembly (.iam)** | Lui-même | Phase 1 → 2 → 3 (Update + Save) |
| **Sub-Assembly (.iam) ouvert seul** | Lui-même = "top de session" | Idem (4 phases sur sa hiérarchie) |
| **Part (.ipt)** | Lui-même | Phase 1 unique : Save direct |
| **Drawing (.idw/.dwg)** ou **Présentation (.ipn)** | Skip ; promotion du **1er .iam Dirty** référencé comme top effectif | Phase 1 + 2 + 3 sur les modèles 3D, drawing/.ipn jamais sauvé |

#### 📊 Telemetry par phase

```
[+] SaveAll ordonne: 8 doc(s) | IPT=5, Sub-IAM=2, Top-IAM update=1/save=1 | skip=3 | 1247ms
[+] SaveActive ordonne: 4 doc(s) | IPT=2, Sub-IAM=1, Top-IAM update=1/save=1 | skip=0 | 638ms
```

### 🔧 Modifications de `Services/InventorSaveService.cs`

- **`DocEntry`** étendu : nouvelles propriétés `Ext` (string) et `IsTopAssembly` (bool)
- **`CollectDirtyDocuments(rootDoc)`** refactoré : détection top via extension du root, promotion du 1er `.iam` si root = drawing/`.ipn`
- **`CollectRecursive`** refactoré : récursion dans `ReferencedDocuments` même pour drawings/`.ipn` (pour collecter leurs modèles 3D), mais skip TOTAL des drawings/`.ipn` du résultat final
- **Nouvel ordering** : `0 = .ipt` / `1 = sub-.iam` / `2 = top-.iam` / `3 = autre`
- **Boucle save 4-phases** appliquée à `SaveActiveDocument` ET `SaveAllDocuments`
- **Telemetry** complète avec stopwatch par opération

### 🧹 Cleanup VS Code (0 erreur, 0 warning)

14 diagnostics analyseurs nettoyés sur `InventorSaveService.cs` :

| Code | Catégorie | Fix |
|------|-----------|-----|
| **RCS1226** ×1 | Roslynator (doc XML) | Wrap `<para>...</para>` dans le bloc IMPORTANT |
| **RCS1001** ×2 | Roslynator (braces) | `{ }` ajoutées aux `if` multi-lignes |
| **IDE0300** ×5 | C# 12 collection expr | `new[] { ... }` → `[...]` (collection literals) |
| **SYSLIB1054** ×1 | LibraryImport | Class `partial` + `[LibraryImport]` pour `CLSIDFromProgID` (les 2 autres P/Invoke restent en `[DllImport]` car IDispatch marshalling non supporté) |
| **CA1860** ×1 | Performance | `.Any()` → `.Length == 0` |
| **CA1840** ×2 | Threading | `Thread.CurrentThread.ManagedThreadId` → `Environment.CurrentManagedThreadId` |
| **CA1822** ×2 | Static | `CollectRecursive` + `CollectDirtyDocuments` rendus `static` |

### 📐 Build

```
.\build-and-release.ps1 -Quick
```
✅ 0 erreur, 0 warning, ~5.5s — exe ~68.5 MB self-contained

---

## v2.1.0 (2026-04-08) — Fix COM Save + Timer

### 🔧 Fix critique: COM IDispatch (Sauvegarde fonctionnelle)
- **Root cause**: `GetActiveObject` avec `MarshalAs(UnmanagedType.IUnknown)` retourne un `__ComObject` nu en .NET 8.
  Les appels `dynamic` et `Type.InvokeMember` echouent tous les deux car le RCW ne connait pas `IDispatch`.
- **Fix**: Changement de `UnmanagedType.IUnknown` → `UnmanagedType.IDispatch` dans les P/Invoke `oleaut32.dll`.
  Avec `IDispatch`, .NET cree un RCW qui supporte `IDispatch::GetIDsOfNames` + `IDispatch::Invoke`,
  ce qui permet le dispatch `dynamic` comme en .NET Framework 4.8.
- **Suppression de `ComHelper`**: La classe de contournement via `Type.InvokeMember` n'est plus necessaire.
  Retour a `dynamic` natif (propre, lisible, meme pattern que XEAT).
- **Verification au connect**: Log `Dynamic OK: Caption = "Autodesk Inventor Professional 2026"` confirme
  que `_inventorApp.Caption` fonctionne via `dynamic` au moment de la connexion.

### 🔧 Fix: Timer AutoSave ne tirait pas
- **Root cause**: `Dispatcher.Invoke()` (synchrone) depuis le thread pool du `System.Timers.Timer`
  provoquait un deadlock quand le thread UI etait occupe.
- **Fix**: `Dispatcher.Invoke()` → `Dispatcher.BeginInvoke()` (asynchrone) pour les handlers
  `SaveCompleted`, `StatusChanged`, et `TimerTick` dans `MainViewModel`.
- **Protection supplementaire**: `try/catch` autour de `TimerTick?.Invoke()` et `TriggerSave()`
  dans `OnTimerElapsed` pour empecher les exceptions de tuer le timer.
- **Diagnostic**: Log `[~] AutoSave: rien a sauvegarder` quand aucun document dirty.

### 🎨 Fix: Context Menu
- **`StaticResource` → `DynamicResource`** pour `MenuDropShadow` dans les `ControlTemplate` du menu
  et sous-menu — evite `StaticResourceHolder` exception au rendu.

### 📋 Historique des bugs COM resolus (v1.0 development)
1. `ole32.dll` au lieu de `oleaut32.dll` → Fix DLL name
2. `EntryPoint` manquant sur P/Invoke renomme → Ajout `EntryPoint = "GetActiveObject"`
3. Native WPF DLLs manquantes dans SingleFile → `-p:IncludeNativeLibrariesForSelfExtract=true`
4. `Specified cast is not valid` avec `dynamic` sur `IUnknown` → Tentative `Convert.ToInt32()`
5. `ComHelper` avec `Type.InvokeMember` → Echec aussi (`MissingMethodException`)
6. **`IUnknown` → `IDispatch`** marshaling = **FIX DEFINITIF** ✅

---

## v1.0.0 (2026-04-08) — Refonte Dark Theme + Stabilite

### 🎨 Theme Sombre
- **ContextMenu**: Template complet (`DarkContextMenu`, `DarkMenuItem`, `DarkSeparator`) qui remplace entierement le rendu Windows par defaut
- **Suppression des artefacts blancs**: Plus de lignes blanches, plus de colonnes icone avec fond clair
- **Checkmarks visibles**: Checkmark vert `#00D26A` sur fond sombre, parfaitement visible
- **Hover avec glow**: Effet `DropShadowEffect` bleu subtil sur survol (comme XEAT)
- **Toggle Switch moderne**: Style ON/OFF avec slide visuel (remplace les anciens ToggleButton texte)
- **ResourceDictionary centralise**: `Styles/DarkTheme.xaml` — palette XNRGY officielle
- **Footer style XEAT**: Copyright blanc gras sur fond `#252536`

### 🔒 Stabilite
- **Global Exception Handlers** (3 niveaux): `AppDomain.UnhandledException`, `DispatcherUnhandledException`, `TaskScheduler.UnobservedTaskException`
- **Fix crash "Ouvrir configuration"**: `SafeOpenSettings()` avec protection complète (thread-safety, try/catch, nettoyage reference)
- **Try/catch sur TOUS les event handlers**: SaveCompleted, StatusChanged, Connected, Disconnected
- **Protection fenetre Settings**: Constructeur protege, chaque click handler protege
- **`args.Handled = true`**: Empeche les exceptions non gerees de tuer l'app

### 📝 Documentation
- **README.md**: Documentation complete (table des matieres, architecture, installation, build, configuration)
- **DevDocs/ARCHITECTURE.md**: Document technique avec structure, flux de donnees, decisions
- **DevDocs/CHANGELOG.md**: Ce fichier

### 🏗️ Architecture
- **Styles externalises**: `Styles/DarkTheme.xaml` merge dans `App.xaml` via ResourceDictionary
- **Styles reutilisables**: `XnrgyPrimaryButton`, `IntervalBtn`, `ToggleSwitch`, `ToggleBtn`, `Card`, `SectionHeader`
- **Pattern identique a XEAT**: Meme palette, memes noms de couleurs, meme approche

---

## v1.0.0 (2026-04-01) — Version Initiale

### Core
- Sauvegarde automatique via API COM Inventor
- Modes SaveActive (document actif + composants) et SaveAll
- Timer configurable (10s a 30 min)
- SilentOperation pour supprimer les popups
- Reconnexion automatique a Inventor
- Protection calculs avec retry

### UI
- Icone systray avec menu contextuel
- Fenetre de configuration (SettingsWindow)
- Notifications ballon

### Build
- Script `build-and-release.ps1` (5 niveaux)
- Installateur standalone (`Setup/`)
- Publish self-contained SingleFile (win-x64)

---

*InventorAutoSave — XNRGY Climate Systems ULC*
