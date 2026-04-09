# Changelog — InventorAutoSave

> Historique des versions et changements
>
> **Auteur**: Mohammed Amine Elgalai — XNRGY Climate Systems ULC

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

### 📋 Historique des bugs COM resolus (v2.0 → v2.1)
1. `ole32.dll` au lieu de `oleaut32.dll` → Fix DLL name
2. `EntryPoint` manquant sur P/Invoke renomme → Ajout `EntryPoint = "GetActiveObject"`
3. Native WPF DLLs manquantes dans SingleFile → `-p:IncludeNativeLibrariesForSelfExtract=true`
4. `Specified cast is not valid` avec `dynamic` sur `IUnknown` → Tentative `Convert.ToInt32()`
5. `ComHelper` avec `Type.InvokeMember` → Echec aussi (`MissingMethodException`)
6. **`IUnknown` → `IDispatch`** marshaling = **FIX DEFINITIF** ✅

---

## v2.0.0 (2026-04-08) — Refonte Dark Theme + Stabilite

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
