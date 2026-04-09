# Architecture — InventorAutoSave

> **Document technique** — Structure du projet et decisions d'architecture
>
> **Version**: 2.0.0 | **Date**: 2026-04-08
> **Auteur**: Mohammed Amine Elgalai — XNRGY Climate Systems ULC

---

## Vue d'Ensemble

InventorAutoSave est une application **WPF/.NET 8** structuree selon le pattern **MVVM** (Model-View-ViewModel). Elle fonctionne comme un service systray silencieux qui sauvegarde automatiquement les documents Autodesk Inventor via l'API COM.

### Philosophie
- **Remplacement robuste du script AHK** — Plus de `SendKeys`, tout passe par l'API COM
- **Zero popup** — `SilentOperation=true` + `doc.Save()` direct
- **Discret** — Systray uniquement, pas de fenetre principale
- **Resilient** — Reconnexion auto, retry, protection calculs, global exception handlers

---

## Structure du Projet

```
InventorAutoSave/
│
├── App.xaml                         # Ressources globales (DarkTheme.xaml)
├── App.xaml.cs                      # Point d'entree: systray, menu dark, exception handlers
│                                    #   - CreateTrayIcon()
│                                    #   - BuildDarkContextMenu() ← template complet, pas de rendu Windows
│                                    #   - SafeOpenSettings() ← protection crash
│                                    #   - Global exception handlers (3 niveaux)
│
├── Models/
│   └── AppSettings.cs               # Modeles de donnees
│       ├── SaveMode (enum)           #   SaveActive | SaveAll
│       ├── AppSettings (class)       #   Config persistee (JSON serializable)
│       └── SaveResult (class)        #   Resultat d'une tentative de sauvegarde
│
├── Services/
│   ├── AutoSaveTimerService.cs      # Timer principal (~230 lignes)
│   │   ├── Start/Stop/ChangeInterval
│   │   ├── TriggerSave()            # Declenchement (timer ou manuel)
│   │   ├── Retry Timer              # Si Inventor en calcul: retry toutes les N secondes
│   │   └── Events: SaveCompleted, StatusChanged, TimerTick
│   │
│   ├── InventorSaveService.cs       # Connexion COM + sauvegarde (~640 lignes)
│   │   ├── TryConnect()             # P/Invoke CLSIDFromProgID + GetActiveObject
│   │   ├── Save(SaveMode)           # Dispatch vers SaveActive ou SaveAll
│   │   ├── SaveActiveDocument()     # Sauvegarde intelligente doc actif + composants dirty
│   │   ├── SaveAllDocuments()       # Sauvegarde tous les docs ouverts (trie par type)
│   │   ├── CollectDirtyDocuments()  # Parcours recursif des composants modifies
│   │   ├── IsInventorCalculating()  # Detection calculs via titre fenetre
│   │   └── GetDocumentCounts()      # Comptage docs ouverts/modifies
│   │
│   ├── SettingsService.cs           # Persistance config.json (~90 lignes)
│   │   ├── Load() → AppSettings
│   │   ├── Save(AppSettings) → config.json
│   │   └── Update(Action<AppSettings>) → atomic update + save
│   │
│   ├── Logger.cs                    # Logging quotidien (~35 lignes)
│   │   └── Log(message, LogLevel) → Logs/InventorAutoSave_YYYYMMDD.log
│   │
│   └── StartupManager.cs           # Raccourci Startup Windows (~120 lignes)
│       ├── EnableStartup()          # Cree raccourci .lnk via WScript.Shell COM
│       ├── DisableStartup()         # Supprime le raccourci
│       └── IsStartupEnabled         # File.Exists sur le .lnk
│
├── ViewModels/
│   └── MainViewModel.cs             # ViewModel principal (~280 lignes)
│       ├── Proprietes observables (IsAutoSaveEnabled, StatusMessage, etc.)
│       ├── Commands: ToggleAutoSave, SaveNow, ChangeInterval, ChangeSaveMode
│       ├── UI Refresh Timer (1s)    # Compte a rebours, reconnexion, compteurs
│       └── RelayCommand (ICommand)  # Implementation MVVM simple
│
├── Views/
│   ├── SettingsWindow.xaml          # Fenetre configuration dark
│   └── SettingsWindow.xaml.cs       # Code-behind avec protection erreurs
│
├── Styles/
│   └── DarkTheme.xaml               # ResourceDictionary theme sombre XNRGY
│       ├── Couleurs de base          # Palette officielle XNRGY
│       ├── DarkContextMenu           # ControlTemplate complet (pas de rendu Windows)
│       ├── DarkMenuItem              # Template avec checkmark vert + hover glow
│       ├── DarkSeparator             # Ligne fine #3A3A54
│       ├── ToggleSwitch              # ON/OFF slide moderne
│       ├── ToggleBtn                 # Toggle texte (mode sauvegarde)
│       ├── XnrgyPrimaryButton        # Bouton bleu avec glow hover
│       ├── IntervalBtn               # Boutons intervalle
│       └── Card, SectionHeader       # Styles carte/section
│
├── Resources/
│   └── InventorAutoSave.ico
│
├── Setup/                           # Sous-projet installateur
│   ├── SetupWindow.xaml(.cs)        # UI dark avec barre progression
│   ├── App.xaml(.cs)                # Point d'entree Setup
│   └── InventorAutoSave.Setup.csproj
│
├── config.json                      # Configuration par defaut
├── build-and-release.ps1            # Script PowerShell (5 niveaux)
├── README.md                        # Documentation utilisateur
└── DevDocs/
    └── ARCHITECTURE.md              # Ce document
```

---

## Flux de Donnees

```
 ┌─────────────────────────────────────────────────────────────────┐
 │                        App.xaml.cs                               │
 │   ┌─────────┐    ┌──────────────┐    ┌───────────────────────┐  │
 │   │ TrayIcon │───→│ ContextMenu  │    │ SettingsWindow (MVVM) │  │
 │   │  systray │    │ (DarkTheme)  │    │  DataContext=ViewModel│  │
 │   └────┬─────┘    └──────┬───────┘    └───────────┬───────────┘  │
 │        │                 │                        │              │
 │        ▼                 ▼                        ▼              │
 │   ┌────────────────────────────────────────────────────────┐    │
 │   │              MainViewModel (MVVM)                       │    │
 │   │  Commands → ToggleAutoSave, SaveNow, ChangeInterval    │    │
 │   │  Properties → IsConnected, NextSave, StatusMessage     │    │
 │   │  UI Timer (1s) → refresh compteurs + reconnexion       │    │
 │   └───────────┬────────────────────┬───────────────────────┘    │
 │               │                    │                             │
 │               ▼                    ▼                             │
 │   ┌─────────────────┐  ┌──────────────────────┐                │
 │   │ AutoSaveTimer   │  │ SettingsService       │                │
 │   │ Service          │  │  Load/Save/Update     │                │
 │   │  Start/Stop      │  │  → config.json        │                │
 │   │  TriggerSave()   │  └──────────────────────┘                │
 │   └────────┬─────────┘                                          │
 │            │                                                     │
 │            ▼                                                     │
 │   ┌─────────────────────────────────────────────────────────┐   │
 │   │            InventorSaveService (COM)                     │   │
 │   │                                                          │   │
 │   │  TryConnect() ──→ ole32.dll P/Invoke                    │   │
 │   │  Save(mode) ───→ SilentOperation=true                   │   │
 │   │                   doc.Save() silencieux                  │   │
 │   │                   Pas de popup, pas de SendKeys          │   │
 │   └─────────────────────────────────────────────────────────┘   │
 └─────────────────────────────────────────────────────────────────┘
```

---

## Decisions Techniques

### 1. Pourquoi P/Invoke au lieu de Marshal.GetActiveObject?
`Marshal.GetActiveObject()` a ete retire dans .NET 5+. On utilise directement les fonctions `CLSIDFromProgID` et `GetActiveObject` de `ole32.dll` via P/Invoke. C'est la methode recommandee pour .NET 6/7/8.

### 2. Pourquoi System.Timers.Timer et pas DispatcherTimer?
Les sauvegardes COM peuvent etre longues (assemblages avec beaucoup de composants). `System.Timers.Timer` execute sur un thread pool, evitant de bloquer l'UI. Le flag `_isSaving` protege contre les sauvegardes concurrentes.

### 3. Pourquoi un ControlTemplate complet pour le ContextMenu?
Le rendu par defaut de WPF pour `MenuItem` utilise un theme Windows qui inclut:
- Une colonne icone avec fond blanc/gris
- Des lignes de separation blanches
- Des checkmarks invisibles sur fond sombre

On ne peut PAS corriger ca avec de simples Setters de couleur. Il FAUT un `ControlTemplate` complet qui remplace entierement le rendu de chaque `MenuItem`.

### 4. Pourquoi SafeOpenSettings() au lieu de OpenSettings()?
Le crash d'origine venait de plusieurs causes possibles:
- Exception dans `InitializeComponent()` du SettingsWindow
- Tentative de creation alors que l'ancien Window n'etait pas GC
- Exception dans les bindings MVVM
- Race condition si appele depuis un thread non-UI

`SafeOpenSettings()` protege avec:
- Verification `Dispatcher.CheckAccess()`
- Nettoyage explicite de `_settingsWindow = null`
- Try/catch complet avec logging
- `args.Handled = true` dans le global handler pour empecher le crash total

### 5. Pourquoi 3 niveaux de Global Exception Handlers?
- `AppDomain.UnhandledException` — Exceptions non gerees dans n'importe quel thread
- `DispatcherUnhandledException` — Exceptions UI thread (WPF specifique)
- `TaskScheduler.UnobservedTaskException` — Exceptions dans les `Task.Run()` non awaites

---

## Corrections par rapport au script AHK

| Bug AHK | Solution |
|---|---|
| `SendKeys(Ctrl+D)` → popup "Save multiple files?" | `doc.Save()` via COM, pas de popup |
| Intervalle resetee a chaque restart | `config.json` persiste |
| Pas de detection doc modifie | `doc.Dirty` verifie avant save |
| SaveAll meme si rien n'est modifie | Verification `Dirty` pour chaque doc |
| Pas de protection pendant les calculs | Detection via titre fenetre + retry |
| Crash silencieux si Inventor ferme | Reconnexion automatique + try/catch COM |
| Aucun log | Logger quotidien avec 4 niveaux |

---

## Build Script (build-and-release.ps1)

Le script PowerShell automatise tout le pipeline de release:

| Niveau | Action | Output |
|---|---|---|
| 1 | Clean | Suppression bin/Release, obj/Release, dist/ |
| 2 | Build Release | dotnet build -c Release |
| 3 | Publish Self-Contained | dotnet publish -r win-x64 --self-contained (SingleFile) |
| 4 | Package dist/ | Copie exe + config + ico + README dans dist/ |
| 5 | Build Setup | Compile l'installateur avec l'exe embarque |

Usage:
```powershell
.\build-and-release.ps1              # Pipeline complet
.\build-and-release.ps1 -Quick       # Build Release uniquement
.\build-and-release.ps1 -SkipSetup   # Sans installateur
```

---

*Document technique — InventorAutoSave v1.0.0 — XNRGY Climate Systems ULC*
