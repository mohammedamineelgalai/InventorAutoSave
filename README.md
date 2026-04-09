
# InventorAutoSave

> **Sauvegarde automatique silencieuse pour Autodesk Inventor** via API COM — remplace le script AutoHotKey legacy par une solution WPF/.NET 8 robuste.
>
> Developpe par **Mohammed Amine Elgalai** — XNRGY Climate Systems ULC
>
> **Version**: v1.0.0 | **Date**: 2026-04-08

---

## Table des Matieres

1. [Description](#description)
2. [Pourquoi remplacer AutoHotKey](#pourquoi-remplacer-autohotkey)
3. [Fonctionnalites](#fonctionnalites)
4. [Modes de Sauvegarde](#modes-de-sauvegarde)
5. [Architecture Technique](#architecture-technique)
6. [Installation](#installation)
7. [Build & Release](#build--release)
8. [Configuration](#configuration)
9. [Utilisation](#utilisation)
10. [Securite & Stabilite](#securite--stabilite)

---

## Description

**InventorAutoSave** est une application systray (barre de notification Windows) qui sauvegarde automatiquement les documents Autodesk Inventor a intervalles reguliers. Elle utilise l'**API COM Inventor** pour effectuer des sauvegardes **silencieuses** — sans aucune popup ni raccourci clavier.

L'application a ete concue pour remplacer le script AutoHotKey (AHK) precedent qui utilisait `SendKeys(Ctrl+D)`, declenchant des popups genantes et des conflits avec Inventor.

---

## Pourquoi remplacer AutoHotKey

| Probleme AHK | Solution InventorAutoSave |
|---|---|
| `SendKeys(Ctrl+D)` declenchait la popup "Save multiple files?" | API COM `doc.Save()` — aucune popup |
| Intervalle reinitialisee a chaque restart (hardcode) | Intervalle persistee dans `config.json` |
| Impossible de detecter si Inventor est en calcul | `SilentOperation` + detection calculs |
| Aucune verification des documents modifies | Verifie `doc.Dirty` avant sauvegarde |
| SaveAll sauvegardait meme les docs non modifies | SaveActive intelligent avec detection recursive |
| Pas de logs ni diagnostic | Logger complet avec fichiers quotidiens |
| Pas de reconnexion automatique si Inventor redemarrait | Reconnexion COM automatique silencieuse |

---

## Fonctionnalites

### Core
- ✅ **Sauvegarde silencieuse via API COM** — pas de `SendKeys`, pas de popup
- ✅ **SilentOperation=true** — supprime toutes les popups Inventor
- ✅ **Reconnexion automatique** — si Inventor redemarrait, l'app se reconnecte
- ✅ **Protection calculs** — differe la sauvegarde si Inventor est en calcul (FEA, rebuild)
- ✅ **Retry automatique** — re-essaie apres le calcul (max 60 retries / ~5 min)

### Modes de sauvegarde
- ✅ **SaveActive intelligent** — sauvegarde le doc actif + composants modifies recursivement
- ✅ **SaveAll ordonne** — sauvegarde dans l'ordre ipt → iam → idw (ordre Inventor)

### UI
- ✅ **Systray** — icone dans la barre de notification, discret et permanent
- ✅ **Menu contextuel dark** — theme sombre complet, pas d'artefacts blancs
- ✅ **Fenetre de configuration** — settings visuels avec toggle switches modernes
- ✅ **Notifications tray** — notification ballon apres chaque sauvegarde (desactivable)
- ✅ **Compte a rebours** — affiche le temps avant le prochain save

### Persistence
- ✅ **config.json** — tous les parametres persistes (intervalle, mode, options)
- ✅ **Demarrage avec Windows** — raccourci automatique dans le dossier Startup
- ✅ **Logs quotidiens** — `Logs/InventorAutoSave_YYYYMMDD.log`

---

## Modes de Sauvegarde

### SaveActive (recommande)
Sauvegarde le document actif et **tous ses composants modifies** de maniere recursive:
- Si le doc actif est une **piece** (.ipt) → sauvegarde uniquement cette piece
- Si le doc actif est un **assemblage** (.iam) → parcourt recursivement tous les sous-assemblages et pieces modifies (dirty), puis sauvegarde l'assemblage
- Si le doc actif est un **dessin** (.idw/.dwg) → sauvegarde le dessin uniquement

**Avantage**: Ne sauvegarde que ce qui est pertinent → **plus rapide**

### SaveAll
Sauvegarde **tous les documents ouverts** dans Inventor, dans l'ordre:
1. Parts (.ipt) → composants de base
2. Assemblies (.iam) → referencent les parts
3. Drawings (.idw/.dwg) → referencent les assemblages

**Note**: Sauvegarde meme les docs d'autres projets s'ils sont ouverts.

---

## Architecture Technique

```
InventorAutoSave/
│
├── App.xaml / App.xaml.cs           # Point d'entree, systray, menu contextuel dark
│
├── Models/
│   └── AppSettings.cs               # Config persistee (SaveMode, intervalle, options)
│
├── Services/
│   ├── AutoSaveTimerService.cs      # Timer avec retry, protection calculs
│   ├── InventorSaveService.cs       # Connexion COM + sauvegarde silencieuse (~640 lignes)
│   ├── Logger.cs                    # Logger fichier quotidien
│   ├── SettingsService.cs           # Persistance config.json
│   └── StartupManager.cs           # Raccourci Startup Windows
│
├── ViewModels/
│   └── MainViewModel.cs             # MVVM: commandes, proprietes observables
│
├── Views/
│   ├── SettingsWindow.xaml           # Fenetre configuration dark
│   └── SettingsWindow.xaml.cs
│
├── Styles/
│   └── DarkTheme.xaml               # ResourceDictionary theme sombre complet
│
├── Resources/
│   └── InventorAutoSave.ico         # Icone application
│
├── Setup/                           # Projet installateur standalone
│   ├── SetupWindow.xaml(.cs)        # UI installation (barre de progression, options)
│   └── InventorAutoSave.Setup.csproj
│
├── config.json                      # Configuration par defaut
├── build-and-release.ps1            # Script build + publish + packaging
└── InventorAutoSave.csproj          # Projet WPF .NET 8
```

### Technologies
| Composant | Technologie |
|---|---|
| Framework | .NET 8 (WPF) |
| UI Pattern | MVVM |
| Systray | Hardcodet.NotifyIcon.Wpf |
| COM Interop | P/Invoke ole32.dll (CLSIDFromProgID + GetActiveObject) |
| Persistence | System.Text.Json → config.json |
| Publish | Self-contained single-file (win-x64) |

---

## Installation

### Via l'installateur (recommande)
1. Executer `InventorAutoSave.Setup.exe`
2. Choisir les options:
   - ✅ Demarrer avec Windows
   - ✅ Raccourci Menu Demarrer
   - ✅ Lancer apres installation
3. L'app s'installe dans `%APPDATA%\XNRGY\InventorAutoSave\`

### Installation manuelle
1. Copier `InventorAutoSave.exe` + `config.json` dans un dossier
2. Lancer l'exe — l'icone apparait dans la barre de notification
3. Clic droit → "Demarrer avec Windows" pour activer le demarrage auto

---

## Build & Release

```powershell
# Build + Publish + Package complet
.\build-and-release.ps1

# Build uniquement (rapide)
.\build-and-release.ps1 -Quick

# Sans l'installateur
.\build-and-release.ps1 -SkipSetup
```

### Niveaux du build script:
1. **Clean** — Suppression des artefacts precedents
2. **Build Release** — `dotnet build -c Release`
3. **Publish Self-Contained** — `dotnet publish` (exe standalone win-x64, ~130 MB)
4. **Package dist/** — Copie dans `dist\InventorAutoSave_v1.0.0\`
5. **Build Setup** — Compile l'installateur avec l'exe embarque

---

## Configuration

### config.json
```json
{
  "SaveIntervalSeconds": 180,
  "SaveMode": "SaveActive",
  "EnableAutoSave": true,
  "ShowNotifications": true,
  "NotificationDurationSeconds": 3,
  "StartWithWindows": false,
  "SafetyChecks": true,
  "RetryDelaySeconds": 5
}
```

| Parametre | Default | Description |
|---|---|---|
| `SaveIntervalSeconds` | 180 (3 min) | Intervalle entre chaque sauvegarde |
| `SaveMode` | SaveActive | Mode: `SaveActive` ou `SaveAll` |
| `EnableAutoSave` | true | AutoSave actif au demarrage |
| `ShowNotifications` | true | Notifications ballon apres save |
| `SafetyChecks` | true | Reporter si Inventor en calcul |
| `RetryDelaySeconds` | 5 | Delai retry si calcul en cours |
| `StartWithWindows` | false | Raccourci demarrage Windows |

---

## Utilisation

### Menu contextuel (clic droit sur l'icone systray)
- **Desactiver/Activer AutoSave** — pause/reprise du timer
- **Sauvegarder maintenant** — sauvegarde immediate
- **Mode de sauvegarde** — Document actif / Tous les documents
- **Intervalle de sauvegarde** — 10s a 30 min
- **Notifications** — Activer/Desactiver les notifications tray
- **Protection calculs** — Reporter si Inventor calcule
- **Demarrer avec Windows** — Raccourci Startup
- **Ouvrir la configuration** — Fenetre settings complete
- **Quitter** — Fermer l'application

### Double-clic sur l'icone
Ouvre la fenetre de configuration avec tous les parametres.

---

## Securite & Stabilite

### Protection contre les crashes
- **Global Exception Handlers** — `AppDomain.UnhandledException`, `DispatcherUnhandledException`, `TaskScheduler.UnobservedTaskException`
- **Try/Catch sur chaque operation** — COM, fichiers, UI
- **Reconnexion automatique** — Si la connexion COM est perdue

### Protection Inventor
- **SilentOperation=true** — Supprime toutes les popups Inventor
- **Detection calculs** — Mots-cles: Calculating, Rebuilding, Processing, etc.
- **Retry intelligent** — Max 60 retries (~5 min), puis force la sauvegarde
- **Verification doc.Dirty** — Ne sauvegarde que les documents modifies
- **Verification doc.IsModifiable** — Ignore les documents en lecture seule
- **Anti-boucle infinie** — HashSet pour les assemblages circulaires

### Logs
Fichiers de log quotidiens dans `Logs/InventorAutoSave_YYYYMMDD.log` avec niveaux: DEBUG, INFO, WARNING, ERROR.

---

## Auteur

**Mohammed Amine Elgalai**
- Email: mohammedamine.elgalai@xnrgy.com
- Entreprise: XNRGY Climate Systems ULC
- Date: 2026

---

*Ce projet fait partie des **Smart Tools Amine** — Suite d'outils d'automatisation pour le departement engineering de XNRGY.*
