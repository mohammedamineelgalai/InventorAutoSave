# InventorAutoSave — Copilot Instructions

> **Projet**: InventorAutoSave v1.0.0
> **Auteur**: Mohammed Amine Elgalai — XNRGY Climate Systems ULC
> **Stack**: C# 12, .NET 8.0, WPF, COM Interop (Autodesk Inventor)
> **Date**: Avril 2026 (v1.0.0 — 26 avril 2026)

---

## RÈGLE ABSOLUE — LANGUE DE COMMUNICATION

```
╔══════════════════════════════════════════════════════════════════════════════╗
║                                                                              ║
║  [!!!] TOUJOURS RÉPONDRE EN FRANÇAIS                                        ║
║                                                                              ║
║  L'utilisateur communique en FRANÇAIS.                                       ║
║  L'agent DOIT TOUJOURS répondre en FRANÇAIS.                                ║
║                                                                              ║
║  - Toutes les explications en français                                       ║
║  - Tous les résumés en français                                              ║
║  - Toutes les questions en français                                          ║
║  - Toute l'analyse interne (raisonnement, planification) en français         ║
║  - Les noms de variables/classes/méthodes restent en anglais (convention code)║
║  - Les commentaires de code restent en anglais (convention code)             ║
║  - Les messages de commit restent en anglais (convention Git)                ║
║                                                                              ║
║  NE JAMAIS répondre en anglais quand l'utilisateur parle en français.        ║
║  NE JAMAIS analyser ou raisonner en anglais puis répondre en français.       ║
║  TOUT le processus de réflexion et de réponse doit être en FRANÇAIS.         ║
║                                                                              ║
╚══════════════════════════════════════════════════════════════════════════════╝
```

---

## RÉSUMÉ DU PROJET

Application WPF systray qui sauvegarde automatiquement les documents Autodesk Inventor
via l'API COM. Remplace un ancien script AutoHotKey par une solution robuste avec:
- Icône systray avec menu contextuel dark complet
- Fenêtre de configuration (SettingsWindow)
- Connexion COM via `oleaut32.dll` avec marshaling `IDispatch`
- Timer configurable (10s à 30min)
- Installateur self-contained (Setup project)

---

## ARCHITECTURE

```
InventorAutoSave/
├── App.xaml.cs              # Point d'entrée, tray icon, menu contextuel, event handlers
├── Models/
│   └── AppSettings.cs       # Configuration (SaveMode, intervalle, notifications, etc.)
├── ViewModels/
│   └── MainViewModel.cs     # MVVM — coordonne services, commandes, propriétés bindées
├── Views/
│   ├── SettingsWindow.xaml   # Fenêtre de configuration (dark theme)
│   └── SettingsWindow.xaml.cs
├── Services/
│   ├── InventorSaveService.cs  # Connexion COM + sauvegarde silencieuse
│   ├── AutoSaveTimerService.cs # Timer avec marshaling STA
│   ├── SettingsService.cs      # Persistance config.json
│   ├── StartupManager.cs       # Raccourci démarrage Windows
│   └── Logger.cs               # Logging fichier
├── Styles/
│   └── DarkTheme.xaml       # Thème sombre complet (392 lignes)
├── Resources/
│   └── InventorAutoSave.ico
├── Setup/                   # Sous-projet installateur
│   ├── InventorAutoSave.Setup.csproj
│   ├── SetupWindow.xaml(.cs)
│   └── App.xaml(.cs)
├── build-and-release.ps1    # Pipeline build 5 niveaux
└── config.json              # Configuration runtime
```

---

## RÈGLES CRITIQUES — COM INTEROP INVENTOR

```
╔══════════════════════════════════════════════════════════════════╗
║                                                                  ║
║  [!!!] CES RÈGLES ONT ÉTÉ DÉCOUVERTES APRÈS 7 BUGS EN CHAÎNE  ║
║  NE JAMAIS LES VIOLER — SINON CRASH COM GARANTI                ║
║                                                                  ║
╚══════════════════════════════════════════════════════════════════╝
```

### 1. MARSHALING IDispatch (PAS IUnknown)

```csharp
// CORRECT — oleaut32.dll avec IDispatch
[DllImport("oleaut32.dll", EntryPoint = "GetActiveObject", PreserveSig = false)]
private static extern void GetActiveObject(
    [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
    IntPtr pvReserved,
    [MarshalAs(UnmanagedType.IDispatch)] out object ppunk);  // ✅ IDispatch

// INTERDIT — IUnknown cause "Specified cast is not valid"
// [MarshalAs(UnmanagedType.IUnknown)] out object ppunk  // ❌ CRASH
```

### 2. JAMAIS dynamic POUR LES OBJETS COM

```csharp
// CORRECT — object + ComInvoke helper
private object? _inventorApp;  // ✅ object
var doc = ComInvoke.GetProp(_inventorApp, "ActiveDocument");

// INTERDIT — dynamic perd le dispatch COM
// private dynamic? _inventorApp;  // ❌ CRASH
// var doc = _inventorApp.ActiveDocument;  // ❌ CRASH
```

### 3. ComInvoke HELPER — UNIQUE FAÇON D'APPELER COM

```csharp
// La classe ComInvoke utilise Type.InvokeMember avec BindingFlags.GetProperty
// C'est LA SEULE méthode fiable pour accéder aux propriétés COM Inventor
public static class ComInvoke
{
    public static object? GetProp(object? target, string name);
    public static void SetProp(object? target, string name, object value);
    public static object? Call(object? target, string method, params object[] args);
    public static string GetString(object? target, string name);
    public static bool GetBool(object? target, string name);
    public static int GetInt(object? target, string name);
    public static object? GetItem(object? collection, int index);
}
```

### 4. THREAD STA OBLIGATOIRE POUR COM

```csharp
// CORRECT — Marshaler vers le thread UI (STA) pour les appels COM
Application.Current.Dispatcher.Invoke(() => {
    var result = _inventorService.TriggerSave();
});

// INTERDIT — Appeler COM depuis un thread MTA (ThreadPool/Task.Run)
// Task.Run(() => _inventorService.TriggerSave());  // ❌ COM CRASH
```

### 5. Marshal.ReleaseComObject POUR LES OBJETS TEMPORAIRES

```csharp
// CORRECT — Libérer les RCW temporaires pour éviter corruption canal RPC
var docs = ComInvoke.GetProp(_inventorApp, "Documents");
try
{
    int count = ComInvoke.GetInt(docs, "Count");
    return count;
}
finally
{
    if (docs != null) Marshal.ReleaseComObject(docs);
}

// INTERDIT — Laisser le GC gérer les RCW COM
// var docs = ComInvoke.GetProp(_inventorApp, "Documents");
// return ComInvoke.GetInt(docs, "Count");  // ❌ GC détruit le RCW aléatoirement
```

### 6. BeginInvoke (PAS Invoke) POUR LES EVENT HANDLERS

```csharp
// CORRECT — BeginInvoke évite le deadlock Timer → UI
_timerService.SaveCompleted += (s, e) =>
    App.Current.Dispatcher.BeginInvoke(() => OnSaveCompleted(e));

// INTERDIT — Invoke cause deadlock quand timer fire depuis ThreadPool
// _timerService.SaveCompleted += (s, e) =>
//     App.Current.Dispatcher.Invoke(() => OnSaveCompleted(e));  // ❌ DEADLOCK
```

### 7. STRATÉGIE 4-PHASES OBLIGATOIRE POUR SAUVEGARDE

```
╔══════════════════════════════════════════════════════════════════╗
║  [!!!] NE JAMAIS sauvegarder en un seul bloc indéterministe.    ║
║  Cause des erreurs de SEGMENT BREP au prochain ouvrage du       ║
║  sub-assembly seul (corruption silencieuse des références).      ║
╚══════════════════════════════════════════════════════════════════╝
```

Validée par doc Autodesk officielle (DevBlog Tip #8 + MVP J.S. Hould +
`Document.Save2(SaveDependents=True)`).

| Phase | Action | Order |
|-------|--------|-------|
| **1** | `Save()` toutes les **parts (.ipt) Dirty** | 0 |
| **2** | `Save()` tous les **sub-assemblies (.iam) Dirty** (non-top) | 1 |
| **3** | `Update()` **PUIS** `Save()` sur le **Top Assembly** | 2 |
| **Skip** | **Total** des `.idw / .dwg / .ipn` (jamais sauvés directement) | — |

**Intelligence contextuelle** : le top s'adapte à `ActiveDocument` :
- Top `.iam` actif → lui-même = top
- Sub `.iam` ouvert seul → lui-même = "top de session"
- `.ipt` actif → Phase 1 unique
- Drawing/`.ipn` actif → promouvoir le **1er `.iam` Dirty référencé** comme top effectif

Champs critiques de `DocEntry` : `Ext` (string) + `IsTopAssembly` (bool).

Telemetry obligatoire :
```
[+] SaveAll ordonne: 8 doc(s) | IPT=5, Sub-IAM=2, Top-IAM update=1/save=1 | skip=3 | 1247ms
```

Voir `DevDocs/HANDOFF_InventorAutoSave_SegmentFix.md` pour la spec complète.

### 8. P/Invoke — partial class + LibraryImport

```csharp
// CORRECT — partial class + LibraryImport pour les marshalings simples
public partial class InventorSaveService
{
    [LibraryImport("ole32.dll", StringMarshalling = StringMarshalling.Utf16)]
    private static partial int CLSIDFromProgID(string lpszProgID, out Guid pclsid);

    // ATTENTION — IDispatch marshaling NON supporté par LibraryImport
    // → garder en [DllImport] pour GetActiveObject
    [DllImport("oleaut32.dll", EntryPoint = "GetActiveObject", PreserveSig = false)]
    private static extern void GetActiveObject(
        [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
        IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IDispatch)] out object ppunk);
}
```

---

## PRÉFÉRENCES VISUELLES DU DÉVELOPPEUR

```
╔══════════════════════════════════════════════════════════════════╗
║                                                                  ║
║  Mohammed Amine DÉTESTE les textes gris (0xAAAAAA, 0x888888).    ║
║  TOUJOURS utiliser des couleurs claires qui se lisent bien      ║
║  sur fond sombre (0x12121C).                                     ║
║                                                                  ║
╚══════════════════════════════════════════════════════════════════╝
```

### PALETTE COULEURS DARK THEME

| Usage | Couleur | HEX | Notes |
|-------|---------|-----|-------|
| **Fond principal** | Noir profond | `0x12121C` | Background Window |
| **Fond header** | Gris foncé | `0x1E1E2E` | Headers, cartes |
| **Fond cartes** | Gris moyen | `0x252538` | Cards, borders |
| **Texte principal** | Blanc | `White` | Titres, labels principaux |
| **Texte secondaire** | Bleu acier clair | `0xB8C7D6` | Sous-titres, descriptions |
| **Texte muted** | Bleu clair | `0x90CAF9` | Footer, copyright, tech info |
| **Accent principal** | Bleu | `0x4A7FBF` | Bordures, indicateurs |
| **Succès** | Vert | `0x00D26A` | Checkmarks, ON |
| **Connecté** | Vert Microsoft | `0x107C10` | Status Inventor connecté |
| **Déconnecté** | Rouge | `0xE81123` | Status Inventor déconnecté |
| **Avertissement** | Orange | `0xFF8C00` | Warnings |

### COULEURS INTERDITES

- `0xAAAAAA` — Gris moyen → REMPLACÉ par `0xB8C7D6` (bleu acier clair)
- `0x888888` — Gris foncé → REMPLACÉ par `0x90CAF9` (bleu clair)
- `0x666666` — Gris très foncé → INTERDIT sur fond sombre
- Tout gris neutre (0x999, 0x777, 0xAAA, 0xBBB) → Utiliser des bleus/steel blue

### STYLE FOOTER (RÉFÉRENCE XEAT)

```xml
<!-- Les deux côtés du footer DOIVENT avoir le même Foreground="0x90CAF9" -->
<!-- Pas de Bold d'un côté et pas de l'autre -->
<TextBlock Text="© 2026 Mohammed Amine Elgalai — XNRGY Climate Systems ULC"
           Foreground="0x90CAF9" FontSize="10"/>
<TextBlock Text="WPF .NET 8.0 | COM Interop | v1.0.0"
           Foreground="0x90CAF9" FontSize="10"/>
```

### LARGEUR FENÊTRE

- Width: `620` (minimum pour que titre + copyright soient entièrement visibles)
- MinWidth: `550`
- MinHeight: `600`

---

## EMOJIS — RÈGLES D'UTILISATION

### DEUX CONTEXTES DIFFERENTS

**INTERFACES UI** (XAML, Menu contextuel, MessageBox, Boutons, Titres, ShowBalloonTip):
- TOUS les emojis professionnels sont AUTORISES

**CODE BACKEND** — Logger, Console, fichiers .log, commentaires CS:
- INTERDITS — Utiliser les marqueurs ASCII

### Emojis AUTORISES dans les INTERFACES UI

| Emoji | Usage |
|-------|-------|
| ✅ | Succes/Selection |
| ❌ | Erreur/Annulation |
| ⚠️ | Avertissement |
| ℹ️ | Information |
| ❓ | Question |
| 🔄 | Statut/Refresh |
| 📄 | Fichier |
| 📁 | Dossier |
| 🔍 | Recherche |
| 📋 | Liste |
| ⏳ | Attente/Loading |
| ⚙️ | Configuration |
| 💡 | Astuce/Tip |
| 🗄️ | Vault/Database |
| 👤 | Utilisateur |
| 📡 | Connexion |
| 🛠️ | Outils |
| 📐 | Dimensions |
| 📊 | Statistiques |
| 📥 | Download/Import |
| 📤 | Upload/Export |
| 🔒 | Securite/Lock |
| 🔓 | Deverrouille |
| 💾 | Sauvegarde |
| 🔔 | Notifications |
| 🛡️ | Protection |
| ⏱️ | Timer/Intervalle |
| 🔧 | Réglages |
| ⚡ | Rapide/Énergie |
| ▶️ | Démarrer/Play |
| ⏸️ | Pause/Stop |

### Emojis INTERDITS (Non-professionnels - JAMAIS utiliser)

**VISAGES ET EMOTIONS:**
😊 🙂 😄 😁 😆 😂 🤣 🥲 😅 😍 🥰 😢 😭 😔 😒 😠 😡 😎 🥳 🤗 🤔 🙄 😴
😀 😃 😉 😋 😛 😜 🤪 😝 🤑 🤭 🤫 🤥 😶 😐 😑 😬 🙃 😌 😏 😣 😥 😮
🤐 😯 😪 😫 🥱 😴 🤤 😷 🤒 🤕 🤢 🤮 🤧 🥵 🥶 🥴 😵 🤯 🤠 🥳 🥸 😎

**GESTES ET MAINS:**
👍 👎 👏 🤝 ✌️ 🙌 👋 🤚 🖐️ ✋ 🖖 👌 🤌 🤏 ✊ 👊 🤛 🤜 👈 👉 👆 👇
☝️ 🫵 💪 🦾 🖕 🤙 🫶 🫱 🫲 🫳 🫴 🫷 🫸 🙏

**COEURS ET AMOUR:**
❤️ 🧡 💛 💚 💙 💜 🖤 🤍 🤎 💔 💞 💓 💗 💕 💖 💘 💝 💟

**ANIMAUX:**
🐶 🐱 🐭 🐹 🐰 🦊 🐻 🐼 🐨 🐯 🦁 🐮 🐷 🐸 🐵 🙈 🙉 🙊 🐔 🐧 🐦 🐤
🦆 🦅 🦉 🦇 🐺 🐗 🐴 🦄 🐝 🐛 🦋 🐌 🐞 🐜 🦟 🦗 🕷️ 🦂

**NOURRITURE:**
🍎 🍐 🍊 🍋 🍌 🍉 🍇 🍓 🍈 🍒 🍑 🍍 🥝 🍅 🍆 🥑 🥦 🥬 🥒
🌶️ 🌽 🥕 🧄 🧅 🥔 🍠 🥐 🍞 🥖 🧀 🥚 🍳 🥞 🥓 🥩
🍗 🍖 🌭 🍔 🍟 🍕 🥪 🌮 🌯 🥗 🍝 🍜 🍲 🍛
🍣 🍱 🥟 🍤 🍙 🍚 🍘 🍥 🥠 🍢 🍡 🍧 🍨 🍦 🥧 🧁 🍰 🎂 🍮 🍭
🍬 🍫 🍿 🍩 🍪 🥜 🍯 🥛 🍼 ☕ 🍵 🧃 🥤 🍶 🍺 🍻 🥂 🍷 🥃

**NATURE ET METEO:**
🌧️ ☀️ ⛅ 🌙 🌈 ⚡ 🌪️ 🌊 🔥 💧 🌍 🌎 🌏 💫 ⭐ 🌟 ✨ ☄️ 🌑 🌒 🌓
🌔 🌕 🌖 🌗 🌘 🌚 🌝 🌞 🌛 🌜 ☁️ ❄️ 💨 🌫️ 🌀

**DIVERTISSEMENT ET CELEBRATION:**
🎉 🎊 🎈 🎁 🎀 🎄 🎃 🎗️ 🎟️ 🎫 🎖️ 🏆 🏅 🥇 🥈 🥉 ⚽ 🏀 🏈 ⚾ 🎾
🏐 🎱 🏓 🏸 🏒 🏑 🏏 🎣 🥊 🥋 🎽 🛹
🛷 ⛸️ 🎿 ⛷️ 🏂 🎮 🕹️ 🎰 🎲 🧩 🎭 🎨 🎬 🎤 🎧 🎼 🎹 🥁 🎷
🎺 🎸 🎻 🎯

**TRANSPORT:**
🚀 🛸 🚁 🛩️ ✈️ 🛫 🛬 🚂 🚃 🚄 🚅 🚆 🚇 🚈 🚉 🚊 🚝 🚞 🚋 🚌 🚍 🚎
🚐 🚑 🚒 🚓 🚔 🚕 🚖 🚗 🚘 🚙 🚚 🚛 🚜 🏎️ 🏍️ 🛵 🦽 🦼 🛺 🚲 🛴
🚏 ⛽ 🚨 🚥 🚦 🛑 🚧 ⚓ ⛵ 🛶 🚤 🛳️ ⛴️ 🛥️ 🚢

**SYMBOLES DIVERS:**
💯 💢 💥 💦 💨 💣 💬 💭 💤 ♨️ 💈

**DRAPEAUX:**
Tous les drapeaux sont interdits

### Marqueurs ASCII pour le CODE BACKEND

| Contexte UI | Marqueur Code | Usage |
|-------------|---------------|-------|
| ❌ | [-] | Erreur/Echec |
| ✅ | [+] | Succes |
| ⚠️ | [!] | Avertissement |
| 🔄 | [>] | Traitement en cours |
| 📁 | [i] | Information/Dossier |
| ⏳ | [~] | Attente |
| 📋 | [No] | Numero/Liste |
| 🔍 | [?] | Recherche |

### Exemples Code

```csharp
// CORRECT - Dans le code CS (logs)
Logger.Info("[+] Connexion Vault etablie");
Logger.Error("[-] Erreur lors du checkout");
Logger.Warning("[!] Fichier existe deja");
Logger.Debug("[>] Traitement du fichier en cours...");
```

---

## BUILD & DÉPLOIEMENT

```
╔══════════════════════════════════════════════════════════════════════════════╗
║                                                                              ║
║  [!!!] RÈGLE ABSOLUE — COMPILATION                                          ║
║                                                                              ║
║  TOUJOURS compiler via le script build-and-release.ps1                       ║
║  Le script nettoie les fichiers obj\ dupliqués après chaque build/publish.   ║
║                                                                              ║
║  COMMANDES AUTORISÉES:                                                       ║
║    .\build-and-release.ps1 -Quick       # Build Release rapide (app seule)  ║
║    .\build-and-release.ps1 -BuildOnly   # Build sans publish ni setup       ║
║    .\build-and-release.ps1 -SkipSetup   # Build + Publish sans installateur ║
║    .\build-and-release.ps1 -SetupOnly   # Setup seul (après un publish)     ║
║    .\build-and-release.ps1 -Auto        # Full pipeline automatique         ║
║    .\build-and-release.ps1              # Full pipeline interactif          ║
║    .\build-and-release.ps1 -Clean       # Nettoyage seul                    ║
║                                                                              ║
║  COMMANDES INTERDITES (créent des fichiers dupliqués dans obj\):             ║
║    dotnet build                          ❌ INTERDIT                         ║
║    dotnet build -c Release               ❌ INTERDIT                         ║
║    dotnet build -c Debug                 ❌ INTERDIT                         ║
║    dotnet publish                        ❌ INTERDIT                         ║
║    dotnet publish -c Release -r win-x64  ❌ INTERDIT                         ║
║    dotnet clean                          ❌ INTERDIT (utiliser -Clean)       ║
║                                                                              ║
║  POURQUOI: Les commandes dotnet directes créent des fichiers                 ║
║  AssemblyInfo.cs dupliqués dans obj\ que VS Code détecte comme erreurs.      ║
║  Le script nettoie automatiquement ces artefacts après chaque étape.         ║
║                                                                              ║
╚══════════════════════════════════════════════════════════════════════════════╝
```

### Build Script (build-and-release.ps1)

```powershell
# 5 niveaux:
# 1. Clean         — Suppression artefacts (tolère fichiers verrouillés)
# 2. Build Release — dotnet build -c Release
# 3. Publish       — dotnet publish (self-contained, SingleFile, win-x64)
# 4. Package dist/ — Copie dans dist\InventorAutoSave_v1.0.0\
# 5. Build Setup   — Compile l'installateur avec exe embarqué

.\build-and-release.ps1              # Full pipeline
.\build-and-release.ps1 -SkipSetup   # Sans installateur
.\build-and-release.ps1 -Quick       # Build Release uniquement
```

### RÈGLE PUBLISH — DOSSIER publish/ À LA RACINE

```powershell
# CORRECT — Publier dans publish/ à la racine du projet
$PUBLISH_DIR = Join-Path $PSScriptRoot "publish"
dotnet publish -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -p:PublishDir="$PUBLISH_DIR"

# INTERDIT — Publier dans bin\ (mélange artefacts build et publish)
# dotnet publish -c Release -r win-x64  # ❌ Crée bin\...\publish\
```

### Setup Project

Le sous-projet `Setup/` embarque le `.exe` principal comme `EmbeddedResource`:
```xml
<!-- Setup/InventorAutoSave.Setup.csproj -->
<EmbeddedResource Include="..\publish\InventorAutoSave.exe"
                  LogicalName="InventorAutoSave.exe"/>
```

**CRITIQUE**: Le chemin DOIT correspondre exactement au `$PUBLISH_DIR` du build script.

### Installation

```
Répertoire cible: %APPDATA%\XNRGY\InventorAutoSave\
Fichiers installés:
  - InventorAutoSave.exe (self-contained, ~68 MB)
  - config.json
  - Resources\InventorAutoSave.ico
```

---

## SYNCHRONISATION SETTINGS (MENU ↔ FENÊTRE)

### Architecture bidirectionnelle

Le système de settings utilise une synchronisation bidirectionnelle:

1. **Menu contextuel → SettingsWindow**: Via `_viewModel.OnSettingsChangedExternally()`
   qui déclenche `PropertyChanged` sur `Settings`, capté par la fenêtre.

2. **SettingsWindow → Menu contextuel**: Via `menu.Opened += RefreshContextMenuState()`
   qui relit `_settingsService.Current` à chaque ouverture du menu.
   Aussi via `_settingsWindow.Closed += RefreshContextMenuState()`.

### Règles

- TOUJOURS passer par les commandes ViewModel (`ChangeSaveModeCommand`, etc.)
  quand possible, car elles déclenchent `OnPropertyChanged(nameof(Settings))`.
- Si on modifie `_settingsService.Update()` directement (ex: depuis le menu),
  TOUJOURS appeler `_viewModel.OnSettingsChangedExternally()` après.
- Les sous-items du menu (SaveActive, SaveAll, NotifOn, etc.) sont stockés
  comme champs de `App` pour permettre `RefreshContextMenuState()`.

---

## PATTERNS MVVM

### ViewModel → Services

```
MainViewModel
  ├── InventorSaveService  (COM connection + save)
  ├── AutoSaveTimerService (timer + STA marshaling)
  └── SettingsService      (config.json persistence)
```

### Commandes disponibles

| Commande | Action |
|----------|--------|
| `ToggleAutoSaveCommand` | Active/désactive le timer |
| `SaveNowCommand` | Sauvegarde immédiate |
| `ChangeIntervalCommand` | Change l'intervalle (param: int seconds) |
| `ChangeSaveModeCommand` | Change le mode (param: SaveMode enum) |
| `ToggleNotificationsCommand` | Toggle notifications |
| `ToggleSafetyChecksCommand` | Toggle protection calculs |

### Data Binding (SettingsWindow)

```
IsInventorConnected   → InventorStatusText, InventorStatusColor
IsAutoSaveEnabled     → AutoSaveButtonText, AutoSaveButtonColor
ActiveDocumentName    → Nom du document actif
TotalDocuments        → Compteur documents ouverts
DirtyDocuments        → Compteur documents modifiés
NextSaveText          → Compte à rebours (MM:SS)
LastSaveText          → "Il y a Xs" / "Il y a X min"
StatusMessage         → Messages d'état
```

---

## HISTORIQUE DES BUGS RÉSOLUS (v1.0.0 development)

### Chaîne de bugs COM (7 problèmes en cascade)

| # | Bug | Cause | Fix |
|---|-----|-------|-----|
| 1 | DLL introuvable | `ole32.dll` au lieu de `oleaut32.dll` | Corriger le nom DLL |
| 2 | EntryPoint manquant | P/Invoke renommé sans `EntryPoint` | Ajouter `EntryPoint = "GetActiveObject"` |
| 3 | Native DLLs manquantes | SingleFile n'inclut pas WPF natives | `-p:IncludeNativeLibrariesForSelfExtract=true` |
| 4 | XAML crash | `StaticResource` forward-reference | `DynamicResource` |
| 5 | Cast invalid | `dynamic` sur IUnknown COM | Tentative Convert, échec |
| 6 | MissingMethod | `ComHelper` avec `InvokeMember` | Échec aussi |
| 7 | **FIX DÉFINITIF** | `IUnknown` marshaling | **`IDispatch`** marshaling |

### Autres bugs résolus

| Bug | Cause | Fix |
|-----|-------|-----|
| Deadlock timer | `Dispatcher.Invoke()` depuis ThreadPool | `Dispatcher.BeginInvoke()` |
| Save ne marche pas | `dynamic` dispatch échoue sur COM nested | `ComInvoke` static helper |
| RPC channel corrupt | GC détruit RCW temporaires | `Marshal.ReleaseComObject` + refresh 5s |
| Checkboxes invisibles | Dark theme incomplet | ControlTemplate complet pour CheckBox |
| Lignes blanches menu | ContextMenu rendu Windows par défaut | Templates custom `DarkContextMenu` |
| Settings désync | Menu et fenêtre pas synchronisés | `RefreshContextMenuState()` + events |
| Build installe ancien exe | Chemin publish avec sous-dossier | `-p:PublishDir` explicite |

---

## DÉPENDANCES

| Package | Version | Usage |
|---------|---------|-------|
| **Hardcodet.NotifyIcon.Wpf** | 1.1.0 | Icône systray + ballon notifications |
| **.NET 8.0** (net8.0-windows) | 8.0+ | Runtime WPF |
| **oleaut32.dll** | System | COM Interop P/Invoke |

---

## INFORMATIONS DÉVELOPPEUR

```
👤 Mohammed Amine Elgalai
   CAD Designer & Automation Development
   XNRGY Climate Systems ULC
📧 mohammedamine.elgalai@xnrgy.com
💬 Teams: @Mohammed Amine Elgalai
```

---

## CHECKLIST AVANT MODIFICATION

- [ ] Vérifier qu'aucun `dynamic` n'est utilisé pour les objets COM
- [ ] Tout appel COM passe par `ComInvoke` helper
- [ ] Les appels COM sont sur le thread UI (STA)
- [ ] Les RCW temporaires sont libérés avec `Marshal.ReleaseComObject`
- [ ] Pas de `Dispatcher.Invoke` dans les event handlers timer → utiliser `BeginInvoke`
- [ ] Pas de gris (0xAAAAAA, 0x888888) dans l'UI → utiliser bleu acier/bleu clair
- [ ] Les marqueurs logs sont ASCII: `[+]` `[-]` `[!]` `[i]` (pas d'emojis dans le code backend)
- [ ] Emojis interfaces: uniquement professionnels (⚙️ 💡 📊 ✅ ❌ etc.) — JAMAIS visages/gestes/cœurs/animaux/transport
- [ ] Version = 1.0.0 partout (titre, footer, tooltip, csproj, readme)
- [ ] **Sauvegarde** : respecte la stratégie 4-phases (IPT → Sub-IAM → Update+Save Top, skip drawings/.ipn)
- [ ] **Top assembly** identifié via `ActiveDocument.FullFileName` (intelligence contextuelle)
- [ ] `DocEntry` doit avoir `Ext` + `IsTopAssembly`
- [ ] Telemetry obligatoire : `IPT=X, Sub-IAM=Y, Top-IAM update=Z/save=W | skip=N | Tms`
- [ ] `LibraryImport` uniquement pour marshalings simples ; `[DllImport]` pour `IDispatch`
- [ ] Le chemin publish n'a PAS de sous-dossier `publish\`
- [ ] **COMPILER UNIQUEMENT via `.\build-and-release.ps1`** — JAMAIS `dotnet build` / `dotnet publish` directement
- [ ] Toute modification de settings propage vers menu ET fenêtre

---

*Mettre à jour ce fichier quand les workflows évoluent.*
