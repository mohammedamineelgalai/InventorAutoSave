# ============================================================
# build-and-release.ps1
# InventorAutoSave v2.0 - XNRGY Climate Systems ULC
# Build + Publish + Packaging en une seule commande
#
# NIVEAUX:
#   Niveau 1 - Build Release        : dotnet build -c Release
#   Niveau 2 - Publish self-contained: dotnet publish (exe standalone, pas de .NET requis)
#   Niveau 3 - Package dist/        : Copie tout dans dist\InventorAutoSave_v2.0\
#   Niveau 4 - Build Setup          : Build InventorAutoSave.Setup.exe si present
#
# USAGE:
#   .\build-and-release.ps1              # Build + Publish + Package
#   .\build-and-release.ps1 -SkipSetup   # Sans builder l'installateur
#   .\build-and-release.ps1 -Quick       # Build Release uniquement (rapide)
# ============================================================

param(
    [switch]$SkipSetup,
    [switch]$Quick
)

$ErrorActionPreference = "Stop"

# ============================================================
# CONFIGURATION
# ============================================================
$VERSION       = "2.0.0"
$APP_NAME      = "InventorAutoSave"
$COMPANY       = "XNRGY Climate Systems ULC"
$MAIN_PROJECT  = Join-Path $PSScriptRoot "$APP_NAME.csproj"
$SETUP_PROJECT = Join-Path $PSScriptRoot "Setup\InventorAutoSave.Setup.csproj"
$DIST_DIR      = Join-Path $PSScriptRoot "dist\${APP_NAME}_v${VERSION}"
$PUBLISH_DIR   = Join-Path $PSScriptRoot "bin\Publish"

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================
function Write-Header([string]$text) {
    Write-Host ""
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host "  $text" -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
}

function Write-Step([string]$text) {
    Write-Host ""
    Write-Host "  >> $text" -ForegroundColor Yellow
}

function Write-OK([string]$text) {
    Write-Host "  [+] $text" -ForegroundColor Green
}

function Write-Fail([string]$text) {
    Write-Host "  [-] $text" -ForegroundColor Red
}

function Write-Info([string]$text) {
    Write-Host "  [i] $text" -ForegroundColor Gray
}

# ============================================================
# DEBUT
# ============================================================
Write-Header "$APP_NAME v$VERSION - Build Script"
Write-Info "Repertoire : $PSScriptRoot"
Write-Info "Date       : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# ============================================================
# NIVEAU 1 - CLEAN
# ============================================================
Write-Header "NIVEAU 1 - Clean"

Write-Step "Suppression des artefacts precedents..."
$cleanDirs = @("bin\Release", "obj\Release", "bin\Publish", $DIST_DIR)
foreach ($dir in $cleanDirs) {
    $fullDir = Join-Path $PSScriptRoot $dir
    if (Test-Path $fullDir) {
        Remove-Item $fullDir -Recurse -Force
        Write-Info "Supprime: $dir"
    }
}
Write-OK "Clean termine"

# ============================================================
# NIVEAU 2 - BUILD RELEASE
# ============================================================
Write-Header "NIVEAU 2 - Build Release"

Write-Step "Compilation en mode Release..."
$buildOutput = dotnet build $MAIN_PROJECT -c Release --nologo 2>&1
$buildExitCode = $LASTEXITCODE

$errors   = $buildOutput | Where-Object { $_ -match ": error" }
$warnings = $buildOutput | Where-Object { $_ -match ": warning" }

if ($errors.Count -gt 0) {
    Write-Fail "ERREURS DE COMPILATION:"
    $errors | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
    Write-Fail "Build ECHOUE - Arret."
    exit 1
}

Write-OK "Build Release reussi ($($warnings.Count) warning(s))"

if ($Quick) {
    Write-Header "Mode Quick - Build uniquement"
    Write-OK "Termine en $($stopwatch.Elapsed.TotalSeconds.ToString('F1'))s"
    exit 0
}

# ============================================================
# NIVEAU 3 - PUBLISH SELF-CONTAINED
# ============================================================
Write-Header "NIVEAU 3 - Publish Self-Contained (x64)"

Write-Step "Publication standalone (pas de .NET requis sur le PC cible)..."
Write-Info "Target: win-x64, SingleFile, trimmed"

$publishOutput = dotnet publish $MAIN_PROJECT `
    -c Release `
    -r win-x64 `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -o $PUBLISH_DIR `
    --nologo 2>&1

$publishExitCode = $LASTEXITCODE

if ($publishExitCode -ne 0) {
    $pubErrors = $publishOutput | Where-Object { $_ -match ": error" }
    if ($pubErrors.Count -gt 0) {
        Write-Fail "ERREURS PUBLISH:"
        $pubErrors | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
        Write-Fail "Publish ECHOUE - Arret."
        exit 1
    }
}

$mainExe = Join-Path $PUBLISH_DIR "$APP_NAME.exe"
if (-not (Test-Path $mainExe)) {
    Write-Fail "EXE non trouve apres publish: $mainExe"
    exit 1
}

$exeSize = [math]::Round((Get-Item $mainExe).Length / 1MB, 1)
Write-OK "Publish reussi -> $APP_NAME.exe ($exeSize MB)"

# ============================================================
# NIVEAU 4 - PACKAGING dist/
# ============================================================
Write-Header "NIVEAU 4 - Packaging dist\"

Write-Step "Creation du dossier de distribution..."
New-Item -ItemType Directory -Path $DIST_DIR -Force | Out-Null

# Copier l'exe principal
Copy-Item $mainExe $DIST_DIR

# Copier les fichiers de support (ICO, config par defaut)
$supportFiles = @(
    "Resources\InventorAutoSave.ico",
    "config.json"
)
foreach ($file in $supportFiles) {
    $src = Join-Path $PSScriptRoot $file
    if (Test-Path $src) {
        $destDir = Join-Path $DIST_DIR (Split-Path $file -Parent)
        if ($destDir -ne $DIST_DIR) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }
        Copy-Item $src $destDir
        Write-Info "Copie: $file"
    }
}

# Creer README.txt
$readmePath = Join-Path $DIST_DIR "README.txt"
@"
InventorAutoSave v$VERSION - $COMPANY
======================================

INSTALLATION:
  Lancer InventorAutoSave.Setup.exe (installateur graphique)
  OU copier InventorAutoSave.exe dans un dossier de votre choix.

DEMARRAGE AVEC WINDOWS:
  L'installateur propose d'ajouter l'application au demarrage Windows.
  Vous pouvez aussi utiliser SettingsWindow > "Demarrer avec Windows".

UTILISATION:
  L'application s'installe dans la barre de notification Windows (systray).
  Clic droit sur l'icone bleue pour acceder aux options.
  Double-clic pour ouvrir la fenetre de configuration.

MODES DE SAUVEGARDE:
  - Document actif (recommande): Sauvegarde le doc actif + ses composants
    modifies (pieces, sous-assemblages). Rapide, sans popup.
  - Tous les documents: Sauvegarde tous les docs ouverts. Plus lent.

SUPPORT: mohammedamine.elgalai@xnrgy.com
"@ | Set-Content $readmePath -Encoding UTF8

Write-OK "Package cree: dist\${APP_NAME}_v${VERSION}\"

# Lister le contenu
Write-Info "Contenu du package:"
Get-ChildItem $DIST_DIR -Recurse | ForEach-Object {
    $rel = $_.FullName.Replace($DIST_DIR + "\", "")
    $size = if ($_.PSIsContainer) { "[dir]" } else { "$([math]::Round($_.Length/1KB, 1)) KB" }
    Write-Host "    $rel  ($size)" -ForegroundColor Gray
}

# ============================================================
# NIVEAU 5 - BUILD INSTALLATEUR (si present)
# ============================================================
if (-not $SkipSetup) {
    Write-Header "NIVEAU 5 - Build Installateur"

    if (Test-Path $SETUP_PROJECT) {
        Write-Step "Compilation de InventorAutoSave.Setup.exe..."

        $setupOutput = dotnet build $SETUP_PROJECT -c Release --nologo 2>&1
        $setupErrors = $setupOutput | Where-Object { $_ -match ": error" }

        if ($setupErrors.Count -gt 0) {
            Write-Fail "Erreurs Setup (non bloquant - app principale OK):"
            $setupErrors | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkYellow }
        } else {
            # Copier le Setup.exe dans dist/
            $setupExe = Get-ChildItem (Join-Path $PSScriptRoot "Setup\bin\Release") -Filter "*.exe" -Recurse | Select-Object -First 1
            if ($setupExe) {
                Copy-Item $setupExe.FullName $DIST_DIR
                Write-OK "Setup.exe copie dans dist\"
            }
        }
    } else {
        Write-Info "Projet Setup non trouve (Setup\InventorAutoSave.Setup.csproj)"
        Write-Info "Utiliser -SkipSetup pour ignorer cette etape"
    }
}

# ============================================================
# RESUME FINAL
# ============================================================
$stopwatch.Stop()
Write-Header "BUILD COMPLETE"
Write-OK "Version       : v$VERSION"
Write-OK "Temps total   : $($stopwatch.Elapsed.TotalSeconds.ToString('F1'))s"
Write-OK "Output        : dist\${APP_NAME}_v${VERSION}\"

$distSize = (Get-ChildItem $DIST_DIR -Recurse | Measure-Object -Property Length -Sum).Sum
Write-OK "Taille dist   : $([math]::Round($distSize/1MB, 1)) MB"

Write-Host ""
Write-Host "  Pour lancer l'app directement:" -ForegroundColor White
Write-Host "  .\dist\${APP_NAME}_v${VERSION}\${APP_NAME}.exe" -ForegroundColor Cyan
Write-Host ""
