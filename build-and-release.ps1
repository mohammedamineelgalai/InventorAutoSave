#######################################################################
#  InventorAutoSave - BUILD & RELEASE SCRIPT
#  Version: 1.0.0
#  Author: Mohammed Amine Elgalai - XNRGY Climate Systems ULC
#  Date: 2026-04-09
#
#  Usage:
#    .\build-and-release.ps1                    # Build Release + Publish + Package
#    .\build-and-release.ps1 -Quick             # Build Release uniquement (rapide)
#    .\build-and-release.ps1 -BuildOnly         # [RECOMMANDE DEV] Clean + Build UNIQUEMENT
#    .\build-and-release.ps1 -SkipSetup         # Build + Publish + Package (sans installateur)
#    .\build-and-release.ps1 -SetupOnly         # Build uniquement l'installateur
#    .\build-and-release.ps1 -Clean             # Clean seul (supprime artefacts)
#    .\build-and-release.ps1 -Auto              # [ROBO MODE] Tout automatique
#######################################################################

param(
    [switch]$Quick,
    [switch]$BuildOnly,
    [switch]$SkipSetup,
    [switch]$SetupOnly,
    [switch]$Clean,
    [switch]$Auto
)

$ErrorActionPreference = "Stop"

# ============================================================
# CONFIGURATION
# ============================================================
$VERSION       = "1.0.0"
$APP_NAME      = "InventorAutoSave"
$COMPANY       = "XNRGY Climate Systems ULC"
$MAIN_PROJECT  = Join-Path $PSScriptRoot "$APP_NAME.csproj"
$SETUP_PROJECT = Join-Path $PSScriptRoot "Setup\InventorAutoSave.Setup.csproj"
$DIST_DIR      = Join-Path $PSScriptRoot "dist\${APP_NAME}_v${VERSION}"
$PUBLISH_DIR   = Join-Path $PSScriptRoot "publish"

Push-Location $PSScriptRoot

# ============================================================
# MODE BUILDONLY
# ============================================================
if ($BuildOnly) {
    $SkipSetup = $true
    $SetupOnly = $false
    $Clean = $true
}

# ============================================================
# MODE AUTO (ROBO)
# ============================================================
if ($Auto) {
    Write-Host ""
    Write-Host "  ================================================================" -ForegroundColor Magenta
    Write-Host "  [ROBO MODE] SEQUENCE AUTOMATIQUE COMPLETE ACTIVEE" -ForegroundColor Magenta
    Write-Host "  ================================================================" -ForegroundColor Magenta
    Write-Host "    1. Kill instances $APP_NAME" -ForegroundColor White
    Write-Host "    2. Clean artefacts (bin/obj/dist)" -ForegroundColor White
    Write-Host "    3. Build Release (dotnet build)" -ForegroundColor White
    Write-Host "    4. Publish Self-Contained (SingleFile x64)" -ForegroundColor White
    Write-Host "    5. Package dist/ (exe + config + ico + README)" -ForegroundColor White
    Write-Host "    6. Build Installateur (Setup.exe embarque)" -ForegroundColor White
    Write-Host "  ================================================================" -ForegroundColor Magenta
    Write-Host ""
    $Clean = $true
    $SkipSetup = $false
}

# Calculer les etapes
$StepKill    = $Auto
$StepClean   = $true
$StepBuild   = $true
$StepPublish = (-not $Quick -and -not $BuildOnly)
$StepPackage = (-not $Quick -and -not $BuildOnly)
$StepSetup   = ((-not $Quick -and -not $BuildOnly -and -not $SkipSetup) -or $SetupOnly)

# Mode Clean seul
if ($Clean -and -not $Auto -and -not $BuildOnly -and -not $Quick -and -not $SetupOnly -and -not $StepPublish) {
    $StepBuild = $false
    $StepPublish = $false
    $StepPackage = $false
    $StepSetup = $false
}

$TotalSteps = 0
if ($StepKill)    { $TotalSteps++ }
if ($StepClean)   { $TotalSteps++ }
if ($StepBuild)   { $TotalSteps++ }
if ($StepPublish) { $TotalSteps++ }
if ($StepPackage) { $TotalSteps++ }
if ($StepSetup)   { $TotalSteps++ }

$CurrentStep = 0

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================
function Show-Header {
    Write-Host ""
    Write-Host "  ================================================================" -ForegroundColor Cyan
    $title = "  $APP_NAME v$VERSION - BUILD & RELEASE"
    Write-Host $title -ForegroundColor Cyan
    if ($Auto) {
        Write-Host "  Mode: [ROBO] AUTOMATIQUE COMPLET" -ForegroundColor Magenta
    } elseif ($BuildOnly) {
        Write-Host "  Mode: [BUILD] Compilation uniquement" -ForegroundColor Yellow
    } elseif ($Quick) {
        Write-Host "  Mode: [QUICK] Build Release rapide" -ForegroundColor Yellow
    } elseif ($SkipSetup) {
        Write-Host "  Mode: Build + Publish (sans Setup)" -ForegroundColor Yellow
    } elseif ($SetupOnly) {
        Write-Host "  Mode: [SETUP] Installateur uniquement" -ForegroundColor Yellow
    } else {
        Write-Host "  $COMPANY" -ForegroundColor Cyan
    }
    Write-Host "  ================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  [i] Repertoire : $PSScriptRoot" -ForegroundColor Gray
    Write-Host "  [i] Date       : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
    Write-Host "  [i] Etapes     : $TotalSteps" -ForegroundColor Gray
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

function Write-Warn([string]$text) {
    Write-Host "  [!] $text" -ForegroundColor Yellow
}

# ============================================================
# NETTOYAGE POST-BUILD — Supprime les artefacts obj\ qui causent
# des "Duplicate AssemblyInfo" dans VS Code.
# Appelé après CHAQUE build/publish pour corriger le problème
# même si un agent a fait "dotnet build" directement (INTERDIT).
# ============================================================
function Remove-ObjDuplicates {
    param([string]$Context = "post-build")
    $cleaned = $false

    # 1. Supprimer obj\Debug (si on build en Release, Debug est residuel)
    #    OmniSharp/Roslyn peut recreer obj\Debug en arriere-plan — on le supprime
    foreach ($objDebug in @("obj\Debug", "Setup\obj\Debug")) {
        $fullPath = Join-Path $PSScriptRoot $objDebug
        if (Test-Path $fullPath) {
            Remove-Item $fullPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Info "[$Context] Nettoyage $objDebug (artefact residuel)"
            $cleaned = $true
        }
    }

    # 2. Supprimer obj\...\win-x64 (cree par dotnet publish -r win-x64)
    foreach ($objRoot in @("obj", "Setup\obj")) {
        $winX64 = Join-Path $PSScriptRoot "$objRoot\Release\net8.0-windows\win-x64"
        if (Test-Path $winX64) {
            Remove-Item $winX64 -Recurse -Force -ErrorAction SilentlyContinue
            Write-Info "[$Context] Nettoyage $objRoot\...\win-x64 (artefact publish)"
            $cleaned = $true
        }
    }

    # 3. Arreter le build server pour empecher la recreation de obj\Debug
    #    Le Roslyn build server maintient un cache qui peut recreer obj\Debug
    try {
        $null = dotnet build-server shutdown 2>&1
        Write-Info "[$Context] Build server arrete (evite recreation obj\Debug)"
    } catch { }

    # 4. Seconde passe — re-supprimer obj\Debug au cas ou le shutdown l'a recree
    foreach ($objDebug in @("obj\Debug", "Setup\obj\Debug")) {
        $fullPath = Join-Path $PSScriptRoot $objDebug
        if (Test-Path $fullPath) {
            Remove-Item $fullPath -Recurse -Force -ErrorAction SilentlyContinue
            $cleaned = $true
        }
    }

    if ($cleaned) {
        Write-OK "[$Context] Artefacts dupliques nettoyes"
    }
}

function Stop-ExistingInstances {
    $script:CurrentStep++
    Write-Host ""
    Write-Host "  [$script:CurrentStep/$TotalSteps] Arret des instances existantes..." -ForegroundColor Yellow
    $killed = $false
    $maxRetries = 3
    for ($retry = 1; $retry -le $maxRetries; $retry++) {
        $processes = Get-Process -Name $APP_NAME -ErrorAction SilentlyContinue
        if ($processes) {
            foreach ($proc in $processes) {
                try { $proc.Kill(); $proc.WaitForExit(3000) } catch { }
            }
            Write-Host "        [+] $($processes.Count) instance(s) arretee(s)" -ForegroundColor Green
            $killed = $true
        }
        $null = cmd /c "taskkill /F /T /IM $APP_NAME.exe 2>nul"
        Start-Sleep -Milliseconds 500
        $stillRunning = Get-Process -Name $APP_NAME -ErrorAction SilentlyContinue
        if (-not $stillRunning) { break }
        if ($retry -lt $maxRetries) {
            Write-Host "        [!] Processus encore actif, tentative $retry/$maxRetries..." -ForegroundColor Yellow
            Start-Sleep -Seconds 1
        }
    }
    $finalCheck = Get-Process -Name $APP_NAME -ErrorAction SilentlyContinue
    if ($finalCheck) {
        Write-Fail "Impossible de fermer l'application apres $maxRetries tentatives"
        Write-Fail "Fermez l'application manuellement et relancez le script"
        Pop-Location; exit 1
    }
    if (-not $killed) { Write-Host "        [+] Aucune instance en cours" -ForegroundColor Gray }
    Start-Sleep -Seconds 1
}

# ============================================================
# DEBUT
# ============================================================
Show-Header
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# ============================================================
# ETAPE: KILL INSTANCES (Mode Auto uniquement)
# ============================================================
if ($StepKill) { Stop-ExistingInstances }

# ============================================================
# ETAPE: CLEAN
# ============================================================
if ($StepClean) {
    $CurrentStep++
    Write-Host ""
    Write-Host "  [$CurrentStep/$TotalSteps] Nettoyage des artefacts..." -ForegroundColor Yellow
    $cleanDirs = @("bin", "obj", "publish", "Setup\bin", "Setup\obj", $DIST_DIR)
    $cleanedCount = 0
    foreach ($dir in $cleanDirs) {
        $fullDir = if ([System.IO.Path]::IsPathRooted($dir)) { $dir } else { Join-Path $PSScriptRoot $dir }
        if (Test-Path $fullDir) {
            try {
                Remove-Item $fullDir -Recurse -Force -ErrorAction Stop
                Write-Host "        [+] Supprime: $dir" -ForegroundColor Gray
                $cleanedCount++
            } catch {
                Write-Host "        [!] Impossible de supprimer $dir (fichier verrouille?) - on continue..." -ForegroundColor Yellow
            }
        }
    }
    if ($cleanedCount -eq 0) {
        Write-Host "        [+] Rien a nettoyer" -ForegroundColor Gray
    } else {
        Write-OK "Clean termine ($cleanedCount dossier(s) supprime(s))"
    }
}

# Mode Clean seul - on s'arrete la
if (-not $StepBuild) {
    $stopwatch.Stop()
    Write-Host ""
    Write-Host "  ================================================================" -ForegroundColor Cyan
    Write-Host "                          TERMINE" -ForegroundColor Cyan
    Write-Host "  ================================================================" -ForegroundColor Cyan
    Write-OK "Nettoyage termine en $($stopwatch.Elapsed.TotalSeconds.ToString('F1'))s"
    Write-Host ""
    Pop-Location; exit 0
}

# ============================================================
# ETAPE: BUILD RELEASE
# ============================================================
$CurrentStep++
Write-Host ""
Write-Host "  [$CurrentStep/$TotalSteps] Compilation en mode Release..." -ForegroundColor Yellow

$buildOutput = dotnet build $MAIN_PROJECT -c Release --nologo 2>&1
$null = $LASTEXITCODE  # Consumed to avoid PSScriptAnalyzer warning

$errors   = $buildOutput | Where-Object { $_ -match ": error" }
$warnings = $buildOutput | Where-Object { $_ -match ": warning" }

if ($errors.Count -gt 0) {
    Write-Fail "ERREURS DE COMPILATION:"
    Write-Host ""
    Write-Host "        +-- ERREURS ---------------------------------------------------" -ForegroundColor Red
    foreach ($e in $errors) {
        if ($e -match "([^\\]+\.cs)\((\d+),\d+\):\s+error\s+(CS\d+):\s+(.+)") {
            Write-Host "        |  [$($Matches[3])] $($Matches[1]) (ligne $($Matches[2]))" -ForegroundColor Red
            Write-Host "        |           $($Matches[4].Trim())" -ForegroundColor DarkRed
        } else {
            Write-Host "        |  $($e.ToString().Trim())" -ForegroundColor Red
        }
    }
    Write-Host "        +--------------------------------------------------------------" -ForegroundColor Red
    Write-Fail "Build ECHOUE - Arret."
    Pop-Location; exit 1
}

if ($warnings.Count -gt 0) {
    Write-Host "        [+] Compilation reussie ($($warnings.Count) warning(s))" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "        +-- WARNINGS --------------------------------------------------" -ForegroundColor Yellow
    foreach ($w in $warnings) {
        if ($w -match "([^\\]+\.cs)\((\d+),\d+\):\s+warning\s+(CS\d+):\s+(.+)") {
            Write-Host "        |  [$($Matches[3])] $($Matches[1]) (ligne $($Matches[2]))" -ForegroundColor Yellow
            Write-Host "        |           $($Matches[4].Trim())" -ForegroundColor DarkYellow
        } else {
            Write-Host "        |  $($w.ToString().Trim())" -ForegroundColor Yellow
        }
    }
    Write-Host "        +--------------------------------------------------------------" -ForegroundColor Yellow
} else {
    Write-OK "Compilation reussie (0 warning(s))"
}

# Nettoyer les artefacts dupliques apres chaque build
Remove-ObjDuplicates -Context "post-build"

$buildSummary = "$($errors.Count) erreur(s), $($warnings.Count) avertissement(s)"

# Mode Quick ou BuildOnly - on s'arrete apres le build
if ($Quick -or $BuildOnly) {
    $stopwatch.Stop()
    Write-Host ""
    Write-Host "  ================================================================" -ForegroundColor Cyan
    Write-Host "                          TERMINE" -ForegroundColor Cyan
    Write-Host "  ================================================================" -ForegroundColor Cyan
    Write-Host "  Build : $buildSummary" -ForegroundColor DarkGray
    Write-OK "Temps total : $($stopwatch.Elapsed.TotalSeconds.ToString('F1'))s"
    Write-Host ""
    Pop-Location; exit 0
}

# ============================================================
# ETAPE: PUBLISH SELF-CONTAINED
# ============================================================
if ($StepPublish) {
    $CurrentStep++
    Write-Host ""
    Write-Host "  [$CurrentStep/$TotalSteps] Publication Self-Contained (win-x64, SingleFile)..." -ForegroundColor Yellow
    Write-Info "Target: win-x64 | SingleFile | Self-Contained"
    Write-Info "Output: $PUBLISH_DIR"

    $publishOutput = dotnet publish $MAIN_PROJECT -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true -p:PublishDir="$PUBLISH_DIR" --nologo 2>&1
    $publishExitCode = $LASTEXITCODE

    if ($publishExitCode -ne 0) {
        $pubErrors = $publishOutput | Where-Object { $_ -match ": error" }
        if ($pubErrors.Count -gt 0) {
            Write-Fail "ERREURS PUBLISH:"
            Write-Host ""
            Write-Host "        +-- ERREURS PUBLISH -------------------------------------------" -ForegroundColor Red
            foreach ($e in $pubErrors) {
                Write-Host "        |  $($e.ToString().Trim())" -ForegroundColor Red
            }
            Write-Host "        +--------------------------------------------------------------" -ForegroundColor Red
            Write-Fail "Publish ECHOUE - Arret."
            Pop-Location; exit 1
        }
    }

    $mainExe = Join-Path $PUBLISH_DIR "$APP_NAME.exe"
    if (-not (Test-Path $mainExe)) {
        Write-Fail "EXE non trouve apres publish: $mainExe"
        Pop-Location; exit 1
    }

    $exeSize = [math]::Round((Get-Item $mainExe).Length / 1MB, 1)
    $exeDate = (Get-Item $mainExe).LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
    Write-OK "Publish reussi -> $APP_NAME.exe ($exeSize MB)"
    Write-Info "Date: $exeDate"

    # Nettoyer les artefacts dupliques apres publish (win-x64, Debug, ref/refint)
    Remove-ObjDuplicates -Context "post-publish"
}

# ============================================================
# ETAPE: PACKAGING dist/
# ============================================================
if ($StepPackage) {
    $CurrentStep++
    Write-Host ""
    Write-Host "  [$CurrentStep/$TotalSteps] Packaging dist\..." -ForegroundColor Yellow

    Write-Step "Creation du dossier de distribution..."
    New-Item -ItemType Directory -Path $DIST_DIR -Force | Out-Null

    $mainExe = Join-Path $PUBLISH_DIR "$APP_NAME.exe"
    Copy-Item $mainExe $DIST_DIR
    $exeSize = [math]::Round((Get-Item $mainExe).Length / 1MB, 1)
    Write-Host "        [+] $APP_NAME.exe ($exeSize MB)" -ForegroundColor Green

    $supportFiles = @(
        @{ Path = "Resources\InventorAutoSave.ico"; Desc = "Icone application" },
        @{ Path = "config.json"; Desc = "Configuration par defaut" }
    )

    $copiedFiles = 1
    foreach ($file in $supportFiles) {
        $src = Join-Path $PSScriptRoot $file.Path
        if (Test-Path $src) {
            $parentDir = Split-Path $file.Path -Parent
            if ($parentDir -ne "" -and $parentDir -ne ".") {
                $destSubDir = Join-Path $DIST_DIR $parentDir
                New-Item -ItemType Directory -Path $destSubDir -Force | Out-Null
            }
            Copy-Item $src (Join-Path $DIST_DIR $file.Path)
            $fileSize = [math]::Round((Get-Item $src).Length / 1KB, 1)
            Write-Host "        [+] $($file.Path) ($fileSize KB) - $($file.Desc)" -ForegroundColor Gray
            $copiedFiles++
        } else {
            Write-Host "        [!] Introuvable: $($file.Path)" -ForegroundColor Yellow
        }
    }

    $readmePath = Join-Path $DIST_DIR "README.txt"
    @"
$APP_NAME v$VERSION - $COMPANY

INSTALLATION:
  Lancer InventorAutoSave.Setup.exe (installateur graphique)
  OU copier InventorAutoSave.exe dans un dossier de votre choix.

DEMARRAGE AVEC WINDOWS:
  Utiliser la fenetre de configuration > Demarrer avec Windows.

UTILISATION:
  L'application s'installe dans la barre de notification Windows (systray).
  Clic droit sur l'icone pour acceder aux options.
  Double-clic pour ouvrir la fenetre de configuration.

MODES DE SAUVEGARDE:
  - Document actif (recommande): Sauvegarde le doc actif + composants modifies.
  - Tous les documents: Sauvegarde tous les docs ouverts. Plus lent.

INTERVALLES: 10s, 30s, 1min, 2min, 5min, 10min, 15min, 30min
REPERTOIRE: %APPDATA%\XNRGY\InventorAutoSave\
CONTACT: mohammedamine.elgalai@xnrgy.com
"@ | Set-Content $readmePath -Encoding UTF8
    $copiedFiles++
    Write-Host "        [+] README.txt" -ForegroundColor Gray

    Write-OK "Package cree: dist\${APP_NAME}_v${VERSION}\ ($copiedFiles fichiers)"

    Write-Host ""
    Write-Info "Contenu du package:"
    Get-ChildItem $DIST_DIR -Recurse | ForEach-Object {
        $rel = $_.FullName.Replace($DIST_DIR + "\", "")
        $size = if ($_.PSIsContainer) { "[dir]" } else { "$([math]::Round($_.Length/1KB, 1)) KB" }
        Write-Host "        $rel  ($size)" -ForegroundColor Gray
    }
}

# ============================================================
# ETAPE: BUILD INSTALLATEUR
# ============================================================
if ($StepSetup) {
    $CurrentStep++
    Write-Host ""
    Write-Host "  [$CurrentStep/$TotalSteps] Build Installateur (SingleFile embarque)..." -ForegroundColor Yellow

    if (Test-Path $SETUP_PROJECT) {
        $mainExeForSetup = Join-Path $PUBLISH_DIR "$APP_NAME.exe"
        if (-not (Test-Path $mainExeForSetup)) {
            Write-Fail "EXE principal introuvable: $mainExeForSetup"
            Write-Fail "Le Setup ne peut pas embarquer l'application."
        } else {
            $mainExeDate = (Get-Item $mainExeForSetup).LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
            $mainExeSize = [math]::Round((Get-Item $mainExeForSetup).Length / 1MB, 1)
            Write-Info "EXE a embarquer: $mainExeForSetup"
            Write-Info "  Date: $mainExeDate | Taille: $mainExeSize MB"

            $setupPublishDir = Join-Path $PSScriptRoot "Setup\bin\Release\net8.0-windows\win-x64"

            $setupOutput = dotnet publish $SETUP_PROJECT -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true -p:PublishDir="$setupPublishDir" --nologo 2>&1

            $setupErrors = $setupOutput | Where-Object { $_ -match ": error" }
            $setupWarnings = $setupOutput | Where-Object { $_ -match ": warning" }

            if ($setupErrors.Count -gt 0) {
                Write-Fail "Erreurs Setup (non bloquant - app principale OK):"
                Write-Host ""
                Write-Host "        +-- ERREURS SETUP ---------------------------------------------" -ForegroundColor Red
                foreach ($e in $setupErrors) {
                    Write-Host "        |  $($e.ToString().Trim())" -ForegroundColor DarkYellow
                }
                Write-Host "        +--------------------------------------------------------------" -ForegroundColor Red
            } else {
                $setupExe = Join-Path $setupPublishDir "InventorAutoSave.Setup.exe"
                if (Test-Path $setupExe) {
                    Copy-Item $setupExe $DIST_DIR -Force
                    $setupSize = [math]::Round((Get-Item $setupExe).Length / 1MB, 1)

                    if ($setupSize -lt $mainExeSize) {
                        Write-Warn "Taille installateur ($setupSize MB) < App ($mainExeSize MB)"
                        Write-Warn "Le EXE n'est peut-etre pas embarque correctement"
                    } else {
                        Write-OK "Verification OK: Installateur contient l'app embarquee"
                    }

                    Write-OK "Setup.exe copie dans dist\ ($setupSize MB)"

                    if ($setupWarnings.Count -gt 0) {
                        Write-Host ""
                        Write-Host "        +-- WARNINGS SETUP -------------------------------------------" -ForegroundColor Yellow
                        foreach ($w in $setupWarnings) {
                            if ($w -match "([^\\]+\.cs)\((\d+),\d+\):\s+warning\s+(CS\d+):\s+(.+)") {
                                Write-Host "        |  [$($Matches[3])] $($Matches[1]) (ligne $($Matches[2]))" -ForegroundColor Yellow
                                Write-Host "        |           $($Matches[4].Trim())" -ForegroundColor DarkYellow
                            } else {
                                Write-Host "        |  $($w.ToString().Trim())" -ForegroundColor Yellow
                            }
                        }
                        Write-Host "        +--------------------------------------------------------------" -ForegroundColor Yellow
                    }
                } else {
                    Write-Fail "Installateur non trouve apres publish"
                }
            }

            # Nettoyer les artefacts dupliques apres build Setup
            Remove-ObjDuplicates -Context "post-setup"
        }
    } else {
        Write-Warn "Projet Setup non trouve: $SETUP_PROJECT"
    }
}

# ============================================================
# RESUME FINAL
# ============================================================
$stopwatch.Stop()

Write-Host ""
Write-Host "  ================================================================" -ForegroundColor Cyan
Write-Host "                       BUILD COMPLETE" -ForegroundColor Cyan
Write-Host "  ================================================================" -ForegroundColor Cyan

Write-OK "Version       : v$VERSION"
Write-OK "Temps total   : $($stopwatch.Elapsed.TotalSeconds.ToString('F1'))s"
if ($buildSummary) {
    Write-Host "  Build : $buildSummary" -ForegroundColor DarkGray
}

if (Test-Path $DIST_DIR) {
    $distFiles = Get-ChildItem $DIST_DIR -Recurse -File
    $distSize = ($distFiles | Measure-Object -Property Length -Sum).Sum
    Write-OK "Taille dist   : $([math]::Round($distSize/1MB, 1)) MB"
    Write-OK "Output        : dist\${APP_NAME}_v${VERSION}\"

    Write-Host ""
    Write-Info "Fichiers livres:"
    Get-ChildItem $DIST_DIR -File | ForEach-Object {
        $fSize = [math]::Round($_.Length / 1MB, 1)
        if ($fSize -lt 0.1) {
            $fSize = "$([math]::Round($_.Length / 1KB, 1)) KB"
        } else {
            $fSize = "$fSize MB"
        }
        Write-Host "        $($_.Name)  ($fSize)" -ForegroundColor Gray
    }
}

Write-Host ""
Write-Host "  Pour lancer l'app directement:" -ForegroundColor White
Write-Host "  .\dist\${APP_NAME}_v${VERSION}\${APP_NAME}.exe" -ForegroundColor Cyan
Write-Host ""

Pop-Location
