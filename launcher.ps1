# ==============================================================================
#  LAUNCHER - Mailbox Alert Toolkit
#  Esegue setup SOLO la prima volta (o se qualcosa manca).
#  Nelle esecuzioni successive verifica e lancia direttamente il main script.
# ==============================================================================
#
#  USO:
#    .\launcher.ps1                    -> setup-if-needed + esecuzione
#    .\launcher.ps1 -ForceSetup        -> forza re-setup completo
#    .\launcher.ps1 -SkipRun           -> solo setup, non esegue il main
#    .\launcher.ps1 -Trends            -> genera report storico (analytics)
# ==============================================================================

param(
    [switch]$ForceSetup,
    [switch]$SkipRun,
    [switch]$Trends
)

$ErrorActionPreference = "Stop"
$ScriptRoot = $PSScriptRoot
if (-not $ScriptRoot) { $ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }

# --------------------------------------------------------------------------
# PERCORSI
# --------------------------------------------------------------------------
$venvPath     = Join-Path $ScriptRoot "venv"
$venvPython   = Join-Path $venvPath  "Scripts\python.exe"
$setupMarker  = Join-Path $ScriptRoot ".setup_complete"
$mainScript   = Join-Path $ScriptRoot "mailbox-alert.ps1"
$trendsScript = Join-Path $ScriptRoot "generate_trends.py"
$configFile   = Join-Path $ScriptRoot "config.json"

$requiredFolders = @(
    "reports\user",
    "reports\other",
    "history",
    "logs",
    "credentials"
)

$requiredPsModules = @(
    "ExchangeOnlineManagement",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Identity.DirectoryManagement"
)

# --------------------------------------------------------------------------
# FUNZIONI DI VERIFICA
# --------------------------------------------------------------------------
function Test-PsModule {
    param([string]$Name)
    return [bool](Get-Module -ListAvailable -Name $Name)
}

function Test-PythonInVenv {
    if (-not (Test-Path $venvPython)) { return $false }
    & $venvPython -c "import openpyxl, sqlite3" 2>$null
    return ($LASTEXITCODE -eq 0)
}

function Test-FoldersExist {
    foreach ($f in $requiredFolders) {
        if (-not (Test-Path (Join-Path $ScriptRoot $f))) { return $false }
    }
    return $true
}

function Test-SetupComplete {
    if ($ForceSetup) { return $false }
    if (-not (Test-Path $setupMarker)) { return $false }
    if (-not (Test-PythonInVenv))      { return $false }
    if (-not (Test-FoldersExist))      { return $false }
    foreach ($m in $requiredPsModules) {
        if (-not (Test-PsModule $m))   { return $false }
    }
    return $true
}

# --------------------------------------------------------------------------
# SETUP
# --------------------------------------------------------------------------
function Invoke-Setup {
    Write-Host "`n======================================================" -ForegroundColor Cyan
    Write-Host   "|   SETUP MAILBOX ALERT TOOLKIT - PRIMA INSTALLAZIONE  |" -ForegroundColor Cyan
    Write-Host   "======================================================" -ForegroundColor Cyan

    # 1) Cartelle
    Write-Host "`n[1/5] Verifica struttura cartelle..." -ForegroundColor Yellow
    foreach ($f in $requiredFolders) {
        $p = Join-Path $ScriptRoot $f
        if (-not (Test-Path $p)) {
            New-Item -ItemType Directory -Path $p -Force | Out-Null
            Write-Host "      Creata: $f" -ForegroundColor Gray
        } else {
            Write-Host "      OK:     $f" -ForegroundColor DarkGray
        }
    }

    # 2) Moduli PowerShell
    Write-Host "`n[2/5] Verifica moduli PowerShell..." -ForegroundColor Yellow
    foreach ($m in $requiredPsModules) {
        if (Test-PsModule $m) {
            Write-Host "      OK:     $m" -ForegroundColor DarkGray
        } else {
            Write-Host "      Install: $m ..." -ForegroundColor Gray
            Install-Module $m -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Host "      OK:     $m" -ForegroundColor Green
        }
    }

    # 3) Python di sistema
    Write-Host "`n[3/5] Verifica Python di sistema..." -ForegroundColor Yellow
    $py = Get-Command python -ErrorAction SilentlyContinue
    if (-not $py) {
        $py = Get-Command python3 -ErrorAction SilentlyContinue
    }
    if (-not $py) {
        throw "Python non trovato nel PATH. Installalo da https://www.python.org/ e riprova."
    }
    $pyVersion = & $py.Source --version 2>&1
    Write-Host "      OK:     $pyVersion ($($py.Source))" -ForegroundColor DarkGray

    # 4) Virtual environment
    Write-Host "`n[4/5] Verifica/creazione virtual environment..." -ForegroundColor Yellow
    if (-not (Test-Path $venvPython) -or $ForceSetup) {
        if ((Test-Path $venvPath) -and $ForceSetup) {
            Remove-Item $venvPath -Recurse -Force
        }
        Write-Host "      Creazione venv in: $venvPath" -ForegroundColor Gray
        & $py.Source -m venv $venvPath
        if ($LASTEXITCODE -ne 0) { throw "Errore nella creazione del venv." }
    }
    Write-Host "      OK:     venv presente" -ForegroundColor DarkGray

    Write-Host "      Aggiornamento pip..." -ForegroundColor Gray
    & $venvPython -m pip install --upgrade pip --quiet 2>&1 | Out-Null

    Write-Host "      Installazione openpyxl nel venv..." -ForegroundColor Gray
    & $venvPython -m pip install openpyxl --quiet 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Errore installazione openpyxl." }
    Write-Host "      OK:     openpyxl installato" -ForegroundColor Green

    # 5) Marker setup completato
    Write-Host "`n[5/5] Salvataggio marker setup..." -ForegroundColor Yellow
    @{
        completedAt   = (Get-Date -Format "o")
        toolkitVersion = "1.0"
    } | ConvertTo-Json | Set-Content -Path $setupMarker -Encoding UTF8
    Write-Host "      OK:     $setupMarker" -ForegroundColor DarkGray

    Write-Host "`n[OK] Setup completato con successo!`n" -ForegroundColor Green
}

# --------------------------------------------------------------------------
# MAIN
# --------------------------------------------------------------------------
Write-Host "`n--------------------------------------------------------"
Write-Host " Mailbox Alert Toolkit - Launcher"
Write-Host " $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')"
Write-Host "--------------------------------------------------------"

if (Test-SetupComplete) {
    Write-Host "[OK] Setup gia' completato. Avvio rapido." -ForegroundColor Green
} else {
    Invoke-Setup
}

if ($SkipRun) {
    Write-Host "Parametro -SkipRun: termino senza eseguire il main." -ForegroundColor Yellow
    exit 0
}

# Modalita' Trends (analytics on-demand) - genera e invia per mail
if ($Trends) {
    Write-Host "`n-> Generazione report TRENDS storici..." -ForegroundColor Cyan
    $historyDb  = Join-Path $ScriptRoot "history\mailbox_history.db"
    $trendsXlsx = Join-Path $ScriptRoot ("reports\Trends_$(Get-Date -Format 'yyyy-MM-dd').xlsx")
    & $venvPython $trendsScript $historyDb $trendsXlsx
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERR] Generazione trends fallita, exit code $LASTEXITCODE" -ForegroundColor Red
        exit 1
    }
    Write-Host "[OK] Report trends generato: $trendsXlsx" -ForegroundColor Green

    # Invio mail con il file in allegato
    Write-Host "`n-> Invio mail con il trends in allegato..." -ForegroundColor Cyan
    $cfg = Get-Content $configFile -Raw | ConvertFrom-Json
    $credFile = Join-Path $ScriptRoot "credentials\smtp-cred.xml"
    if (-not (Test-Path $credFile)) {
        Write-Host "[WARN] Credenziali SMTP non trovate ($credFile). Esegui prima il task giornaliero per salvarle." -ForegroundColor Yellow
        exit 0
    }
    $smtpCred = Import-Clixml $credFile

    $today   = Get-Date -Format 'yyyy-MM-dd'
    $subject = "Report Trends storici mailbox $($cfg.tenant.domain) - $today"
    $body    = @"
<html><body style="font-family:Arial,sans-serif;font-size:11pt;color:#333;">
<h2 style="color:#1F3864;">Report trends settimanale</h2>
<p>In allegato il report storico aggiornato a oggi ($today).</p>
<p>Contiene 8 fogli di analytics:</p>
<ul>
  <li>Riepilogo Tenant (serie temporale mailbox e archivio, con grafico evoluzione)</li>
  <li>Top Crescita Mailbox 30gg e 12 mesi</li>
  <li>Top Crescita Archivio 30gg e 12 mesi</li>
  <li>Caselle inattive (oltre 90gg)</li>
  <li>Storico Mailbox per Casella (pivot)</li>
  <li>Storico Archivio per Casella (pivot)</li>
</ul>
<p style="color:#777;font-size:9pt;">Mail automatica - mailbox-alert-toolkit (modalita' Trends)</p>
</body></html>
"@

    try {
        Send-MailMessage `
            -From       $cfg.mail.from `
            -To         $cfg.mail.to `
            -Subject    $subject `
            -Body       $body `
            -BodyAsHtml `
            -SmtpServer $cfg.mail.smtpServer `
            -Port       $cfg.mail.smtpPort `
            -UseSsl `
            -Credential $smtpCred `
            -Attachments $trendsXlsx `
            -Encoding   UTF8
        Write-Host "[OK] Mail trends inviata a $($cfg.mail.to)" -ForegroundColor Green
    } catch {
        Write-Host "[ERR] Errore invio mail trends: $_" -ForegroundColor Red
        exit 1
    }

    exit 0
}

# Esecuzione normale: lancio del main script
Write-Host "`n-> Esecuzione del main script..." -ForegroundColor Cyan
& $mainScript -VenvPython $venvPython -ConfigFile $configFile -ScriptRoot $ScriptRoot
