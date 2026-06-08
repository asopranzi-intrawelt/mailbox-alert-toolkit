# ==============================================================================
#  MAILBOX ALERT - MAIN SCRIPT
#  Chiamato da launcher.ps1 (setup garantito prima dell'esecuzione).
#  Non eseguire direttamente: usa il launcher.
# ==============================================================================

param(
    [Parameter(Mandatory=$true)] [string]$VenvPython,
    [Parameter(Mandatory=$true)] [string]$ConfigFile,
    [Parameter(Mandatory=$true)] [string]$ScriptRoot
)

$ErrorActionPreference = "Stop"

# --------------------------------------------------------------------------
# CARICAMENTO CONFIG
# --------------------------------------------------------------------------
$cfg = Get-Content $ConfigFile -Raw | ConvertFrom-Json

$soglia           = $cfg.thresholds.warningPercent
$sogliaCrit       = $cfg.thresholds.criticalPercent
$adminMail        = $cfg.mail.to
$mittente         = $cfg.mail.from
$smtpServer       = $cfg.mail.smtpServer
$smtpPort         = $cfg.mail.smtpPort
$tenantUpn        = $cfg.tenant.adminUpn
$maxFiles         = $cfg.retention.maxFilesPerType
$includeFolders   = $cfg.collection.includeFolderStats
$includeFwd       = $cfg.collection.includeForwardingInfo
$includeHold      = $cfg.collection.includeHoldInfo

# Notifiche utente (con default sicuri se la sezione manca nel config)
$notifyUsers      = $true
$skipResource     = $true
$itSupport        = $adminMail
if ($cfg.notifications) {
    if ($null -ne $cfg.notifications.notifyMailboxOwners)   { $notifyUsers  = $cfg.notifications.notifyMailboxOwners }
    if ($null -ne $cfg.notifications.skipResourceMailboxes) { $skipResource = $cfg.notifications.skipResourceMailboxes }
    if ($cfg.notifications.itSupportAddress)                { $itSupport    = $cfg.notifications.itSupportAddress }
}

# --------------------------------------------------------------------------
# HELPER: parsing dimensioni Exchange (compatibile con oggetti REST/Deserialized)
# --------------------------------------------------------------------------
# Il modulo ExchangeOnlineManagement moderno restituisce oggetti deserializzati
# (Deserialized.Microsoft.Exchange.Data.ByteQuantifiedSize) che NON hanno piu'
# il metodo .ToBytes(). I valori arrivano come stringhe del tipo:
#   "49.5 GB (53,150,121,984 bytes)"
#   "Unlimited"
# Questa funzione estrae il valore in byte da qualsiasi forma.
function Get-SizeBytes {
    param($val)
    if ($null -eq $val) { return [int64]0 }
    $s = "$val"
    if ([string]::IsNullOrWhiteSpace($s)) { return [int64]0 }
    if ($s -match "(?i)unlimited")        { return [int64]0 }
    # Forma tipica: "49.5 GB (53,150,121,984 bytes)"
    if ($s -match "\(([\d.,]+)\s*bytes\)") {
        $digits = $matches[1] -replace '[^\d]', ''
        if ($digits) { return [int64]$digits }
    }
    # Forma alternativa senza parentesi: "49.5 GB"
    if ($s -match "^([\d.,]+)\s*([KMGT]?B)\b") {
        $num = $matches[1] -replace ',', ''
        try {
            $v = [double]::Parse($num, [System.Globalization.CultureInfo]::InvariantCulture)
            switch ($matches[2]) {
                "B"  { return [int64]$v }
                "KB" { return [int64]($v * 1KB) }
                "MB" { return [int64]($v * 1MB) }
                "GB" { return [int64]($v * 1GB) }
                "TB" { return [int64]($v * 1TB) }
            }
        } catch { }
    }
    return [int64]0
}

# --------------------------------------------------------------------------
# PERCORSI
# --------------------------------------------------------------------------
$data      = Get-Date -Format "yyyy-MM-dd"
$dataLabel = Get-Date -Format "dd/MM/yyyy HH:mm"

$reportDirUser  = Join-Path $ScriptRoot "reports\user"
$reportDirOther = Join-Path $ScriptRoot "reports\other"
$historyDir     = Join-Path $ScriptRoot "history"
$logsDir        = Join-Path $ScriptRoot "logs"
$credDir        = Join-Path $ScriptRoot "credentials"

$jsonPath    = Join-Path $historyDir   "mailbox_data_$data.json"
$historyDb   = Join-Path $historyDir   "mailbox_history.db"
$reportUser  = Join-Path $reportDirUser  "Report_UserMailbox_$data.xlsx"
$reportOther = Join-Path $reportDirOther "Report_AltreMailbox_$data.xlsx"
$logFile     = Join-Path $logsDir       "run_$data.log"
$credFile    = Join-Path $credDir       "smtp-cred.xml"

$pyScript    = Join-Path $ScriptRoot "generate_mailbox_report.py"

# Logging su file + console
Start-Transcript -Path $logFile -Append -ErrorAction SilentlyContinue | Out-Null

# --------------------------------------------------------------------------
# CONNESSIONI
# --------------------------------------------------------------------------
# Determina se usare l'autenticazione via app registration (certificato).
# Modalita' raccomandata per task schedulati: nessuna interazione browser,
# nessuna scadenza di refresh token. Richiede setup lato Azure AD una volta sola.
$useAppAuth = $false
if ($cfg.appRegistration -and $cfg.appRegistration.enabled -eq $true `
    -and $cfg.appRegistration.clientId `
    -and $cfg.appRegistration.tenantId `
    -and $cfg.appRegistration.certificateThumbprint) {
    $useAppAuth = $true
}

# Importante: Graph va caricato PRIMA di ExchangeOnlineManagement.
# ExchangeOnlineManagement carica versioni vecchie di Azure.Identity /
# Microsoft.Identity.Client che mandano in errore Microsoft.Graph se
# vengono caricate dopo (TypeLoadException su GetTokenAsync).
Write-Host "`n-> Connessione a Microsoft Graph..." -ForegroundColor Cyan
Import-Module Microsoft.Graph.Authentication            -ErrorAction Stop
Import-Module Microsoft.Graph.Users                     -ErrorAction Stop
Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction SilentlyContinue

if ($useAppAuth) {
    Write-Host "    Modalita' app registration (certificato) - unattended" -ForegroundColor DarkGray
    Write-Host "    ClientId:   $($cfg.appRegistration.clientId)" -ForegroundColor DarkGray
    Write-Host "    TenantId:   $($cfg.appRegistration.tenantId)" -ForegroundColor DarkGray
    Write-Host "    Thumbprint: $($cfg.appRegistration.certificateThumbprint)" -ForegroundColor DarkGray
    Connect-MgGraph -ClientId             $cfg.appRegistration.clientId `
                    -TenantId             $cfg.appRegistration.tenantId `
                    -CertificateThumbprint $cfg.appRegistration.certificateThumbprint `
                    -NoWelcome
} else {
    Write-Host "    Modalita' interattiva (login browser)" -ForegroundColor DarkGray
    Connect-MgGraph -Scopes "User.Read.All","Organization.Read.All" -NoWelcome
}

Write-Host "-> Connessione a Exchange Online..." -ForegroundColor Cyan
Import-Module ExchangeOnlineManagement -ErrorAction Stop

if ($useAppAuth) {
    $orgName = $cfg.appRegistration.organization
    if (-not $orgName) { $orgName = $cfg.tenant.domain }
    Write-Host "    Modalita' app registration (certificato) - unattended" -ForegroundColor DarkGray
    Write-Host "    Organization: $orgName" -ForegroundColor DarkGray
    Connect-ExchangeOnline -AppId                $cfg.appRegistration.clientId `
                           -CertificateThumbprint $cfg.appRegistration.certificateThumbprint `
                           -Organization          $orgName `
                           -ShowBanner:$false
} else {
    Write-Host "    Modalita' interattiva (login browser)" -ForegroundColor DarkGray
    Connect-ExchangeOnline -UserPrincipalName $tenantUpn -ShowBanner:$false
}

# --------------------------------------------------------------------------
# MAPPATURA LICENZE
# --------------------------------------------------------------------------
Write-Host "`n[1/5] Caricamento mappa licenze..." -ForegroundColor Cyan

# Mappa skuPartNumber -> nome leggibile.
# Gli skuPartNumber sono codici testuali stabili (es. "ENTERPRISEPACK", "SPB")
# che Microsoft non cambia mai, a differenza dei GUID che cambiano per ogni tenant.
# Lista basata sul reference Microsoft "Product names and service plan identifiers".
$skuPartNumberNames = @{
    # Microsoft 365 / Office 365 - Business
    "SPB"                          = "Microsoft 365 Business Premium"
    "O365_BUSINESS_PREMIUM"        = "Microsoft 365 Business Standard"
    "O365_BUSINESS_ESSENTIALS"     = "Microsoft 365 Business Basic"
    "O365_BUSINESS"                = "Microsoft 365 Apps for business"
    "SMB_BUSINESS_PREMIUM"         = "Microsoft 365 Business Premium (legacy)"
    "SMB_BUSINESS_ESSENTIALS"      = "Microsoft 365 Business Basic (legacy)"
    "SMB_BUSINESS"                 = "Microsoft 365 Apps for business (legacy)"

    # Microsoft 365 / Office 365 - Enterprise
    "STANDARDPACK"                 = "Office 365 E1"
    "ENTERPRISEPACK"               = "Office 365 E3"
    "ENTERPRISEPREMIUM"            = "Office 365 E5"
    "ENTERPRISEPREMIUM_NOPSTNCONF" = "Office 365 E5 (no PSTN)"
    "ENTERPRISEPACKWITHOUTPROPLUS" = "Office 365 E3 (no Office)"
    "SPE_E3"                       = "Microsoft 365 E3"
    "SPE_E5"                       = "Microsoft 365 E5"
    "SPE_F1"                       = "Microsoft 365 F1"
    "SPE_F3"                       = "Microsoft 365 F3"
    "DESKLESSPACK"                 = "Office 365 F1 (legacy)"
    "OFFICESUBSCRIPTION"           = "Microsoft 365 Apps for enterprise"

    # Exchange Online
    "EXCHANGESTANDARD"             = "Exchange Online (Piano 1)"
    "EXCHANGEENTERPRISE"           = "Exchange Online (Piano 2)"
    "EXCHANGEDESKLESS"             = "Exchange Online Kiosk"
    "EXCHANGE_S_ARCHIVE"           = "Exchange Online Archiving"
    "EXCHANGE_S_FOUNDATION"        = "Exchange Foundation"

    # Visio
    "VISIOCLIENT"                  = "Visio Plan 2"
    "VISIOONLINE_PLAN1"            = "Visio Plan 1"
    "VISIO_PLAN1_DEPT"             = "Visio Plan 1"
    "VISIOENTERPRISE"              = "Visio Plan 2 (Enterprise)"

    # Project
    "PROJECTPROFESSIONAL"          = "Project Plan 3"
    "PROJECTPREMIUM"               = "Project Plan 5"
    "PROJECT_P1"                   = "Project Plan 1"
    "PROJECTESSENTIALS"            = "Project Online Essentials"

    # Power Platform
    "POWER_BI_STANDARD"            = "Power BI (free)"
    "POWER_BI_PRO"                 = "Power BI Pro"
    "POWER_BI_PRO_DEPT"            = "Power BI Pro"
    "PBI_PREMIUM_PER_USER"         = "Power BI Premium per User"
    "FLOW_FREE"                    = "Power Automate Free"
    "FLOW_PER_USER"                = "Power Automate per User"
    "POWERAUTOMATE_ATTENDED_RPA"   = "Power Automate per User con RPA"
    "POWERAPPS_VIRAL"              = "Power Apps Plan 2 Trial"
    "POWERAPPS_PER_USER"           = "Power Apps per User Plan"

    # Teams
    "TEAMS_EXPLORATORY"            = "Microsoft Teams Exploratory"
    "TEAMS_FREE"                   = "Microsoft Teams Free"
    "MCO_TEAMS_IW"                 = "Microsoft Teams (Trial)"
    "MEETING_ROOM"                 = "Microsoft Teams Rooms Standard"
    "Microsoft_Teams_Rooms_Pro"    = "Microsoft Teams Rooms Pro"

    # Security / Mobility
    "INTUNE_A"                     = "Intune"
    "EMS"                          = "Enterprise Mobility + Security E3"
    "EMSPREMIUM"                   = "Enterprise Mobility + Security E5"
    "AAD_PREMIUM"                  = "Azure AD Premium P1"
    "AAD_PREMIUM_P2"               = "Azure AD Premium P2"
    "AAD_BASIC"                    = "Azure AD Basic"
    "ATP_ENTERPRISE"               = "Defender for Office 365 Plan 1"
    "THREAT_INTELLIGENCE"          = "Defender for Office 365 Plan 2"
    "WIN_DEF_ATP"                  = "Defender for Endpoint Plan 2"
    "WINDEFATP"                    = "Defender for Endpoint Plan 1"

    # Windows
    "WIN10_PRO_ENT_SUB"            = "Windows 10/11 Enterprise E3"
    "WIN10_VDA_E5"                 = "Windows 10/11 Enterprise E5"

    # SharePoint / OneDrive
    "SHAREPOINTSTANDARD"           = "SharePoint Online (Piano 1)"
    "SHAREPOINTENTERPRISE"         = "SharePoint Online (Piano 2)"
    "SHAREPOINTSTORAGE"            = "SharePoint Storage Quota"
    "WACONEDRIVESTANDARD"          = "OneDrive for Business (Piano 1)"
    "WACONEDRIVEENTERPRISE"        = "OneDrive for Business (Piano 2)"

    # Copilot e altri
    "Microsoft_365_Copilot"        = "Copilot for Microsoft 365"
    "M365_COPILOT_BUSINESS"        = "Copilot Business"
    "VIVA"                         = "Microsoft Viva Suite"
    "STREAM"                       = "Microsoft Stream"
    "FORMS_PLAN_E5"                = "Microsoft Forms (Plan E5)"
    "DYN365_ENTERPRISE_PLAN1"      = "Dynamics 365 Customer Engagement Plan"
    "CRMSTANDARD"                  = "Dynamics 365 Customer Service Pro"
}

# Costruisco la mappa SkuId(GUID) -> nome friendly leggendo gli SKU
# effettivamente acquistati dal tenant. Richiede permesso Organization.Read.All.
$skuMap = @{}
try {
    $subscribedSkus = Get-MgSubscribedSku -All -ErrorAction Stop
    foreach ($s in $subscribedSkus) {
        $partNumber = $s.SkuPartNumber
        $guid       = $s.SkuId.ToString().ToLower()
        $friendly   = if ($skuPartNumberNames.ContainsKey($partNumber)) {
            $skuPartNumberNames[$partNumber]
        } else {
            $partNumber   # fallback: usa il codice testuale, che e' comunque leggibile
        }
        $skuMap[$guid] = $friendly
    }
    Write-Host "    SKU del tenant identificati: $($skuMap.Count)" -ForegroundColor Green
    $sconosciuti = $subscribedSkus | Where-Object { -not $skuPartNumberNames.ContainsKey($_.SkuPartNumber) }
    if ($sconosciuti) {
        Write-Host "    Nota: SKU senza nome friendly nella mappa interna (mostrati come SkuPartNumber):" -ForegroundColor DarkGray
        foreach ($sk in $sconosciuti) {
            Write-Host "      - $($sk.SkuPartNumber)" -ForegroundColor DarkGray
        }
    }
} catch {
    Write-Warning "Impossibile leggere SubscribedSku (serve Organization.Read.All): $_"
    Write-Warning "Continuo con i GUID; per nomi leggibili rilanciare con consenso al nuovo scope."
}

# Mappa UPN utente -> stringa di licenze leggibili
$licenseMap = @{}
try {
    $mgUsers = Get-MgUser -All `
        -Property "UserPrincipalName,AssignedLicenses" `
        -Filter  "assignedLicenses/`$count ne 0" `
        -ConsistencyLevel eventual -CountVariable c
    foreach ($u in $mgUsers) {
        $names = @()
        foreach ($lic in $u.AssignedLicenses) {
            $id = $lic.SkuId.ToString().ToLower()
            $names += if ($skuMap.ContainsKey($id)) { $skuMap[$id] } else { $id }
        }
        $licenseMap[$u.UserPrincipalName.ToLower()] = ($names -join ", ")
    }
    Write-Host "    Mappati $($licenseMap.Count) utenti con licenza." -ForegroundColor Green
} catch {
    Write-Warning "Impossibile recuperare le licenze: $_"
}

# --------------------------------------------------------------------------
# RACCOLTA DATI MAILBOX
# --------------------------------------------------------------------------
Write-Host "`n[2/5] Raccolta dati mailbox..." -ForegroundColor Cyan

$allTypes  = @("UserMailbox","SharedMailbox","RoomMailbox","EquipmentMailbox","SchedulingMailbox")
$mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails $allTypes

$allData    = [System.Collections.Generic.List[hashtable]]::new()
$alerts     = [System.Collections.Generic.List[hashtable]]::new()   # mailbox principale sopra soglia
$archAlerts = [System.Collections.Generic.List[hashtable]]::new()   # archivio online sopra soglia
$total   = $mailboxes.Count
$i       = 0

foreach ($mb in $mailboxes) {
    $i++
    Write-Progress -Activity "Analisi mailbox ($i/$total)" `
                   -Status $mb.PrimarySmtpAddress `
                   -PercentComplete (($i / $total) * 100)

    try {
        # --- Mailbox stats ---
        $stats    = Get-MailboxStatistics -Identity $mb.Identity -ErrorAction Stop
        $usedBytes = Get-SizeBytes $stats.TotalItemSize
        $deletedBytes = if ($stats.TotalDeletedItemSize) { Get-SizeBytes $stats.TotalDeletedItemSize } else { 0 }

        $limitBytes = Get-SizeBytes $mb.ProhibitSendQuota

        $usedGB     = [math]::Round($usedBytes    / 1GB, 2)
        $deletedGB  = [math]::Round($deletedBytes / 1GB, 2)
        $limitGB    = if ($limitBytes -gt 0) { [math]::Round($limitBytes / 1GB, 2) } else { $null }
        $pct        = if ($limitBytes -gt 0) { [math]::Round(($usedBytes / $limitBytes) * 100, 1) } else { $null }

        # --- Folder stats (Inbox / Sent / Deleted) ---
        $inboxGB = $null; $sentGB = $null
        if ($includeFolders) {
            try {
                $folders = Get-MailboxFolderStatistics -Identity $mb.Identity `
                    -FolderScope Inbox -ErrorAction SilentlyContinue
                $inbox = $folders | Where-Object { $_.FolderType -eq "Inbox" } | Select-Object -First 1
                if ($inbox -and $inbox.FolderAndSubfolderSize) {
                    $inboxGB = [math]::Round((Get-SizeBytes $inbox.FolderAndSubfolderSize) / 1GB, 2)
                }

                $foldersSent = Get-MailboxFolderStatistics -Identity $mb.Identity `
                    -FolderScope SentItems -ErrorAction SilentlyContinue
                $sent = $foldersSent | Where-Object { $_.FolderType -eq "SentItems" } | Select-Object -First 1
                if ($sent -and $sent.FolderAndSubfolderSize) {
                    $sentGB = [math]::Round((Get-SizeBytes $sent.FolderAndSubfolderSize) / 1GB, 2)
                }
            } catch { }
        }

        # --- Archive stats ---
        $archEnabled = "No"; $archUsedGB = $null; $archQuotaGB = $null; $archPct = $null
        if ($mb.ArchiveStatus -eq "Active") {
            try {
                $archStats = Get-MailboxStatistics -Identity $mb.Identity -Archive -ErrorAction Stop
                $archBytes = Get-SizeBytes $archStats.TotalItemSize
                $archLimitBytes = Get-SizeBytes $mb.ArchiveQuota
                $archEnabled  = "Si'"
                $archUsedGB   = [math]::Round($archBytes / 1GB, 2)
                $archQuotaGB  = if ($archLimitBytes -gt 0) { [math]::Round($archLimitBytes / 1GB, 2) } else { $null }
                $archPct      = if ($archLimitBytes -gt 0) { [math]::Round(($archBytes / $archLimitBytes) * 100, 1) } else { $null }
            } catch {
                $archEnabled = "Si' (errore)"
            }
        }

        # --- Forwarding ---
        $fwdActive = "No"
        if ($includeFwd) {
            if ($mb.ForwardingAddress -or $mb.ForwardingSmtpAddress) {
                $fwdActive = if ($mb.DeliverToMailboxAndForward) { "Si' (copia)" } else { "Si'" }
            }
        }

        # --- Hold / GAL ---
        $litHold = if ($includeHold -and $mb.LitigationHoldEnabled) { "Si'" } else { "No" }
        $hiddenGAL = if ($mb.HiddenFromAddressListsEnabled) { "Si'" } else { "No" }

        # --- Last access (prefer LastUserAccessTime, fallback LastLogonTime) ---
        $lastAccess = $null
        if ($stats.LastUserAccessTime)   { $lastAccess = $stats.LastUserAccessTime }
        elseif ($stats.LastLogonTime)    { $lastAccess = $stats.LastLogonTime }

        $lastAccessStr = if ($lastAccess) { $lastAccess.ToString("dd/MM/yyyy HH:mm") } else { "Mai" }
        $daysInactive  = if ($lastAccess) { [math]::Round(((Get-Date) - $lastAccess).TotalDays, 0) } else { $null }

        # --- License ---
        $upnKey = $mb.PrimarySmtpAddress.ToLower()
        $lic    = if ($licenseMap.ContainsKey($upnKey)) { $licenseMap[$upnKey] } else { "" }

        # --- Tipo ---
        $tipo = $mb.RecipientTypeDetails.ToString()

        # --- Created date ---
        $created = if ($mb.WhenCreated) { $mb.WhenCreated.ToString("dd/MM/yyyy") } else { "" }

        $record = @{
            Date              = $data
            Casella           = $mb.PrimarySmtpAddress.ToString()
            DisplayName       = $mb.DisplayName
            Tipo              = $tipo
            Licenza           = $lic
            CreataIl          = $created
            UsatoGB           = $usedGB
            QuotaGB           = $limitGB
            PercMailbox       = $pct
            NumeroEmail       = $stats.ItemCount
            EliminatiGB       = $deletedGB
            InboxGB           = $inboxGB
            InviatiGB         = $sentGB
            ArchivioAbilitato = $archEnabled
            ArchivioUsatoGB   = $archUsedGB
            ArchivioQuotaGB   = $archQuotaGB
            PercArchivio      = $archPct
            UltimoAccesso     = $lastAccessStr
            GiorniInattivita  = $daysInactive
            InoltroAttivo     = $fwdActive
            LitigationHold    = $litHold
            NascostaGAL       = $hiddenGAL
        }
        $allData.Add($record)

        # Warning sulla mailbox principale: blocca ricezione, impatto immediato
        if ($null -ne $pct -and $pct -ge $soglia) {
            $alerts.Add($record)
            Write-Host ("  [M]  {0,-45} {1,5}%   {2} GB / {3} GB" -f `
                $mb.PrimarySmtpAddress, $pct, $usedGB, $limitGB) -ForegroundColor Yellow
        }

        # Warning sull'archivio online: problema separato, non blocca ricezione
        if ($null -ne $archPct -and $archPct -ge $soglia) {
            $archAlerts.Add($record)
            Write-Host ("  [A]  {0,-45} {1,5}%   {2} GB / {3} GB  (archivio)" -f `
                $mb.PrimarySmtpAddress, $archPct, $archUsedGB, $archQuotaGB) -ForegroundColor Magenta
        }

    } catch {
        Write-Host "  [X] Errore su $($mb.PrimarySmtpAddress): $_" -ForegroundColor Red
    }
}
Write-Progress -Activity "Analisi mailbox" -Completed

Write-Host ("`n    Elaborate {0} caselle.  Mailbox in warning: {1}.  Archivi in warning: {2}." -f `
    $allData.Count, $alerts.Count, $archAlerts.Count) -ForegroundColor Green

# --------------------------------------------------------------------------
# GENERAZIONE EXCEL + SCRITTURA STORICO SQLITE (via Python nel venv)
# --------------------------------------------------------------------------
Write-Host "`n[3/5] Generazione report Excel e aggiornamento storico..." -ForegroundColor Cyan

$allData | ConvertTo-Json -Depth 4 | Set-Content -Path $jsonPath -Encoding UTF8

$pyOut = & $VenvPython $pyScript $jsonPath $reportUser $reportOther $historyDb 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "    Errore generazione report: $pyOut" -ForegroundColor Red
} else {
    Write-Host "    $pyOut" -ForegroundColor Green
}

Remove-Item $jsonPath -ErrorAction SilentlyContinue

# --------------------------------------------------------------------------
# RETENTION (max N file per cartella, sui piu' recenti)
# --------------------------------------------------------------------------
Write-Host "`n[4/5] Applicazione retention ($maxFiles file per tipo)..." -ForegroundColor Cyan

function Invoke-Retention {
    param([string]$Folder, [int]$Max)
    $files = Get-ChildItem $Folder -Filter "*.xlsx" -File |
             Sort-Object LastWriteTime -Descending
    if ($files.Count -gt $Max) {
        $toDelete = $files | Select-Object -Skip $Max
        foreach ($f in $toDelete) {
            Remove-Item $f.FullName -Force
            Write-Host "    Rimosso: $($f.Name)" -ForegroundColor DarkGray
        }
        Write-Host "    Eliminati $($toDelete.Count) file in $Folder" -ForegroundColor Yellow
    } else {
        Write-Host "    $Folder : $($files.Count)/$Max file (nessuna pulizia)" -ForegroundColor DarkGray
    }
}

Invoke-Retention -Folder $reportDirUser  -Max $maxFiles
Invoke-Retention -Folder $reportDirOther -Max $maxFiles

# Pulizia log oltre 90 giorni
Get-ChildItem $logsDir -Filter "run_*.log" -File |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-90) } |
    Remove-Item -Force -ErrorAction SilentlyContinue

# --------------------------------------------------------------------------
# INVIO MAIL
# --------------------------------------------------------------------------
Write-Host "`n[5/5] Invio mail di report..." -ForegroundColor Cyan

if (Test-Path $credFile) {
    $cred = Import-Clixml $credFile
} else {
    Write-Host "    Prima esecuzione: salvataggio credenziali SMTP..." -ForegroundColor Yellow
    $cred = Get-Credential -Message "Inserisci le credenziali SMTP per $mittente"
    $cred | Export-Clixml $credFile
}

# --- Tabella 1: caselle con MAILBOX PRINCIPALE sopra soglia ---
$righeMb = ""
foreach ($a in $alerts) {
    $bg = if ($a.PercMailbox -ge $sogliaCrit) { "#FFCCCC" } else { "#FFF3CD" }
    $righeMb += "<tr style='background:$bg'>
        <td>$($a.Casella)</td>
        <td>$($a.Tipo)</td>
        <td>$($a.UsatoGB) GB</td>
        <td>$($a.QuotaGB) GB</td>
        <td><b>$($a.PercMailbox)%</b></td>
    </tr>"
}
$tabellaMb = if ($alerts.Count -gt 0) { @"
<h3 style='color:#c0392b;margin-top:18px'>Caselle con mailbox principale sopra il $soglia%</h3>
<p style='font-size:12px;color:#555;margin-top:0'>Questo tipo di warning impedisce, oltre il limite, la ricezione di nuovi messaggi.</p>
<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;font-size:13px;width:100%'>
  <thead style='background:#2E75B6;color:white'>
    <tr><th>Casella</th><th>Tipo</th><th>Usato</th><th>Quota</th><th>% Occupato</th></tr>
  </thead><tbody>$righeMb</tbody>
</table>
"@ } else {
"<p style='color:#155724;background:#d4edda;padding:10px;border-radius:4px;margin-top:18px'>Nessuna mailbox principale sopra il $soglia%.</p>"
}

# --- Tabella 2: ARCHIVI ONLINE sopra soglia ---
$righeArc = ""
foreach ($a in $archAlerts) {
    $bg = if ($a.PercArchivio -ge $sogliaCrit) { "#FFCCCC" } else { "#FFF3CD" }
    $righeArc += "<tr style='background:$bg'>
        <td>$($a.Casella)</td>
        <td>$($a.Tipo)</td>
        <td>$($a.ArchivioUsatoGB) GB</td>
        <td>$($a.ArchivioQuotaGB) GB</td>
        <td><b>$($a.PercArchivio)%</b></td>
    </tr>"
}
$tabellaArc = if ($archAlerts.Count -gt 0) { @"
<h3 style='color:#8e44ad;margin-top:18px'>Archivi online sopra il $soglia%</h3>
<p style='font-size:12px;color:#555;margin-top:0'>Questo tipo di warning riguarda l'archivio online (cassetta di archivio Exchange). Non blocca la ricezione di nuovi messaggi sulla mailbox principale, ma impedisce all'utente di archiviare ulteriori contenuti e blocca l'auto-archiviazione se configurata.</p>
<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;font-size:13px;width:100%'>
  <thead style='background:#7B4F9F;color:white'>
    <tr><th>Casella</th><th>Tipo</th><th>Archivio Usato</th><th>Archivio Quota</th><th>% Archivio</th></tr>
  </thead><tbody>$righeArc</tbody>
</table>
"@ } else {
"<p style='color:#155724;background:#d4edda;padding:10px;border-radius:4px;margin-top:18px'>Nessun archivio online sopra il $soglia%.</p>"
}

$corpo = @"
<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#212529">
<h2 style="color:#1F3864;border-bottom:2px solid #2E75B6;padding-bottom:6px">
  Report Spazio Mailbox - $($cfg.tenant.domain)
</h2>
<p>Data: <b>$dataLabel</b> | Caselle analizzate: <b>$($allData.Count)</b><br>
Mailbox principali in warning: <b>$($alerts.Count)</b> | Archivi in warning: <b>$($archAlerts.Count)</b></p>
$tabellaMb
$tabellaArc
<p style="margin-top:24px">In allegato i due report Excel del giorno.<br>
Per analisi storiche multi-anno usa: <code>.\launcher.ps1 -Trends</code></p>
<p style="color:#6c757d;font-size:11px;border-top:1px solid #dee2e6;padding-top:8px;margin-top:20px">
Generato automaticamente | Mailbox Alert Toolkit
</p></body></html>
"@

$allegati = @()
if (Test-Path $reportUser)  { $allegati += $reportUser }
if (Test-Path $reportOther) { $allegati += $reportOther }

$oggetto = if (($alerts.Count + $archAlerts.Count) -gt 0) {
    "[Mailbox: $($alerts.Count) | Archivio: $($archAlerts.Count)] Report Mailbox $($cfg.tenant.domain) - $data"
} else {
    "[OK] Report Mailbox $($cfg.tenant.domain) - $data"
}

$mailParams = @{
    To = $adminMail; From = $mittente; Subject = $oggetto
    Body = $corpo; BodyAsHtml = $true
    SmtpServer = $smtpServer; Port = $smtpPort; UseSsl = $true
    Credential = $cred
    Encoding = ([System.Text.Encoding]::UTF8)
}
if ($allegati.Count -gt 0) { $mailParams.Attachments = $allegati }

Send-MailMessage @mailParams
Write-Host "    Mail IT inviata a $adminMail con $($allegati.Count) allegati" -ForegroundColor Green

# --------------------------------------------------------------------------
# NOTIFICHE PERSONALI ALLE CASELLE IN WARNING/CRITICAL
# --------------------------------------------------------------------------
if ($notifyUsers) {
    Write-Host "`n[6/6] Invio notifiche personali alle caselle in warning..." -ForegroundColor Cyan

    $resourceTypes = @("RoomMailbox","EquipmentMailbox","SchedulingMailbox")
    $sentMb        = 0
    $sentArc       = 0
    $skipResMb     = 0
    $skipResArc    = 0
    $errCount      = 0

    # Funzione helper: invia la mail di notifica con i parametri standardizzati
    function Send-UserNotification {
        param($To, $Subject, $Body)
        Send-MailMessage `
            -To         $To `
            -From       $script:mittente `
            -Subject    $Subject `
            -Body       $Body `
            -BodyAsHtml `
            -SmtpServer $script:smtpServer `
            -Port       $script:smtpPort `
            -UseSsl `
            -Credential $script:cred `
            -Encoding   ([System.Text.Encoding]::UTF8)
    }

    # ------------------------------------------------------------------
    # LOOP 1: NOTIFICHE MAILBOX PRINCIPALE
    # Una mail per ogni casella in $alerts che NON sia una risorsa.
    # ------------------------------------------------------------------
    foreach ($r in $alerts) {
        if ($skipResource -and ($resourceTypes -contains $r.Tipo)) {
            $skipResMb++
            continue
        }

        $pct = [double]$r.PercMailbox
        $isCritical = $pct -ge $sogliaCrit

        if ($isCritical) {
            $subj    = "URGENTE - La tua casella di posta e' quasi piena"
            $color   = "#c0392b"
            $title   = "La tua casella di posta e' quasi piena"
            $urgency = "Intervieni subito per evitare di non poter piu' ricevere nuovi messaggi: oltre il limite la casella smette di ricevere posta."
            $bg      = "#FADBD8"
        } else {
            $subj    = "Avviso - La tua casella di posta sta raggiungendo il limite"
            $color   = "#b9770e"
            $title   = "La tua casella di posta sta raggiungendo il limite"
            $urgency = "Ti suggeriamo di iniziare a fare pulizia prima di arrivare al limite, oltre il quale la casella smette di ricevere nuovi messaggi."
            $bg      = "#FCF3CF"
        }

        $body = @"
<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#212529;max-width:640px;margin:0 auto">
<div style="background:$bg;border-left:4px solid $color;padding:14px 18px;margin-bottom:18px">
<h2 style="color:$color;margin:0 0 8px 0">$title</h2>
<p style="margin:0">$urgency</p>
</div>
<p>Ciao $($r.DisplayName),</p>
<p>La tua casella di posta <strong>$($r.Casella)</strong> ha raggiunto il <strong>$($r.PercMailbox)% di occupazione</strong> ($($r.UsatoGB) GB usati su $($r.QuotaGB) GB di quota).</p>
<p style="font-size:13px;color:#555">Questa notifica riguarda esclusivamente la <strong>casella di posta principale</strong>. Se ricevi anche una notifica separata sull'archivio online, sono problemi distinti e vanno gestiti separatamente.</p>
<h3 style="color:#1F3864;margin-top:22px">Cosa puoi fare</h3>
<ul>
  <li>elimina i messaggi vecchi che non ti servono piu', in particolare quelli con allegati pesanti</li>
  <li>svuota la cartella posta eliminata (e svuota anche elementi recuperabili se vuoi liberare subito spazio)</li>
  <li>se hai l'archivio online attivo (e non e' anch'esso pieno), sposta i messaggi storici nell'archivio</li>
  <li>per qualsiasi dubbio o richiesta di aiuto, scrivi a <a href="mailto:$itSupport" style="color:#2E75B6">$itSupport</a></li>
</ul>
<p style="color:#6c757d;font-size:11px;border-top:1px solid #dee2e6;padding-top:10px;margin-top:24px">
Messaggio automatico inviato dal sistema di monitoraggio delle caselle di posta del tenant $($cfg.tenant.domain).<br>
Generato il $dataLabel. Per assistenza: $itSupport
</p>
</body></html>
"@

        try {
            Send-UserNotification -To $r.Casella -Subject $subj -Body $body
            $sentMb++
            Write-Host ("    [M] Notifica MAILBOX  inviata a {0,-42} {1,5}%  [{2}]" -f `
                $r.Casella, $pct, $(if ($isCritical) {"CRITICAL"} else {"WARNING"})) -ForegroundColor Yellow
        } catch {
            $errCount++
            Write-Host "    Errore invio notifica MAILBOX a $($r.Casella): $_" -ForegroundColor Red
        }
    }

    # ------------------------------------------------------------------
    # LOOP 2: NOTIFICHE ARCHIVIO ONLINE
    # Una mail per ogni casella in $archAlerts che NON sia una risorsa.
    # ------------------------------------------------------------------
    foreach ($r in $archAlerts) {
        if ($skipResource -and ($resourceTypes -contains $r.Tipo)) {
            $skipResArc++
            continue
        }

        $pct = [double]$r.PercArchivio
        $isCritical = $pct -ge $sogliaCrit

        if ($isCritical) {
            $subj    = "URGENTE - Il tuo archivio online e' quasi pieno"
            $color   = "#7B4F9F"
            $title   = "Il tuo archivio online e' quasi pieno"
            $urgency = "Intervieni il prima possibile: oltre il limite l'archivio non potra' piu' contenere nuovi messaggi spostati dalla mailbox principale e l'auto-archiviazione si blocchera'."
            $bg      = "#E8DAEF"
        } else {
            $subj    = "Avviso - Il tuo archivio online sta raggiungendo il limite"
            $color   = "#8e44ad"
            $title   = "Il tuo archivio online sta raggiungendo il limite"
            $urgency = "L'archivio online si sta riempiendo. Conviene iniziare a fare pulizia prima di arrivare al limite, oltre il quale non potrai piu' archiviare nuovi messaggi."
            $bg      = "#F4ECF7"
        }

        $body = @"
<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#212529;max-width:640px;margin:0 auto">
<div style="background:$bg;border-left:4px solid $color;padding:14px 18px;margin-bottom:18px">
<h2 style="color:$color;margin:0 0 8px 0">$title</h2>
<p style="margin:0">$urgency</p>
</div>
<p>Ciao $($r.DisplayName),</p>
<p>Il tuo archivio online associato alla casella <strong>$($r.Casella)</strong> ha raggiunto il <strong>$($r.PercArchivio)%</strong> ($($r.ArchivioUsatoGB) GB su $($r.ArchivioQuotaGB) GB).</p>
<p style="font-size:13px;color:#555">Questa notifica riguarda esclusivamente l'<strong>archivio online</strong>, che e' una cassetta postale separata dalla tua mailbox principale di lavoro quotidiano. Non blocca la ricezione di nuovi messaggi sulla casella principale, ma blocca l'archiviazione (manuale e automatica) di ulteriori contenuti.</p>
<h3 style="color:#1F3864;margin-top:22px">Cosa puoi fare</h3>
<ul>
  <li>apri Outlook e nella sezione "Archivio online" (sotto la tua casella principale) elimina le cartelle o i messaggi che non servono piu'</li>
  <li>svuota la cartella posta eliminata anche all'interno dell'archivio</li>
  <li>se l'archivio contiene email molto vecchie che vuoi comunque conservare, valuta con il reparto IT la possibilita' di esportarle in PST o aumentare la quota</li>
  <li>per qualsiasi dubbio o richiesta di aiuto, scrivi a <a href="mailto:$itSupport" style="color:#2E75B6">$itSupport</a></li>
</ul>
<p style="color:#6c757d;font-size:11px;border-top:1px solid #dee2e6;padding-top:10px;margin-top:24px">
Messaggio automatico inviato dal sistema di monitoraggio delle caselle di posta del tenant $($cfg.tenant.domain).<br>
Generato il $dataLabel. Per assistenza: $itSupport
</p>
</body></html>
"@

        try {
            Send-UserNotification -To $r.Casella -Subject $subj -Body $body
            $sentArc++
            Write-Host ("    [A] Notifica ARCHIVIO inviata a {0,-42} {1,5}%  [{2}]" -f `
                $r.Casella, $pct, $(if ($isCritical) {"CRITICAL"} else {"WARNING"})) -ForegroundColor Magenta
        } catch {
            $errCount++
            Write-Host "    Errore invio notifica ARCHIVIO a $($r.Casella): $_" -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-Host ("    Notifiche mailbox:  {0} inviate, {1} risorse saltate" -f $sentMb,  $skipResMb)  -ForegroundColor Yellow
    Write-Host ("    Notifiche archivio: {0} inviate, {1} risorse saltate" -f $sentArc, $skipResArc) -ForegroundColor Magenta
    if ($errCount -gt 0) {
        Write-Host ("    Errori invio:       {0}" -f $errCount) -ForegroundColor Red
    }
} else {
    Write-Host "`n[6/6] Notifiche personali disabilitate in config.json (notifyMailboxOwners = false)" -ForegroundColor DarkGray
}

# --------------------------------------------------------------------------
# DISCONNESSIONE
# --------------------------------------------------------------------------
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Disconnect-MgGraph | Out-Null

Write-Host "`n[OK] Esecuzione completata.`n" -ForegroundColor Green
Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
