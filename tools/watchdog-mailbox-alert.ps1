<#
    Guardiano del toolkit, finche' il servizio vive su questa postazione.

    Osserva il fallimento invece della rotazione della password: controlla che
    l'ultima esecuzione sia fresca e che le attivita' pianificate non stiano
    fallendo l'avvio. Copre quindi la rotazione della password, ma anche le
    altre cause di fermo, e non dipende dal ricordarsi di eseguirlo al momento
    giusto.

    Non invia posta di proposito: la credenziale SMTP e' cifrata DPAPI e nel
    guasto tipico e' proprio la posta a non partire, quindi un avviso via mail
    tacerebbe esattamente quando serve. Avvisa a video e scrive un file di stato.

    Va registrato con tipo di accesso Interattivo, senza password memorizzata,
    altrimenti erediterebbe il guasto che deve sorvegliare.
#>
param(
    [int]$GiorniTolleranza = 1,
    [switch]$Silenzioso
)

$ErrorActionPreference = 'Stop'
$radice  = Split-Path -Parent $PSScriptRoot
$logDir  = Join-Path $radice 'logs'
$statoFile = Join-Path $radice '_notes\watchdog-stato.txt'
$problemi = @()

# 1. Freschezza dell'ultima esecuzione: e' il log a provare che lo script e'
#    partito davvero, mentre LastRunTime dell'attivita' si aggiorna anche quando
#    l'avvio fallisce subito dopo.
$ultimo = Get-ChildItem $logDir -Filter 'run_*.log' -ErrorAction SilentlyContinue |
          Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $ultimo) {
    $problemi += "Nessun log di esecuzione in $logDir."
} else {
    $eta = [int]((Get-Date) - $ultimo.LastWriteTime).TotalDays
    if ($eta -gt $GiorniTolleranza) {
        $problemi += "Ultima esecuzione $eta giorni fa ($($ultimo.Name)). Il servizio e' fermo."
    }
}

# 2. Fallimenti di avvio delle due attivita'. 0x8007052E e' la password
#    memorizzata non piu' valida, tipicamente dopo un cambio password o un
#    riavvio da aggiornamento.
foreach ($nome in 'Mailbox alert Intrawelt', 'MailboxTrendsIntrawelt') {
    $t = Get-ScheduledTask -TaskName $nome -ErrorAction SilentlyContinue
    if (-not $t) { $problemi += "Attivita' '$nome' non registrata."; continue }
    if ($t.State -eq 'Disabled') { $problemi += "Attivita' '$nome' disabilitata." }

    $xp = "*[EventData[Data[@Name='TaskName']='\$nome']]"
    $ev = Get-WinEvent -LogName 'Microsoft-Windows-TaskScheduler/Operational' `
            -FilterXPath $xp -MaxEvents 10 -ErrorAction SilentlyContinue |
          Where-Object { $_.Id -eq 101 -and $_.TimeCreated -gt (Get-Date).AddDays(-3) }
    if ($ev) {
        $q = ($ev | Measure-Object).Count
        $problemi += "Attivita' '$nome': $q fallimenti di avvio negli ultimi 3 giorni. Se l'errore e' 0x8007052E la password memorizzata non e' piu' valida: riaprire il progetto e reimmetterla."
    }
}

# 3. Scadenza del certificato di app registration, che fermerebbe il servizio
#    in un modo diverso e con lo stesso silenzio.
try {
    $cfg = Get-Content (Join-Path $radice 'config.json') -Raw | ConvertFrom-Json
    $tp  = $cfg.appRegistration.certificateThumbprint
    $c   = Get-ChildItem Cert:\CurrentUser\My -ErrorAction SilentlyContinue |
           Where-Object Thumbprint -eq $tp
    if (-not $c) { $problemi += "Certificato $tp non trovato in Cert:\CurrentUser\My." }
    elseif (($c.NotAfter - (Get-Date)).TotalDays -lt 60) {
        $problemi += "Certificato in scadenza il $($c.NotAfter.ToString('yyyy-MM-dd'))."
    }
} catch { $problemi += "config.json non leggibile: $($_.Exception.Message)" }

# Esito
New-Item -ItemType Directory -Force -Path (Split-Path $statoFile) | Out-Null
$ora = Get-Date -Format 'yyyy-MM-dd HH:mm'
if ($problemi.Count -eq 0) {
    "[$ora] OK - servizio fresco, nessun fallimento di avvio." | Set-Content $statoFile -Encoding UTF8
    Write-Host "OK: nessun problema." -ForegroundColor Green
    exit 0
}

$testo = "MAILBOX ALERT TOOLKIT - IL SERVIZIO RICHIEDE ATTENZIONE`n`n" +
         (($problemi | ForEach-Object { "- $_" }) -join "`n") +
         "`n`nProgetto: $radice"
"[$ora] PROBLEMA`n$testo" | Set-Content $statoFile -Encoding UTF8
Write-Host $testo -ForegroundColor Red

if (-not $Silenzioso) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($testo, 'Mailbox Alert Toolkit',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
}
exit 1
