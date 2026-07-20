---
generated-from-commit: a8eab4c
generated-from-branch: main
generated-date: 2026-06-18
covers-paths:
  - "*.xml"
  - launcher.ps1
  - config.example.json
last-verified-commit: a8eab4c
---

# Deployment

> Commit, push e deploy restano operazioni manuali dell'utente.

## Livelli

Non esistono livelli di staging né ambienti separati. Il toolkit gira direttamente in produzione sul tenant Microsoft 365 di Intrawelt (`intrawelt.com`), su una singola macchina Windows 11 con Windows Task Scheduler. Non c'è CI/CD né pipeline di rilascio automatizzata.

L'ambiente di produzione è la cartella `D:\mailbox-alert-toolkit\` sulla macchina designata. I file generati (report Excel, database SQLite, log) rimangono locali su quella macchina e non vengono caricati in nessun sistema remoto. Le uniche comunicazioni di rete sono le connessioni verso Microsoft Graph API, Exchange Online, e il server SMTP `smtp.office365.com:587`.

## Comandi

Il setup iniziale avviene eseguendo `launcher.ps1` con il flag `-SkipRun` la prima volta, per installare i moduli PowerShell e creare il *venv* Python senza lanciare il main script:

```powershell
powershell.exe -ExecutionPolicy Bypass -File "D:\mailbox-alert-toolkit\launcher.ps1" -SkipRun
```

L'esecuzione normale (quella che gira da Task Scheduler ogni giorno alle 07:00) è:

```powershell
powershell.exe -ExecutionPolicy Bypass -NonInteractive -WindowStyle Hidden -File "D:\mailbox-alert-toolkit\launcher.ps1"
```

Il report trends si genera on-demand o da Task Scheduler ogni lunedì alle 08:00:

```powershell
powershell.exe -ExecutionPolicy Bypass -NonInteractive -WindowStyle Hidden -File "D:\mailbox-alert-toolkit\launcher.ps1" -Trends
```

Per forzare il re-setup (aggiornamento moduli PS, ricostruzione *venv*):

```powershell
powershell.exe -ExecutionPolicy Bypass -File "D:\mailbox-alert-toolkit\launcher.ps1" -ForceSetup -SkipRun
```

L'importazione dei task schedulati in Windows Task Scheduler avviene una sola volta per macchina:

```powershell
schtasks /create /xml "D:\mailbox-alert-toolkit\Mailbox alert Intrawelt.xml" /tn "Mailbox alert Intrawelt" /ru "DOMAIN\username" /rp "password"
schtasks /create /xml "D:\mailbox-alert-toolkit\MailboxTrendsIntrawelt.xml" /tn "MailboxTrendsIntrawelt" /ru "DOMAIN\username" /rp "password"
```

Il SID utente richiesto nei file XML si ottiene con `whoami /user`. Se il SID cambia (cambio macchina, cambio account), i file XML vanno aggiornati prima dell'importazione.

## Variabili d'ambiente e segreti

Non si usano variabili d'ambiente di sistema. La configurazione vive interamente in `config.json` (gitignored). I campi sensibili sono:

| Campo | Dove vive | Note |
|---|---|---|
| `appRegistration.clientId` | `config.json` | GUID app registration Entra ID |
| `appRegistration.tenantId` | `config.json` | GUID tenant Microsoft 365 |
| `appRegistration.certificateThumbprint` | `config.json` | Thumbprint del cert nel Windows Certificate Store |
| Chiave privata del certificato | Windows Certificate Store | Non su disco come file |
| Credenziali SMTP | `credentials/smtp-cred.xml` | DPAPI, leggibile solo dall'utente Windows corrente |

Il certificato pubblico `app-cert.cer` è caricato sull'app registration in Azure; la chiave privata rimane nel *certificate store* della macchina e non deve mai uscirne. Il file `app-cert.cer` è tracciato in git ma è inerte senza la chiave privata.
