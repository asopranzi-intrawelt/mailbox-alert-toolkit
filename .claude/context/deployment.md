---
generated-from-commit: 152deac
generated-from-branch: main
generated-date: 2026-06-18
covers-paths:
  - "*.xml"
  - launcher.ps1
  - config.example.json
last-verified-commit: 152deac
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

## Stato di esercizio e migrazione prevista su VM Proxmox

Annotazione del 2026-10-28, successiva alla generazione della scheda: la collocazione descritta sopra, cioè il task su `PC-ALESSIO` sotto l'utenza interattiva `Utente`, ha prodotto un fermo di ventitré giorni. Il task `Mailbox alert Intrawelt` è configurato con `LogonType=Password` e `RunLevel=Highest`, quindi conserva una password dell'utenza Windows; dopo il riavvio da Windows Update del 2026-10-06 quella credenziale non è più valida e Utilità di pianificazione rifiuta l'avvio ogni mattina con errore `0x8007052E`, ERROR_LOGON_FAILURE, senza che `launcher.ps1` venga mai eseguito e quindi senza che venga prodotto alcun log. La diagnosi completa, con le misure e il loro perimetro, sta nel work-log alla data; la decisione sul rimedio immediato sta in ADR-003 del registro delle decisioni.

La destinazione decisa in ADR-002 è una macchina virtuale Proxmox dedicata, a scopo unico, progettata con il materiale del progetto `D:\network-design`. Il requisito che la governa è la continuità del servizio: la postazione di lavoro non è un luogo dove collocare un processo che non deve interrompersi, perché si riavvia per aggiornamenti, dipende dalla sessione di una persona e non è osservabile dall'esterno.

Tre vincoli condizionano la migrazione e vanno risolti prima di spegnere il task sulla postazione, perché nessuno dei tre si trasferisce con una copia di cartella. Il certificato di app registration `CN=MailboxAlertToolkit`, thumbprint `<thumbprint-cert-app>`, vive in `Cert:\CurrentUser\My` dell'utente `Utente` con la chiave privata e scade il 2028-05-21: sulla VM va reinstallato come PFX esportato, oppure sostituito da un certificato nuovo caricato sull'app registration in Entra ID, e la scelta fra le due vie determina se il vecchio resti valido durante la transizione. La credenziale SMTP in `credentials/smtp-cred.xml` è cifrata con DPAPI e quindi decifrabile solo dall'utente Windows che l'ha creata sulla macchina che l'ha creata: va rigenerata sulla VM, non copiata, e questo è anche il motivo per cui il principal del task non può passare a `S4U` come rimedio rapido, perché sotto quel tipo di logon DPAPI non dispone della chiave che sblocca la master key dell'utente. Il terzo vincolo è lo storico: `history/mailbox_history.db` contiene la serie dal 2026-05-21 e i confronti di crescita dei report dipendono da essa, quindi va trasferito insieme al codice e non ricostruito.

Resta inoltre da colmare l'assenza che ha reso l'incidente invisibile. Il solo canale che segnalava il funzionamento del toolkit era la mail che il toolkit stesso invia, quindi il suo fermo si manifesta come silenzio, che è indistinguibile da una giornata senza allarmi. Sulla VM serve un controllo di vivacità esterno al processo, che osservi la freschezza dell'ultima esecuzione invece della sola raggiungibilità della macchina.
