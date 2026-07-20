---
generated-from-commit: a8eab4c
generated-from-branch: main
generated-date: 2026-06-18
covers-paths:
  - "**/*.ps1"
  - "**/*.py"
  - config.example.json
last-verified-commit: a8eab4c
---

# Stack applicativo

> Documento di recupero più importante: tracciato, perché un collega che clona deve vederlo.

## Stack e runtime

Il toolkit è un'applicazione PowerShell-Python ibrida senza framework web né dipendenze esterne oltre a quelle elencate. Non esiste un file `requirements.txt` esplicito: le dipendenze Python vengono installate da `launcher.ps1` direttamente nel *venv* isolato sotto `venv/`.

Lato PowerShell, lo script richiede almeno la versione 5.1, quella inclusa in Windows 11. I moduli necessari sono `ExchangeOnlineManagement` (versione 3.x o superiore) e il sottosistema `Microsoft.Graph.*` nella tripletta `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`, `Microsoft.Graph.Identity.DirectoryManagement`. Il loro ordine di caricamento è critico: Graph va importato *prima* di ExchangeOnlineManagement per evitare `TypeLoadException` su versioni vecchie degli assembly.

Lato Python, il runtime richiesto è 3.10 o superiore. L'unica dipendenza esterna è `openpyxl` (ultima versione disponibile), installata nel *venv* da `launcher.ps1`. Il modulo `sqlite3` è built-in e non richiede installazione.

| Componente | Versione minima | Utilizzo |
|---|---|---|
| PowerShell | 5.1 | Orchestrazione, connessioni Exchange/Graph, mail SMTP |
| Python | 3.10 | Generazione Excel, scrittura SQLite |
| openpyxl | ultima | Creazione e formattazione file `.xlsx` |
| sqlite3 | built-in Python | Database storico `mailbox_history.db` |
| ExchangeOnlineManagement | 3.x+ | `Get-Mailbox`, `Get-MailboxStatistics`, `Get-MailboxFolderStatistics` |
| Microsoft.Graph.Authentication | ultima | Auth cert-based verso Microsoft Graph |
| Microsoft.Graph.Users | ultima | `Get-MgUser` per licenze assegnate |
| Microsoft.Graph.Identity.DirectoryManagement | ultima | `Get-MgSubscribedSku` per mappa SKU tenant |
| Windows Task Scheduler | built-in Windows | Esecuzione automatica via XML importato |

## Alternative deliberatamente escluse

Un database relazionale completo come PostgreSQL o SQL Server è stato escluso: lo storico è in SQLite perché non richiede installazione, è portabile, e le query di analytics su di esso sono sufficientemente veloci per un volume di dati dell'ordine di migliaia di righe annue. La generazione Excel via COM (`Microsoft.Office.Interop.Excel`) è stata esclusa in favore di `openpyxl` perché la prima richiede un'installazione di Office sulla macchina di esecuzione, cosa non garantita su un server di schedulazione. Non esiste una dipendenza da `pandas` o `numpy`: le operazioni sui dati sono abbastanza semplici da essere gestite in Python puro.

## Flussi di codice e ruolo architetturale dei file

L'entry point è `launcher.ps1`, che esegue un setup idempotente alla prima esecuzione e poi trasferisce il controllo a `mailbox-alert.ps1`. La comunicazione tra PowerShell e Python avviene tramite *stdout*: PowerShell serializza i dati raccolti in JSON e li passa a `generate_mailbox_report.py` via *pipeline* di processo, senza file intermedi su disco. Il Python legge lo *stdin*, scrive i file Excel in `reports/user/` e `reports/other/`, aggiorna il database SQLite in `history/`, e risponde con una riga di riepilogo su *stdout* che PowerShell stampa nel log.

Il flusso per i trend è separato e on-demand: `launcher.ps1 -Trends` chiama `generate_trends.py` direttamente, passando il percorso del database come argomento. Non c'è interazione con Exchange o Graph in questa modalità.

```
launcher.ps1
  └─ Invoke-Setup()         [prima esecuzione: moduli PS, venv Python]
  └─ mailbox-alert.ps1      [esecuzione normale]
       ├─ Microsoft.Graph   [licenze utenti]
       ├─ ExchangeOnline    [statistiche mailbox]
       ├─ JSON → stdin → generate_mailbox_report.py
       │    └─ Excel (user, other) + SQLite upsert
       ├─ Retention policy  [pulizia vecchi file]
       └─ SMTP              [mail IT + notifiche personali]
  └─ generate_trends.py     [solo con -Trends]
       └─ SQLite → Excel (8 fogli analytics)
```

## Riferimenti a snippet

| Simbolo | File | Nota |
|---|---|---|
| `Get-SizeBytes` | `mailbox-alert.ps1` | Parser per stringhe dimensione Exchange deserializzate |
| `Invoke-Setup` | `launcher.ps1` | Setup idempotente: moduli PS, venv, openpyxl |
| `init_db` | `generate_mailbox_report.py` | Crea schema SQLite e indici |
| `upsert_today` | `generate_mailbox_report.py` | `INSERT OR REPLACE` idempotente per data+email |
| `get_growth_30gg` | `generate_mailbox_report.py` | Delta 30gg da SQLite per colonna Δ Excel |
| `write_sheet` | `generate_mailbox_report.py` | Genera un foglio Excel con formattazione completa |
| `pct_color` | `generate_mailbox_report.py` | Colorazione cella in base a percentuale occupazione |
