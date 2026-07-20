---
generated-from-commit: a8eab4c
generated-from-branch: main
generated-date: 2026-06-18
covers-paths:
  - "**/*.ps1"
  - "**/*.py"
  - "examples/**"
last-verified-commit: a8eab4c
---

# Test di sviluppo

## Test runner e comandi

Non esiste un framework di test automatizzato. Non ci sono file di test, né *test runner*, né CI. Il toolkit è verificato manualmente in tre modalità.

*Verifica del setup* senza connessioni di rete, per controllare che l'installazione dei moduli e la creazione del *venv* siano corretti:

```powershell
.\launcher.ps1 -SkipRun
```

*Esecuzione completa manuale* per verificare la raccolta dati, la generazione Excel e il flusso mail:

```powershell
.\launcher.ps1
```

*Generazione trends* per verificare la lettura del database SQLite e la produzione del workbook analytics:

```powershell
.\launcher.ps1 -Trends
```

Il log dell'esecuzione viene scritto in `logs/mailbox-alert-YYYY-MM-DD.log` come trascrizione PowerShell completa. È il primo posto dove guardare in caso di errore.

## Rotte e dati mockati

Non esistono mock né stub. Non c'è una modalità *dry-run* che simuli le chiamate a Exchange o Graph senza effettuarle. I file in `examples/` (`ESEMPIO_Report_UserMailbox.xlsx`, `ESEMPIO_Report_AltreMailbox.xlsx`, `ESEMPIO_Trends.xlsx`) sono output reali anonimi usati come riferimento visivo per verificare formattazione, colori e struttura dei fogli, ma non sono generati da dati di test controllati.

L'unico modo di testare la generazione Excel senza accesso a Exchange è passare un JSON costruito a mano direttamente a `generate_mailbox_report.py`:

```powershell
'[{"email":"test@example.com","displayName":"Test User","tipo":"UserMailbox",...}]' | python venv/Scripts/python.exe generate_mailbox_report.py
```

Questa modalità non è documentata né strutturata: è un'opzione di sviluppo manuale.

## Hook e controlli di qualità

Non ci sono *linter*, *type checker* né *pre-commit hook* configurati. La qualità del codice si verifica eseguendo gli script e controllando il log di trascrizione. Il controllo visivo del proofing dei report Excel avviene tramite screenshot, secondo la regola `manual-screenshots.md`: si chiede all'utente di catturare lo schermo con Screenpresso e si legge l'immagine dalla cartella `%USERPROFILE%\Pictures\Screenpresso`.
