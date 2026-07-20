---
generated-from-commit: a8eab4c
generated-from-branch: main
generated-date: 2026-06-18
covers-paths:
  - mailbox-alert.ps1
  - launcher.ps1
  - generate_mailbox_report.py
  - generate_trends.py
  - config.example.json
last-verified-commit: a8eab4c
---

# Design e sicurezza applicativa

## Paradigmi di software design

Il toolkit segue un pattern di *pipeline* lineare e deterministica: ogni fase produce un output esplicito che diventa l'input della fase successiva, senza stato condiviso globale né callback. Non esiste un framework di dependency injection né un sistema di plugin. La separazione delle responsabilità è netta: PowerShell si occupa di tutto ciò che richiede l'autenticazione a Microsoft 365 (raccolta dati, invio mail, gestione credenziali), Python si occupa di tutto ciò che richiede la generazione di file binari Excel e la persistenza strutturata su SQLite.

La comunicazione tra i due layer avviene tramite un contratto JSON esplicito: PowerShell costruisce un array di oggetti con 22 campi per ogni mailbox e lo serializza su *stdout*, Python lo legge da *stdin*, lo valida implicitamente nella deserializzazione e procede. Non esiste versioning formale di questo contratto: modifiche ai campi richiedono aggiornamento coordinato di entrambi gli script.

Il codice è procedurale, non orientato agli oggetti. Le funzioni helper (`Get-SizeBytes`, `Invoke-Setup`, `Invoke-Retention`, `pct_color`, `growth_color`, `write_sheet`) sono estratte per leggibilità e riuso interno, non come astrazione verso interfacce. Non esiste un sistema di configurazione dinamica a runtime: `config.json` viene letto una sola volta all'avvio e i valori vengono usati come costanti per l'intera esecuzione.

## Sicurezza applicativa

La gestione dell'autenticazione verso Microsoft 365 distingue due modalità. In produzione, quella raccomandata, si usa l'*app registration* in Microsoft Entra ID con un certificato *self-signed* a due anni: nessun segreto condiviso, nessun token di refresh da rinnovare, nessun login browser. PowerShell carica il certificato dal *certificate store* locale per thumbprint e lo usa per firmare il JWT di autenticazione. La modalità interattiva (browser popup) è un fallback per lo sviluppo e non è compatibile con task schedulati *unattended*.

Le credenziali SMTP sono cifrate con *DPAPI* (Data Protection API di Windows) e salvate in `credentials/smtp-cred.xml`. Il file è leggibile solo dall'utente Windows che lo ha generato, sulla stessa macchina. Se il file non esiste alla prima esecuzione, PowerShell chiede la password interattivamente e salva il file cifrato. Su un'altra macchina o con un altro account utente, il file è illeggibile e il toolkit chiede nuovamente le credenziali.

I segreti non versionati sono: `config.json` (contiene clientId, tenantId, thumbprint del certificato), `app-cert.cer` (certificato pubblico caricato su Azure), file `*.pfx`/`*.key`/`*.p12`, `credentials/smtp-cred.xml`. Il `.gitignore` esclude esplicitamente tutti questi percorsi. Il file versionato `config.example.json` contiene solo segnaposto senza valori reali.

Il toolkit non espone superfici di rete: non ha un server HTTP, non ascolta su porte, non accetta input dall'esterno. L'unica superficie di input è `config.json`, letto da disco prima dell'esecuzione. I dati raccolti da Exchange transitano in memoria come oggetti PowerShell e vengono scritti su disco solo come file Excel e come righe SQLite, mai come testo in chiaro in posizioni accessibili dall'esterno.

Non viene eseguita validazione esplicita degli input utente perché non ci sono input utente a runtime: il toolkit è completamente automatizzato. La validazione delle dimensioni mailbox avviene nella funzione `Get-SizeBytes`, che gestisce i vari formati stringa prodotti dagli oggetti Exchange deserializzati e restituisce `$null` per stringhe non parsabili, senza interrompere l'esecuzione.

## Diagrammi

Nessun diagramma in `diagrams/` al commit corrente. Il flusso architetturale è descritto testualmente in `STACK.md` e nel `README.md`.
