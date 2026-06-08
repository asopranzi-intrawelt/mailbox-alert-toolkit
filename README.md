# Mailbox alert toolkit per intrawelt.com

Toolkit per il monitoraggio dello spazio occupato dalle caselle di posta del tenant Microsoft 365 di intrawelt.com.

Lo scopo è ricevere ogni giorno una mail di riepilogo a it@intrawelt.com con allegati i due report Excel del giorno (uno per le user mailbox associate a una licenza, uno per le altre caselle come shared, sale riunioni, risorse), avvisando in evidenza quando una qualsiasi casella supera l'80% di spazio occupato. Mailbox principale e archivio online vengono trattati come due cose distinte in tutto il sistema. Ogni casella in warning (80%) o critical (95%) riceve anche una mail di notifica personale separata per il problema della mailbox principale e per quello dell'archivio. In parallelo viene mantenuto uno storico di lungo periodo (database SQLite) che permette analisi multi-anno e che alimenta un report trends generato on-demand.

L'installazione di riferimento è sul PC fisico denominato PC-ALESSIO, in `D:\mailbox-alert-toolkit\` (SSD interno). Le mail di alert vengono inviate dall'indirizzo `mailer@intrawelt.com`, una casella dedicata creata appositamente per le notifiche automatiche.

L'autenticazione a Microsoft Graph e Exchange Online avviene tramite **app registration con certificato** (modalità non interattiva), così il task schedulato di Windows funziona senza nessun popup di login e senza scadenze di refresh token.

---

## Indice

1. Cosa fa il toolkit, in concreto
2. Perché è strutturato così (scelte di design)
3. Cosa contiene il pacchetto
4. Requisiti
5. Setup App Registration in Azure AD (autenticazione unattended)
6. Setup iniziale passo passo su Windows 11
7. Come eseguirlo manualmente
8. Schedulazione con utilità di pianificazione di Windows 11
9. Struttura della cartella di lavoro
10. File di configurazione config.json
11. Metriche raccolte per ogni casella
12. I due report Excel giornalieri
13. Notifiche personali agli utenti (mailbox separata dall'archivio)
14. Lo storico SQLite e perché esiste
15. Il report trends
16. Retention policy dei file Excel
17. Portabilità del virtual environment Python
18. Gestione delle credenziali e sicurezza
19. Logging
20. Tutte le modalità del launcher
21. Gestione delle licenze e SKU
22. Note tecniche per gli sviluppatori (perché PowerShell si comporta così)
23. Troubleshooting
24. Manutenzione e aggiornamenti

---

## 1. Cosa fa il toolkit, in concreto

A ogni esecuzione (manuale o schedulata) il toolkit:

1. si connette a Microsoft Graph (usando l'account amministratore configurato) e a Exchange Online
2. recupera la lista degli SKU acquistati dal tenant per costruire dinamicamente la mappa licenze leggibile
3. recupera la lista di tutte le mailbox del tenant (user, shared, room, equipment, scheduling)
4. per ognuna di esse raccoglie 22 metriche (spazio usato, quota, percentuale, archivio, inbox, posta inviata, eliminati, inoltri, hold, ultimo accesso, eccetera)
5. calcola la crescita degli ultimi 30 giorni confrontando con lo storico salvato in passato
6. genera due file Excel del giorno, uno per le user mailbox e uno per le altre caselle
7. aggiorna il database SQLite con il record del giorno (una riga per ogni casella)
8. applica la retention policy mantenendo solo gli ultimi 365 file Excel per tipo
9. invia una mail riassuntiva a it@intrawelt.com con allegati i due file Excel del giorno e DUE tabelle HTML distinte: una per le mailbox principali in warning, una per gli archivi online in warning
10. invia mail di notifica personale separate: chi ha la mailbox in warning riceve una mail dedicata, chi ha l'archivio in warning ne riceve una diversa, chi ha entrambi i problemi riceve due mail distinte

La generazione di un report trends storico è separata e on-demand: si lancia con un parametro dedicato e produce un workbook con 8 fogli di analytics.

---

## 2. Perché è strutturato così (scelte di design)

### Perché due report Excel separati per user e altre mailbox

Le user mailbox e le altre caselle hanno significato diverso: le prime sono associate a una licenza Microsoft 365 (quindi a un costo), le seconde sono caselle funzionali (info@, sale riunioni, risorse). Separarle rende più facile dare responsabilità di gestione a persone diverse e fare analisi distinte.

### Perché un launcher separato dal main script

Il launcher fa setup idempotente: la prima volta installa moduli PowerShell, crea il virtual environment Python e installa le dipendenze, crea la struttura di cartelle, salva un marker. Dalla seconda volta in poi verifica solo che tutto sia in ordine e parte direttamente. Questo evita di rifare ogni giorno operazioni inutili e rende il toolkit auto-riparante: se cancelli per sbaglio il venv, alla prossima esecuzione viene ricreato.

### Perché un virtual environment Python e non installazione globale

Il venv tiene isolate le dipendenze Python (la libreria openpyxl che serve per generare gli Excel) nella cartella stessa del toolkit. Significa che puoi spostare l'intera cartella su un altro disco, su una chiavetta USB o su un'altra macchina, senza dover reinstallare nulla a livello di sistema operativo.

### Perché 365 file Excel per tipo

Uno per ogni giorno per un anno. Permette ispezione manuale immediata di un giorno specifico, audit, e l'invio come allegato senza dipendere dal database. Oltre i 365 vengono cancellati automaticamente partendo dal più vecchio.

### Perché in più anche un database SQLite

Leggere 365 file Excel per fare statistiche è scomodo, lento e fragile. Il database SQLite affiancato accumula tutto lo storico indefinitamente (cresce di pochi megabyte all'anno) e permette query istantanee per analytics. I due sistemi convivono: gli Excel per la consultazione umana e l'audit, il database per l'analisi.

### Perché il report trends è on-demand e non automatico

Il report trends è pesante (interroga tutto lo storico, contiene grafici e pivot complete) e ha senso lanciarlo periodicamente (ogni settimana o ogni mese), non ogni giorno. Quindi è separato e si attiva con il parametro `-Trends`.

### Perché mailbox principale e archivio sono trattati come due cose distinte

Sono due problemi tecnicamente diversi che richiedono interventi diversi:

- la mailbox principale che si riempe oltre il limite BLOCCA la ricezione di nuovi messaggi (impatto immediato e visibile per l'utente)
- l'archivio online che si riempe non blocca la ricezione di posta, ma blocca l'archiviazione manuale e automatica (impatto sulla gestione dei messaggi storici)

Vengono trattati come problemi distinti in console, nel report Excel, nella mail al reparto IT (due tabelle separate), e nelle notifiche personali agli utenti (due mail separate se entrambi in warning).

### Perché le notifiche personali partono in automatico

L'idea è che IT non debba fare da postino. Quando una casella raggiunge la soglia, la persona proprietaria riceve direttamente una mail con messaggio adattato al livello di urgenza e suggerimenti pratici su cosa fare. Le caselle senza destinatario umano (sale riunioni, attrezzature, scheduling) sono escluse di default.

### Perché un account mittente dedicato `mailer@intrawelt.com`

In Microsoft 365 SMTP AUTH richiede che il `From` coincida con l'utente autenticato. Usare un account dedicato significa: non sporcare la posta inviata di `it@intrawelt.com` con mail automatiche, isolare le configurazioni di SMTP AUTH e MFA su una casella sola, poter ruotare le credenziali del mittente senza toccare l'account amministratore.

---

## 3. Cosa contiene il pacchetto

```
mailbox-alert-toolkit/
    launcher.ps1                       entry point: lanci sempre questo
    mailbox-alert.ps1                  main script chiamato dal launcher
    generate_mailbox_report.py         generatore xlsx + scrittura storico
    generate_trends.py                 generatore report analytics on-demand
    config.json                        tutte le impostazioni in un posto
    README.md                          questo documento
    README.docx                        stesso contenuto del README in formato Word
    examples/
        ESEMPIO_Report_UserMailbox.xlsx
        ESEMPIO_Report_AltreMailbox.xlsx
        ESEMPIO_Trends.xlsx
```

I file dentro `examples/` sono campioni generati con dati fittizi per mostrare il formato dei report. Puoi cancellarli quando non ti servono più.

---

## 4. Requisiti

Sulla macchina Windows 11 (PC-ALESSIO) dove fai girare il toolkit ti servono:

- PowerShell 5.1 o superiore, già presente in Windows 11
- Python 3.10 o superiore, installato e presente nel PATH (durante l'installazione spunta "Add Python to PATH"). Scaricalo da https://www.python.org/
- connessione internet per il primo download dei moduli PowerShell e dei pacchetti Python
- un account Microsoft 365 con privilegi sufficienti sul tenant:
    - ruolo Exchange Administrator (per leggere le statistiche di tutte le mailbox)
    - permessi `User.Read.All` e `Organization.Read.All` su Microsoft Graph (per leggere licenze degli utenti e SKU acquistati dal tenant)
- una casella mittente dedicata `mailer@intrawelt.com` che invierà sia il report giornaliero a it@intrawelt.com sia le notifiche personali alle caselle in warning. Requisiti specifici di questa casella:
    - deve essere una user mailbox (non shared) con almeno una licenza Exchange Online
    - deve avere SMTP AUTH abilitato (Microsoft lo disabilita di default sui tenant moderni)
    - deve avere una password nota perché ti verrà chiesta al primo lancio
    - non deve avere MFA abilitato. Per questa casella di servizio si crea una eccezione specifica nelle Conditional Access

Tutto il resto (moduli PowerShell, virtual environment Python, libreria openpyxl, sqlite3) viene installato dal launcher al primo avvio.

### Verificare e abilitare SMTP AUTH su mailer@intrawelt.com

Tramite portale:

1. accedi a https://admin.exchange.microsoft.com
2. vai su destinatari, poi cassette postali
3. seleziona la casella mailer@intrawelt.com
4. nella scheda cassetta postale, sezione app di posta elettronica, controlla che "SMTP autenticato" sia attivo
5. se è disattivato, attivalo e salva

Tramite PowerShell:

```powershell
Set-CASMailbox -Identity mailer@intrawelt.com -SmtpClientAuthenticationDisabled $false
```

Per verificare velocemente che SMTP AUTH funzioni prima di lanciare il toolkit:

```powershell
$cred = Get-Credential mailer@intrawelt.com
Send-MailMessage -To "it@intrawelt.com" -From "mailer@intrawelt.com" `
    -Subject "Test SMTP" -Body "Test" `
    -SmtpServer "smtp.office365.com" -Port 587 -UseSsl -Credential $cred
```

Se questo comando funziona, anche il toolkit funzionerà.

---

## 5. Setup App Registration in Azure AD (autenticazione unattended)

Questa configurazione viene fatta UNA VOLTA SOLA all'inizio e permette al task schedulato di funzionare senza alcun popup di login, senza scadenze di refresh token, e senza dover essere loggati interattivamente sulla macchina.

L'idea: invece di accedere come utente, lo script si presenta a Microsoft come un'**applicazione registrata** che si autentica con un **certificato** salvato sul PC. Microsoft riconosce l'app, verifica il certificato, e concede l'accesso senza chiedere nulla a nessuno.

Tutti i passi sotto sono per il portale Microsoft 365 / Entra in **italiano**. I corrispondenti inglesi sono indicati tra parentesi per riferimento.

### Passo 1: registra l'app in Microsoft Entra

1. accedi a https://entra.microsoft.com
2. menu di sinistra → **Identità** → **Applicazioni** → **Registrazioni app** ("App registrations")
3. clicca **"Nuova registrazione"** ("New registration")
4. nome: `Mailbox Alert Toolkit intrawelt`
5. tipi di account supportati: lascia **"Solo gli account in questa directory organizzativa"** ("Accounts in this organizational directory only")
6. URI di reindirizzamento: lascia vuoto
7. clicca **Registra** ("Register")

Sulla pagina che si apre (la "Panoramica" dell'app), copia e salva due valori:

- **ID applicazione (client)** ("Application (client) ID") → diventerà `appRegistration.clientId` nel config
- **ID della directory (tenant)** ("Directory (tenant) ID") → diventerà `appRegistration.tenantId`

### Passo 2: assegna i permessi applicativi

1. nell'app appena creata → menu di sinistra → **Autorizzazioni API** ("API permissions")
2. clicca **"Aggiungi un'autorizzazione"** ("Add a permission")
3. seleziona **"Microsoft Graph"**
4. importante: scegli **"Autorizzazioni applicazione"** ("Application permissions"), NON "Autorizzazioni delegate". È fondamentale, perché l'app deve agire per conto suo, non per conto di un utente
5. nella ricerca digita `User.Read.All`, spunta la casella
6. nella ricerca digita `Organization.Read.All`, spunta la casella
7. clicca **"Aggiungi autorizzazioni"** ("Add permissions")
8. clicca di nuovo **"Aggiungi un'autorizzazione"**
9. seleziona **"Office 365 Exchange Online"** (potrebbe essere sotto "API usate dalla mia organizzazione")
10. **Autorizzazioni applicazione** → cerca `Exchange.ManageAsApp` → spunta
11. **Aggiungi autorizzazioni**

Adesso vedi tre permessi nella lista, ognuno con lo "Stato" vuoto o "Non concesso" (yellow warning):

- Microsoft Graph / User.Read.All
- Microsoft Graph / Organization.Read.All
- Office 365 Exchange Online / Exchange.ManageAsApp

Per attivarli devi cliccare **"Concedi consenso amministratore per intrawelt"** ("Grant admin consent for intrawelt") in cima alla pagina. Questa azione richiede che il tuo account abbia il ruolo di **Amministratore globale** (Global Admin). Se non ce l'hai, chiedi a chi ce l'ha di cliccare quel pulsante una sola volta.

Dopo il consenso, tutti e tre i permessi mostrano un check verde "Concesso per intrawelt".

### Passo 3: genera un certificato self-signed sul PC-ALESSIO

Apri PowerShell come amministratore sul PC-ALESSIO e lancia (UN SOLO COMANDO, copia tutto in una volta):

```powershell
$cert = New-SelfSignedCertificate `
    -Subject "CN=MailboxAlertToolkit" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

$cert.Thumbprint
Export-Certificate -Cert $cert -FilePath "D:\mailbox-alert-toolkit\app-cert.cer" | Out-Null
```

Il comando:

- crea un certificato auto-firmato valido 2 anni nel "magazzino certificati personale" dell'utente Windows con cui sei loggato
- stampa il **Thumbprint** (40 caratteri esadecimali) → diventerà `appRegistration.certificateThumbprint` nel config
- esporta la parte PUBBLICA del certificato in `D:\mailbox-alert-toolkit\app-cert.cer`. È quella che caricherai in Azure. La parte privata resta solo sul PC

**Importante**: il comando deve essere lanciato con lo STESSO utente Windows che farà girare il task schedulato. Il certificato resta legato al profilo Windows di quell'utente. Se più tardi cambi utente del task, devi rifare il certificato con quell'utente.

### Passo 4: carica il certificato pubblico sull'app

1. torna su https://entra.microsoft.com → Identità → Applicazioni → Registrazioni app
2. apri l'app `Mailbox Alert Toolkit intrawelt`
3. menu di sinistra → **Certificati e segreti** ("Certificates & secrets")
4. tab **"Certificati"** ("Certificates")
5. clicca **"Carica certificato"** ("Upload certificate")
6. seleziona `D:\mailbox-alert-toolkit\app-cert.cer`
7. descrizione: `PC-ALESSIO mailbox alert task`
8. clicca **Aggiungi**

Vedi il certificato con il suo Thumbprint, che deve coincidere con quello stampato al Passo 3.

### Passo 5: assegna all'app un ruolo che le permetta di leggere le mailbox

`Exchange.ManageAsApp` da solo autorizza l'accesso ma serve anche un ruolo Entra che dica cosa l'app può fare. La via più semplice è assegnarle il ruolo **"Lettore globale"** ("Global Reader"), che è in sola lettura su tutto il tenant.

1. https://entra.microsoft.com → **Identità** → **Ruoli e amministratori** → **Tutti i ruoli** ("All roles")
2. nella barra di ricerca scrivi **"Lettore globale"** (in inglese: "Global Reader")
3. clicca il ruolo "Lettore globale"
4. clicca **"Aggiungi assegnazioni"** ("Add assignments")
5. nella casella di ricerca cerca `Mailbox Alert Toolkit intrawelt`
6. selezionala e clicca **Aggiungi**

In alternativa, se preferisci privilegi più ristretti, puoi usare il ruolo **"Amministratore destinatari Exchange"** ("Exchange Recipient Administrator"), che è sufficiente per leggere statistiche mailbox.

### Passo 6: trova il "dominio iniziale" del tenant

Connect-ExchangeOnline con cert auth richiede il nome del tenant nel formato `qualcosa.onmicrosoft.com`. Per trovarlo:

1. https://entra.microsoft.com → **Identità** → **Impostazioni** → **Nomi di dominio** ("Domain names")
2. nella lista, cerca quello con la colonna "Tipo" = `Iniziale` ("Initial"). Sarà tipo `esempio.onmicrosoft.com`

Questo valore va in `appRegistration.organization` nel config.

### Passo 7: compila config.json

Apri `D:\mailbox-alert-toolkit\config.json` con il blocco note e verifica la sezione `appRegistration`:

```json
"appRegistration": {
  "enabled":               true,
  "clientId":              "00000000-0000-0000-0000-000000000000",
  "tenantId":              "00000000-0000-0000-0000-000000000000",
  "certificateThumbprint": "INSERIRE_THUMBPRINT_CERTIFICATO",
  "organization":          "esempio.onmicrosoft.com"
}
```

Sostituisci i quattro valori con quelli ottenuti dai passi precedenti. Se `enabled` è `true`, lo script usa l'autenticazione via certificato. Se lo metti a `false` (o togli del tutto la sezione `appRegistration`), torna a usare il login interattivo via browser (utile per debug ma non funziona in task unattended).

### Passo 8: testa la connessione

Apri PowerShell come l'utente che ha generato il certificato e lancia:

```powershell
cd D:\mailbox-alert-toolkit
.\launcher.ps1
```

Dovresti vedere:

```
-> Connessione a Microsoft Graph...
    Modalita' app registration (certificato) - unattended
    ClientId:   00000000-0000-0000-0000-000000000000
    TenantId:   00000000-0000-0000-0000-000000000000
    Thumbprint: INSERIRE_THUMBPRINT_CERTIFICATO
-> Connessione a Exchange Online...
    Modalita' app registration (certificato) - unattended
    Organization: esempio.onmicrosoft.com
```

Se vedi questo, l'autenticazione applicativa funziona. Lo script prosegue normalmente raccogliendo i dati.

### Cosa fare se i token o il certificato cambiano

- il certificato scade dopo 2 anni (parametro `-NotAfter` del passo 3). Quando si avvicina la scadenza basta rifare i Passi 3 e 4 con un nuovo certificato e aggiornare `certificateThumbprint` nel config
- se sposti il toolkit su un altro PC, devi rifare il Passo 3 sul nuovo PC (esportare il certificato non basta, perché la chiave privata è legata al profilo Windows) e ricaricare la parte pubblica al Passo 4
- se vuoi vedere i certificati installati: `Get-ChildItem Cert:\CurrentUser\My`
- per rimuovere il vecchio certificato: `Remove-Item "Cert:\CurrentUser\My\<thumbprint>"`

---

## 6. Setup iniziale passo passo su Windows 11

### Passo 1: scegli dove mettere il toolkit

Il percorso di installazione previsto è `D:\mailbox-alert-toolkit\`, ovvero direttamente sulla radice dell'SSD interno di PC-ALESSIO.

Crea la cartella se non esiste:

1. apri esplora file
2. naviga su `E:\`
3. clicca col destro, nuovo, cartella, e chiamala `MailboxAlert`

### Passo 2: scompatta lo zip

Lo zip al suo interno contiene una cartella `mailbox-alert-toolkit/`. Quando estrai con esplora file, Windows crea automaticamente una cartella di destinazione con il nome dello zip e ci mette dentro il contenuto, risultando in una struttura annidata `D:\mailbox-alert-toolkit\mailbox-alert-toolkit\`. Per evitare l'annidamento:

1. apri lo zip con doppio click (esplora file ti fa entrare senza estrarre)
2. entra nella cartella `mailbox-alert-toolkit` che vedi dentro lo zip
3. seleziona tutto il suo contenuto (Ctrl+A)
4. trascina (o copia incolla) direttamente dentro `D:\mailbox-alert-toolkit\`

A questo punto dovresti avere:

```
D:\mailbox-alert-toolkit\launcher.ps1
D:\mailbox-alert-toolkit\mailbox-alert.ps1
D:\mailbox-alert-toolkit\generate_mailbox_report.py
D:\mailbox-alert-toolkit\generate_trends.py
D:\mailbox-alert-toolkit\config.json
D:\mailbox-alert-toolkit\README.md
D:\mailbox-alert-toolkit\README.docx
D:\mailbox-alert-toolkit\examples\...
```

### Passo 3: verifica che Python sia installato

Apri il prompt dei comandi e scrivi:

```
python --version
```

Se vedi qualcosa come Python 3.11.x o superiore sei a posto. Se invece esce un errore "Python non riconosciuto", devi installarlo:

1. scarica l'installer da https://www.python.org/downloads/
2. lancialo
3. importante: nella prima schermata spunta "Add python.exe to PATH"
4. clicca install now
5. al termine, riapri il prompt e verifica di nuovo

### Passo 4: configura config.json

Apri `D:\mailbox-alert-toolkit\config.json` con il blocco note (o un editor migliore come Visual Studio Code) e verifica i valori. Il file è già preconfigurato per intrawelt.com, ma puoi cambiare a piacere:

- la soglia di warning (`warningPercent`, default 80)
- il numero di file Excel da mantenere (`maxFilesPerType`, default 365)
- l'indirizzo del destinatario del report IT (`mail.to`)

### Passo 5: primo avvio del launcher

Apri PowerShell come amministratore:

1. premi il tasto Windows
2. scrivi powershell
3. clicca col destro su "Windows PowerShell" e scegli "Esegui come amministratore"
4. accetta il prompt UAC

Vai nella cartella del toolkit e lancia:

```powershell
cd D:\mailbox-alert-toolkit
powershell -ExecutionPolicy Bypass -File .\launcher.ps1
```

Oppure abilita una volta per tutte gli script firmati:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

Al primo avvio il launcher esegue in sequenza:

1. crea la struttura di cartelle (`reports\user`, `reports\other`, `history`, `logs`, `credentials`)
2. installa i moduli PowerShell necessari (`ExchangeOnlineManagement`, `Microsoft.Graph.Users`, `Microsoft.Graph.Authentication`) nell'ambito CurrentUser
3. verifica che Python sia presente
4. crea il virtual environment Python in `D:\mailbox-alert-toolkit\venv\`
5. aggiorna pip dentro il venv
6. installa openpyxl dentro il venv
7. scrive il marker `.setup_complete`
8. chiama `mailbox-alert.ps1`, che apre tre prompt interattivi:
    - **login Microsoft Graph**: si apre una finestra del browser per il login moderno. Usa l'account amministratore configurato in config.json (tenant.adminUpn). La prima volta in assoluto Graph chiederà il consenso ai permessi `User.Read.All` e `Organization.Read.All`. Clicca Accept
    - **login Exchange Online**: stessa cosa
    - **credenziali SMTP**: una finestra di Windows chiede username e password. Importante: qui inserisci le credenziali di `mailer@intrawelt.com`, NON di it@intrawelt.com. Vengono salvate cifrate nel file `credentials\smtp-cred.xml`, leggibile solo dallo stesso utente Windows che le ha salvate sulla stessa macchina (protezione DPAPI)

Tutto questo richiede una decina di minuti il primo giro, principalmente per il download dei moduli PowerShell. Le esecuzioni successive partono in pochi secondi.

### Verifica che sia andato tutto bene

Al termine vedi un messaggio finale verde "[OK] Esecuzione completata." e dovresti aver ricevuto una mail su it@intrawelt.com con due allegati Excel. Apri uno dei due e verifica che ci siano le caselle, le percentuali, e che le colorazioni (verde, giallo, rosso) abbiano senso.

Se qualcosa va storto, leggi il file di log in `D:\mailbox-alert-toolkit\logs\run_YYYY-MM-DD.log`.

---

## 7. Come eseguirlo manualmente

Una volta fatto il setup iniziale, per qualsiasi esecuzione successiva ti basta:

```powershell
cd D:\mailbox-alert-toolkit
.\launcher.ps1
```

Il launcher si accorge che il setup è già completato e parte direttamente con l'analisi delle mailbox.

Per generare il report trends storico:

```powershell
.\launcher.ps1 -Trends
```

Questa modalità non manda mail, genera solo il file `reports\Trends_YYYY-MM-DD.xlsx`.

---

## 8. Schedulazione con utilità di pianificazione di Windows 11

Per farlo girare automaticamente, ti consiglio la modalità "importa attività" da XML: è ripetibile, documentata, non richiede di cliccare 30 caselle, ed è quello che è stato fatto nell'installazione di riferimento.

In alternativa puoi creare il task a mano ("crea attività") con i parametri equivalenti, ma è facile sbagliare. Le impostazioni critiche sono:

- `Esegui indipendentemente dalla connessione utente` (logon type = Password, non solo "Quando l'utente è connesso"): permette al task di girare anche quando nessuno è loggato fisicamente sulla macchina, fondamentale per task pianificati di notte
- `Esegui con i privilegi più elevati`
- nella scheda **Impostazioni** lasciare DISATTIVATO `Usa motore di pianificazione unificato` (in inglese: "Use new scheduling engine"). Con `LogonType = Password` questa combinazione causa problemi
- il task deve girare con lo STESSO utente Windows che ha generato il certificato dell'app registration e salvato le credenziali SMTP, perché:
  - il certificato è in `Cert:\CurrentUser\My` di quel profilo
  - il file `credentials\smtp-cred.xml` è cifrato con DPAPI di quel profilo
  - se cambi utente, sia il certificato che le credenziali smettono di funzionare

### XML del task giornaliero

Salva il blocco sotto come `MailboxAlertIntrawelt.xml` da qualche parte (per esempio nella stessa cartella `D:\mailbox-alert-toolkit\`). Modifica il `<UserId>` (vedi paragrafo successivo per come trovarlo) e poi importalo nell'utilità di pianificazione.

```xml
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2026-05-21T16:17:42.7214953</Date>
    <Author>PC-ALESSIO\Utente</Author>
    <Description>Report giornaliero spazio mailbox tenant intrawelt.com</Description>
    <URI>\Mailbox alert Intrawelt</URI>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>2026-05-22T07:00:00</StartBoundary>
      <Enabled>true</Enabled>
      <RandomDelay>PT15M</RandomDelay>
      <ScheduleByDay>
        <DaysInterval>1</DaysInterval>
      </ScheduleByDay>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-21-3006074265-4287421424-1368263433-1000</UserId>
      <LogonType>Password</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>false</UseUnifiedSchedulingEngine>
    <WakeToRun>true</WakeToRun>
    <ExecutionTimeLimit>PT2H</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT15M</Interval>
      <Count>3</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell.exe</Command>
      <Arguments>-ExecutionPolicy Bypass -NonInteractive -WindowStyle Hidden -File "D:\mailbox-alert-toolkit\launcher.ps1"</Arguments>
      <WorkingDirectory>D:\mailbox-alert-toolkit</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
```

Punti chiave del XML giornaliero:

- `StartBoundary 2026-05-22T07:00:00` + `RandomDelay PT15M`: parte fra le 07:00 e le 07:15 ogni giorno (il random delay serve a non avere tanti task che partono allo stesso microsecondo)
- `UseUnifiedSchedulingEngine: false`: critico, non toccarlo
- `WakeToRun: true`: se la macchina è in sospensione la sveglia per eseguire (se invece è spenta non parte)
- `ExecutionTimeLimit: PT2H`: il task non gira più di 2 ore (anche se in realtà ne basta 5-10 minuti)
- `RunOnlyIfNetworkAvailable: true`: senza rete non ha senso provare
- `RestartOnFailure: 15 min, 3 volte`: se fallisce, ritenta dopo 15 minuti per un massimo di 3 volte
- gli argomenti `-NonInteractive -WindowStyle Hidden` sono OK perché lo script usa cert auth (vedi sezione 5), non richiede mai interazione utente

### Come trovare il SID dell'utente Windows

L'XML usa il SID (Security Identifier) dell'utente nel campo `<UserId>`. Per scoprire il tuo:

```powershell
whoami /user
```

Stampa qualcosa tipo:

```
NOME UTENTE        SID
================== ==============================================
pc-alessio\utente  S-1-5-21-3006074265-4287421424-1368263433-1000
```

Copia il valore della colonna SID e mettilo nel XML al posto di quello di esempio.

In alternativa puoi lasciare il campo come `<UserId>NOMEPC\NOMEUTENTE</UserId>` (es. `<UserId>PC-ALESSIO\Utente</UserId>`) e Windows lo risolve automaticamente all'import.

### XML del task settimanale per i trends

Salva come `MailboxTrendsIntrawelt.xml` e importa allo stesso modo:

```xml
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2026-05-21T17:35:00.0000000</Date>
    <Author>PC-ALESSIO\Utente</Author>
    <Description>Report trends settimanale spazio mailbox tenant intrawelt.com (analytics storiche)</Description>
    <URI>\Mailbox trends Intrawelt</URI>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>2026-05-25T08:00:00</StartBoundary>
      <Enabled>true</Enabled>
      <RandomDelay>PT15M</RandomDelay>
      <ScheduleByWeek>
        <DaysOfWeek>
          <Monday />
        </DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-21-3006074265-4287421424-1368263433-1000</UserId>
      <LogonType>Password</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>false</UseUnifiedSchedulingEngine>
    <WakeToRun>true</WakeToRun>
    <ExecutionTimeLimit>PT30M</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT15M</Interval>
      <Count>3</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell.exe</Command>
      <Arguments>-ExecutionPolicy Bypass -NonInteractive -WindowStyle Hidden -File "D:\mailbox-alert-toolkit\launcher.ps1" -Trends</Arguments>
      <WorkingDirectory>D:\mailbox-alert-toolkit</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
```

Differenze rispetto al task giornaliero:

- `URI: \Mailbox trends Intrawelt` (nome univoco)
- `ScheduleByWeek` con `<Monday />`: lunedì invece che giornaliero
- `StartBoundary 2026-05-25T08:00:00`: il primo lunedì utile alle 08:00
- `Arguments` aggiunge ` -Trends` in coda
- `ExecutionTimeLimit: PT30M`: 30 minuti bastano, il trends non si connette a niente
- `RunOnlyIfNetworkAvailable: false`: il trends legge solo dal database locale

### Come importare il task

1. apri **utilità di pianificazione** (premi Windows, scrivi "utilità di pianificazione")
2. nel pannello di destra (azioni) clicca **Importa attività**
3. seleziona il file XML salvato in precedenza
4. si apre il dialog del task con tutti i campi precompilati
5. nella scheda **Generale** verifica che l'utente associato sia quello giusto. Se ti dice "Utente sconosciuto" o se il SID è diverso, clicca **Cambia utente o gruppo** e seleziona il tuo
6. nella scheda **Trigger** verifica orari e ricorrenza
7. nella scheda **Azioni** verifica che il percorso `D:\mailbox-alert-toolkit\launcher.ps1` sia corretto (modificalo se hai installato altrove)
8. clicca OK
9. Windows ti chiede la password dell'utente associato: inseriscila

A questo punto vedi il task nella libreria con stato "Pronto" e la prossima esecuzione pianificata.

### Test manuale del task subito dopo l'import

Senza aspettare il prossimo trigger:

1. seleziona il task nella lista
2. click destro → **Esegui**
3. la colonna "Stato" diventa "In esecuzione"
4. dopo qualche minuto (5-8 per il task giornaliero, meno per il trends), controlla:
    - colonna "Esito ultima esecuzione" → deve essere `(0x0)` = successo
    - scheda "Cronologia" → l'ultimo evento "Task completed" deve avere codice operativo `(0)` invece di `(2)`
    - per il task giornaliero: nuova mail su `it@intrawelt.com` e nuovi file in `reports\user\` e `\other\`
    - per il task trends: nuovo file `reports\Trends_YYYY-MM-DD.xlsx`
    - in `logs\run_YYYY-MM-DD.log` deve esserci la riga `[OK] Esecuzione completata.`

### Abilitare la cronologia del task scheduler

Di default in Windows 11 la cronologia degli eventi del task scheduler è disabilitata a livello globale. Per attivarla (consigliato per debug):

1. apri utilità di pianificazione
2. pannello azioni a destra → **Abilita tutta la cronologia attività**
3. conferma il prompt UAC

Da quel momento ogni task (questo e qualsiasi altro) traccia eventi di start/stop, errori, codici operativi, visibili nella scheda "Cronologia" del task. Impostazione globale, non interferisce con nulla.

### Lo stato "In esecuzione" appare bloccato

Capita che l'UI dell'utilità di pianificazione mostri "In esecuzione" anche dopo che il task è già terminato (cache UI). Per verificare se è davvero in esecuzione:

1. apri Task Manager (Ctrl+Shift+Esc), tab **Dettagli**, cerca processi `powershell.exe` associati all'utente del task
2. se non ci sono processi, il task è già finito, basta premere F5 sull'utilità di pianificazione o cambiare cartella e tornare per far aggiornare lo stato
3. nella scheda "Cronologia" del task vedi l'evento "Task completed" con il timestamp reale di fine

---

## 9. Struttura della cartella di lavoro

Dopo qualche giorno di esecuzione la cartella sarà popolata così:

```
D:\mailbox-alert-toolkit\
    launcher.ps1
    mailbox-alert.ps1
    generate_mailbox_report.py
    generate_trends.py
    config.json
    README.md
    README.docx
    .setup_complete                    marker JSON con data del setup
    venv\                              virtual environment Python (~30 MB)
        Scripts\python.exe
        Lib\site-packages\openpyxl\...
    reports\
        user\
            Report_UserMailbox_2026-05-20.xlsx
            ... (max 365 file)
        other\
            Report_AltreMailbox_2026-05-20.xlsx
            ... (max 365 file)
        Trends_2026-05-20.xlsx         se hai lanciato -Trends
    history\
        mailbox_history.db             database SQLite con tutto lo storico
    logs\
        run_2026-05-20.log             vecchi log oltre 90 giorni vengono ripuliti
    credentials\
        smtp-cred.xml                  credenziali SMTP cifrate DPAPI
    examples\                          campioni demo (puoi cancellarli)
```

---

## 10. File di configurazione config.json

Tutte le impostazioni in un unico file, modificabile con un editor di testo. Non serve toccare il codice degli script.

```json
{
  "tenant": {
    "domain":   "intrawelt.com",
    "adminUpn": "it@intrawelt.com"
  },
  "appRegistration": {
    "enabled":               true,
    "clientId":              "00000000-0000-0000-0000-000000000000",
    "tenantId":              "00000000-0000-0000-0000-000000000000",
    "certificateThumbprint": "INSERIRE_THUMBPRINT_CERTIFICATO",
    "organization":          "esempio.onmicrosoft.com"
  },
  "mail": {
    "to":         "it@intrawelt.com",
    "from":       "mailer@intrawelt.com",
    "smtpServer": "smtp.office365.com",
    "smtpPort":   587
  },
  "thresholds": {
    "warningPercent":  80,
    "criticalPercent": 95
  },
  "retention": {
    "maxFilesPerType": 365,
    "keepHistoryDb":   true
  },
  "collection": {
    "includeFolderStats":   true,
    "includeForwardingInfo": true,
    "includeHoldInfo":       true
  },
  "notifications": {
    "notifyMailboxOwners":   true,
    "skipResourceMailboxes": true,
    "itSupportAddress":      "it@intrawelt.com"
  }
}
```

Significato dei campi:

- tenant.domain: dominio del tenant, etichetta nei report e nelle mail
- tenant.adminUpn: UPN amministratore usato per login interattivo se appRegistration è disabilitato
- appRegistration.enabled: se true, lo script si autentica via app registration con certificato (vedi sezione 5). Se false, fallback al login browser interattivo. DEVE essere true per task schedulati
- appRegistration.clientId: ID applicazione (client) dell'app registration in Entra
- appRegistration.tenantId: ID directory (tenant)
- appRegistration.certificateThumbprint: 40 caratteri esadecimali del certificato (stampato dal comando New-SelfSignedCertificate al Passo 3 della sezione 5)
- appRegistration.organization: dominio iniziale del tenant nel formato `xxx.onmicrosoft.com` (vedi Passo 6 della sezione 5)
- mail.to: destinatario del report giornaliero IT
- mail.from: mittente di tutte le mail in uscita (mailer@intrawelt.com con SMTP AUTH)
- mail.smtpServer e mail.smtpPort: smtp.office365.com e 587
- thresholds.warningPercent: default 80
- thresholds.criticalPercent: default 95
- retention.maxFilesPerType: default 365
- retention.keepHistoryDb: placeholder, lascialo true
- collection.includeFolderStats: raccoglie Inbox e Posta inviata (più chiamate API)
- collection.includeForwardingInfo: segnala caselle con inoltro attivo
- collection.includeHoldInfo: segnala caselle con litigation hold
- notifications.notifyMailboxOwners: default true. Mettilo a false per disabilitare temporaneamente le notifiche agli utenti
- notifications.skipResourceMailboxes: default true. Salta sale/equipment
- notifications.itSupportAddress: indirizzo mostrato nelle notifiche all'utente

Dopo aver modificato config.json non serve riavviare nulla.

---

## 11. Metriche raccolte per ogni casella

Per ogni mailbox vengono raccolte 22 metriche, che finiscono sia negli Excel del giorno sia nel database SQLite.

| colonna | descrizione | utilità |
|---|---|---|
| Casella | indirizzo email primario | identificatore univoco |
| Display Name | nome visualizzato | leggibilità report |
| Tipo | UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox, SchedulingMailbox | separazione categorie |
| Licenza | licenze Microsoft 365 assegnate (nome leggibile, vedi sezione 20) | cost analysis per casella |
| Creata il | data creazione mailbox | distinguere mailbox storiche da nuove |
| Usato (GB) | spazio totale occupato della mailbox principale | metrica principale |
| Quota (GB) | limite ProhibitSendQuota | calcolo della percentuale |
| % Occupato | percentuale di riempimento mailbox principale | trigger del warning mailbox |
| Δ 30gg (GB) | differenza in GB rispetto a 30 giorni fa | trend di crescita |
| N° Email | conteggio totale messaggi | densità contenuti |
| Inbox (GB) | dimensione totale cartella posta in arrivo | chi non smista la posta |
| Inviati (GB) | dimensione totale cartella posta inviata | volume outbound |
| Eliminati (GB) | spazio occupato dai deleted items | comportamento di pulizia |
| Arch. Att. | archivio online abilitato (Si'/No) | feature in uso |
| Arch. Usato (GB) | spazio occupato nell'archivio online | gestione archive |
| Arch. Quota (GB) | quota archivio | calcolo % archivio |
| % Archivio | percentuale occupata nell'archivio | trigger del warning archivio |
| Ultimo Accesso | timestamp ultimo accesso utente | individuare caselle abbandonate |
| Inattività (gg) | giorni passati dall'ultimo accesso | metrica di inattività |
| Inoltro | inoltro automatico attivo (sì, sì con copia, no) | sicurezza, regole sospette |
| Lit. Hold | litigation hold attivo | compliance |
| Hidden GAL | mailbox nascosta dalla rubrica globale | governance |

La colonna **% Occupato** decide il colore di sfondo della riga (mailbox principale): verde sotto 80%, giallo tra 80 e 95%, rosso sopra 95%.

La colonna **Δ 30gg (GB)** ha colorazione propria: arancione se cresciuta di più di 2 GB in 30 giorni, gialla tra 0.5 e 2, azzurra se diminuita, neutra altrimenti.

La colonna **% Archivio** è colorata indipendentemente dalla mailbox principale, perché può capitare che la casella sia ok ma l'archivio sia pieno.

---

## 12. I due report Excel giornalieri

Ogni esecuzione produce due file:

```
reports\user\Report_UserMailbox_2026-05-20.xlsx
reports\other\Report_AltreMailbox_2026-05-20.xlsx
```

Caratteristiche:

- una sola scheda con 22 colonne (vedi sopra)
- intestazione blu fissata in cima quando scorri (freeze panes attivo)
- autofiltro su tutte le colonne per ordinamento e ricerca
- riga di riepilogo in fondo con conteggi separati: `Totale: X | Mailbox: A warning, B critical | Archivio: C warning, D critical (su N con archivio) | Inattive >=90gg: Z`
- font Arial size 9 per il corpo, size 10 grassetto per l'header
- colonne larghezza ottimizzata per stampa A4 orizzontale

Il file Excel `Report_UserMailbox_*` contiene solo le caselle di tipo UserMailbox (quelle con licenza). Il file `Report_AltreMailbox_*` contiene tutto il resto: shared, room, equipment, scheduling.

---

## 13. Notifiche personali agli utenti (mailbox separata dall'archivio)

Oltre alla mail di report per IT, ogni casella che supera la soglia riceve una mail di notifica personale, mandata da `mailer@intrawelt.com` direttamente al proprietario.

Le notifiche di mailbox principale e archivio sono TRATTATE COME DUE EVENTI SEPARATI, perché sono problemi tecnici diversi:

- mailbox principale piena: blocca la ricezione di nuovi messaggi. Subject e contenuto della mail enfatizzano questo impatto
- archivio online pieno: non blocca la ricezione, ma blocca l'archiviazione di altri messaggi. Subject e contenuto sono diversi, evidenziano lo specifico problema

### Comportamento

| stato della casella | mail ricevute dall'utente |
|---|---|
| solo mailbox sopra soglia | 1 mail: "Avviso/URGENTE - La tua casella di posta..." |
| solo archivio sopra soglia | 1 mail: "Avviso/URGENTE - Il tuo archivio online..." |
| entrambi sopra soglia | 2 mail separate, una per problema |
| nessuno sopra soglia | nessuna mail |

### Tono adattivo: warning vs critical

- tra 80% e 94%: subject "Avviso", banner giallo (mailbox) o viola chiaro (archivio), tono di promemoria
- al 95% o oltre: subject "URGENTE", banner rosso (mailbox) o viola scuro (archivio), tono di intervento immediato

Il valore della soglia critical viene letto da `thresholds.criticalPercent` in config.json, default 95.

### Contenuto della mail al destinatario

Mail in italiano, HTML semplice, contiene:

- saluto personalizzato con il display name della persona
- riga con valore e percentuale specifici del problema (mailbox o archivio, non entrambi)
- frase esplicativa che chiarisce ALL'UTENTE che la mail riguarda solo uno dei due ambiti (e che, se il caso, riceverà anche una notifica separata per l'altro)
- lista di azioni suggerite specifiche per il tipo di problema
- riferimento al supporto IT
- NESSUN allegato Excel: quelli sono solo per IT

### Caselle escluse dalla notifica

Per default vengono escluse:

- RoomMailbox (sale riunioni)
- EquipmentMailbox (attrezzature)
- SchedulingMailbox

Queste non hanno destinatario umano, inviarle è solo rumore. Le shared mailbox invece ricevono la notifica normalmente, perché tipicamente sono presidiate da utenti delegati. Se vuoi includere anche le risorse, metti `skipResourceMailboxes` a `false` in config.json.

### Disabilitare temporaneamente le notifiche

Metti `notifyMailboxOwners` a `false` in config.json. La prossima esecuzione manderà solo la mail al reparto IT.

### Cosa controllare se le notifiche non arrivano

1. controlla il log in `logs\run_*.log`: la sezione finale elenca quante notifiche mailbox e quante archivio sono state inviate, eventuali errori per casella
2. verifica che `mailer@intrawelt.com` esista, abbia licenza, abbia SMTP AUTH abilitato, e non abbia MFA bloccante
3. verifica nella sezione monitoraggio messaggi dell'Exchange admin center che le mail siano effettivamente uscite

### Cosa vede IT nella mail di report

La mail al reparto IT contiene due tabelle distinte:

- prima tabella (header blu): "Caselle con mailbox principale sopra il 80%". Mostra casella, tipo, usato, quota, percentuale
- seconda tabella (header viola): "Archivi online sopra il 80%". Mostra casella, tipo, usato archivio, quota archivio, percentuale archivio

Il subject della mail IT è tipo `[Mailbox: 3 | Archivio: 8] Report Mailbox intrawelt.com - 2026-05-20`, così a colpo d'occhio sai cosa sta arrivando.

---

## 14. Lo storico SQLite e perché esiste

Il file `history\mailbox_history.db` è un database SQLite che accumula uno snapshot completo per ogni esecuzione. Ogni riga rappresenta una casella in una data.

Schema della tabella `mailbox_history`:

- chiave primaria composta da `snapshot_date` + `email` (idempotente: rilanciare lo script nello stesso giorno sovrascrive il record)
- 22 colonne corrispondenti alle metriche raccolte
- due indici: per data e per email

### Come consultare il database

Scarica un client gratuito come DB Browser for SQLite da https://sqlitebrowser.org/, aprilo, fai file > apri database > scegli `history\mailbox_history.db`. Da lì puoi scorrere tutte le tabelle, esportare in CSV, fare query SQL personalizzate.

Esempi di query utili:

```sql
-- crescita media giornaliera del tenant negli ultimi 90 giorni
select snapshot_date, round(sum(used_gb), 2) as totale_mailbox_gb,
                     round(sum(case when lower(archivio_abilitato) like 's%'
                                    then archivio_used_gb else 0 end), 2) as totale_archivio_gb
from mailbox_history
where snapshot_date >= date('now', '-90 days')
group by snapshot_date
order by snapshot_date;

-- top 10 caselle per crescita mailbox negli ultimi 30 giorni
with primo as (
  select email, used_gb from mailbox_history
  where snapshot_date = (select min(snapshot_date) from mailbox_history
                          where snapshot_date >= date('now', '-30 days'))
), ultimo as (
  select email, used_gb from mailbox_history
  where snapshot_date = (select max(snapshot_date) from mailbox_history)
)
select u.email, round(u.used_gb - p.used_gb, 2) as crescita_gb
from primo p join ultimo u using(email)
order by crescita_gb desc
limit 10;

-- caselle inattive che occupano più di 5 GB (candidati per archiviazione)
select email, display_name, used_gb, archivio_used_gb, giorni_inattivita
from mailbox_history
where snapshot_date = (select max(snapshot_date) from mailbox_history)
  and giorni_inattivita >= 90
  and used_gb > 5
order by used_gb desc;

-- archivi più pieni
select email, archivio_used_gb, archivio_quota_gb, pct_archivio
from mailbox_history
where snapshot_date = (select max(snapshot_date) from mailbox_history)
  and lower(archivio_abilitato) like 's%'
order by pct_archivio desc
limit 20;
```

---

## 15. Il report trends

Lo lanci on-demand con:

```powershell
.\launcher.ps1 -Trends
```

Genera un file `reports\Trends_YYYY-MM-DD.xlsx` con 8 fogli, separando mailbox e archivio come due dimensioni distinte:

1. **Riepilogo Tenant**: una riga per ogni giorno presente nel database, con numero caselle, totale spazio mailbox, percentuale media mailbox, conteggio mailbox in warning e critical. In più: numero archivi attivi, totale spazio archivio, percentuale media archivio, conteggio archivi in warning e critical. In alto a destra un grafico a linee con l'evoluzione dello spazio totale mailbox vs archivio nel tempo
2. **Top Crescita Mailbox 30gg**: le 30 caselle che sono cresciute di più (mailbox principale) negli ultimi 30 giorni
3. **Top Crescita Mailbox 12mesi**: stesso ma su 12 mesi
4. **Top Crescita Archivio 30gg**: top 30 per crescita archivio negli ultimi 30 giorni (solo caselle con archivio attivo)
5. **Top Crescita Archivio 12mesi**: top 30 per crescita archivio negli ultimi 12 mesi
6. **Inattive**: tutte le caselle con più di 90 giorni di inattività, con spazio occupato mailbox E archivio. Quelle con totale > 5 GB sono evidenziate in giallo (candidati per archiviazione o dismissione)
7. **Storico Mailbox per Casella**: pivot completa, una riga per casella, una colonna per data, valore in GB della mailbox principale
8. **Storico Archivio per Casella**: pivot della stessa forma ma per l'archivio (solo caselle con archivio attivo)

I fogli "Top Crescita" mostrano per ogni casella:

- data e valore in GB del punto iniziale (il più vecchio snapshot nella finestra)
- data e valore in GB del punto finale (l'ultimo snapshot disponibile)
- delta in GB (valore assoluto della crescita)
- delta in percentuale rispetto al valore iniziale
- delta medio per giorno in MB, calcolato sui giorni effettivi tra i due snapshot

Note matematiche importanti:

- il riferimento "oggi" usato dai Top Crescita è il MAX(snapshot_date) del database, non la data di sistema. Quindi se il task non gira da qualche giorno, il trends rimane comunque coerente con i dati effettivi
- il "delta medio per giorno" usa i giorni EFFETTIVI tra il primo e l'ultimo snapshot della finestra. Se il database ha 60 giorni di storia e tu chiedi i "Top Crescita 12 mesi", il calcolo usa 60 giorni come divisore, non 365

Il file è autocontenuto: puoi mandarlo in mail a un manager, salvarlo come allegato a un report mensile, condividerlo per analisi di capacity e costi.

### Invio automatico del trends via mail (opzionale)

Di default `launcher.ps1 -Trends` genera solo il file in `reports\Trends_YYYY-MM-DD.xlsx`, senza inviarlo via mail. Se vuoi che il trends settimanale arrivi automaticamente come allegato sulla casella IT (riusando le stesse credenziali SMTP già configurate per le mail giornaliere), basta sostituire il blocco "Modalita' Trends" in `launcher.ps1` con quello esteso documentato sotto.

Cerca nel file `launcher.ps1` queste righe (verso la fine, intorno alla riga 176):

```powershell
# Modalita' Trends (analytics on-demand)
if ($Trends) {
    Write-Host "`n-> Generazione report TRENDS storici..." -ForegroundColor Cyan
    $historyDb  = Join-Path $ScriptRoot "history\mailbox_history.db"
    $trendsXlsx = Join-Path $ScriptRoot ("reports\Trends_$(Get-Date -Format 'yyyy-MM-dd').xlsx")
    & $venvPython $trendsScript $historyDb $trendsXlsx
    Write-Host "[OK] Report trends generato: $trendsXlsx" -ForegroundColor Green
    exit 0
}
```

E sostituiscile con questo blocco esteso:

```powershell
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
```

Cosa cambia rispetto al codice originale:

- aggiunta verifica esplicita di errore alla generazione del file Excel (se Python fallisce, esce con codice 1 invece di proseguire silenziosamente)
- aggiunto blocco di invio mail che riusa: il `config.json` per destinatari/SMTP, e il file `credentials\smtp-cred.xml` già esistente (cifrato DPAPI dal task giornaliero)
- se le credenziali SMTP non esistono ancora, il task termina con un warning ma exit 0 (il file Excel è comunque salvato)

Dopo la modifica, alla prossima esecuzione di `launcher.ps1 -Trends` (manuale o schedulata):

1. genera il file `reports\Trends_YYYY-MM-DD.xlsx`
2. invia una mail a `it@intrawelt.com` con quel file in allegato e subject `Report Trends storici mailbox intrawelt.com - YYYY-MM-DD`

Lasci il file Excel locale comunque salvato per storico, oltre alla copia in allegato.

---

## 16. Retention policy dei file Excel

Alla fine di ogni esecuzione, il main script applica la retention:

1. legge la lista di tutti i .xlsx in `reports\user\`
2. li ordina per data di modifica decrescente (dal più recente al più vecchio)
3. se ce ne sono più di `maxFilesPerType` (default 365), cancella tutti quelli oltre il 365° posto
4. stessa cosa per `reports\other\`

Il database SQLite NON viene mai pulito: conserva tutto lo storico finché c'è spazio su disco. Un anno di esecuzioni giornaliere per casella pesa pochi kilobyte. Per un tenant con 100 caselle e 5 anni di storico parliamo di pochi megabyte.

I log oltre i 90 giorni vengono cancellati automaticamente a fine esecuzione.

Per cambiare la retention, modifica `maxFilesPerType` in config.json. Non scendere sotto 2 (altrimenti hai solo l'attuale e ieri). Per disattivare la retention metti un numero molto grande, tipo 99999.

---

## 17. Portabilità del virtual environment Python

Il virtual environment è la cartella `venv\` accanto agli script. Contiene una copia isolata di Python con dentro openpyxl. Pesa circa 30 MB.

Vantaggi:

- nulla viene installato a livello di sistema
- puoi spostare l'intera cartella `MailboxAlert\` (compreso il venv) su un altro disco, su una chiavetta, su un'altra macchina, e funziona, a patto che la macchina destinazione abbia una versione major di Python compatibile (3.10, 3.11, 3.12) installata
- per disinstallare il toolkit basta cancellare l'intera cartella

Se la versione di Python cambia (es. passi da 3.11 a 3.12), il venv potrebbe smettere di funzionare. Lancia `.\launcher.ps1 -ForceSetup` per ricreare il venv pulito.

---

## 18. Gestione delle credenziali e sicurezza

Le credenziali SMTP vengono chieste al primo avvio e salvate in `credentials\smtp-cred.xml`.

Il salvataggio usa `Export-Clixml`, che si appoggia a DPAPI di Windows: il file è cifrato e leggibile solo dallo stesso utente Windows, sulla stessa macchina, che lo ha creato. Anche un amministratore della macchina con un altro account non lo può decifrare.

Quindi:

- è ragionevolmente sicuro per uso interno
- non è portabile: se sposti la cartella su un'altra macchina, il file delle credenziali smetterà di funzionare e ti verrà richiesto di reinserire al primo avvio sulla nuova macchina

Per Exchange Online e Microsoft Graph la modalità raccomandata è l'autenticazione via certificato (app registration), descritta in dettaglio nella sezione 5. In quella modalità:

- nessun token utente da rinnovare, nessuna scadenza periodica
- il certificato self-signed sul PC vale 2 anni di default
- il task gira completamente unattended senza popup browser

Se invece `appRegistration.enabled` è impostato a `false` in config.json, lo script torna alla modalità login interattivo via browser. In quel caso i moduli Microsoft.Graph e ExchangeOnlineManagement mantengono un refresh token nella cache del profilo Windows dell'utente, normalmente valido 90 giorni. Quando scade, la prossima esecuzione del task schedulato fallisce perché non può aprire un browser. La fix è loggarsi sulla macchina con quell'utente, lanciare a mano `.\launcher.ps1` per rinfrescare il token, e da lì il task riprende per altri 90 giorni. Per evitare questo ciclo è preferibile usare cert auth.

---

## 19. Logging

Ogni esecuzione produce un file in `logs\run_YYYY-MM-DD.log` con il transcript completo: connessioni, ogni mailbox analizzata, errori, generazione Excel, retention, invio mail.

Quando vedi warning sulla console, in particolare:

- `[M]` significa "warning mailbox principale" per quella casella
- `[A]` significa "warning archivio" per quella casella

Una stessa casella può comparire con entrambi i marker se è sopra entrambe le soglie.

Stesso schema per il blocco di invio notifiche personali: due loop separati, uno per le notifiche mailbox e uno per le notifiche archivio. Il count finale stampa entrambi i totali.

I log oltre i 90 giorni vengono cancellati automaticamente.

---

## 20. Tutte le modalità del launcher

| comando | descrizione |
|---|---|
| `.\launcher.ps1` | comportamento normale: setup se serve, poi esecuzione completa |
| `.\launcher.ps1 -ForceSetup` | rifà tutto il setup da zero, ricrea il venv |
| `.\launcher.ps1 -SkipRun` | esegue solo il setup, non lancia il main |
| `.\launcher.ps1 -Trends` | non manda mail, genera solo il report analytics storico |

I parametri sono combinabili: `.\launcher.ps1 -ForceSetup -SkipRun` rifà setup pulito senza eseguire analisi.

---

## 21. Gestione delle licenze e SKU

SKU significa "Stock Keeping Unit", è il codice prodotto. In Microsoft 365 ogni licenza ha:

- uno **SkuId** che è un GUID lungo, univoco a livello globale (uguale per tutti i tenant del mondo)
- uno **SkuPartNumber** che è un codice testuale stabile, tipo `O365_BUSINESS_PREMIUM` o `SPB`

Per ottenere nomi leggibili nella colonna "Licenza" dei report, lo script:

1. al login Graph chiede anche lo scope `Organization.Read.All`
2. legge la lista degli SKU effettivamente acquistati dal tenant con `Get-MgSubscribedSku`
3. per ognuno, prende il SkuPartNumber e lo confronta con una mappa interna di nomi friendly (circa 70 SKU comuni mappati)
4. se uno SkuPartNumber non è nella mappa, lo script mostra il SkuPartNumber stesso (che è comunque leggibile, es. `EXCHANGESTANDARD`)

Allo step `[1/5] Caricamento mappa licenze` viene stampata in console:

- quanti SKU del tenant sono stati identificati
- la lista degli SkuPartNumber per cui non c'è ancora un nome friendly nella mappa interna

Se vedi nella console qualche SkuPartNumber mostrato così com'è (senza una versione friendly tipo "Microsoft 365 ..."), comunicalo e si può aggiungere alla mappa nello script. Lista completa SKU Microsoft: https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference

---

## 22. Note tecniche per gli sviluppatori (perché PowerShell si comporta così)

Questa sezione spiega i punti tecnici emersi durante lo sviluppo, utili se in futuro modifichi tu lo script.

### Encoding dei file .ps1

PowerShell 5.1 di Windows legge i file .ps1 come ANSI a meno che non abbiano il BOM (Byte Order Mark) all'inizio. Se in uno script ci sono caratteri UTF-8 (frecce, emoji, accenti) e il file non ha BOM, il parser interpreta i byte UTF-8 come singoli caratteri ANSI e fallisce con errori del tipo "Carattere di terminazione mancante nella stringa".

Soluzione: i file .ps1 del toolkit sono salvati con UTF-8 BOM E contengono solo caratteri ASCII puri (accenti italiani sostituiti con apostrofi: "è" → "e'", "più" → "piu'"). Questo li rende immuni a problemi di encoding su qualsiasi locale.

### Ordine di import dei moduli Microsoft.Graph ed Exchange Online

Il modulo `ExchangeOnlineManagement` precarica versioni vecchie degli assembly `Azure.Identity` e `Microsoft.Identity.Client`. Se carichi prima Exchange e poi Microsoft.Graph, Graph trova in memoria versioni vecchie e crasha con `TypeLoadException` su metodi tipo `GetTokenAsync`.

Soluzione: nel main script Graph viene importato e connesso PRIMA di ExchangeOnlineManagement. Così Graph carica i suoi assembly nuovi, e quando dopo carichi Exchange si adatta a quelli già presenti.

### ByteQuantifiedSize deserializzato

Il modulo ExchangeOnlineManagement moderno (versione 3.x+) usa REST API sotto al cofano e restituisce **oggetti deserializzati** (`Deserialized.Microsoft.Exchange.Data.ByteQuantifiedSize`) invece degli oggetti reali. Gli oggetti deserializzati conservano i valori ma non i metodi originali, quindi `.Value.ToBytes()` non esiste più.

I valori arrivano come stringhe formato `"49.5 GB (53,150,121,984 bytes)"` o `"Unlimited"`.

Soluzione: lo script ha una funzione `Get-SizeBytes` che parsa la stringa, gestendo tutti i casi (null, vuoto, Unlimited, formato con bytes in parentesi, formato senza parentesi con KB/MB/GB/TB). Usa `CultureInfo.InvariantCulture` per non dipendere dalla cultura del PC.

### Parentesi delle chiamate function dentro metodi .NET

In PowerShell, dentro una chiamata a un metodo .NET come `[math]::Round(...)`, il parser interpreta gli argomenti separati da virgole. Una function call PowerShell come `Get-SizeBytes $x` lì dentro confonde il parser:

```powershell
# Sbagliato: PowerShell vede 4 argomenti separati da spazio
[math]::Round(Get-SizeBytes $x / 1GB, 2)

# Giusto: la function call e' parenthesata
[math]::Round((Get-SizeBytes $x) / 1GB, 2)
```

Regola generale: ogni volta che passi a un metodo .NET il risultato di una function PowerShell, racchiudi la function call tra parentesi.

### Convention "Si'" invece di "Sì"

Visto che tutti gli accenti italiani sono stati sostituiti per ASCII compatibility, il campo `ArchivioAbilitato` nel JSON arriva al Python come `"Si'"` (con apostrofo), non `"Sì"`. Lo script Python usa una funzione `is_archive_enabled()` tollerante che matcha qualsiasi forma iniziando per "S" (Sì, Si, Si', yes, true).

### Versioni multiple di ExchangeOnlineManagement convivono

Su Windows 11 è frequente avere installata `ExchangeOnlineManagement 3.7.0` a livello sistema (`C:\Program Files\WindowsPowerShell\Modules\`) preinstallata o messa da update di Windows. Quando lanci `Update-Module ExchangeOnlineManagement -Force`, PowerShell non riesce a sovrascrivere la versione di sistema (richiede privilegi più alti che spesso non bastano) e installa la nuova versione a livello utente in `C:\Users\<nomeutente>\Documents\WindowsPowerShell\Modules\`. Quando il main script chiama `Import-Module ExchangeOnlineManagement`, PowerShell carica automaticamente la versione più recente disponibile (la nuova installata a livello utente), quindi tutto funziona. I warning del tipo `AVVISO: The version '2.2.5' of module 'PowerShellGet' is currently in use. Retry the operation after closing the applications` durante `Update-Module` sono ininfluenti: il modulo viene comunque aggiornato, semplicemente PowerShellGet stesso non si auto-aggiorna mentre è caricato.

### Errore IDX12729 con cert auth e ExchangeOnlineManagement 3.7.0

Con la versione 3.7.0 del modulo (precaricata da Windows) e autenticazione via certificato, può capitare l'errore `IDX12729: Unable to decode the header '[PII of type 'System.String' is hidden].' as Base64Url encoded string`. È un bug noto della libreria di parsing JWT contenuta in quella versione del modulo. La fix è aggiornare il modulo:

```powershell
Update-Module ExchangeOnlineManagement -Force
```

Dopo questo comando, accanto alla 3.7.0 viene installata una versione più recente (3.9.x) a livello utente, che PowerShell preferisce automaticamente al prossimo lancio.

---

## 23. Troubleshooting

### "IDX12729: Unable to decode the header" sulla connessione a Exchange Online

Bug del modulo `ExchangeOnlineManagement` 3.7.0 con cert auth. Fix:

```powershell
Update-Module ExchangeOnlineManagement -Force
Get-Module -ListAvailable ExchangeOnlineManagement | Select-Object Version, Path
```

Devi vedere due versioni: la 3.7.0 a livello sistema e la 3.9.x a livello utente. La 3.9.x viene preferita automaticamente. Vedi sezione 22 "Versioni multiple di ExchangeOnlineManagement convivono" per i dettagli.

### "InteractiveBrowserCredential authentication failed" nel log del task schedulato

Il task gira con `-NonInteractive` ma `Connect-MgGraph` sta tentando il login via browser, che non può aprirsi. Significa che `appRegistration.enabled` non è impostato a true in config.json, oppure i valori non sono corretti.

Controlla:

1. `D:\mailbox-alert-toolkit\config.json` ha la sezione `appRegistration` con `enabled: true` e tutti e quattro i campi compilati (clientId, tenantId, certificateThumbprint, organization)
2. il certificato è installato nel magazzino dell'utente Windows che fa girare il task. Lancia da PowerShell come quell'utente: `Get-ChildItem Cert:\CurrentUser\My | Where-Object Thumbprint -eq "INSERIRE_THUMBPRINT_CERTIFICATO"` e devi vedere il certificato
3. l'app registration in Entra ha il "consenso amministratore" concesso a tutti e tre i permessi (User.Read.All, Organization.Read.All, Exchange.ManageAsApp)
4. l'app registration ha il ruolo "Lettore globale" assegnato in Ruoli e amministratori

Vedi la sezione 5 per la procedura completa.

### Il certificato dell'app registration è scaduto

`Connect-MgGraph` o `Connect-ExchangeOnline` falliscono con un errore tipo "certificate has expired". I certificati self-signed creati dal Passo 3 della sezione 5 valgono 2 anni di default. Per rinnovare:

1. genera un nuovo certificato (rifai il Passo 3 della sezione 5)
2. carica la nuova parte pubblica sull'app registration (rifai il Passo 4)
3. aggiorna `appRegistration.certificateThumbprint` in config.json con il nuovo thumbprint
4. il vecchio certificato lo puoi lasciare sull'app come backup, oppure rimuoverlo dal portale e dal magazzino certificati di Windows



L'execution policy di PowerShell è restrittiva. Apri PowerShell come amministratore e lancia una volta:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

oppure usa sempre `-ExecutionPolicy Bypass` quando lanci il launcher.

### "python non riconosciuto come comando interno o esterno"

Python non è nel PATH. Riavvia il prompt, oppure reinstalla Python spuntando "Add to PATH".

### "Connect-MgGraph: insufficient privileges to complete the operation"

L'account amministratore non ha i permessi per leggere licenze o SubscribedSku. Serve il ruolo Global Reader oppure User Administrator + accesso a Organization in Graph. Chiedi a un Global Admin di assegnarti il ruolo o di pre-approvare il consenso ai permessi `User.Read.All` e `Organization.Read.All`.

### "SMTP authentication required but disabled"

La casella mailer@intrawelt.com ha SMTP AUTH disattivato. Riattivalo come spiegato nella sezione 4.

### "Send-MailMessage cmdlet is obsolete"

È un warning, non un errore. Il comando funziona ancora su PowerShell 5.1 e 7.x. In futuro si potrà migrare a Send-MgUserMail (Graph API).

### Le credenziali SMTP non vengono lette quando gira il task

Il task sta girando con un utente Windows diverso da quello che ha salvato le credenziali. Lancia il setup con l'utente del task, oppure cambia l'utente del task in utilità di pianificazione.

### Il task gira ma non riceviamo la mail

Controlla in ordine:

1. il log in `logs\run_*.log` per vedere se l'invio risulta riuscito
2. la cartella posta indesiderata della casella it@intrawelt.com
3. il pannello di Exchange admin center, sezione monitoraggio messaggi, per vedere se la mail è stata accettata e consegnata
4. che SMTP AUTH sia abilitato sulla casella mittente

### Le notifiche personali non arrivano agli utenti, ma IT riceve il report

Possibili cause in ordine di probabilità:

1. `notifyMailboxOwners` è messo a false in config.json
2. la casella mailer@intrawelt.com non ha SMTP AUTH abilitato o ha MFA che blocca
3. nel log run_*.log la sezione "Notifiche personali" elenca gli errori specifici
4. la casella destinataria è un room/equipment/scheduling e quindi viene saltata di default

### Un utente si lamenta di ricevere troppi alert

Le notifiche partono ogni volta che il task viene eseguito (di default ogni giorno) finché la casella resta sopra l'80%. Se vuoi attenuare:

- aumenta `warningPercent` da 80 a 85
- contatta l'utente direttamente per aiutarlo a fare pulizia
- metti temporaneamente `notifyMailboxOwners` a false

### Le license appaiono come GUID lunghi anziché nomi leggibili

Sta succedendo che `Get-MgSubscribedSku` non viene eseguito (manca lo scope Organization.Read.All) oppure lo SkuPartNumber non è nella mappa interna dello script. Controlla la sezione 20 e il log all'avvio: c'è la lista degli SKU sconosciuti che si possono aggiungere alla mappa.

### Voglio resettare tutto e ripartire pulito

Cancella `.setup_complete`, la cartella `venv\` e la cartella `credentials\`, poi rilancia:

```powershell
.\launcher.ps1
```

Lo storico (`history\mailbox_history.db`) e i report già generati restano intatti.

### Voglio cancellare anche lo storico

Cancella anche `history\mailbox_history.db` e tutti gli xlsx in `reports\user\` e `reports\other\`.

### Quanto dura un'esecuzione tipica

Su intrawelt (circa 100 caselle) circa 5-8 minuti. Il tempo è dominato dalle chiamate API a Exchange Online. Se vuoi velocizzare, metti `includeFolderStats` a false: perdi le colonne Inbox (GB) e Inviati (GB) ma guadagni metà del tempo.

---

## 24. Manutenzione e aggiornamenti

### Cose da controllare ogni tanto

- una volta al mese, controlla che la mail arrivi davvero
- ogni 60-80 giorni esegui manualmente il launcher con l'utente che fa girare il task per rinfrescare il token di Exchange/Graph
- ogni 6 mesi controlla la dimensione di `history\mailbox_history.db`: se cresce in modo anomalo, qualcosa non va

### Aggiornamento moduli PowerShell

```powershell
Update-Module ExchangeOnlineManagement -Force
Update-Module Microsoft.Graph.Users -Force
Update-Module Microsoft.Graph.Authentication -Force
```

### Aggiornamento di Python

Se aggiorni Python sulla macchina, lancia `.\launcher.ps1 -ForceSetup` per ricreare il venv con la nuova versione.

### Aggiornamento del toolkit

Quando arriva una nuova versione del toolkit:

1. backup di `config.json` e di `history\mailbox_history.db`
2. sovrascrivi i file .ps1 e .py con quelli nuovi
3. lancia `.\launcher.ps1 -ForceSetup -SkipRun` per riallineare il venv
4. ricarica il backup di `config.json` se hai personalizzazioni
5. al primo run normale, se lo schema del database è cambiato, il generatore aggiunge le colonne nuove automaticamente

Per qualsiasi dubbio o estensione (autenticazione tramite app registration, invio del trends via mail settimanale, dashboard Power BI sopra il database SQLite, aggiunta SKU mancanti alla mappa), chiedi.
