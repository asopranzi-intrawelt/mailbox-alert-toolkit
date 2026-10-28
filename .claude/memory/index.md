# Snapshot di sincronizzazione

> Da leggere per primo a inizio sessione. Fotografa lo stato del progetto al commit di riferimento e mappa ogni scheda al suo stato di verifica. È la fonte di verità su cosa è fatto, non le spunte del diario.

## Stato

```
Branch attivo:         main
Commit di riferimento: ec9883f
Data snapshot:         2026-10-28
```

## Stato di verifica delle schede

| Scheda | last-verified | Stato |
|---|---|---|
| STACK.md | 152deac | aggiornata |
| design-and-security.md | 152deac | aggiornata |
| deployment.md | 152deac | aggiornata |
| dev-testing.md | 152deac | aggiornata |
| current-work.md | 152deac | aggiornata |
| roadmap.md | 152deac | aggiornata |

## Punto di ripresa

Il servizio è tornato pronto il 2026-10-28: la password è stata reimmessa sulle due attività pianificate e `Set-ScheduledTask` ha restituito `Ready`. L'effetto non è misurato, perché nessuna esecuzione è stata osservata dopo la correzione.

Sei voci restano aperte, in ordine di urgenza, da riprendere con `/riprendi`.

La prima è la verifica che il servizio sia davvero ripartito: si guarda il log del 2026-10-29 in `logs/` e si controlla che `NumberOfMissedRuns` smetta di crescere. Finché non avviene, lo stato è causa nota ed effetto non misurato.

La seconda è la campagna di phishing accertata il 2026-10-28, con dominio di destinazione `mastbusiness.com` dietro un pulsante che si presenta come Microsoft 365. Vanno fatte due cose che in questo giro non sono state fatte per decisione dell'utente: bloccare il dominio sul filtro di posta, e una ricerca messaggi sul tenant per sapere quante caselle lo abbiano ricevuto e se qualcuno abbia premuto.

La terza è la saturazione degli archivi online. Al 2026-10-28 quattro archivi sono oltre la quota e un quinto ha meno di dieci megabyte liberi; la tabella completa con i byte sta in `.claude/memory/caselle.local.md`, che non è tracciato. L'espansione automatica è spenta a livello di organizzazione e richiede licenze che quasi nessuna delle caselle in sofferenza possiede, quindi la decisione è economica e spetta alla direzione. Le licenze registrate vengono dallo snapshot del 2026-10-05 e vanno riverificate prima di qualunque acquisto. Cinque delle undici caselle sono condivise con archivio attivo e nessuna licenza, il che è anche una questione di conformità da guardare a parte.

La quarta è l'adozione del pacchetto `anonymization` e la dichiarazione della visibilità pubblica del repository nel `CLAUDE.md`, proposte in ADR-004 e non ancora attuate.

La quinta è il promemoria chiesto dall'utente: finché il progetto vive sulla postazione, una rotazione della password di Windows deve avvisare che si torna qui a reimmetterla sulle due attività. Non è stato implementato. La via più robusta è un controllo che osservi il fallimento invece della rotazione, cioè una verifica periodica che legga gli eventi 101 con errore `0x8007052E` sulle due attività e avvisi, perché così copre anche le cause diverse dalla rotazione e non dipende dal ricordarsi di eseguirlo al momento giusto.

La sesta è la migrazione su macchina virtuale Proxmox dedicata, decisa in ADR-002, con i tre vincoli che non si trasferiscono copiando una cartella, cioè certificato, credenziale SMTP cifrata DPAPI e database storico, e con il controllo di vivacità esterno da aggiungere. Il contesto per quel lavoro è stato passato a `D:\network-design` in `docs/operativita/servizio-monitoraggio-caselle.md`, in forma anonimizzata.
