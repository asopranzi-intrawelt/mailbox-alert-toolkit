---
generated-from-commit: a8eab4c
generated-from-branch: main
generated-date: 2026-06-18
covers-paths: []
last-verified-commit: a8eab4c
---

# Roadmap

> Direzione e priorità del progetto. Non è il work-log: qui sta dove si va, non cosa è già stato fatto.

## Direzione

Il toolkit è funzionalmente completo per il caso d'uso corrente: monitoraggio giornaliero delle mailbox Exchange Online di Intrawelt, alert automatici e storico SQLite. La direzione di medio termine non è esplicitata nei file di codice né nel README.

## Priorità

Nessuna priorità formale definita al commit corrente. Da verificare con l'utente e documentare qui quando emerge una direzione concreta.

## Idee e ipotesi da verificare

Le seguenti sono spunti inferiti dalla struttura attuale del codice, non decisioni confermate. Vanno promossi a priorità solo dopo conferma esplicita.

Rinnovo automatico del certificato di autenticazione: il certificato *self-signed* ha una validità di due anni dalla sua creazione. Quando si avvicina la scadenza, va rigenerato, ricaricato su Entra ID e aggiornato il thumbprint in `config.json`. Una procedura automatizzata o almeno un alert preventivo ridurrebbe il rischio di interruzione del servizio.

Aggiornamento moduli PowerShell: `ExchangeOnlineManagement` e i moduli `Microsoft.Graph.*` ricevono aggiornamenti regolari. Il toolkit non verifica né aggiorna le versioni automaticamente al di fuori del setup iniziale. Da verificare se il comportamento attuale (usa la versione installata al momento del setup) è sufficiente o se serve un meccanismo di aggiornamento periodico.
