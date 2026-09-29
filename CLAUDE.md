# mailbox-alert-toolkit

> Istruzioni di progetto, versionate. Indice dei satelliti tracciati e procedura di ripresa. Le preferenze personali vivono in `CLAUDE.local.md` (ignorato).

## Cos'e questo progetto

Toolkit di monitoraggio e reportistica delle caselle di posta (Intrawelt): genera alert sullo stato delle mailbox e report/trend in Excel. Eseguito via scheduled task di Windows (gli `.xml` inclusi) che lanciano gli script. Lavoro deterministico in script Python (`generate_mailbox_report.py`, `generate_trends.py`) e PowerShell (`mailbox-alert.ps1`, `launcher.ps1`).

## Dati sensibili (mai versionati)

`credentials/`, `config.json`, `app-cert.cer`, `*.pfx`/`*.key`/`*.p12`, `reports/`, `history/`, `logs/` sono gitignored e non vanno mai committati. Il template di configurazione versionato e `config.example.json`. I dati reali delle caselle restano locali.

## Procedura di ripresa

A inizio sessione si legge `.claude/memory/index.md` (branch, commit di riferimento, stato delle schede, prossima azione), poi `.claude/context/current-work.md` se c'e una feature attiva, e si invoca la skill `sync-context` per il drift schede-codice. Work-log in `.claude/memory/progress.md`, decisioni in `.claude/memory/decisions.md`. La skill `onboard` da la spiegazione completa.

## Standard e strumenti

Allineato allo standard portabile `.claude/PROJECT-SYSTEM.md` (rules, engine skills, catalogo `.claude/templates/PACKAGES.md`). Identita git locale `asopranzi@intrawelt` + alias SSH `github-corp`. Pacchetti esterni non adottati (tool Python/PowerShell piccolo); vale la regola `manual-screenshots` per il proofing visivo. Commit e push restano manuali dell'utente.

Norme caricate su richiesta, una riga per situazione con le parole con cui si presenta, così che il caricamento non dipenda dal ricordare che la norma esista.

- `git worktree list` mostra più di un albero, se ne crea o se ne rimuove uno, si deve decidere da dove leggere la memoria versionata: skill `alberi-di-lavoro`.
- Un recupero web fallisce con 403 o con una pagina di verifica anti-bot, la fonte sta su Reddit o su Discord, serve la trascrizione di un video, si sta per annotare una fonte non letta: skill `fonti-non-recuperabili`.
- Si scrive o si valuta una prova automatica, si chiude un difetto, una verifica manuale smentisce una suite verde, si sta per dichiarare completo un intervento il cui scopo era un effetto misurabile: skill `prove-che-misurano`.
- Si inizializza o si allinea il progetto, oppure cambia il modo in cui si prova e si rilascia, e va deciso come separare test e produzione: skill `separazione-ambienti`.
