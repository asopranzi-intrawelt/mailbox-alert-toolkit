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
